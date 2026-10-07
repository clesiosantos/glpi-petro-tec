# GLPI Petrobras TEC

Projeto de implantação e automação do ambiente GLPI da operação Petrobras/G4F.

## Script de importação da Base 4711

O arquivo `scripts/import_base4711.py` automatiza a carga da planilha **Base_dedados_4711_com_Senior.xlsx**, armazenada na raiz deste repositório, para o GLPI 10.

Principais regras implementadas:

- usa **GERENCIA LOTACAO** somente no campo dinâmico **Gerencia Lotação**;
- usa a mesma gerência no campo dinâmico **Gerencia Lotação**;
- usa **LOCAL** no campo dinâmico **Localização fisica Gerencia**;
- usa **STATUS DA MOBILIZACAO** no campo dinâmico **Status Mobilização**;
- associa o perfil informado em **PERFIL PADRÃO**;
- vincula **PREPOSTO** ao campo dinâmico `Preposto` do usuário para todos os perfis **Posto de Trabalho**;
- usa **Matrícula Senior como login do GLPI**;
- em usuários já existentes, altera somente o login e preserva o mesmo `users_id`, perfis, grupos e vínculos;
- grava em comentários `ITEM PPU: <valor> :: PREPOSTO: <valor>`;
- ignora linhas marcadas como **não mobilizar de imediato**;
- possui modos `--dry-run` e `--apply`.

## Dependências

```bash
dnf -y install python3-openpyxl python3-requests
```

## Variáveis de ambiente

Os tokens não devem ser versionados. Use um arquivo local protegido, por exemplo `/root/.glpi.env`:

```bash
export GLPI_API_URL='https://tec.g4f.sharksolucoes.com.br/apirest.php'
export GLPI_USER_TOKEN='...'
export GLPI_APP_TOKEN='...'
```

## Validação

```bash
source /root/.glpi.env

python3 scripts/import_base4711.py --dry-run
```

## Implantação

Somente após validar o `--dry-run`:

```bash
python3 scripts/import_base4711.py --apply
```

> Não versionar tokens, arquivos `.env` ou planilhas contendo dados pessoais.


## Senha inicial dos usuários

A senha padrão não é versionada no repositório. Defina-a localmente em `/root/.glpi.env`:

```bash
export GLPI_DEFAULT_PASSWORD='SENHA_PADRAO_AQUI'
```

O importador envia `password` e `password2` ao GLPI durante a criação dos usuários. O valor não é exibido nos logs.


## Login integrado com Senior

A regra oficial é:

```text
GLPI.User.name = Matrícula Senior
```

A Matrícula Senior é usada diretamente como login e como chave de integração com o Senior, sem duplicação em outro campo do GLPI.

Para corrigir usuários já implantados, o script localiza cada pessoa pelo nome atual, valida conflito de matrícula e altera o login preservando o mesmo `users_id`. Dessa forma, tickets, perfis, grupos e demais vínculos continuam ligados ao mesmo usuário.

A planilha `Base_dedados_4711_com_Senior.xlsx` deve conter obrigatoriamente a coluna `Matricula Senior` na aba `Posto de Trabalho e Fiscais`. O script usa essa coluna diretamente como login do GLPI.


## Vínculo de Preposto

Para usuários com perfil **Posto de Trabalho**, o importador usa a coluna `PREPOSTO` da Base 4711, localiza o usuário GLPI correspondente pelo nome e grava seu `users_id` no campo dinâmico **Preposto** do container **Agrupamento**.

O campo **Preposto** é uma referência GLPI para `User`; o importador deriva a chave REST no padrão `users_id_<nome_interno>` a partir do próprio cadastro do campo no plugin Fields. Antes de qualquer gravação, o `--dry-run` valida que todos os prepostos foram localizados de forma única.


## Autorização por ITEM PPU

As autorizações dos usuários seguem a regra:

- **Preposto**: entidade **G4F**, perfil **Preposto**, recursividade **Sim**.
- **Posto de Trabalho**: o valor de **ITEM PPU** determina a entidade. Exemplo: `3.3` é associado à entidade cujo nome inicia por `3.3 - `; o perfil é **Posto de Trabalho** e a recursividade é **Sim**.
- Para os Postos de Trabalho já implantados, o importador remove a autorização antiga do perfil **Posto de Trabalho** diretamente na entidade raiz **G4F**, evitando acesso além do escopo do ITEM PPU.
- A entidade da autorização também é definida como entidade padrão do usuário.


## Regra de grupos dos usuários

Os usuários importados pela Base 4711 **não devem possuir associação a Grupo GLPI**. A coluna **GERENCIA LOTACAO** é apenas informação do campo dinâmico **Gerencia Lotação**.

Ao aplicar o importador:
- vínculos existentes em `Group_User` dos usuários da Base 4711 são removidos;
- o grupo padrão do usuário é limpo com `groups_id = 0`;
- nenhum novo vínculo de grupo é criado;
- os objetos de Grupo já existentes no GLPI não são excluídos, pois a correção é sobre a associação do usuário.


## Unificação das subentidades operacionais

A árvore de entidades passa a utilizar uma única subentidade operacional `.1` por ITEM PPU.

Exemplo:

```text
1.1 - Serviços Técnicos de Apoio às Instalações/Predial
├── 1.1.1 - Apoio à Fiscalização e Operacional / Atividades Técnicas e Administrativas
└── 1.1.3 - Atividade de Liderança
```

A antiga subentidade `.2` é migrada para a `.1`:

- o ID da entidade `.1` é preservado;
- a `.1` é renomeada com os dois escopos;
- categorias ITIL associadas à antiga `.2` passam a apontar para a `.1`;
- autorizações **Posto de Trabalho** existentes no pai do ITEM PPU ou na antiga `.2` são migradas para a `.1`, com recursividade **Sim**;
- a entidade padrão dos profissionais é ajustada para a `.1`;
- a antiga entidade `.2` é enviada para a lixeira;
- a subentidade `.3` de Liderança permanece inalterada.

A árvore de categorias não é fundida: seus códigos e nomes são preservados; apenas a associação de entidade da antiga `.2` é migrada para a nova `.1`.

Executar primeiro:

```bash
source /root/.glpi.env
python3 scripts/migrate_unify_operational_entities.py --dry-run
```

Somente após validar:

```bash
python3 scripts/migrate_unify_operational_entities.py --apply
```

O importador principal `import_base4711.py` também foi ajustado para associar novos Postos de Trabalho diretamente à subentidade operacional `<ITEM PPU>.1`.
