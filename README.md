# GLPI Petrobras TEC

Projeto de implantação e automação do ambiente GLPI da operação Petrobras/G4F.

## Script de importação da Base 4711

O arquivo `scripts/import_base4711.py` automatiza a carga da planilha **Base_dedados_4711.xlsx**, armazenada na raiz deste repositório, para o GLPI 10.

Principais regras implementadas:

- cria/garante grupos a partir de **GERENCIA LOTACAO**;
- usa a mesma gerência no campo dinâmico **Gerencia Lotação**;
- usa **LOCAL** no campo dinâmico **Localização fisica Gerencia**;
- usa **STATUS DA MOBILIZACAO** no campo dinâmico **Status Mobilização**;
- associa o perfil informado em **PERFIL PADRÃO**;
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

python3 scripts/import_base4711.py \
  --senior-xlsx ./Acessos_GLPI_Petrobras_G4F.xlsx \
  --dry-run
```

## Implantação

Somente após validar o `--dry-run`:

```bash
python3 scripts/import_base4711.py \
  --senior-xlsx ./Acessos_GLPI_Petrobras_G4F.xlsx \
  --apply
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

Se a `Base_dedados_4711.xlsx` não possuir a coluna `Matricula Senior`, informe uma planilha local de mapeamento com as colunas `Nome` e `Matricula Senior` na aba `Acessos GLPI`.

A planilha de mapeamento não deve ser versionada no repositório.
