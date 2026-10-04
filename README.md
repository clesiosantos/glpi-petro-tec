# GLPI Petrobras TEC

Projeto de implantação e automação do ambiente GLPI da operação Petrobras/G4F.

## Script de importação da Base 4711

O arquivo `scripts/import_base4711.py` automatiza a carga da planilha **Base de dados_4711.xlsx** para o GLPI 10.

Principais regras implementadas:

- cria/garante grupos a partir de **GERENCIA LOTACAO**;
- usa a mesma gerência no campo dinâmico **Gerencia Lotação**;
- usa **LOCAL** no campo dinâmico **Localização fisica Gerencia**;
- usa **STATUS DA MOBILIZACAO** no campo dinâmico **Status Mobilização**;
- associa o perfil informado em **PERFIL PADRÃO**;
- cria logins no padrão `primeironome + 4 dígitos`;
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
  --xlsx "/root/Base de dados_4711.xlsx" \
  --dry-run
```

## Implantação

Somente após validar o `--dry-run`:

```bash
python3 scripts/import_base4711.py \
  --xlsx "/root/Base de dados_4711.xlsx" \
  --apply
```

> Não versionar tokens, arquivos `.env` ou planilhas contendo dados pessoais.
