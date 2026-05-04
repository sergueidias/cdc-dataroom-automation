# cdc-dataroom-automation

Gerador declarativo de dataroom em Python, com criacao automatica de subpastas,
`README.md` por pasta, suporte a `--dry-run` e relatorio final em JSON.

## Arquivos principais

- `structure.json` define a estrutura da arvore
- `create_drive_dataroom.py` cria a hierarquia no sistema de arquivos
- `create_master_risk_register.py` gera o controle central em Excel
- `templates/MASTER_RISK_REGISTER.xlsx` modelo do controle (gerado pelo script acima)
- `requirements.txt` lista as dependencias

## Como usar

Dry run:

```bash
python3 create_drive_dataroom.py --dry-run --target-dir ./tmp-dataroom
```

Execucao real:

```bash
python3 create_drive_dataroom.py --target-dir ./tmp-dataroom --report-file ./tmp-dataroom-report.json
```

Reexecucao idempotente:

```bash
python3 create_drive_dataroom.py --target-dir ./tmp-dataroom --report-file ./tmp-dataroom-report-second-run.json
```

## Gerar o MASTER_RISK_REGISTER

Pre-requisito (uma vez): `pip install -r requirements.txt`.

```bash
python3 create_master_risk_register.py
```

O script cria `templates/MASTER_RISK_REGISTER.xlsx` com duas abas:

- **Controle** — 15 colunas, primeira linha congelada, filtro automatico, 8 linhas Tier 1 ja preenchidas (CNPJs, contrato social, marca, passivos, protestos, bloqueios, patrocinios, recebiveis) e validacao de dados em Status, Prioridade, Red flag, Tipo de risco e Decisao.
- **Listas** — valores permitidos para cada campo controlado.

Caminho de saida customizavel:

```bash
python3 create_master_risk_register.py --output ./outro/local/RISK.xlsx
```
