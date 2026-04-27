# cdc-dataroom-automation

Gerador declarativo de dataroom em Python, com criacao automatica de subpastas,
`README.md` por pasta, suporte a `--dry-run` e relatorio final em JSON.

## Arquivos principais

- `structure.json` define a estrutura da arvore
- `create_drive_dataroom.py` cria a hierarquia no sistema de arquivos
- `requirements.txt` documenta que nao ha dependencias externas

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
