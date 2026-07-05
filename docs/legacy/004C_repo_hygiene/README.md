# Legacy repo hygiene - 004C/004D

Questa cartella conserva file storici/orfani identificati durante l'audit read-only `FZ-HARDENING-004C-REPO-HYGIENE-LEGACY-INVENTARIO` e archiviati nella microfase `FZ-HARDENING-004D-REPO-HYGIENE-CLEANUP-SICURO`.

Questi file non sono configurazioni operative del workflow vivo.

Configurazioni operative vive:

- `config/plano_contas_gerencial.yaml`
- `config/mapeamento_dre.yaml`
- `config/regras_classificacao_produtos.yaml`
- `config/regras_classificacao_comprovantes.yaml`

File archiviati:

- `FANTOZZI_plano_contas_gerencial_proposta.yaml`: bozza storica iniziale, non operativa.
- `FANTOZZI_plano_contas_rules_proposta.yaml`: bozza storica iniziale, non operativa.
- `csv_columns.yaml`: configurazione orfana non usata da codice/test/CLI/docs operative.

Non usare questi file per run, DRE, classificazione prodotti, comprovantes o manifest finali.
