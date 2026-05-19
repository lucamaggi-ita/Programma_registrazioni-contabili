# STATO_ATTUALE.md — Programma Registrazioni Contabili / Fantozzi

Aggiornato: 2026-05-19

## Stato sintetico

La Fase 9 è conclusa fino alla Fase 9O.
La Fase 13 è conclusa fino alla Fase 13W (parziale — diagnostica read-only eseguita, run reale non ancora materializzata).

Ultimo stato confermato:
- Commit `eb8b57d` — linea DRE 13.02 creata, 11.01.005/006 rimappate (pushato).
- Commit `f0186a2` — .gitignore protetto, plano_contas aggiornato (pushato).
- Audit YAML comprovantes: 0 erros, 0 alertas — 30 regole valide.
- Audit mapeamento_dre: 0 erros — 121 contas, nessun dangling ref.

## Stato repository

Branch principale: master

Commit recenti rilevanti:
- `eb8b57d` — Cria linha DRE 13.02 e remapeia distribuicao e antecipacao de lucros
- `f0186a2` — Protege repositorio e atualiza plano de contas
- `93aa321` — Alinha metadados DRE das regras bancarias
- `9b3571c` — Aplica exclusoes controladas nas regras de comprovantes
- `e7492d4` — Executa classificacao real de comprovantes na Fase 13S
- `90f6dc7` — Completa inventario do pacote contabil gerencial na Fase 9O

Comandi iniziali consigliati per ogni sessione:
```powershell
git status -sb
git log --oneline --decorate -5
git rev-list --left-right --count origin/master...HEAD
git diff --name-status
git diff --check
```

## Funzionalità consolidate

### Pipeline provvisoria

- Processa XML reali.
- Genera output in `output/importacao_real` (ignorata da .gitignore).
- Genera report in `reports/importacao_real` (ignorata da .gitignore).
- Mantiene `fase9_executada=false` salvo uso esplicito della Fase 9.

### Fase 9

Chiusa. Funzioni: engine decisionale da CSV, supporto CSV multipli, gate COMODATO/CFOP1949, lanci manuali gerenciais, pacchetto contabile gerencial provvisorio.

### Classificazione bancaria / comprovantes

CLI disponibili:
- `executar-classificacao-comprovantes-real` — run classificazione reale
- `auditar-regras-comprovantes` — audit YAML read-only
- `gerar-pacote-revisao-comprovantes` — pacchetto revisione XLSX

Regole: `config/regras_classificacao_comprovantes.yaml` — 30 regole, 0 erros.

### Bonificação recebida

Conta: `2.05.003`, trattamento: `REDUCAO_CUSTO`. Non è ricavo né costo normale.

### COMODATO

Non entra in DRE. Fuori da report OK.

### Lançamentos manuais gerenciais

Fase 9N conclusa. Controllati via CSV. Pre-livro: due righe tecniche (debito/credito). DRE: impatta una sola volta.

### Piano dei conti e mapeamento DRE

- `plano_contas_gerencial.yaml` — aggiornato, validato.
- `mapeamento_dre.yaml` — 28 linhas_dre, 121 contas. Linea 13.02 aggiunta.
- 11.01.005 (Distribuicao de Lucros) e 11.01.006 (Antecipacao de Lucros) rimappate a 13.02 (entra_dre: false).

### Skill disponibili

| Skill | Funzione |
|---|---|
| `regole-registrazioni-banca` | Gestione regole YAML comprovantes |
| `gerar-pdf-regras-banca` | PDF documentale regole banca |
| `gerar-pdf-regras-xml` | PDF documentale regole NF-e/XML |
| `aggiornamento-piano-conti-dre` | Aggiornamento plano_contas e mapeamento_dre |
| `gerar-pdfs-plano-dre` | PDF documentali plano/DRE |

### PDF documentali

Generati in `reports/documentacao_contabil/` (ignorata da .gitignore):
- `regras_registrazioni_banca.pdf` — regole comprovantes (landscape A4)
- `regras_classificacao_xml.pdf` — regole NF-e/XML (landscape A4)

## Punti aperti (P)

| ID | Problema | Priorità |
|---|---|---|
| P1 | Fallback `99.01` ha `entra_dre: true` — prodotti non classificati entrano in DRE | Governance da decidere |
| P2 | 13W non ancora materializzata con run reale | Blocco prima di Fase 14 |

## Prossimo lavoro

### Chiudere 13W

Eseguire run classificazione reale e verificare REVISAR rimasti:

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli executar-classificacao-comprovantes-real --regras config\regras_classificacao_comprovantes.yaml ...
```

### Fase 14 — Giornale e DRE completo

Integrazione di tutte le fonti (XML + banca + manuali) nel pre-livro diário e DRE gestionale.

CLI già disponibili: `gerar-pre-livro-diario-teste`, `gerar-dre-gerencial-ok-provisoria`.

## Rischi da non dimenticare

1. Non chiamare la prossima macrofase Fase 10 — è Fase 14.
2. Non committare dati reali bancari (coperti da .gitignore: output/importacao_real/, reports/importacao_real/, pendenze_*.txt).
3. Non usare OCR come fonte primaria se esiste OFX/CSV/XLSX.
4. Non consolidare i due CNPJ prima della normalizzazione.
5. Non confondere pagamento bancario con competenza contabile.
6. Non far entrare COMODATO in DRE.
7. Non trattare CFOP 1949 in modo generico.
8. Non riusare run vecchie come base operativa dopo cambi logici.
9. Non creare alias troppo generici.
10. Non fare `git add .` — sempre aggiungere file specifici.
