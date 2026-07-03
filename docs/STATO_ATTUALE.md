# STATO_ATTUALE.md — Programma Registrazioni Contabili / Fantozzi

Aggiornato: 2026-07-03 (allineato dopo ricognizione tecnica completa read-only)

Nota storica: la versione precedente di questo file (2026-05-19) descriveva
uno stato antecedente a diversi commit rilevanti (`bcea1c7`, `01b4d0a`,
`2670acb` sulle registrazioni manuali; `9598e1a`, `1b4e2e9`, `4fc0fdd`,
`4871e12`, `2d69d3a` sugli output visuali). I contenuti sotto sono stati
verificati sul codice/git reale, non solo riportati.

## Stato sintetico

La Fase 9 è conclusa fino alla Fase 9O.
La Fase 13 è conclusa fino alla Fase 13W (parziale — diagnostica read-only eseguita, run reale non ancora materializzata).

Ultimo stato confermato (verificato 2026-07-03):
- Branch operativo: `senza-incrementali`, allineato 1:1 con `origin/senza-incrementali`.
- Strategia operativa corrente: **FULL RUN DA ZERO** — nessuna baseline, nessun workflow incrementale. `executar-run-incremental` non è un percorso operativo valido.
- Registrazioni manuali gerenciais: **già implementate, validate e integrate nella pipeline reale** (non solo progettate) — 39 test automatizzati verdi.
- Output visuali attivi: PDF DRE, dashboard HTML, Excel giornale leggibile (versione "003D" consolidata).
- Un tentativo di dashboard integrata nell'Excel giornale leggibile ("003E") è stato bocciato e revertito in modo pulito — non è una linea attiva da riprendere.

## Stato repository

Branch principale operativo: **senza-incrementali**

Branch storici/non operativi (non usare come base di lavoro):
- `master` e `backup/con-incrementali` — stessa HEAD, linea abbandonata "con incrementali" (baseline + run incrementali retroattivi), superata dalla strategia "senza incrementali".
- `main` — ancestor diretto di `senza-incrementali`, fermo alla Fase 11A, non riceve più commit diretti.
- `main-initial-backup` — snapshot congelato del primo commit del repository.

Commit recenti rilevanti (linea `senza-incrementali`):
- `2d69d3a` — Revert "Migliora dashboard DRE del Giornale Excel leggibile" (bocciata "003E")
- `4871e12` — Migliora dashboard DRE del Giornale Excel leggibile ("003E", revertita da `2d69d3a`)
- `4fc0fdd` — Consolida Giornale Excel leggibile ("003D", **valida e attiva**)
- `1b4e2e9` — Adiciona dashboard HTML gerencial da DRE
- `9598e1a` — Implementa gerador PDF DRE gerencial (FZ-OUTPUT-VISUAL-001 Fase 1)
- `01b4d0a` — Corrige mapeamento DRE da contrapartida tecnica manual
- `2670acb` — Corrige path default de saida CSV e template manuais gerenciais
- `bcea1c7` — Operacionaliza registrazioni manuais

Comandi iniziali consigliati per ogni sessione:
```powershell
git status -sb
git branch --show-current
git log --oneline --decorate -5
git rev-list --left-right --count origin/senza-incrementali...HEAD
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

Fase 9N conclusa. **Già operative tecnicamente**, non solo progettate: template
XLSX operatore → validazione doppio livello (fail-closed) → CSV tecnico →
conversione in due lançamentos standard (debito/credito) → integrazione nel
pré-livro come fonte aggiuntiva → contropartida tecnica `99.99.998` esclusa
per design dal DRE. Coperte da 39 test automatizzati (`test_lancamentos_manuais_gerenciais.py`,
`test_lancamentos_manuais_template.py`). Pre-livro: due righe tecniche (debito/credito). DRE: impatta una sola volta.

Path reali (verificati nel codice, `cli.py`):
```text
Database_Dev\Registrazioni_manuali\registrazioni_manuali_base.xlsx     - template operatore
Database_Dev\Registrazioni_manuali\registrazioni_manuali_tecnico.csv   - CSV tecnico intermediario
```

Limiti noti (governance, non funzionali — vedi anche `docs/LANCAMENTOS_MANUAIS_GERENCIAIS_OPERADOR.md`):
- nessuna anti-duplicazione automatica con XML (delegata alla revisione umana);
- deduplicazione via hash valida solo dentro un singolo caricamento, non tra run/periodi diversi;
- nessun test end-to-end automatico copre l'orchestrazione one-shot completa `executar-lancamentos-manuais-gerenciais` (solo i componenti separati).

### Output visuais (post-run, non ufficiali)

Tutti generati da un run completo esistente, marcati internamente `TESTE_TECNICO_NAO_OFICIAL`:

| Output | Comando | Stato |
|---|---|---|
| PDF DRE | `gerar-pdf-dre --run-dir` | Attivo, 28 test |
| Dashboard HTML | `gerar-dashboard-html --run-dir` | Attivo, 31 test |
| Excel giornale leggibile ("003D") | `gerar-giornale-leggibile --run-dir` | Attivo, versione consolidata (`4fc0fdd`), 29 test |

**"003E"** (tentativo di dashboard integrata nell'Excel giornale leggibile,
commit `4871e12`) è stata **bocciata e revertita** (`2d69d3a`) — confermato che
il codice tracciato oggi è identico alla versione pre-tentativo, nessun residuo.
Non è una linea da riprendere così com'era.

Rischio noto: `reportlab` (usato da `dre_pdf_gerencial.py` e dagli script in
`scripts/`) non è dichiarato tra le dipendenze in `pyproject.toml` — funziona
solo perché già installato nell'ambiente attuale.

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
11. Non usare `master`/`backup/con-incrementali` come base di lavoro o per merge/cherry-pick — è la linea storica "con incrementali" superata dalla strategia "senza incrementali".
12. Per modifiche visuali (PDF/HTML/Excel leggibile), non "inventare" layout — seguire il modello approvato ("003D") o produrre un'anteprima prima del commit. La "003E" è l'esempio concreto di cosa evitare.
