---
name: gerar-giornale-dre-multifonte
description: Esegue una full run multifonte solo quando richiesta esplicitamente. Con baseline approvata e manuali operatore, usare invece executar-run-incremental. Non modifica YAML, non fa commit/push, non inferisce conti.
---

# Skill: gerar-giornale-dre-multifonte

Skill operativa per eseguire il workflow giornale + DRE multifonte in modo
autonomo, sicuro e auditabile. La baseline 007 è un riferimento storico di
confronto, non una precondizione rigida, salvo verifica esplicita del file
`output\multifonte_007\audit_007_baseline_post_006_summary.json`.

**Regola 011B:** questa skill non e il percorso quotidiano per le registrazioni
manuali. Se esiste `output\baselines\baseline_atual.json` approvata e
l'operatore sta registrando manuali, usare `executar-run-incremental` tramite
`lanciamento-registrazioni-manuali`.

Usare `reprocessar-importacao-real` solo per una full run richiesta
esplicitamente dall'operatore o per creare/riapprovare una baseline. In presenza
di baseline approvata, la full run richiede sempre `--full-run`.

**L'operatore deve solo:**
- verificare che i file decisionali banca siano aggiornati;
- confermare il periodo da coprire;
- confermare i risultati finali.

---

## 1. Scopo

Generare una run multifonte operativa:

```text
XML (CNPJ 08884180000123 + 55746722000181)
+ banca/comprovantes approvati (bridge)
+ lançamentos manuais (opzionali)
→ bridge banca
→ pre-livro / giornale gerencial
→ DRE gerencial provisória
→ pacote contábil gerencial
→ manifest / audit trail
```

---

## 2. Quando usare questa skill

Usare SOLO quando:

- si vuole generare il giornale + DRE aggiornato con dati reali multifonte;
- i file decisionali banca (`Analisi_regole_registrazioni_banca_comprovanti.xlsx`
  e `Analisi_regole_registrazioni_banca_conto_corrente.xlsx`) sono già aggiornati
  con decisioni approvate dall'operatore;
- le regole XML sono già approvate e applicate.

---

## 3. Quando NON usare questa skill

Non usare per:

- aggiornare regole XML (usare `regole-registrazioni-xml`);
- aggiornare regole banca (usare `regole-registrazioni-banca`);
- creare o applicare lançamentos manuais operativi quotidiani (usare
  `lanciamento-registrazioni-manuali`, che valida e chiama
  `executar-run-incremental`);
- aggiornare piano dei conti (usare `aggiornamento-piano-conti-dre`);
- generare solo il DRE da un Excel già esistente (usare CLI direttamente);
- audit del giornale già generato (usare `auditar-giornale-multifonte`);
- audit del DRE già generato (usare `auditar-dre-gerencial`).

---

## 4. Precondizioni

Prima di eseguire, verificare:

1. Branch `master` allineato, nessun diff su file tracciati;
2. `bank_raw\_decisoes_regras\Analisi_regole_registrazioni_banca_comprovanti.xlsx`
   esiste e contiene decisioni APROVADO aggiornate;
3. `bank_raw\_decisoes_regras\Analisi_regole_registrazioni_banca_conto_corrente.xlsx`
   esiste e contiene decisioni APROVADO aggiornate;
4. `Database_Dev\XML_fatture_acquisto\` contiene XML reali per entrambi i CNPJ;
5. i CSV multifonte classificati di comprovantes e conta corrente sono indicati
   esplicitamente e il loro path esiste;
6. `config\plano_contas_gerencial.yaml` non è stato modificato senza autorizzazione;
7. `config\mapeamento_dre.yaml` non è stato modificato senza autorizzazione.

---

## 5. File / Input richiesti

| File | Default / Path | Obbligatorio |
|---|---|---|
| Excel decisional comprovantes | `bank_raw\_decisoes_regras\Analisi_regole_registrazioni_banca_comprovanti.xlsx` | sì |
| Excel decisional conta corrente | `bank_raw\_decisoes_regras\Analisi_regole_registrazioni_banca_conto_corrente.xlsx` | sì |
| CSV comprovantes classificati | path validato della baseline multifonte, es. `reports\comprovantes_classificacao\runs\...\comprovantes_classificados.csv` | sì (policy) |
| CSV conta corrente classificata | path validato della baseline multifonte, es. `output\audit_002b\conta_corrente_classificados_combined.csv` | sì (policy) |
| Pasta XML NF-e | `Database_Dev\XML_fatture_acquisto\` | sì |
| CSV manuais tecnico | `Database_Dev\Registrazioni_manuali\registrazioni_manuali_tecnico.csv` | no |
| Manifest validazione manuais | `Database_Dev\Registrazioni_manuali\registrazioni_manuali_validazione_manifest.json` | solo se si usa CSV manuais |
| Piano conti | `config\plano_contas_gerencial.yaml` | letto dalla pipeline |
| Mapeamento DRE | `config\mapeamento_dre.yaml` | letto dalla pipeline |

> **Nota CSV (policy vs CLI):** `--comprovantes-csv` e `--conta-corrente-csv` sono
> tecnicamente opzionali per la CLI (default=None). Questa skill li tratta come
> obbligatori per policy operativa: senza CSV classificati il bridge lavora solo
> sugli Excel decisionali, riducendo la qualità della deduplicazione multifonte.
> Il blocco `BLOCCO_CSV_MULTIFONTE_AUSENTES` è una scelta operativa, non un vincolo
> tecnico del comando.

---

## 6. Comandi copiabili

### A — Precheck Git (obbligatorio)

```powershell
git status -sb
git rev-list --left-right --count origin/master...HEAD
git diff --name-status
git diff --cached --name-status
git diff --name-status config/
git diff --name-status src/
git diff --name-status tests/
```

Atteso:
- branch `master`, divergenza `0 0`;
- nessun diff su `config/`, `src/`, `tests/`;
- nessun file staged.

Se appaiono diff su file tracciati:
```text
BLOCCO: BLOCCO_GIT_STATO_SPORCO
Causa: esistono modifiche tracciate inattese
Azione: fermarsi, riportare `git diff --name-status`, chiedere conferma all'operatore
Divieto: non usare git restore / git reset / git checkout automaticamente
Commit/push: non eseguiti
```

---

### B — Verificare file decisionali banca

```powershell
Test-Path "bank_raw\_decisoes_regras\Analisi_regole_registrazioni_banca_comprovanti.xlsx"
Test-Path "bank_raw\_decisoes_regras\Analisi_regole_registrazioni_banca_conto_corrente.xlsx"
```

Se non esistono:
```text
BLOCCO: BLOCCO_DECISIONALI_BANCA_AUSENTES
Causa: file Excel decisional banca non trovati
Azione: eseguire skill regole-registrazioni-banca per aggiornarli
```

---

### B2 — Verificare CSV multifonte classificati

I CSV multifonte sono obbligatori per policy. Non eseguire il bridge se non sono
indicati esplicitamente o se uno dei path non esiste.

```powershell
Test-Path "<COMPROVANTES_CSV>"
Test-Path "<CONTA_CORRENTE_CSV>"
```

Path di riferimento dalla baseline 007:
- Comprovantes: `reports\comprovantes_classificacao\runs\audit_002c_2026-05-26\comprovantes_classificados.csv`
- Conta corrente: `output\audit_002b\conta_corrente_classificados_combined.csv`

Se mancano o non sono stati indicati:
```text
BLOCCO: BLOCCO_CSV_MULTIFONTE_AUSENTES
Causa: CSV multifonte classificati assenti o non indicati
Nota: --comprovantes-csv e --conta-corrente-csv sono CLI-opzionali ma obbligatori per questa skill
Azione: indicare i CSV comprovantes/conta corrente validati della baseline multifonte
```

---

### C — Bridge banca (adattare TAG e date)

Definire un TAG univoco per fase/periodo:

```powershell
$TAG = "multifonte_009_2026_01_04"
$OUT = "output\$TAG"

$COMPROVANTES_CSV = "<COMPROVANTES_CSV>"
$CONTA_CORRENTE_CSV = "<CONTA_CORRENTE_CSV>"
```

**Guardia isolation — eseguire PRIMA del bridge:**

```powershell
if (Test-Path "$OUT\lancamentos_banca_aprovados_$TAG.xlsx") {
    Write-Error "BLOCCO: BLOCCO_OUTPUT_MULTIFONTE_EXISTENTE — artefatto bridge già presente"
    exit 1
}
if (Test-Path "$OUT\manifest_lancamentos_banca.json") {
    Write-Error "BLOCCO: BLOCCO_OUTPUT_MULTIFONTE_EXISTENTE — manifest bridge già presente"
    exit 1
}
if (Test-Path "$OUT\lancamentos_banca_rejeitados.csv") {
    Write-Error "BLOCCO: BLOCCO_OUTPUT_MULTIFONTE_EXISTENTE — rejeitados bridge già presente"
    exit 1
}

New-Item -ItemType Directory -Force $OUT | Out-Null
```

**Bridge:**

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli gerar-lancamentos-banca-aprovados `
  --comprovantes-csv $COMPROVANTES_CSV `
  --conta-corrente-csv $CONTA_CORRENTE_CSV `
  --comprovantes-xlsx bank_raw\_decisoes_regras\Analisi_regole_registrazioni_banca_comprovanti.xlsx `
  --conta-corrente-xlsx bank_raw\_decisoes_regras\Analisi_regole_registrazioni_banca_conto_corrente.xlsx `
  --saida "$OUT\lancamentos_banca_aprovados_$TAG.xlsx" `
  --periodo-ini YYYY-MM-DD `
  --periodo-fim YYYY-MM-DD
```

Il comando genera automaticamente (nella stessa cartella del `--saida`):
- `$OUT\lancamentos_banca_aprovados_$TAG.xlsx`
- `$OUT\lancamentos_banca_rejeitados.csv`
- `$OUT\manifest_lancamentos_banca.json`

---

### D — Leggere manifest bridge

```powershell
$BRIDGE_MANIFEST = "$OUT\manifest_lancamentos_banca.json"
$bm = Get-Content $BRIDGE_MANIFEST | ConvertFrom-Json

$bm.linhas_aprovadas_incluidas
$bm.deduplicacao.possiveis_duplicados_data_divergente
$bm.deduplicacao.ambiguos_data_divergente

if ($bm.linhas_aprovadas_incluidas -le 0) {
  Write-Error "BLOCCO: BLOCCO_BRIDGE_BANCA_ZERO_APROVADAS"
  exit 1
}

if ($bm.deduplicacao.possiveis_duplicados_data_divergente -gt 0) {
  Write-Error "BLOCCO: BLOCCO_BRIDGE_BANCA_DUPLICADOS_PENDENTES"
  exit 1
}

if ($bm.deduplicacao.ambiguos_data_divergente -gt 0) {
  Write-Error "BLOCCO: BLOCCO_BRIDGE_BANCA_DUPLICADOS_PENDENTES"
  exit 1
}
```

Verificare nel manifest bridge:
- `linhas_aprovadas_incluidas > 0`
- `deduplicacao.possiveis_duplicados_data_divergente = 0`
- `deduplicacao.ambiguos_data_divergente = 0`

> **Riferimento storico (non check bloccante):**
> Nella baseline jan-apr 2026 era stato osservato:
> `deduplicacao.duplicados_bancarios_data_divergente_excluidos = 11`
> Questo valore è specifico della baseline 007. Non costituisce criterio di blocco
> per periodi futuri o per run con dati/regole aggiornati.

Se `linhas_aprovadas_incluidas <= 0`:
```text
BLOCCO: BLOCCO_BRIDGE_BANCA_ZERO_APROVADAS
Causa: manifest bridge indica linhas_aprovadas_incluidas <= 0
Azione: verificare periodo, CSV multifonte classificati e file decisionali banca
Commit/push: non eseguiti
```

Se `deduplicacao.possiveis_duplicados_data_divergente > 0` o
`deduplicacao.ambiguos_data_divergente > 0`:
```text
BLOCCO: BLOCCO_BRIDGE_BANCA_DUPLICADOS_PENDENTES
Causa: duplicati banca non risolti
Azione: riesaminare Excel decisional banca e risolvere i casi
```

---

### E — Run multifonte (solo banca + XML)

Eseguire questo passo solo se l'operatore ha chiesto esplicitamente una full run.
Se esiste baseline approvata, includere `--full-run`. Non usare questo passo per
il normale workflow manuali.

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli reprocessar-importacao-real `
  --full-run `
  --xml-base Database_Dev\XML_fatture_acquisto `
  --saida-output "$OUT\reprocessamento_output" `
  --saida-reports "$OUT\reprocessamento_reports" `
  --lancamentos-banca-aprovados "$OUT\lancamentos_banca_aprovados_$TAG.xlsx" `
  --gerar-pacote-contabil-gerencial
```

---

### F — Run multifonte con manuais (opzionale)

Non usare questa modalita per il workflow operatore "ho finito" della fase 011B.
Quel caso deve chiamare `executar-run-incremental`.

Usare questa modalita solo in una full run esplicitamente richiesta, se il CSV
manuais tecnico e stato preparato **e** se esiste un manifest di validazione
approvato prodotto dalla skill `lanciamento-registrazioni-manuali`.

**Manifest richiesto:**
```text
Database_Dev\Registrazioni_manuali\registrazioni_manuali_validazione_manifest.json
```

**Campi minimi attesi nel manifest:**
```json
{
  "status_validacao": "APROVADO",
  "arquivo_csv_tecnico": "Database_Dev\\Registrazioni_manuali\\registrazioni_manuali_tecnico.csv",
  "rejeitados": 0,
  "pendencias": 0,
  "pendencias_justificadas": false,
  "duplicidade_xml_banca": 0,
  "duplicidades_justificadas": false
}
```

**Regole di validazione del manifest:**
```text
status_validacao deve essere APROVADO
arquivo_csv_tecnico deve corrispondere al CSV usato nella run
rejeitados deve essere 0
pendencias deve essere 0, oppure pendencias_justificadas = true
duplicidade_xml_banca deve essere 0, oppure duplicidades_justificadas = true
```

> **Nota:** il manifest di validazione manuais deve essere prodotto dalla skill
> `lanciamento-registrazioni-manuali` secondo questo schema.
> Se la skill manuale non produce ancora questi campi, la modalità manuais è da considerare
> **NON DISPONIBILE** per questa skill multifonte, salvo approvazione tecnica esplicita.
> Non creare manifest manuali a mano per aggirare il blocco.

Se manca il manifest, i campi non passano, o il manifest non è stato prodotto dalla skill manuale:
```text
BLOCCO: BLOCCO_MANUAIS_NAO_VALIDADOS
Causa: CSV manuais indicato senza manifest di validazione approvato prodotto dalla skill lanciamento-registrazioni-manuali
Azione: eseguire prima la skill lanciamento-registrazioni-manuali e validare il CSV tecnico
Commit/push: non eseguiti
```

```powershell
$MANUAIS_CSV = "Database_Dev\Registrazioni_manuali\registrazioni_manuali_tecnico.csv"

.\.venv\Scripts\python.exe -m fantozzi.cli reprocessar-importacao-real `
  --full-run `
  --xml-base Database_Dev\XML_fatture_acquisto `
  --saida-output "$OUT\reprocessamento_output" `
  --saida-reports "$OUT\reprocessamento_reports" `
  --lancamentos-banca-aprovados "$OUT\lancamentos_banca_aprovados_$TAG.xlsx" `
  --lancamentos-manuais-gerenciais-csv $MANUAIS_CSV `
  --gerar-pacote-contabil-gerencial
```

---

### G — Trovare run_id

```powershell
Get-ChildItem "$OUT\reprocessamento_reports\runs\" | Sort-Object Name | Select-Object -Last 3
```

Oppure leggere `current_run.txt` se generato nella run:
```powershell
Get-Content "$OUT\reprocessamento_reports\current_run.txt"
```

---

### H — Localizzare output principali

Sostituire `{run_id}` con il run_id trovato.

**Output per CNPJ (lançamentos Excel):**
```text
$OUT\reprocessamento_output\runs\{run_id}\08884180000123\lancamentos_fantozzi_real.xlsx
$OUT\reprocessamento_output\runs\{run_id}\55746722000181\lancamentos_fantozzi_real.xlsx
$OUT\reprocessamento_output\runs\{run_id}\BANCA_APROVADA\lancamentos_banca_aprovados.xlsx
```

**Report e manifest della run:**
```text
$OUT\reprocessamento_reports\runs\{run_id}\manifest_run.json
$OUT\reprocessamento_reports\runs\{run_id}\pre_livro_diario_gerencial_teste.xlsx
$OUT\reprocessamento_reports\runs\{run_id}\dre_gerencial_ok_provisoria.xlsx
$OUT\reprocessamento_reports\runs\{run_id}\dre_gerencial_ok_provisoria.txt
$OUT\reprocessamento_reports\runs\{run_id}\relatorios_gerenciais_ok_provisorios.xlsx
$OUT\reprocessamento_reports\runs\{run_id}\relatorios_gerenciais_ok_provisorios.txt
$OUT\reprocessamento_reports\runs\{run_id}\auditoria_pre_livro_diario_gerencial_teste.txt
$OUT\reprocessamento_reports\runs\{run_id}\auditoria_relatorios_gerenciais_ok_provisorios.xlsx
$OUT\reprocessamento_reports\runs\{run_id}\auditoria_relatorios_gerenciais_ok_provisorios.txt
$OUT\reprocessamento_reports\runs\{run_id}\pacote_contabil_gerencial_provisorio\
```

---

### I — Verifiche post-run (obbligatorie)

```powershell
# Git non deve avere nulla staged
git diff --cached --name-status

# YAML non devono essere stati modificati
git diff --name-status config/

# Codice/test non devono essere stati modificati
git diff --name-status src/
git diff --name-status tests/
```

---

## 7. Output attesi

> In questa skill, FASE_PERIODO è rappresentato dalla variabile `$TAG`
> (es. `multifonte_009_2026_01_04`). `$OUT = "output\$TAG"`.

```text
$OUT\
  lancamentos_banca_aprovados_$TAG.xlsx             ← bridge banca approvati
  lancamentos_banca_rejeitados.csv                  ← bridge banca rigettati
  manifest_lancamentos_banca.json                   ← manifest bridge

$OUT\reprocessamento_output\runs\{run_id}\
  08884180000123\lancamentos_fantozzi_real.xlsx      ← lançamentos XML CNPJ 1
  55746722000181\lancamentos_fantozzi_real.xlsx      ← lançamentos XML CNPJ 2
  BANCA_APROVADA\lancamentos_banca_aprovados.xlsx    ← copia operativa del bridge

$OUT\reprocessamento_reports\runs\{run_id}\
  manifest_run.json
  pre_livro_diario_gerencial_teste.xlsx
  dre_gerencial_ok_provisoria.xlsx
  dre_gerencial_ok_provisoria.txt
  relatorios_gerenciais_ok_provisorios.xlsx
  relatorios_gerenciais_ok_provisorios.txt
  auditoria_pre_livro_diario_gerencial_teste.txt
  auditoria_relatorios_gerenciais_ok_provisorios.xlsx
  auditoria_relatorios_gerenciais_ok_provisorios.txt
  pacote_contabil_gerencial_provisorio\
```

---

## 8. Controlli obbligatori post-run

### 8.1 — Leggere manifest bridge

```powershell
$mb = Get-Content "$OUT\manifest_lancamentos_banca.json" | ConvertFrom-Json

$mb.linhas_aprovadas_incluidas
$mb.deduplicacao.possiveis_duplicados_data_divergente
$mb.deduplicacao.ambiguos_data_divergente

if ($mb.linhas_aprovadas_incluidas -le 0) {
  Write-Error "BLOCCO: BLOCCO_BRIDGE_BANCA_ZERO_APROVADAS"
  exit 1
}

if ($mb.deduplicacao.possiveis_duplicados_data_divergente -gt 0) {
  Write-Error "BLOCCO: BLOCCO_BRIDGE_BANCA_DUPLICADOS_PENDENTES"
  exit 1
}

if ($mb.deduplicacao.ambiguos_data_divergente -gt 0) {
  Write-Error "BLOCCO: BLOCCO_BRIDGE_BANCA_DUPLICADOS_PENDENTES"
  exit 1
}
```

Verificare:
- `linhas_aprovadas_incluidas > 0`
- `deduplicacao.possiveis_duplicados_data_divergente = 0`
- `deduplicacao.ambiguos_data_divergente = 0`

---

### 8.2 — Leggere manifest run

```powershell
$RUN_MANIFEST = "$OUT\reprocessamento_reports\runs\{run_id}\manifest_run.json"
$mr = Get-Content $RUN_MANIFEST | ConvertFrom-Json

$mr.status_run
$mr.lancamentos_banca_aprovados.incluido
$mr.lancamentos_banca_aprovados.linhas_lidas

$dreCheckpoint = $mr.checkpoints | Where-Object { $_.descricao -eq "DRE_OK_PROVISORIA_GERADA" }
$pacoteCheckpoint = $mr.checkpoints | Where-Object { $_.descricao -eq "PACOTE_CONTABIL_GERENCIAL_GERADO" }

if ($mr.status_run -ne "PROVISORIO_CONCLUIDO") {
  Write-Error "BLOCCO: BLOCCO_MANIFEST_RUN_INVALIDO"
  exit 1
}

if ($mr.lancamentos_banca_aprovados.incluido -ne $true) {
  Write-Error "BLOCCO: BLOCCO_MANIFEST_RUN_INVALIDO"
  exit 1
}

if ($mr.lancamentos_banca_aprovados.linhas_lidas -le 0) {
  Write-Error "BLOCCO: BLOCCO_MANIFEST_RUN_INVALIDO"
  exit 1
}

if (-not $dreCheckpoint) {
  Write-Error "BLOCCO: BLOCCO_MANIFEST_RUN_INVALIDO"
  exit 1
}

if (-not $pacoteCheckpoint) {
  Write-Error "BLOCCO: BLOCCO_MANIFEST_RUN_INVALIDO"
  exit 1
}
```

Verificare:
- `status_run = PROVISORIO_CONCLUIDO`
- `lancamentos_banca_aprovados.incluido = true`
- `lancamentos_banca_aprovados.linhas_lidas > 0`
- `checkpoints` contiene `descricao = "DRE_OK_PROVISORIA_GERADA"`
- `checkpoints` contiene `descricao = "PACOTE_CONTABIL_GERENCIAL_GERADO"`

Se uno dei check fallisce:
```text
BLOCCO: BLOCCO_MANIFEST_RUN_INVALIDO
Causa: manifest_run.json non conferma status_run, banca inclusa, linhas_lidas o checkpoints
Azione: fermarsi, riportare i valori letti, non proseguire con audit finale
Commit/push: non eseguiti
```

---

### 8.3 — Leggere DRE TXT

```powershell
Get-Content "$OUT\reprocessamento_reports\runs\{run_id}\dre_gerencial_ok_provisoria.txt"
```

Verificare nel TXT solo:
- `Status DRE: OK_PROVISORIO`

> **Nota:** i controlli strutturali DRE (REVISAR = 0, 99.99.998, contas 99.*,
> contrapartidas, bonificações 03.08, COMODATO, M A BRITO) non dipendono dal TXT
> ma dagli XLSX e dalle skill di audit dedicate. Usare:
> - `dre_gerencial_ok_provisoria.xlsx`
> - `auditoria_relatorios_gerenciais_ok_provisorios.xlsx`
> - skill `auditar-dre-gerencial`
> - skill `auditar-giornale-multifonte`

---

## 9. Errori, Blocchi e Cosa Fare

**Formato blocco:**
```text
BLOCCO: <CODICE>
Causa: <descrizione>
File coinvolto: <path>
Azione proposta: <azione sicura>
Commit/push: non eseguiti
```

| Blocco | Causa | Azione |
|---|---|---|
| `BLOCCO_GIT_STATO_SPORCO` | Modifiche tracciate inattese | Fermarsi, riportare `git diff --name-status`, chiedere conferma operatore — non usare git restore/reset automaticamente |
| `BLOCCO_DECISIONALI_BANCA_AUSENTES` | Excel decisional banca mancanti | Eseguire skill `regole-registrazioni-banca` |
| `BLOCCO_CSV_MULTIFONTE_AUSENTES` | CSV multifonte assenti o non indicati (policy skill) | Indicare i CSV comprovantes/conta corrente validati |
| `BLOCCO_OUTPUT_MULTIFONTE_EXISTENTE` | Cartella output contiene già artefatti bridge | Creare TAG nuovo oppure chiedere approvazione esplicita all'operatore prima di sovrascrivere |
| `BLOCCO_BRIDGE_BANCA_DUPLICADOS_PENDENTES` | Duplicati non risolti nel bridge | Riesaminare Excel decisional banca |
| `BLOCCO_BRIDGE_BANCA_ZERO_APROVADAS` | Bridge non ha righe approvate | Verificare periodo e file decisionali |
| `BLOCCO_MANUAIS_NAO_VALIDADOS` | CSV manuais senza manifest approvato | Eseguire prima skill `lanciamento-registrazioni-manuali` |
| `BLOCCO_RUN_ERRORE` | Pipeline interrotta con errore | Leggere output del comando per dettagli |
| `BLOCCO_MANIFEST_RUN_INVALIDO` | manifest_run.json non conferma status_run, banca inclusa, linhas_lidas o checkpoints DRE/pacote | Fermarsi, riportare i valori letti, non proseguire |
| `BLOCCO_DRE_REVISAR_NONZERO` | DRE ha qualsiasi REVISAR > 0 (indipendentemente dalla causa) | Eseguire audit DRE con skill `auditar-dre-gerencial` |
| `BLOCCO_PRE_LIVRO_REVISAR_NON_COMODATO` | pre-livro ha voci REVISAR che non sono COMODATO atteso | Eseguire audit giornale con skill `auditar-giornale-multifonte` e classificare la causa |
| `BLOCCO_99_99_998_IN_DRE` | Conta tecnica nel DRE | Verificare `mapeamento_dre.yaml` (entra_dre=false per 99.99.998) |
| `BLOCCO_OUTPUT_STAGED` | Output reali staged | `git restore --staged <file>` per de-stageare |
| `BLOCCO_YAML_MODIFICATO` | YAML config cambiati | Investigare prima di procedere |

---

## 10. Guardrail Git

```text
NON usare git add .
NON fare git clean / git reset / git restore sui file di lavoro
CONSENTITO solo `git restore --staged <path>` per rimuovere dallo stage output accidentali
NON committare output reali (output/, reports/, Database_Dev/)
NON fare push
NON stageare lancamentos_banca_aprovados*.xlsx
NON stageare manifest_lancamentos_banca.json
```

---

## 11. Guardrail Contabili

```text
NON modificare config/mapeamento_dre.yaml
NON modificare config/plano_contas_gerencial.yaml
NON modificare regras_classificacao_*.yaml
NON modificare codice Python (src/)
NON inserire 99.99.998 manualmente
NON creare regole contabili nuove
NON classificare senza regola esistente
NON forzare lancamentos REVISAR a OK
NON usare --modo definitivo
```

---

## 12. Criteri di Successo

Verifiche organizzate per fonte.

### 12.1 — `manifest_lancamentos_banca.json` (bridge)

```text
[x] linhas_aprovadas_incluidas > 0
[x] deduplicacao.possiveis_duplicados_data_divergente = 0
[x] deduplicacao.ambiguos_data_divergente = 0
```

Solo riferimento storico, non criterio bloccante per run future:
```text
[ ] deduplicacao.duplicados_bancarios_data_divergente_excluidos = 11 (baseline jan-apr 2026)
```

### 12.2 — `manifest_run.json`

```text
[x] status_run = PROVISORIO_CONCLUIDO
[x] lancamentos_banca_aprovados.incluido = true
[x] lancamentos_banca_aprovados.linhas_lidas > 0
[x] checkpoints contiene descricao = "DRE_OK_PROVISORIA_GERADA"
[x] checkpoints contiene descricao = "PACOTE_CONTABIL_GERENCIAL_GERADO"
```

### 12.3 — `dre_gerencial_ok_provisoria.txt`

```text
[x] Status DRE: OK_PROVISORIO
```

### 12.4 — XLSX e skill di audit dedicate

I seguenti controlli strutturali vanno verificati tramite
`auditar-dre-gerencial` e/o `auditar-giornale-multifonte`:

```text
[x] DRE REVISAR = 0
[x] 99.99.998 assente dal DRE
[x] contas 99.* assenti dal DRE
[x] contrapartidas assenti dal pre-livro e DRE
[x] bonificações in 03.08 con sinal +
[x] COMODATO = 0 nel DRE (6 COMODATO restano REVISAR nel pre-livro)
[x] M A BRITO: NF 2684 esclusa da XML, pagamenti banca inclusi
```

### 12.5 — Git (post-run)

```text
[x] git diff --cached = vuoto
[x] git diff config/ = vuoto
[x] git diff src/ = vuoto
[x] git diff tests/ = vuoto
[x] manifest_run.json generato
[x] pacote_contabil_gerencial_provisorio/ generato
```

---

## 13. Criteri di Stop

Fermarsi immediatamente e non proseguire se:

```text
- bridge banca ha possiveis_duplicados > 0 o ambiguos > 0
- cartella output contiene già artefatti bridge (senza approvazione operatore)
- CSV multifonte comprovantes/conta corrente non indicati o inesistenti
- CSV manuais indicato senza manifest di validazione approvato
- pipeline produce errore non recuperabile
- manifest_run.json non conferma status_run, banca inclusa o linhas_lidas
- DRE ha qualsiasi REVISAR > 0 → BLOCCO_DRE_REVISAR_NONZERO: eseguire auditar-dre-gerencial
- pre-livro ha REVISAR non-COMODATO → BLOCCO_PRE_LIVRO_REVISAR_NON_COMODATO: eseguire auditar-giornale-multifonte
- 99.99.998 appare nel DRE
- file tracciati (config/, src/) risultano modificati
- output reali risultano staged
```

---

## 14. Esempio Reale — Baseline 007

La baseline 007 è stata eseguita come `run 2026-05-26_001`.

> **Nota path:** la baseline 007 usa la convenzione `multifonte_007` (senza anno e
> mesi) poiché prodotta prima della standardizzazione. Le run future usano la
> convenzione `multifonte_NNN_YYYY_MM_MM` (es. `multifonte_009_2026_01_04`).

```powershell
# Bridge banca — baseline 007
$TAG_007 = "multifonte_007"
$OUT_007 = "output\$TAG_007"

.\.venv\Scripts\python.exe -m fantozzi.cli gerar-lancamentos-banca-aprovados `
  --comprovantes-csv reports\comprovantes_classificacao\runs\audit_002c_2026-05-26\comprovantes_classificados.csv `
  --conta-corrente-csv output\audit_002b\conta_corrente_classificados_combined.csv `
  --comprovantes-xlsx bank_raw\_decisoes_regras\Analisi_regole_registrazioni_banca_comprovanti.xlsx `
  --conta-corrente-xlsx bank_raw\_decisoes_regras\Analisi_regole_registrazioni_banca_conto_corrente.xlsx `
  --saida "$OUT_007\lancamentos_banca_aprovados_multifonte_007.xlsx" `
  --periodo-ini 2026-01-01 `
  --periodo-fim 2026-04-30

# Risultati bridge osservati nella baseline 007:
#   linhas_aprovadas_incluidas: 205
#   deduplicacao.duplicados_excluidos: 83
#   deduplicacao.duplicados_bancarios_data_divergente_excluidos: 11  ← solo storico
#   deduplicacao.possiveis_duplicados_data_divergente: 0
#   deduplicacao.ambiguos_data_divergente: 0

# Run multifonte (non ripetere, la 007 è approvata — solo come riferimento)
.\.venv\Scripts\python.exe -m fantozzi.cli reprocessar-importacao-real `
  --xml-base Database_Dev\XML_fatture_acquisto `
  --saida-output "$OUT_007\reprocessamento_output" `
  --saida-reports "$OUT_007\reprocessamento_reports" `
  --lancamentos-banca-aprovados "$OUT_007\lancamentos_banca_aprovados_multifonte_007.xlsx" `
  --gerar-pacote-contabil-gerencial

# Risultati attesi (baseline 007):
#   status_run: PROVISORIO_CONCLUIDO
#   lancamentos_banca_aprovados.incluido: true
#   lancamentos_banca_aprovados.linhas_lidas: 205
#   Status DRE: OK_PROVISORIO
#   DRE lancamentos OK: 1040 (XML-08884: 370 + XML-55746: 465 + BANCA: 205)
#   DRE valor total base OK: R$ 1.371.347,44
#   bonificações 03.08: R$ 1.536,12 (24 righe)
```

---

## 15. Troubleshooting Operativo

**La run non trova XML:**
```text
Verificare che Database_Dev\XML_fatture_acquisto\ contenga sottocartelle
con nome = CNPJ (es: 08884180000123\ e 55746722000181\).
```

**Il bridge genera 0 righe approvate:**
```text
Verificare che i file decisionali banca contengano righe con STATUS_DECISAO=APROVADO.
Verificare che i CSV multifonte classificati siano quelli validati per la baseline.
Verificare che il periodo (--periodo-ini / --periodo-fim) coincida con le date presenti.
```

**DRE con REVISAR > 0:**
```text
Qualsiasi REVISAR > 0 nel DRE è un blocco — BLOCCO_DRE_REVISAR_NONZERO.
Azione: eseguire skill auditar-dre-gerencial e fermarsi.
Il DRE deve sempre avere REVISAR = 0 e COMODATO = 0.
```

**pre-livro con REVISAR non-COMODATO:**
```text
Nel pre-livro le 6 voci COMODATO restano REVISAR per design — è normale.
Se appaiono voci REVISAR che NON sono COMODATO: BLOCCO_PRE_LIVRO_REVISAR_NON_COMODATO.
Azione: eseguire skill auditar-giornale-multifonte e classificare la causa.
```

**99.99.998 appare nel DRE:**
```text
Verificare config/mapeamento_dre.yaml: la conta 99.99.998 deve avere entra_dre=false.
Non procedere fino alla risoluzione.
```

**Output staged accidentalmente:**
```text
git diff --cached --name-status
git restore --staged <file>
```

Usare `git restore --staged <file>` solo per de-stageare output accidentali.
Non usare `git restore` sui file di lavoro.

**manifest_run.json non conferma status_run:**
```text
Fermarsi. Riportare il valore letto di status_run e di lancamentos_banca_aprovados.
Non proseguire con l'audit finale finché non è chiaro cosa ha interrotto la pipeline.
```

**Cartella output già esistente con artefatti bridge:**
```text
Creare un TAG nuovo (es. multifonte_009b_2026_01_04) oppure
chiedere autorizzazione esplicita all'operatore prima di sovrascrivere.
```

---

## 16. Cosa Fare / Cosa NON Fare

```text
FARE:
  Verificare Git prima di qualsiasi operazione
  Verificare assenza di artefatti bridge nella cartella output (isolation guard)
  Verificare l'esistenza dei CSV multifonte classificati
  Leggere sempre manifest bridge (§8.1) prima della run
  Controllare deduplicacao.possiveis_duplicados_data_divergente e ambiguos_data_divergente
  Leggere manifest_run.json (§8.2) dopo la run usando i campi corretti
  Verificare Status DRE: OK_PROVISORIO nel TXT (§8.3)
  Delegare i controlli strutturali DRE alle skill dedicate
  Verificare git diff --cached dopo la run

NON FARE:
  NON modificare YAML senza autorizzazione
  NON usare git add .
  NON fare push
  NON stageare output reali
  NON usare --modo definitivo
  NON ignorare blocchi o warning del bridge
  NON trattare la baseline 007 come precondizione rigida: usarla come riferimento storico
  NON inventare path o opzioni non documentate
  NON usare git restore/reset/checkout automaticamente su diff inattesi
  NON sovrascrivere artefatti bridge senza approvazione esplicita dell'operatore
  NON cercare nel manifest_run.json le chiavi status o lancamentos_banca_aprovados_incluidos
  NON usare dre_gerencial_ok_provisoria.txt per controlli strutturali DRE
```

---

## Riferimenti Tecnici

- CLI: `.\.venv\Scripts\python.exe -m fantozzi.cli`
- Bridge banca: `gerar-lancamentos-banca-aprovados`
- Pipeline multifonte full run: `reprocessar-importacao-real --full-run`
- Pipeline manuali incrementale: `executar-run-incremental`
- Baseline 007: `output\multifonte_007\audit_007_baseline_post_006_summary.json`
- Piano conti: `config\plano_contas_gerencial.yaml`
- Mapeamento DRE: `config\mapeamento_dre.yaml`
- Skill correlate: `auditar-giornale-multifonte`, `auditar-dre-gerencial`,
  `lanciamento-registrazioni-manuali`, `workflow-operatore-final`
- Documento di progettazione: `output\skills_008\FZ_SKILLS_GIORNALE_DRE_008_PROGETTAZIONE.md`
