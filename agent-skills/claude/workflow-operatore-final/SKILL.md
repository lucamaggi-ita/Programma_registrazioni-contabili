---
name: workflow-operatore-final
description: Orchestratore end-to-end per operatore non tecnico. Usa fast path incrementale per manuali quando esiste baseline approvata; full run e audit solo quando richiesti esplicitamente. Non modifica YAML, non fa commit/push.
---

# Skill: workflow-operatore-final

Skill di orchestrazione end-to-end per l'intero processo contabile gerencial
multifonte. Guida l'operatore (anche non tecnico) dall'inizio alla fine, nella
sequenza corretta, richiamando le skill specializzate al momento giusto.

**L'operatore deve solo:**
- rispondere alle domande di questa skill a ogni passo;
- fornire le decisioni contabili (classificazioni, approvazioni);
- confermare o bloccare il progresso.

---

## 1. Scopo

Orchestrare l'intero workflow contabile gerencial:

```text
Fase 1: Verifica Git e precondizioni
Fase 2: Aggiornamento regole XML (se necessario)
Fase 3: Aggiornamento regole banca/comprovantes (se necessario)
Fase 4: Lançamentos manuais (se necessario, fast path incrementale se baseline approvata)
Fase 5: Generazione giornale + DRE multifonte solo per full run esplicita
Fase 6: Audit giornale solo se richiesto
Fase 7: Audit DRE solo se richiesto
Fase 8: Verdetto baseline
Fase 9: Consegna pacchetto contábil gerencial
```

---

## 2. Quando usare questa skill

Usare quando si vuole eseguire il processo contabile completo da zero o
aggiornare una baseline esistente con nuovi dati o nuove decisioni.

Se la richiesta e solo:

```text
Fantozzi, voglio fare registrazioni manuali
```

non avviare il workflow completo. Richiamare direttamente
`lanciamento-registrazioni-manuali`: Excel -> validazione ->
`executar-run-incremental` -> riepilogo sintetico.

Usare anche come punto di partenza se non si è sicuri da dove cominciare.

---

## 3. Quando NON usare questa skill

Non usare per:

- eseguire solo una fase specifica (usare la skill dedicata);
- auditare una run già esistente senza rigenerare (usare `auditar-giornale-multifonte`
  o `auditar-dre-gerencial` direttamente);
- aggiornare solo regole XML senza rigenerare (usare `regole-registrazioni-xml`).

---

## 4. Precondizioni Iniziali

Nessun prerequisito tecnico avanzato. La skill guida passo per passo.

Requisiti minimi:
- ambiente Python configurato (`.venv` presente);
- accesso ai file sorgente (XML NF-e, estratti banca, comprovantes);
- YAML di configurazione non modificati senza autorizzazione.

---

## 5. Panoramica Skill Richiamate

| Fase | Skill richiamata | Quando |
|---|---|---|
| Regole XML | `regole-registrazioni-xml` | Se ci sono nuove classificazioni XML da approvare |
| Regole banca | `regole-registrazioni-banca` | Se ci sono nuove classificazioni banca da approvare |
| Excel revisione banca | `importa-revisione-banca` | Se i file decisionali banca devono essere aggiornati |
| Excel revisione XML | `importa-revisione-xml` | Se i file decisionali XML devono essere aggiornati |
| Manuais | `lanciamento-registrazioni-manuali` | Se ci sono spese manuali da aggiungere |
| Giornale + DRE full run | `gerar-giornale-dre-multifonte` | Solo se richiesta full run esplicita o serve nuova baseline |
| Audit giornale | `auditar-giornale-multifonte` | Solo su richiesta esplicita |
| Audit DRE | `auditar-dre-gerencial` | Solo su richiesta esplicita |

---

## 6. Workflow Completo Passo per Passo

---

### FASE 1 — Verifica Git e Precondizioni

**Comandi:**
```powershell
git status -sb
git rev-list --left-right --count origin/master...HEAD
git diff --name-status
git diff --name-status config/
git diff --name-status src/
git diff --name-status tests/
git diff --cached --name-status
```

**Atteso:**
- branch `master`;
- `origin/master...HEAD = 0 0`;
- nessun diff tracciato inatteso;
- nessun diff su `config/`, `src/`, `tests/`;
- nessun file staged.

**Fermarsi se:**
```text
BLOCCO: BLOCCO_GIT_FASE1
Causa: diff su file tracciati o staging inatteso
Azione: fermarsi, riportare il diff, chiedere all'operatore se le modifiche sono attese
Divieto: non usare restore/reset/checkout automaticamente
```

**Domanda all'operatore:**
```text
Ci sono nuove regole XML da aggiornare prima di questa run?
→ Sì → procedere con FASE 2
→ No → passare a FASE 3
```

---

### FASE 2 — Aggiornamento Regole XML (se necessario)

Richiamare: `regole-registrazioni-xml` o `importa-revisione-xml`

**Quando saltare questa fase:**
- regole XML già aggiornate nell'ultima sessione;
- nessuna nuova NF-e da classificare;
- l'operatore conferma che le regole sono attuali.

**Output atteso prima di continuare:**
- `config/regras_classificacao_produtos.yaml` aggiornato e verificato;
- nessun commit non autorizzato.

---

### FASE 3 — Aggiornamento Regole Banca/Comprovantes (se necessario)

Richiamare: `regole-registrazioni-banca` o `importa-revisione-banca`

**Quando saltare questa fase:**
- file decisionali banca già aggiornati;
- l'operatore conferma che le decisioni APROVADO sono attuali per il periodo.

**Verificare che esistano:**
```powershell
Test-Path "bank_raw\_decisoes_regras\Analisi_regole_registrazioni_banca_comprovanti.xlsx"
Test-Path "bank_raw\_decisoes_regras\Analisi_regole_registrazioni_banca_conto_corrente.xlsx"
```

**Output atteso prima di continuare:**
- `config/regras_classificacao_comprovantes.yaml` aggiornato;
- Excel decisionali aggiornati con decisioni APROVADO per il periodo.

**Domanda all'operatore:**
```text
Ci sono lançamentos manuali da aggiungere a questa run?
(spese manuali non presenti in XML né in banca/comprovantes)
→ Sì → procedere con FASE 4
→ No → passare a FASE 5
```

---

### FASE 4 — Lançamentos Manuais (se necessario)

Richiamare: `lanciamento-registrazioni-manuali`

Se esiste `output\baselines\baseline_atual.json` approvata, questa fase deve
usare il fast path incrementale:

```text
Excel operativo -> validar-lancamentos-manuais-gerenciais -> executar-run-incremental
```

Non passare da `gerar-giornale-dre-multifonte` e non chiamare
`reprocessar-importacao-real` per manuali operatore ordinari.

**Quando saltare questa fase:**
- nessuna spesa manuale da aggiungere;
- l'operatore conferma che non ci sono movimenti fuori da XML e banca.

**Output atteso prima di continuare:**
- `Database_Dev\Registrazioni_manuali\registrazioni_manuali_tecnico.csv` generato e validato.
- `Database_Dev\Registrazioni_manuali\registrazioni_manuali_validazione_manifest.json` presente e approvato.

**Nota 011B:** il CSV tecnico manuais viene usato direttamente da
`executar-run-incremental` insieme al manifest di validazione.

Per poter usare i manuais nella FASE 5, non basta il CSV tecnico.
Serve anche il manifest di validazione prodotto dalla skill
`lanciamento-registrazioni-manuali`.

**Manifest richiesto:**
```text
Database_Dev\Registrazioni_manuali\registrazioni_manuali_validazione_manifest.json
```

**Campi minimi:**
```text
status_validacao = APROVADO
arquivo_csv_tecnico = path del CSV usato nella run
rejeitados = 0
pendencias = 0 oppure pendencias_justificadas = true
duplicidade_xml_banca = 0 oppure duplicidades_justificadas = true
```

**Fermarsi se:**
```text
BLOCCO: BLOCCO_MANUAIS_NAO_VALIDADOS
Causa: CSV manuais indicato senza manifest di validazione approvato prodotto dalla skill manuale
Azione: eseguire lanciamento-registrazioni-manuali o non usare manuais in questa run
```

Se l'incrementale termina con successo, il workflow manuali e concluso. Non
proseguire automaticamente a FASE 5, FASE 6 o FASE 7.

---

### FASE 5 — Generazione Giornale + DRE Multifonte

Richiamare: `gerar-giornale-dre-multifonte`

Questa fase e una full run. Eseguirla solo se l'operatore chiede esplicitamente
una full run, se manca una baseline approvata da costruire, oppure se una
modifica strutturale rende necessario rigenerare la baseline.

Se esiste baseline approvata e la richiesta riguarda solo manuali, saltare questa
fase: il percorso corretto e `executar-run-incremental`.

Prima di eseguire la run, definire un TAG univoco della fase/periodo.

**Esempio operativo:**
```powershell
$TAG = "multifonte_009_2026_01_04"
$OUT = "output\$TAG"
```

Tutti gli output della run 009 devono stare sotto `$OUT`.
Non usare `reports\importacao_real` come destinazione principale della run
hardenizzata.

**Domanda all'operatore prima di eseguire:**
```text
Periodo da coprire?
Esempio: 2026-01-01 → 2026-04-30
Risposta: YYYY-MM-DD → YYYY-MM-DD
```

Per la run multifonte operativa, la skill `gerar-giornale-dre-multifonte`
richiede i CSV multifonte classificati. Questi flag sono opzionali per la CLI,
ma obbligatori per policy della skill.

**Input CSV richiesti:**
```text
$COMPROVANTES_CSV = path al comprovantes_classificados.csv validato
$CONTA_CORRENTE_CSV = path al conta_corrente_classificados_combined.csv validato
```

**Esempi di path:**
```text
reports\comprovantes_classificacao\runs\...\comprovantes_classificados.csv
output\audit_002b\conta_corrente_classificados_combined.csv
```

**Fermarsi se:**
```text
BLOCCO: BLOCCO_CSV_MULTIFONTE_AUSENTES
Causa: CSV comprovantes o CSV conta corrente classificato assente/non indicato
Azione: generare o indicare i CSV multifonte validati prima di FASE 5
```

**Sequenza della skill `gerar-giornale-dre-multifonte`:**
1. Precheck Git
2. Verifica file decisionali banca
3. Bridge banca con il periodo indicato
4. Verifica `manifest_lancamentos_banca.json`:
   - `linhas_aprovadas_incluidas > 0`
   - `deduplicacao.possiveis_duplicados_data_divergente = 0`
   - `deduplicacao.ambiguos_data_divergente = 0`
5. Run multifonte completa
6. Verifica post-run (git diff --cached = vuoto)

**Fermarsi se:**
```text
BLOCCO: BLOCCO_BRIDGE_BANCA_ZERO_APROVADAS
Causa: manifest bridge indica linhas_aprovadas_incluidas <= 0
Azione: verificare periodo, CSV multifonte e file decisionali banca
```

**Output obbligatori prima di procedere:**
```text
$OUT\reprocessamento_reports\runs\{run_id}\dre_gerencial_ok_provisoria.xlsx
$OUT\reprocessamento_reports\runs\{run_id}\pre_livro_diario_gerencial_teste.xlsx
$OUT\reprocessamento_reports\runs\{run_id}\manifest_run.json
$OUT\reprocessamento_reports\runs\{run_id}\pacote_contabil_gerencial_provisorio\
```

**Output operativi di importazione attesi:**
```text
$OUT\reprocessamento_output\runs\{run_id}\
  08884180000123\lancamentos_fantozzi_real.xlsx
  55746722000181\lancamentos_fantozzi_real.xlsx
  BANCA_APROVADA\lancamentos_banca_aprovados.xlsx
```

**Fermarsi se la skill `gerar-giornale-dre-multifonte` riporta un BLOCCO.**

---

### FASE 6 — Audit Giornale

Richiamare: `auditar-giornale-multifonte`

Non lanciarla automaticamente dopo il fast path manuali incrementale. L'audit e
una skill separata e va eseguita solo se l'operatore lo chiede.

**Input:** `run_id` e `$OUT` dalla FASE 5.

**Atteso:**
```text
verdetto: APROVADO o APROVADO_COM_RESSALVAS (nessun P1)
```

**Fermarsi se verdetto = REPROVADO:**
```text
BLOCCO: BLOCCO_AUDIT_GIORNALE_REPROVADO
Causa: finding P1 nel giornale multifonte
Azione: segnalare all'utente — non procedere al DRE
```

---

### FASE 7 — Audit DRE

Richiamare: `auditar-dre-gerencial`

Non lanciarla automaticamente dopo il fast path manuali incrementale. L'audit e
una skill separata e va eseguita solo se l'operatore lo chiede.

**Input:** `run_id` e `$OUT` dalla FASE 5.

**Atteso:**
```text
verdetto: APROVADO o APROVADO_COM_RESSALVAS (nessun P1)
```

**Fermarsi se verdetto = REPROVADO:**
```text
BLOCCO: BLOCCO_AUDIT_DRE_REPROVADO
Causa: finding P1 nel DRE gerencial
Azione: segnalare all'utente — non approvare baseline
```

---

### FASE 8 — Verdetto Baseline

Raccogliere i verdetti dalle FASI 6 e 7.

**Criteri per verdetto baseline:**

```text
BASELINE_APROVADA
  Audit giornale: APROVADO o APROVADO_COM_RESSALVAS
  Audit DRE: APROVADO o APROVADO_COM_RESSALVAS
  Nessun P1 in nessun audit.

BASELINE_APROVADA_COM_RESSALVAS
  Uno o più P2 in audit giornale o DRE, con decisione esplicita dell'operatore
  di accettare ogni finding P2 residuo.

BASELINE_REPROVADA
  Uno o più P1 in qualsiasi audit.
  Non consegnare il pacchetto come baseline approvata.
  Segnalare all'utente i finding bloccanti.
```

**Report baseline da produrre:**
```text
run_id               : {run_id}
data_execucao        : {data}
periodo              : {periodo_ini} — {periodo_fim}
audit_giornale       : APROVADO / APROVADO_COM_RESSALVAS / REPROVADO
audit_dre            : APROVADO / APROVADO_COM_RESSALVAS / REPROVADO
finding_p1           : nessuno / [lista]
finding_p2           : nessuno / [lista]
verdetto_baseline    : BASELINE_APROVADA / ... / BASELINE_REPROVADA
```

---

### FASE 9 — Consegna Pacchetto Contábil Gerencial

Eseguire solo se `verdetto_baseline = BASELINE_APROVADA` o `BASELINE_APROVADA_COM_RESSALVAS`.

**Output da consegnare all'operatore/contador:**

```text
$OUT\reprocessamento_reports\runs\{run_id}\
  dre_gerencial_ok_provisoria.xlsx
  dre_gerencial_ok_provisoria.txt
  pre_livro_diario_gerencial_teste.xlsx
  relatorios_gerenciais_ok_provisorios.xlsx
  relatorios_gerenciais_ok_provisorios.txt
  auditoria_pre_livro_diario_gerencial_teste.txt
  auditoria_relatorios_gerenciais_ok_provisorios.xlsx
  auditoria_relatorios_gerenciais_ok_provisorios.txt
  pacote_contabil_gerencial_provisorio\
  manifest_run.json
  audit_giornale_multifonte.md   ← se generato dalla FASE 6
  audit_dre_gerencial.md         ← se generato dalla FASE 7
```

Il pacchetto operativo della 009 va cercato sotto `$OUT`, non sotto
`reports\importacao_real`.

**Nota:** il commit/push del pacchetto o della baseline NON è compito
di questa skill. Richiede autorizzazione esplicita dell'utente.

---

## 7. Quando Fermarsi

```text
STOP obbligatorio in queste condizioni:
- Fase 1: Git sporco con diff su file tracciati
- Fase 3: file decisionali banca assenti
- Fase 4: CSV manuais senza manifest approvato dalla skill manuale
  → BLOCCO_MANUAIS_NAO_VALIDADOS
- Fase 5: CSV multifonte comprovantes/conta corrente assenti o non indicati
  → BLOCCO_CSV_MULTIFONTE_AUSENTES
- Fase 5: bridge banca con linhas_aprovadas_incluidas <= 0
  → BLOCCO_BRIDGE_BANCA_ZERO_APROVADAS
- Fase 5: bridge banca con deduplicacao.possiveis_duplicados_data_divergente > 0
  o deduplicacao.ambiguos_data_divergente > 0
- Fase 5: run multifonte con errore non recuperabile
- Fase 5: DRE con qualsiasi REVISAR > 0
  → BLOCCO_DRE_REVISAR_NONZERO
- Fase 5/6: pre-livro con REVISAR non-COMODATO
  → BLOCCO_PRE_LIVRO_REVISAR_NON_COMODATO
- Fase 6: audit giornale REPROVADO con P1
- Fase 7: audit DRE REPROVADO con P1
- Qualsiasi fase: YAML modificati senza autorizzazione
- Qualsiasi fase: output reali staged
```

DRE: qualsiasi `REVISAR > 0` è blocker.
Pre-livro: eventuali `REVISAR` sono accettabili solo se `COMODATO` atteso e
fuori DRE.

---

## 8. Quando Chiedere Revisione Manuale

```text
Chiedere revisione manuale all'operatore quando:
- ci sono nuovi movimenti banca non coperti da regole → usare regole-registrazioni-banca
- ci sono nuovi prodotti XML non classificati → usare regole-registrazioni-xml
- ci sono spese manuali da aggiungere → usare lanciamento-registrazioni-manuali
- audit giornale riporta finding P2 → l'operatore decide se accettare o bloccare
- audit DRE riporta finding P2 → l'operatore decide se accettare o bloccare
```

---

## 9. Quando Non Procedere al DRE

```text
Non procedere al DRE se:
- audit giornale REPROVADO
- bridge banca ha duplicati non risolti
- run multifonte è terminata con errore
- YAML non sono allineati con l'ultima approvazione
```

---

## 10. Output da Consegnare al Contador/Operatore

```text
Base path:
$OUT\reprocessamento_reports\runs\{run_id}\

File contábil operativi:
  dre_gerencial_ok_provisoria.xlsx
  dre_gerencial_ok_provisoria.txt
  pre_livro_diario_gerencial_teste.xlsx
  relatorios_gerenciais_ok_provisorios.xlsx
  relatorios_gerenciais_ok_provisorios.txt

File di audit:
  auditoria_pre_livro_diario_gerencial_teste.txt
  auditoria_relatorios_gerenciais_ok_provisorios.xlsx
  auditoria_relatorios_gerenciais_ok_provisorios.txt
  audit_giornale_multifonte.md
  audit_dre_gerencial.md

Pacchetto:
  pacote_contabil_gerencial_provisorio\

Manifest tracciabilità:
  manifest_run.json
  $OUT\manifest_lancamentos_banca.json
```

---

## 11. Guardrail Git (validi per tutto il workflow)

```text
NON usare git add .
NON fare git clean / git reset / git restore su file di lavoro
Consentito solo `git restore --staged <path>` per rimuovere dallo stage output accidentali
NON committare output reali (output/, reports/)
NON fare push
NON stageare file della run
VERIFICARE git diff --cached dopo ogni fase
```

---

## 12. Guardrail Contabili (validi per tutto il workflow)

```text
NON modificare YAML senza conferma utente
NON inferire conti
NON creare regole nuove senza decisione utente
NON forzare lancamentos REVISAR a OK
NON ignorare finding P1 negli audit
NON usare --modo definitivo
```

---

## 13. Criteri di Successo del Workflow Completo

```text
[x] Git allineato alla partenza e alla fine
[x] Regole XML aggiornate (o confermate attuali)
[x] Regole banca aggiornate (o confermate attuali)
[x] Bridge banca: linhas_aprovadas_incluidas > 0
[x] Bridge banca: deduplicacao.possiveis_duplicados_data_divergente = 0
[x] Bridge banca: deduplicacao.ambiguos_data_divergente = 0
[x] Run multifonte: manifest_run.status_run = PROVISORIO_CONCLUIDO
[x] Run multifonte: manifest_run.lancamentos_banca_aprovados.incluido = true
[x] Run multifonte: manifest_run.lancamentos_banca_aprovados.linhas_lidas > 0
[x] Run multifonte: manifest_run.checkpoints contém descricao = "DRE_OK_PROVISORIA_GERADA"
[x] Run multifonte: manifest_run.checkpoints contém descricao = "PACOTE_CONTABIL_GERENCIAL_GERADO"
[x] Audit giornale: APROVADO o APROVADO_COM_RESSALVAS, solo se richiesto
[x] Audit DRE: APROVADO o APROVADO_COM_RESSALVAS, solo se richiesto
[x] DRE: 99.99.998 = 0, COMODATO = 0, bonificações in 03.08
[x] Pacote contábil gerencial generato
[x] Nessun output staged
[x] Nessun YAML modificato senza autorizzazione
```

---

## Riferimenti Tecnici

- Skill giornale + DRE: `gerar-giornale-dre-multifonte`
- Skill audit giornale: `auditar-giornale-multifonte`
- Skill audit DRE: `auditar-dre-gerencial`
- Skill manuais: `lanciamento-registrazioni-manuali`
- Skill regole XML: `regole-registrazioni-xml`
- Skill regole banca: `regole-registrazioni-banca`
- Baseline 007: `output/multifonte_007/audit_007_baseline_post_006_summary.json`
- Documento di progettazione: `output/skills_008/FZ_SKILLS_GIORNALE_DRE_008_PROGETTAZIONE.md`
- Workflow operativo docs: `docs/WORKFLOW_OPERATORE_FINALE.md`
