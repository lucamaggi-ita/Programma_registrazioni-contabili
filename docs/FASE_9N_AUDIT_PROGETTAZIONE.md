# FANTOZZI - Fase 9N: audit e progettazione

Data: 2026-05-16
Stato: AUDIT / DESIGN - nessun codice funzionale implementato
Branch atteso: master
Commit base atteso: 9bf7fce Inclui bonificacao recebida no DRE gerencial

---

## 0. Scopo

La Fase 9N introduce lancamentos manuais gerenciais controlados senza editare a mano gli XLSX finali.

La fase deve:

- leggere un CSV manuale esplicito passato da CLI;
- validarlo prima di generare qualunque output contabile;
- convertire le righe approvate in lancamentos compatibili con la pipeline attuale;
- marcare ogni riga con origine `MANUAL_GERENCIAL`;
- generare audit CSV/TXT dei manuali applicati;
- integrare i manuali in pre-livro, relatorios gerenciais OK e DRE consolidata;
- registrare la sezione dedicata nel manifest;
- non modificare alias, classificazioni prodotto, YAML di config o modo definitivo.

Fuori scope: estratto conto bancario, comprovantes, OCR, matching bancario, modo definitivo.

---

## 1. Pre-check eseguito

Comandi eseguiti prima della progettazione:

```powershell
git status -sb
git rev-list --left-right --count origin/master...HEAD
git log --oneline --decorate -5
git diff --name-status
git diff --check
```

Risultato:

- `HEAD -> master, origin/master` su `9bf7fce`.
- `origin/master...HEAD = 0 0`.
- Nessuna diff tracciata.
- `git diff --check` pulito.
- Untracked solo quelli attesi: `Database_Dev/`, `output/importacao_real/`, `reports/importacao_real/`, `pendenze_dettaglio_2026-05-14_005.txt`.

Nota: `git status` mostra warning di permesso su `C:\Users\Recanto do Sossego\.config\git\ignore`; non blocca il lavoro.

---

## 2. Decisioni contabili da preservare

Questa fase non riapre decisioni gia prese.

- I 2 CNPJ restano consolidati come una sola impresa gestionale.
- COMODATO resta fuori DRE e controllo extracontabile.
- BONIFICACAO_RECEBIDA resta in DRE/pre-livro come `2.05.003` con `REDUCAO_CUSTO`.
- `C_PENDENCIA_VERDADEIRA = 0`: nessuna nuova fase alias/classificazione.
- `fase9_executada` non deve essere usato per 9N. La 9N e un complemento gerencial manuale, non la Fase 9 di writeback classificazioni.

---

## 3. Audit tecnico del codice attuale

### 3.1 Pipeline

Riferimenti:

- `src/fantozzi/reprocessamento_pipeline.py:90` definisce `PipelineContext`.
- `src/fantozzi/reprocessamento_pipeline.py:130` definisce `reprocessar_importacao_real_provisoria`.
- `src/fantozzi/reprocessamento_pipeline.py:223` valorizza `context.lancamentos_xlsx` dopo l'import XML.
- `src/fantozzi/reprocessamento_pipeline.py:532` genera il pre-livro.
- `src/fantozzi/reprocessamento_pipeline.py:713` genera i relatorios OK.
- `src/fantozzi/reprocessamento_pipeline.py:754` genera la DRE.
- `src/fantozzi/reprocessamento_pipeline.py:405` aggiorna il manifest partendo da `StepResult`.
- `src/fantozzi/reprocessamento_pipeline.py:630` contiene lo step Fase 9, distinto e da non riusare per 9N.

Osservazione:

La pipeline e gia step-based. Il punto naturale per 9N e dopo `OPERACOES_ESPECIAIS_AUDITADAS` e prima di `PRE_LIVRO_GERADO`, cosi i manuali entrano in pre-livro, relatorios OK e DRE senza toccare la classificazione XML.

### 3.2 Lettura dei lancamentos

Riferimenti:

- `src/fantozzi/excel_writer.py:20` definisce `HEADERS`.
- `src/fantozzi/excel_writer.py:55` scrive gli XLSX di lancamentos.
- `src/fantozzi/pipeline_paths.py` risolve ricorsivamente gli `.xlsx` sotto una run.

Osservazione:

I report correnti leggono XLSX con aba `Lancamentos` e colonne `HEADERS`. Per integrarsi con il minimo rischio, la 9N deve convertire il CSV manuale in un XLSX tecnico con la stessa struttura, scritto dentro la run di `output`.

### 3.3 Pre-livro

Riferimenti:

- `src/fantozzi/pre_livro_diario.py:109` genera il pre-livro.
- `src/fantozzi/pre_livro_diario.py:263` legge gli XLSX di lancamentos.
- `src/fantozzi/pre_livro_diario.py:338` accumula il resumo por conta.
- `src/fantozzi/pre_livro_diario.py:427` costruisce auditoria.
- `src/fantozzi/pre_livro_diario.py:574` scrive le abas di lancamentos.

Osservazione:

Il pre-livro oggi non ha logica di partita doppia. Lavora su righe gestionali singole per conta. La 9N puo rappresentare ogni lancamento manuale a partita doppia come due righe tecniche:

- una riga `DEBITO` sulla conta debito;
- una riga `CREDITO` sulla conta credito.

Per il DRE, entrera solo la riga la cui conta e mappata come risultato.

### 3.4 Relatorios OK

Riferimenti:

- `src/fantozzi/relatorios_gerenciais_ok.py:103` genera i relatorios OK.
- `src/fantozzi/relatorios_gerenciais_ok.py:298` legge gli XLSX.
- `src/fantozzi/relatorios_gerenciais_ok.py:401` contiene il gate BONIFICACAO_RECEBIDA.

Osservazione:

I manuali devono arrivare con `status_revisao = OK`, senza `tipo_operacao_especial`, e con `origem = MANUAL_GERENCIAL`. In questo modo entrano nella base OK senza interferire con COMODATO/BONIFICACAO.

### 3.5 DRE

Riferimenti:

- `src/fantozzi/dre_gerencial_provisoria.py:94` genera la DRE.
- `src/fantozzi/dre_gerencial_provisoria.py:233` legge `mapeamento_dre.yaml`.
- `src/fantozzi/dre_gerencial_provisoria.py:251` adatta una riga alla base DRE.
- `src/fantozzi/dre_gerencial_provisoria.py:307` raggruppa i valori DRE.

Osservazione:

La DRE considera contas mappate in `mapeamento_dre.yaml`. Quindi, per 9N v1, non serve inventare logica DRE nuova: le righe manuali entrano nella DRE solo se la conta tecnica della riga e mappata e `entra_dre = true`.

Attenzione:

Oggi la DRE usa `abs(valor)` e poi applica il segno della linha DRE. Per evitare doppia interpretazione del segno, il CSV 9N deve accettare solo `valor` positivo. L'effetto gestionale viene determinato dalla conta e dal mapeamento, non da valori negativi nel CSV.

Guardrail DRE approvato per v1:

- 0 contas della partita doppia mappate DRE con `entra_dre = true`: il lancamento entra solo in pre-livro/audit, non in DRE.
- 1 conta mappata DRE con `entra_dre = true`: entra una sola volta nel DRE.
- 2 contas mappate DRE con `entra_dre = true`: bloccare il CSV con errore chiaro. La v1 non tenta compensazioni automatiche.

---

## 4. Disegno proposto

### 4.1 Nuova flag CLI

Aggiungere a `reprocessar-importacao-real`:

```text
--lancamentos-manuais-gerenciais-csv <path>
```

Comportamento:

- senza flag: pipeline invariata;
- con flag: valida il CSV, genera artefatti manuali e integra i manuali nella run;
- flag ripetibile solo se serve in futuro; per la v1 e sufficiente un singolo CSV.

La flag deve essere distinta da `--fase9-csv` per non confondere 9N con la Fase 9 di classificazione.

### 4.2 Nuovo modulo

Creare:

```text
src/fantozzi/lancamentos_manuais_gerenciais.py
tests/test_lancamentos_manuais_gerenciais.py
```

Responsabilita del modulo:

- leggere CSV `;` con `utf-8-sig`;
- validare schema e righe;
- normalizzare data, texto e valor;
- generare chave deduplicacao stabile;
- convertire ogni riga manuale in due `Lancamento`;
- scrivere `lancamentos_manuais_gerenciais_aplicados.csv`;
- scrivere `auditoria_lancamentos_manuais_gerenciais.txt`;
- scrivere un XLSX tecnico compatibile con `HEADERS`.

### 4.3 Posizione degli output

Input reale: mai committato.

Output tecnico dentro la run:

```text
output/importacao_real/runs/<run_id>/MANUAL_GERENCIAL/lancamentos_manuais_gerenciais.xlsx
```

Audit dentro reports:

```text
reports/importacao_real/runs/<run_id>/lancamentos_manuais_gerenciais_aplicados.csv
reports/importacao_real/runs/<run_id>/auditoria_lancamentos_manuais_gerenciais.txt
```

Motivo:

I report gia leggono ricorsivamente gli XLSX della run. Scrivere un XLSX tecnico sotto `output_run_dir` permette a pre-livro, relatorios OK e DRE di includere i manuali senza creare scorciatoie parallele.

### 4.4 Schema CSV v1

Separatore: `;`.

Colonne obbligatorie:

```text
data
historico
documento
fornecedor_ou_contraparte
conta_debito_id
conta_debito_nome
conta_credito_id
conta_credito_nome
valor
```

Colonne opzionali:

```text
cnpj_origem_opcional
centro_custo_opcional
observacao
```

Derivati dal sistema:

```text
fonte = MANUAL_GERENCIAL
status = OK
tipo_lancamento = MANUAL_GERENCIAL
chave_deduplicacao
linha_origem_csv
```

### 4.5 Validazioni bloccanti

Fallire l'intero processo se:

- CSV non esiste o non e file;
- colonna obbligatoria assente;
- colonna duplicata;
- CSV vuoto;
- riga completamente vuota in mezzo a dati;
- `data` vuota o non `YYYY-MM-DD`;
- `historico` vuoto;
- `valor` vuoto, zero, negativo o non parsabile;
- `valor` con segno negativo o zero anche se formalmente parsabile;
- `conta_debito_id` vuota;
- `conta_credito_id` vuota;
- `conta_debito_id == conta_credito_id`;
- `conta_debito_nome` o `conta_credito_nome` vuoto;
- conta id non presente nel plano de contas;
- nome conta incompatibile con plano de contas;
- entrambe le contas della riga mappate in DRE con `entra_dre = true`;
- duplicata nel CSV stesso.

### 4.6 Deduplicazione v1

Chave suggerita:

```text
data
documento
historico normalizado
conta_debito_id
conta_credito_id
valor
fornecedor_ou_contraparte normalizado
```

Se due righe hanno la stessa chave, il CSV fallisce.

Non implementare ancora deduplicazione contro XML, banco o comprovantes.

### 4.7 Conversione partita doppia

Per ogni riga CSV accettata:

Riga tecnica debito:

- `conta_id = conta_debito_id`;
- `conta_nome = conta_debito_nome`;
- `valor = valor positivo`;
- `origem = MANUAL_GERENCIAL`;
- `status_revisao = OK`;
- `tipo_lancamento = MANUAL_GERENCIAL_DEBITO`;
- `chave_origem = chave_deduplicacao + ":D"`;
- `observacoes` include documento, contrapartida credito e osservazione utente.

Riga tecnica credito:

- `conta_id = conta_credito_id`;
- `conta_nome = conta_credito_nome`;
- `valor = valor positivo`;
- `origem = MANUAL_GERENCIAL`;
- `status_revisao = OK`;
- `tipo_lancamento = MANUAL_GERENCIAL_CREDITO`;
- `chave_origem = chave_deduplicacao + ":C"`;
- `observacoes` include documento, contrapartida debito e osservazione utente.

Il CSV applicato deve mantenere una riga per lancamento manuale originale, non due, per audit umano piu leggibile. L'XLSX tecnico contiene invece le due righe necessarie alla composizione contabile.

### 4.8 CNPJ e consolidamento

`cnpj_origem_opcional` e solo audit/origine. Non separa gestionalmente l'impresa.

Se valorizzato, popolare `cnpj_empresa`; se vuoto, usare:

```text
CONSOLIDADO_GERENCIAL
```

Il path dell'XLSX tecnico puo rimanere sotto `MANUAL_GERENCIAL`, cosi i report identificano chiaramente la fonte manuale invece di fingere che sia XML di uno dei due CNPJ.

---

## 5. Integrazione pipeline

### 5.1 Context

Aggiungere a `PipelineContext`:

```python
lancamentos_manuais_gerenciais_csv_path: Path | None = None
```

oppure una tupla se si decide di supportare piu file in v1.

### 5.2 Step

Nuovo checkpoint suggerito:

```text
LANCAMENTOS_MANUAIS_GERENCIAIS_APLICADOS
```

Posizione:

```text
IMPORTACAO_XML_CONCLUIDA
DIAGNOSTICOS_CONCLUIDOS
OPERACOES_ESPECIAIS_AUDITADAS
LANCAMENTOS_MANUAIS_GERENCIAIS_APLICADOS
PRE_LIVRO_GERADO
AUDITORIA_PRE_LIVRO_CONCLUIDA
PACOTE_REVISAO_CONTABIL_GERADO
...
```

Lo step deve:

1. validare il CSV;
2. scrivere XLSX tecnico in `output_run_dir/MANUAL_GERENCIAL`;
3. aggiungere l'XLSX a `context.lancamentos_xlsx`;
4. scrivere CSV/TXT audit in `reports_run_dir`;
5. aggiornare manifest.

Se la flag non e presente, non eseguire lo step oppure eseguirlo come no-op solo se serve per registrare `"executado": false`.

### 5.3 Manifest

Senza flag:

```json
"lancamentos_manuais_gerenciais": {
  "executado": false
}
```

Con flag:

```json
"lancamentos_manuais_gerenciais": {
  "executado": true,
  "input_csv": "...",
  "output_xlsx_tecnico": "...",
  "output_csv": "...",
  "auditoria_txt": "...",
  "linhas_lidas": 3,
  "lancamentos_aceitos": 3,
  "linhas_tecnicas_geradas": 6,
  "total_manual": "1234.56",
  "fonte": "MANUAL_GERENCIAL"
}
```

Nota tecnica:

`StepResult` oggi aggiorna solo `totais`, `classificacao`, `avisos` e `relatorios_gerados`. Per 9N serve una piccola estensione conservativa, per esempio:

```python
manifest_updates: Mapping[str, object] = field(default_factory=dict)
```

Poi `_atualizar_manifesto_step` fonde questi campi top-level nel manifest.

### 5.4 Dry-run

Per `--dry-run`, non leggere ne validare il CSV manuale in v1. Il dry-run attuale calcola solo run e tappe senza creare file. Deve restare leggero e senza effetti collaterali.

Questo comportamento deve essere documentato in test o codice: se l'utente passa il CSV insieme a `--dry-run`, la pipeline mostra solo la previsione della run e non applica/valida i manuali.

---

## 6. Rischi e mitigazioni

| Rischio | Mitigazione |
|---|---|
| Doppio impatto DRE per debito e credito entrambi mappati | Test specifico con una conta resultado e una patrimonial; documentare che v1 richiede contas coerenti. |
| Manuali finiscono nella revisione classificazione prodotti | Inserire lo step dopo diagnostici e prima di pre-livro; cosi non alimenta candidati prodotto. |
| Confusione con Fase 9 | Flag separata, manifest separato, non toccare `fase9_executada`. |
| Duplicati nel CSV | Chave dedupe obbligatoria e blocco dell'intero CSV. |
| Input reale committato | Nessun path implicito; solo fixture sintetiche in `tests/`. |
| Nome conta incoerente | Validazione contro `ChartOfAccounts`. |
| Segni contabili ambigui | `valor` sempre positivo; effetto DRE deciso dal mapeamento. |

---

## 7. Test minimi

Unitari nuovo modulo:

- legge CSV valido `;` con decimal brasileiro;
- legge decimal con punto;
- rifiuta colonna obbligatoria assente;
- rifiuta colonna duplicata;
- rifiuta data invalida;
- rifiuta valor zero/negativo;
- rifiuta debito uguale credito;
- rifiuta conta inesistente;
- rifiuta nome conta incompatibile;
- rifiuta duplicata interna;
- preserva `linha_origem_csv`;
- genera `chave_deduplicacao` stabile;
- converte una riga CSV in due righe `Lancamento`.

Pipeline/CLI:

- parser accetta `--lancamentos-manuais-gerenciais-csv`;
- senza flag, manifest e outputs restano invariati salvo eventuale sezione `executado: false`;
- con CSV valido, manifest registra `executado: true`;
- con CSV valido, genera XLSX tecnico, CSV applicato e audit TXT;
- con CSV valido, pre-livro include le righe manuali;
- con CSV valido, DRE include solo manuali con conta mappata;
- con CSV invalido, pipeline fallisce con errore chiaro e manifest `ERRO`;
- COMODATO resta fuori DRE;
- BONIFICACAO_RECEBIDA continua in DRE come `2.05.003`.

Comandi di test previsti:

```powershell
.\.venv\Scripts\python.exe -m pytest tests\test_lancamentos_manuais_gerenciais.py
.\.venv\Scripts\python.exe -m pytest tests\test_reprocessamento_pipeline.py tests\test_cli.py
.\.venv\Scripts\python.exe -m pytest
```

---

## 8. Criteri di accettazione

La Fase 9N e pronta quando:

- senza flag la pipeline mantiene il comportamento precedente;
- con flag CSV valido, i manuali entrano come `MANUAL_GERENCIAL`;
- DRE e pre-livro sono rigenerati, non editati a mano;
- audit CSV/TXT dei manuali e creato;
- manifest registra la sezione 9N;
- duplicati e dati invalidi bloccano il processo;
- `fase9_executada` resta indipendente;
- COMODATO e BONIFICACAO mantengono il trattamento deciso;
- full suite passa;
- nessun output reale o dato reale viene committato.

---

## 9. Sequenza consigliata per implementazione

1. Creare `lancamentos_manuais_gerenciais.py` con parser/validator puro e test unitari.
2. Aggiungere writer XLSX tecnico riusando `write_lancamentos_to_excel`.
3. Aggiungere audit CSV/TXT dei manuali.
4. Estendere `StepResult` con `manifest_updates`.
5. Aggiungere flag CLI e campo nel `PipelineContext`.
6. Inserire lo step 9N dopo audit operazioni speciali e prima del pre-livro.
7. Aggiungere test CLI/pipeline.
8. Eseguire test mirati e full suite.
9. Solo dopo test, valutare run reale senza CSV e run con CSV locale di prova non committato.
