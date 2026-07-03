# Workflow Operatore Finale — Fantozzi

Versione: 1.0 — 2026-05-24
Progetto: Programma_registrazioni_contabili

**NOTA OPERATIVA 2026-07-03:** il branch operativo corrente è
`senza-incrementali` (non `master`). Per le registrazioni manuali gerenciais
(Step 8) la fonte di verità aggiornata è
`docs/LANCAMENTOS_MANUAIS_GERENCIAIS_OPERADOR.md`: lo Step 8 di questo
documento è stato riallineato di conseguenza, ma in caso di dubbio prevale
sempre quel documento.

---

## Cosa produce questo workflow

Alla fine del workflow avrai:

```
DRE gerencial                        conto economico gestionale
Pre-livro / Livro Diario gerencial   giornale gestionale
Lancamentos_OK                       movimenti classificati OK
Lancamentos_REVISAR                  movimenti da revisionare (separati)
Pacote contabil gerencial            cartella con tutti i file finali + audit
```

Non e contabilita fiscale ufficiale. E un sistema di contabilita gestionale
per analisi interna e confronto con il contador.

---

## Cosa entra nel DRE

Entrano nel DRE:

1. **XML acquisti** — classificati automaticamente dalle regole prodotti
2. **Lancamentos manuais gerenciais** — spese inserite dall'operatore tramite template
3. **Spese in contanti** — inserite tramite template manuali
4. **Comprovantes bancari con regola approvata e sicura** — entrano nel DRE attraverso il flusso banca/comprovantes (Step 5–7)
5. **Comprovantes/spese bancarie senza regola o eccezionali** — solo se inseriti manualmente come lancamentos manuais gerenciais (Step 8)

Non entrano nel DRE:

- Tutto l'estratto conto bancario in automatico
- Tutti i comprovantes non revisionati
- Righe con status REVISAR
- COMODATO (beni in comodato d'uso)
- Contropartida tecnica 99.99.998 (e solo bilanciamento tecnico)
- Entrate bancarie (ricavi/vendite vengono da LeChef, non dalla banca)
- Duplicati di spese gia presenti negli XML

---

## Regole contabili consolidate

| Voce | Trattamento |
|---|---|
| COMODATO | Fuori DRE — eventuale report beni di terzi separato |
| BONIFICACAO_RECEBIDA FORTE | Entra nel DRE come riduzione costi — conto 2.05.003 |
| BONIFICACAO_RECEBIDA non FORTE | Non creare regola automatica — lasciare BLOQUEADO/REVISAR |
| INSS nei comprovantes | ESCLUDERE dal flusso comprovantes/DRE |
| Entrate bancarie | Fuori scope — NON usare come ricavi |
| Manuali/contanti | Entrano nel DRE tramite template lancamentos manuais |
| 99.99.998 Contrapartida Tecnica | Non entra nel DRE — e solo bilanciamento tecnico del pre-livro |

---

## Prima di iniziare

### Checklist operatore

- [ ] XML acquisti presenti in `Database_Dev\XML_fatture_acquisto\`
- [ ] `bank_raw\` presente se si usano comprovantes bancari
- [ ] Nessun file Excel aperto durante i comandi PowerShell
- [ ] Backup locale dei dati reali effettuato

### Checklist tecnica

Aprire PowerShell nella cartella del progetto ed eseguire:

```powershell
git status -sb
git rev-list --left-right --count origin/senza-incrementali...HEAD
git diff --name-status
```

**Semafori:**

| Colore | Condizione | Azione |
|---|---|---|
| VERDE | Nessun diff tracciato inatteso, origin/senza-incrementali...HEAD = 0 0 | Procedere |
| GIALLO | Solo output locali ignorati (es. graphify-out/) | Procedere con attenzione |
| ROSSO | File src/ o config/ modificati, o branch non allineato | **FERMARSI — supporto tecnico** |

---

## Step 1 — Generare run provvisoria XML

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli reprocessar-importacao-real `
  --xml-base Database_Dev\XML_fatture_acquisto `
  --saida-output output\importacao_real `
  --saida-reports reports\importacao_real `
  --modo provisorio
```

**Output generato in:**

```
reports\importacao_real\runs\<RUN_ID>\
```

**Come trovare il RUN_ID:** e la cartella piu recente in `reports\importacao_real\runs\`
oppure leggere l'output del comando (es. `2026-05-24_003`).

**Controlli dopo il comando:**

- `manifest_run.json` presente nella cartella della run
- `status_run: PROVISORIO_CONCLUIDO`
- `fase9_executada: false`
- CSV di revisione per CNPJ presenti

Se il comando termina con errore: **fermarsi e chiedere supporto tecnico.**

---

## Step 2 — Generare Excel revisione XML

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli importar-revisao-xml --run-id <RUN_ID>
```

Sostituire `<RUN_ID>` con il valore trovato nello Step 1 (es. `2026-05-24_003`).

**Output:**

```
Database_Dev\_decisoes_regras\xml\Analisi_regole_registrazioni_xml_prodotti.xlsx
Database_Dev\_decisoes_regras\xml\manifest_importa_revisao_xml.json
```

**Regola:** Se appare un messaggio di BLOCCO, fermarsi.
Non usare `--allow-stale-run` senza supporto tecnico.

---

## Step 3 — Revisionare Excel XML

Aprire il file:

```
Database_Dev\_decisoes_regras\xml\Analisi_regole_registrazioni_xml_prodotti.xlsx
```

Compilare la colonna `STATUS_DECISAO` solo per le righe sicure.

**Valori ammessi:**

| Valore | Significato |
|---|---|
| `CRIAR_ALIAS` | Creare alias per questo prodotto |
| `CRIAR_REGRA` | Creare regola per questo fornitore |
| `CLASSIFICAR` | Classificare direttamente |
| `BLOQUEADO` | Lasciare bloccato — non applicabile |
| `IGNORAR` | Ignorare questa riga |
| `REVISAR` | Da rivedere piu tardi |

**Regole obbligatorie:**

- Le righe `BLOQUEADO` possono restare solo `BLOQUEADO`, `IGNORAR` o `REVISAR`
- Non trasformare COMODATO in un costo normale
- Non trasformare BONIFICACAO non FORTE in regola automatica
- Se nessuna riga ha valore approvabile: **saltare Step 4 direttamente**

---

## Step 4 — Applicare decisioni XML (solo se esistono)

**Questo step richiede Claude Code o supporto tecnico. Non e autonomo.**

**Se il numero di righe APROVADO/CRIAR_ALIAS/CRIAR_REGRA/CLASSIFICAR = 0:**

```
Non fare nulla. Passare direttamente allo Step 5.
```

**Se ci sono decisioni approvate:**

1. Comunicarlo al tecnico / aprire Claude Code
2. Usare la skill `regole-registrazioni-xml`
3. Auditare il diff YAML prima di applicare
4. Committare solo dopo audit approvato

---

## Step 5 — Generare Excel revisione banca/comprovantes

Indicare la cartella del mese da analizzare:

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli importar-revisao-banca `
  --bank-raw-dir bank_raw\2026\2026-04
```

Oppure con anno e mese separati:

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli importar-revisao-banca `
  --ano 2026 `
  --mes 04
```

**Output:**

```
bank_raw\_decisoes_regras\Analisi_regole_registrazioni_banca_comprovanti.xlsx
bank_raw\_decisoes_regras\Analisi_regole_registrazioni_banca_conto_corrente.xlsx
bank_raw\_decisoes_regras\manifest_importa_revisao_banca.json
```

**Regole operative:**

- Entrate bancarie: fuori scope, non usare come ricavi
- INSS: segnare ESCLUDERE se gia deciso in precedenza
- `RISCO_DUPLICIDADE_XML = DUVIDA`: non approvare come regola automatica

---

## Step 6 — Revisionare Excel banca/comprovantes

Aprire entrambi gli Excel:

```
Analisi_regole_registrazioni_banca_comprovanti.xlsx
Analisi_regole_registrazioni_banca_conto_corrente.xlsx
```

Approvare solo le righe sicure. Lasciare `PENDENTE` o `NAO_CRIAR_REGRA` le righe dubbie.

**Regola importante:**

I comprovantes non entrano nel DRE solo perche presenti nei file banca/comprovantes.

Un comprovante entra nel DRE se ha una regola approvata e sicura che lo classifica come spesa DRE,
oppure se viene inserito come lancamento manual gerencial (Step 8).

Il template manuali non sostituisce le regole automatiche: serve per spese senza regola,
casi eccezionali, contanti o comprovantes approvati manualmente.

Un comprovante **non deve entrare nel DRE** se:

- e il pagamento di una fattura XML gia registrata;
- ha `RISCO_DUPLICIDADE_XML = DUVIDA`;
- resta REVISAR/PENDENTE/NAO_CRIAR_REGRA.

---

## Step 7 — Applicare decisioni banca (solo se esistono)

**Questo step richiede Claude Code o supporto tecnico. Non e autonomo.**

**Se il numero di righe APROVADO/CLASSIFICAR = 0:**

```
Non fare nulla. Passare direttamente allo Step 8.
```

**Se ci sono decisioni approvate:**

1. Comunicarlo al tecnico / aprire Claude Code
2. Usare la skill `regole-registrazioni-banca`
3. Auditare il diff YAML prima di applicare
4. Committare solo dopo audit approvato

---

## Step 8 — Lancamentos manuais gerenciais

Usare questo step per:

- Spese senza XML (affitto, servizi, prestatori, spese varie)
- Spese in contanti
- Comprovantes bancari senza regola automatica sicura, approvati manualmente per il DRE
- Correzioni gestionali approvate

**Non usare per:**

- Spese gia presenti negli XML acquisti
- Pagamenti di fatture XML gia registrate
- Movimenti bancari raw non revisionati
- Righe dubbie o non confermate

Se non ci sono spese manuali: **saltare al Step 10.**

### 8a — Generare il template

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli gerar-template-lancamentos-manuais-gerenciais
```

Il comando accetta opzionalmente `--saida` per un path diverso, ma il default
corretto (confermato nel codice, `cli.py::DEFAULT_REGISTRAZIONI_MANUALI_BASE`)
è già:

```
Database_Dev\Registrazioni_manuali\registrazioni_manuali_base.xlsx
```

**Output atteso:**

```
[OK] Template gerado: ...
[INFO] Categorias disponiveis: 9
[INFO] Contrapartida tecnica: 99.99.998 - Contrapartida Tecnica Lancamentos Manuais
Exit code: 0
```

Se vedi `[BLOQUEIO]` e il comando termina con `Exit code: 1`: **fermarsi — supporto tecnico.**

### 8b — Compilare il template

Aprire il file generato e compilare **solo queste colonne visibili** nell'aba `lancamentos_manuais`:

| Colonna | Descrizione | Obbligatoria |
|---|---|---|
| `DATA` | Data della spesa in formato YYYY-MM-DD (es. 2026-04-15) | Si |
| `CONTA_ESCOLHIDA` | Conto selezionato dal dropdown `ID - NOME` (piano dei conti gestionale) | Si |
| `VALOR` | Valore positivo (es. 150,00) | Si |
| `DESCRICAO_DA_DESPESA` | Descrizione concreta di cosa è stato comprato/contrattato (es. "Conserto pia cozinha") | Si |
| `OBSERVACAO` | Note, payment_id, riferimento comprovante, arquivo_origem | No |

Le colonne `CONTA_ID_APROVADA` e `CONTA_NOME_APROVADA` sono calcolate
automaticamente via VLOOKUP a partire da `CONTA_ESCOLHIDA` — **non modificarle
a mano.** Restano inoltre nascoste e calcolate dal sistema: `CONTA_CREDITO_ID`,
`CONTA_CREDITO_NOME` (sempre `99.99.998`), `STATUS_VALIDACAO`,
`MENSAGEM_VALIDACAO`, `CHAVE_PREVIEW`.

**Colonne obsolete — non compaiono più nel template e non vanno cercate o
reintrodotte:** `HISTORICO`, `DOCUMENTO`, `FORNECEDOR_OU_CONTRAPARTE`,
`CATEGORIA_GERENCIAL`, `CNPJ_ORIGEM`, `CENTRO_CUSTO`, `FORMA_PAGAMENTO`,
`BANCO`, `CAIXA`, `CARTAO`. Il template descrive solo **cosa** è stato
comprato/contrattato e il valore, non come è stato pagato o da quale conto
patrimoniale.

**Convenzione per spese in contanti o comprovante bancario approvato** — usare
il campo libero `OBSERVACAO`, ad esempio:

```
OBSERVACAO = pagamento em dinheiro (contanti)
```

oppure

```
OBSERVACAO = payment_id: <id> / arquivo_origem: Analisi_regole_...xlsx / pagamento verificado
```

**Non lasciare righe vuote nel mezzo dei dati.**

---

## Step 9 — Validare manuali e generare CSV tecnico

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli validar-lancamentos-manuais-gerenciais `
  --entrada Database_Dev\Registrazioni_manuali\registrazioni_manuali_base.xlsx `
  --saida-csv Database_Dev\Registrazioni_manuali\registrazioni_manuali_tecnico.csv `
  --relatorio Database_Dev\Registrazioni_manuali\validacao_lancamentos_manuais.txt
```

**Output atteso:**

```
[OK] CSV tecnico gerado: ...
[INFO] Linhas lidas: N
[INFO] Linhas validas: N
```

**Se il comando termina con `[ERRO]`:**

- Aprire il file di relatorio indicato
- Correggere gli errori nel template Excel
- Rilanciare la validazione
- Non procedere con un CSV errato

**Avisos** (non bloccanti): leggere e verificare, ma il CSV viene generato.
Verificare in particolare l'avviso letterale: *"ATENCAO: o sistema nao
garante que esta despesa ja nao exista em XML."* — la responsabilità di non
duplicare resta dell'operatore (dettagli in
`docs/LANCAMENTOS_MANUAIS_GERENCIAIS_OPERADOR.md`, sezione "O que o sistema
NAO garante automaticamente").

**Alternativa raccomandata — comando one-shot:** invece di eseguire Step 8a +
9 + 10 separatamente, è disponibile un comando unico che valida il template,
genera il CSV tecnico ed esegue la run completa con pacote in un solo passo:

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli executar-lancamentos-manuais-gerenciais
```

La pipeline sottostante di questo comando
(`validar_lancamentos_manuais_gerenciais_operador()` →
`reprocessar_importacao_real_provisoria()`) è coperta dal test e2e
`tests/test_lancamentos_manuais_e2e.py`. Non è invece coperto da un test e2e
diretto l'entry point argparse del comando one-shot in sé (gap noto e non
bloccante, non un'istruzione di non-uso — vedi
`docs/LANCAMENTOS_MANUAIS_GERENCIAIS_OPERADOR.md`).

---

## Step 10 — Run finale multi-fonte con pacote

**Con spese manuali (CSV validato presente):**

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli reprocessar-importacao-real `
  --xml-base Database_Dev\XML_fatture_acquisto `
  --saida-output output\importacao_real `
  --saida-reports reports\importacao_real `
  --modo provisorio `
  --lancamentos-manuais-gerenciais-csv Database_Dev\Registrazioni_manuali\registrazioni_manuali_tecnico.csv `
  --gerar-pacote-contabil-gerencial
```

**Senza spese manuali:**

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli reprocessar-importacao-real `
  --xml-base Database_Dev\XML_fatture_acquisto `
  --saida-output output\importacao_real `
  --saida-reports reports\importacao_real `
  --modo provisorio `
  --gerar-pacote-contabil-gerencial
```

**Nota:** il flag `--gerar-pacote-contabil-gerencial` e obbligatorio per ottenere il pacote finale con tutti i file organizzati.

**Output generato in:**

```
reports\importacao_real\runs\<RUN_ID_FINAL>\pacote_contabil_gerencial_provisorio\
```

**Controlli dopo il comando:**

- `MANUAL_GERENCIAL\lancamentos_manuais_gerenciais.xlsx` presente (se CSV manuali usato)
- `06_AUDITORIA_LANCAMENTOS_MANUAIS_GERENCIAIS.txt` presente nel pacote
- DRE generato
- Pre-livro generato
- `Lancamentos_REVISAR` separati

---

## Step 11 — Controlli finali

Verificare nel pacote:

- [ ] COMODATO fuori DRE (presente in `04_AUDITORIA_COMODATO_FORA_DRE.txt`)
- [ ] BONIFICACAO_RECEBIDA FORTE in conto 2.05.003 (audit `05_AUDITORIA_BONIFICACAO_RECEBIDA_DRE.txt`)
- [ ] BONIFICACAO non FORTE non trasformata in regola
- [ ] REVISAR fuori DRE (presenti in foglio separato nel pre-livro)
- [ ] Manuali inclusi se CSV passato (audit `06_AUDITORIA_LANCAMENTOS_MANUAIS_GERENCIAIS.txt`)
- [ ] Contrapartida 99.99.998 non appare nelle voci DRE
- [ ] Comprovantes con regola sicura inclusi correttamente se Step 7 applicato
- [ ] Comprovantes che pagano fatture XML gia registrate non entrano nel DRE
- [ ] Spese manuali non duplicano spese XML gia registrate
- [ ] Entrate bancarie non usate come ricavi
- [ ] Git pulito (`git status -sb`)

---

## Step 12 — Pacote finale: dove trovare i file

Il pacote si trova in:

```
reports\importacao_real\runs\<RUN_ID_FINAL>\pacote_contabil_gerencial_provisorio\
```

**Contenuto del pacote:**

| File | Contenuto |
|---|---|
| `00_INDICE_PACOTE_CONTABIL_GERENCIAL.txt` | Indice e riepilogo del pacote |
| `01_DRE_GERENCIAL_CONSOLIDADO.xlsx` | DRE gerencial finale |
| `02_PRE_LIVRO_DIARIO_GERENCIAL.xlsx` | Pre-livro / Livro Diario gerencial |
| `03_RELATORIOS_GERENCIAIS_OK.xlsx` | Report lancamentos OK per CNPJ e mese |
| `04_AUDITORIA_COMODATO_FORA_DRE.txt` | Audit beni in comodato esclusi dal DRE |
| `05_AUDITORIA_BONIFICACAO_RECEBIDA_DRE.txt` | Audit bonificacoes nel DRE |
| `06_AUDITORIA_LANCAMENTOS_MANUAIS_GERENCIAIS.txt` | Audit spese manuali inserite |
| `manifest_pacote_contabil_gerencial.json` | Manifest tecnico del pacote |

**Da inviare al contador:**

- `01_DRE_GERENCIAL_CONSOLIDADO.xlsx`
- `02_PRE_LIVRO_DIARIO_GERENCIAL.xlsx`
- Foglio `Lancamentos_REVISAR` dal pre-livro
- File audit 04, 05, 06

**Per uso interno:**

- Tutti i file del pacote + il manifest

**Tutti i file restano locali.** Non vengono committati in Git.

---

## Semafori e punti di stop

### VERDE — Procedere

- manifest_run.json presente
- Template manuali generato con exit code 0
- Validazione manuali: 0 errori
- Run finale: `PROVISORIO_CONCLUIDO`
- Pacote generato
- Git pulito

### GIALLO — Procedere con attenzione

- Righe REVISAR presenti (e normale)
- Avisos non bloccanti nel report manuali
- Manuali senza CNPJ
- Nessuna nuova decisione applicata

### ROSSO — Fermarsi immediatamente

- Hash YAML stale
- Messaggio BLOCCO in qualsiasi comando
- CSV manuali con errori (exit code 1 su validazione)
- Contrapartida tecnica 99.99.998 assente dal piano conti
- DRE include 99.99.998
- COMODATO entra nel DRE
- Entrate bancarie diventano ricavi
- `git diff` mostra modifiche a `src/` o `config/` non attese

---

## Errori comuni

| Errore | Causa | Soluzione |
|---|---|---|
| Comando bloccato / nessun output | Excel con lo stesso file aperto | Chiudere Excel e riprovare |
| `RUN_ID non trovato` | RUN_ID errato o run non eseguita | Verificare cartella `reports\importacao_real\runs\` |
| `CONTA_ESCOLHIDA inexistente` | Conto scritto a mano diverso dalla lista, o non presente nel piano conti | Scegliere dalla lista a discesa nel template (colonna `CONTA_ESCOLHIDA`) |
| `VALOR deve ser maior que zero` | Valore zero o negativo | Inserire valore positivo |
| `DATA invalida` | Data non in formato YYYY-MM-DD | Correggere in formato 2026-04-15 |
| `Duplicata interna` | Stessa riga inserita due volte nel template | Rimuovere la riga duplicata |
| `Linha vazia no meio dos dados` | Riga vuota tra dati nel template | Rimuovere la riga vuota |
| Pacote non generato | Manca `--gerar-pacote-contabil-gerencial` | Aggiungere il flag al comando |
| `[BLOQUEIO] Contrapartida tecnica ausente` | 99.99.998 assente dal piano conti | Supporto tecnico — non procedere |

---

## Cosa non deve fare l'operatore

```
Non modificare i file YAML in config/
Non modificare il piano dei conti a mano
Non fare git add, git commit, git push
Non usare --allow-stale-run senza supporto tecnico
Non forzare righe REVISAR per azzerare le pendenze
Non inventare categorie non presenti nella lista a discesa
Non inserire spese gia negli XML
Non usare l'estratto bancario per registrare ricavi
Non compilare le colonne tecniche del template (CONTA_*, STATUS_*)
Non committare file da Database_Dev/, reports/, output/, bank_raw/
```

---

## Glossario

| Termine | Significato operativo |
|---|---|
| **DRE gerencial** | Conto economico gestionale — mostra ricavi, costi e risultato |
| **Pre-livro diario gerencial** | Bozza del giornale contabile gestionale — file `pre_livro_diario_gerencial_teste.xlsx` |
| **Livro Diario gerencial** | Nome ufficiale del giornale — stesso file del pre-livro |
| **Pacote contabil gerencial** | Cartella con tutti i file finali organizzati — `pacote_contabil_gerencial_provisorio/` |
| **Lancamento manual gerencial** | Spesa inserita manualmente dall'operatore tramite template |
| **Contrapartida tecnica** | Conto 99.99.998 usato solo per bilanciamento tecnico — non e un conto reale |
| **REVISAR** | Movimento non classificato che resta fuori dal DRE fino a revisione |
| **BLOQUEADO** | Riga che il sistema non puo classificare automaticamente |
| **APROVADO** | Decisione approvata dall'operatore nel foglio di revisione |
| **FORA_DRE** | Movimento che non entra nel DRE (comodato, contrapartida tecnica, ecc.) |
| **RUN_ID** | Identificatore univoco di ogni elaborazione (es. 2026-05-24_003) |

**Nota sui suffissi tecnici:**

- `_teste` in `pre_livro_diario_gerencial_teste.xlsx` e nomenclatura storica, non indica un file di prova
- `_provisorio` in `pacote_contabil_gerencial_provisorio` indica che la Fase 9 non e stata eseguita, non che i dati siano inaffidabili
- `ok_provisoria` in `dre_gerencial_ok_provisoria.xlsx` indica che il DRE e basato sui soli movimenti OK, esclusi i REVISAR
