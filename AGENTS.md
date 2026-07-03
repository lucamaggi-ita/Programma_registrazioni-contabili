# AGENTS.md — Programma_registrazioni_contabili (Fantozzi)

Creato: 2026-07-03, dopo ricognizione tecnica completa read-only del progetto.
Aggiornato: 2026-07-03, con regole generali di qualità tecnica, audit e chiusura fase.

Questo file è la versione tool-agnostica delle istruzioni operative. Per il
contesto completo (identità, workflow, skill, filosofia del progetto) fare
riferimento a `CLAUDE.md` nella root del repository, che resta la fonte
principale. Qui sono raccolte solo le regole operative essenziali, valide per
qualunque agente AI che lavori su questo repository.

---

## Identità e principio fondamentale

Le decisioni contabili appartengono esclusivamente all'utente. Un agente che
lavora qui:

- non inferisce conti;
- non crea regole di classificazione autonomamente;
- non modifica YAML/config senza conferma esplicita;
- non fa commit/push senza ordine esplicito;
- se manca una regola, il movimento resta in revisione;
- non introduce fallback silenziosi;
- non spaccia correzioni manuali o ipotesi come decisioni definitive.

Regola madre:

```text
Prima contrattualizza, poi programma.
```

Ogni decisione funzionale deve diventare un contratto verificabile in almeno uno
di questi luoghi:

- codice;
- YAML/config;
- test;
- documentazione operatore;
- manifest/provenance;
- audit/report;
- output etichettato e non ambiguo.

Nessuna decisione importante deve restare solo in chat.

---

## Dati mancanti, stati e fonti

Per XML, banca, comprovanti, registrazioni manuali, pré-livro, DRE, manifest e
output:

- un dato mancante non deve mai diventare automaticamente un valore contabile valido;
- se un dato manca, usare uno stato esplicito come `MANCANTE`, `NON_VERIFICATO`,
  `NON_DISPONIBILE`, `N.A.` o equivalente già previsto dal progetto;
- ogni dato importante deve avere, quando applicabile:
  - valore;
  - fonte;
  - stato;
  - copertura;
  - motivo;
  - eventuale regola/YAML/test che lo giustifica.

Sono vietati:

- fallback silenziosi;
- hardcode nascosti per regole stabili;
- inferenze contabili automatiche non approvate;
- correzioni manuali trattate come definitive senza documentazione.

---

## Precheck obbligatorio

Prima di qualunque lavoro non banale su codice, dati, YAML, output,
documentazione, pipeline, manifest, DRE, banca, comprovanti o registrazioni
manuali:

```powershell
git status -sb
git branch --show-current
git rev-list --left-right --count origin/senza-incrementali...HEAD
git diff --name-status
git diff --cached --name-status
```

Condizione attesa ordinaria:

- branch operativo `senza-incrementali`;
- allineamento `0 0` con `origin/senza-incrementali`, salvo mandato esplicito diverso;
- nessun diff tracked;
- nessun diff staged;
- solo untracked già censiti o dichiarati.

Mai lavorare sopra uno stato misto staged/unstaged senza prima classificarlo
esplicitamente.

Se ci sono file tracked modificati, staged o unstaged, produrre una tabella:

| File | Origine/mandato | Stato | Destino |
|---|---|---|---|
| path | patch/MD/mandato | pronto/parziale/dubbio | commit mirato / patch nominata / scarto autorizzato |

Nessun file tracked modificato può restare “lasciato lì” senza destino esplicito.

---

## Branch

- Branch operativo attuale: **senza-incrementali**.
- Verificare sempre con `git branch --show-current`, non assumere.
- Non usare `master` o `backup/con-incrementali` come base di lavoro: sono la
  vecchia linea “con incrementali” (baseline + run incrementali retroattivi),
  superata dalla strategia attuale “full run da zero”.
- Usare `master` o `backup/con-incrementali` solo su decisione esplicita
  dell'utente.
- `main` e `main-initial-backup` sono rami storici/di backup, non operativi.
- Non hardcodare `master` in documentazione, script o comandi operativi nuovi.

---

## Commit, push e perimetro

Regole obbligatorie:

- commit piccoli, a perimetro stretto: un commit, un tema;
- nessun `git add .`;
- nessun commit misto con file di temi diversi;
- nessun output reale committato (`output/`, `reports/`, `runs/`, PDF/HTML/Excel generati);
- nessun merge, push, reset, switch, checkout distruttivo, restore, clean o stash
  senza via libera esplicita dell'utente;
- nessun push se esistono file tracked modificati o commit locali non spiegati;
- prima di consolidare modifiche locali importanti o rischiose, valutare una patch
  di sicurezza nominata.

Distinguere sempre:

- codice: `src/`, `tests/`;
- configurazione: `config/`;
- documentazione/audit: `docs/`, `CLAUDE.md`, `AGENTS.md`;
- input dati: `Database_Dev/`, `bank_raw/`;
- output temporanei o di run: `output/`, `reports/`, `runs/`.

---

## Metodo obbligatorio per ogni modifica

Prima di modificare file non banali, dichiarare:

1. problema;
2. decisione funzionale;
3. contratto da preservare o implementare;
4. file/logiche da toccare;
5. test previsti;
6. audit previsti;
7. rischi;
8. cosa non verrà toccato.

Durante il lavoro:

1. leggere `AGENTS.md` e `CLAUDE.md`;
2. leggere YAML/documentazione/test pertinenti;
3. validare lo stato reale nel codice, non fidarsi solo di memoria o documenti stale;
4. modificare solo i file autorizzati dal perimetro concordato;
5. preservare il contratto contabile salvo mandato esplicito diverso.

Dopo il lavoro, dichiarare:

1. modifiche effettuate;
2. contratti implementati o preservati;
3. test eseguiti;
4. audit eseguiti;
5. strumenti non disponibili o non pertinenti;
6. rischi residui;
7. stato repository;
8. file modificati e destino.

---

## Qualità tecnica del codice

Ogni lavoro non banale su codice Python, CLI, pipeline, import, manifest, DRE,
banca, comprovanti, registrazioni manuali o output deve rispettare standard
elevati di qualità tecnica.

Principi obbligatori:

- codice semplice, leggibile, modulare e mantenibile;
- minor numero ragionevole di righe senza sacrificare chiarezza, testabilità o sicurezza;
- vietate duplicazioni non motivate;
- vietati workaround opachi;
- vietata complessità inutile;
- naming, struttura e commenti devono rendere chiaro il motivo delle scelte tecniche;
- ogni ottimizzazione o refactoring deve preservare il contratto funzionale;
- nessun cambio semantico senza dichiarazione esplicita.

Separare le responsabilità:

- parsing XML;
- parsing banca;
- parsing comprovanti;
- parsing template manuali;
- classificazione;
- regole YAML/config;
- logica contabile;
- DRE;
- pré-livro;
- manifest/provenance;
- CLI;
- output PDF/HTML/Excel;
- audit/report.

Non mescolare queste responsabilità senza motivo documentato.

---

## Test e audit proporzionati al rischio

Per ogni modifica non banale, eseguire test proporzionati al rischio.

Verifiche da valutare, dove disponibili e pertinenti:

- `py_compile` o controllo sintattico dei file Python toccati;
- test mirati sui moduli toccati;
- test di regressione rilevanti;
- test CLI se cambia argparse o comando utente;
- test e2e se cambia orchestrazione/pipeline;
- linter, formatter, type checker o static analysis solo se già presenti/configurati;
- `git diff --check`.

Se uno strumento non è disponibile, non configurato o non pertinente, dichiararlo
nel report finale invece di fingere che sia stato eseguito.

Esempi di test per area:

- manuali:
  - `pytest tests/test_cli.py -k manuais -v`
  - `pytest tests/test_lancamentos_manuais_gerenciais.py tests/test_lancamentos_manuais_template.py -v`
  - `pytest tests/test_lancamentos_manuais_e2e.py -v`
- DRE:
  - `pytest tests/test_dre_gerencial_provisoria.py -v`
- reprocessamento/pipeline:
  - `pytest tests/test_reprocessamento_pipeline.py -v`
- banca/comprovanti:
  - `pytest tests/test_importa_revisao_banca.py -v`
  - `pytest tests/test_payment_rules.py -v`
  - `pytest tests/test_lancamentos_banca.py -v`
- CLI:
  - `pytest tests/test_cli.py -v` oppure filtro mirato coerente col comando toccato.

La full suite è utile quando proporzionata, ma errori preesistenti devono essere
distinti chiaramente da regressioni introdotte.

Il report finale dei test deve indicare:

- comando eseguito;
- esito;
- area coperta;
- area non coperta;
- eventuali errori preesistenti;
- rischio residuo.

---

## Performance, memoria e uso risorse

L'efficienza va sempre considerata, ma le analisi pesanti devono essere
proporzionate al rischio.

Profilazione CPU, analisi memoria, benchmark, tracing dettagliato o dump sono
obbligatori solo quando almeno una delle condizioni seguenti è vera:

- la modifica tocca pipeline di import massivo;
- la modifica tocca batch storici;
- la modifica tocca generazione workbook/PDF/HTML pesanti;
- la modifica introduce nuovi cicli su dataset grandi;
- la modifica introduce caching, caricamenti massivi o trasformazioni costose;
- esistono timeout, rallentamenti, memoria anomala o regressioni osservate;
- l'utente richiede esplicitamente audit prestazionale;
- il contratto tecnico della fase lo prevede.

In questi casi il report deve indicare:

- data/ora e contesto dell'analisi;
- comando o strumento usato;
- versione dello strumento se rilevante;
- dati principali rilevati;
- moduli/funzioni coinvolte;
- conclusione;
- eventuali azioni correttive.

---

## Dump, trace e dati sensibili

I dump completi di memoria non sono una verifica ordinaria.

Possono essere prodotti solo quando sono tecnicamente necessari, proporzionati al
problema e autorizzati dall'utente, perché possono contenere:

- dati sensibili;
- path locali;
- contenuti di workbook;
- dati bancari;
- dati fiscali;
- informazioni operative riservate.

Quando un dump o trace può contenere dati sensibili:

- dichiarare il rischio;
- evitare allegati inutili;
- preferire metriche aggregate, estratti sanitizzati o report sintetici;
- documentare dove il file è stato salvato;
- non committare dump, trace pesanti o artefatti diagnostici salvo autorizzazione esplicita.

---

## YAML/config e regole contabili

Regole obbligatorie:

- non modificare YAML/config senza conferma esplicita dell'utente;
- prima di cambiare una regola, verificare piano conti, mapeamento DRE, test e documentazione;
- nessuna nuova regola di classificazione deve essere dedotta solo per similarity o intuizione;
- se una regola manca, lasciare il movimento in revisione;
- ogni modifica YAML deve avere test/audit di congruenza proporzionati.

File centrali:

- `config/plano_contas_gerencial.yaml`;
- `config/mapeamento_dre.yaml`;
- `config/regras_classificacao_produtos.yaml`;
- `config/regras_classificacao_comprovantes.yaml`.

Quando si toccano YAML/config, verificare dove pertinente:

- `conta_id` inesistenti;
- `dre_linha_id` inesistenti;
- mapping DRE incoerenti;
- duplicati;
- esclusioni con destinazione contabile indebita;
- conti tecnici fuori DRE;
- provenance/manifest aggiornati.

---

## Fasi visuali: PDF / HTML / Excel / dashboard

- Se una fase visuale già pushata viene bocciata, preferire un revert pulito
  piuttosto che patch correttive sopra codice bocciato.
- Esempio storico del progetto: “003E” `4871e12` → revert `2d69d3a`, confermato
  zero residui di codice dopo il revert.
- Per layout visuali, non inventare: seguire il modello già approvato oppure
  produrre un'anteprima e attendere conferma prima del commit.
- Output visuali non sono automaticamente ufficiali: devono essere etichettati e
  approvati.

---

## Documentazione

La documentazione può essere operativa, storica o stale. Non fidarsi del solo
titolo o della data.

Fonti aggiornate principali:

- `CLAUDE.md`;
- `AGENTS.md`;
- `docs/STATO_ATTUALE.md`;
- `docs/LANCAMENTOS_MANUAIS_GERENCIAIS_OPERADOR.md`;
- `docs/DECISIONI_CONTABILI.md`;
- `docs/GOVERNANCA_REGRAS_CONTABEIS.md`;
- `docs/DIRETRIZ_FONTES_BANCARIAS_COMPROVANTES.md`;
- `docs/WORKFLOW_CONTABIL_GERENCIAL.md`.

Documenti noti come storici/stale o da trattare con cautela:

- `docs/CONTINUITA_2026-05-19.md`;
- `docs/GRAPHIFY_AUDIT_SUMMARY.md`;
- vecchi riferimenti a `master`;
- vecchi layout visuali revertiti;
- bozze `FANTOZZI_*_proposta.yaml`.

Se una documentazione contraddice codice/test/YAML aggiornati, segnalarlo prima
di agire.

---

## Chiusura fase: regola obbligatoria

Nessuna fase può essere considerata chiusa se `git status -sb` mostra file
tracked modificati non classificati.

Prima di qualunque commit, push, cambio fase, handoff o chiusura:

```powershell
git status -sb
git diff --name-status
git diff --cached --name-status
```

Se esistono file tracked modificati, staged o unstaged, produrre la tabella:

| File | Origine/mandato | Stato | Destino |
|---|---|---|---|
| path | patch/MD/mandato | pronto/parziale/dubbio | commit mirato / patch nominata / scarto autorizzato |

Regole obbligatorie:

- nessun file tracked modificato può restare senza destino esplicito;
- nessun `git add .`;
- nessun commit misto;
- nessun push se esistono commit locali o file tracked modificati non spiegati;
- nessun reset, restore, checkout, clean o stash senza autorizzazione esplicita;
- se una modifica non è pronta per commit, creare patch nominata o mandato di continuità;
- se una modifica sembra da scartare, chiedere autorizzazione esplicita prima di scartarla.

Formula obbligatoria:

```text
FASE CHIUSA solo se:
1. commit/push necessari completati oppure patch nominata creata;
2. git status -sb controllato;
3. nessun file tracked modificato resta senza classificazione;
4. eventuali untracked sono dichiarati come output temporanei, audit, patch o file da ignorare.
```

---

## Report finale minimo

Ogni fase deve chiudere con un report che includa almeno:

1. branch;
2. commit hash, se creato;
3. conferma push, se eseguito;
4. file modificati;
5. file staged;
6. untracked rilevanti;
7. test eseguiti e risultato;
8. audit eseguiti;
9. rischi residui;
10. conferma che non sono stati toccati file fuori perimetro;
11. stato `origin/senza-incrementali...HEAD`, se pertinente;
12. `git diff --name-status`;
13. `git diff --cached --name-status`.

Se il lavoro è solo read-only, dichiarare esplicitamente:

```text
Nessun file modificato.
Nessun commit.
Nessun push.
```
