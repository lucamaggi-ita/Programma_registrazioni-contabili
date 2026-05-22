# Graphify Audit Summary — Fantozzi

Data: 2026-05-22  
Progetto: `Programma_registrazioni_contabili`  
Modalità esecuzione: controllata / no-cache / output locale non committato

## 1. Scopo dell'audit

Questo documento riassume l'uso di Graphify come strumento di mappatura tecnica del progetto Fantozzi.

Graphify è stato usato per:

- visualizzare dipendenze tra moduli Python;
- identificare nodi ad alto impatto;
- raggruppare il progetto in community tecniche;
- individuare connessioni utili per audit e manutenzione futura;
- supportare lo sviluppo futuro di Fantozzi e delle skill operative.

Graphify non è fonte di verità contabile.  
Le fonti di verità restano:

- decisioni esplicite dell'utente;
- YAML validati;
- documenti di governance;
- test automatici;
- codice reale del progetto.

## 2. Stato finale della pipeline Graphify

La pipeline Graphify è stata completata con successo.

Output finale locale in `graphify-out/`:

- `graph.html`
- `graph.json`
- `GRAPH_REPORT.md`
- `.graphify_detect.json`
- `.graphify_ast.json`
- `.graphify_semantic.json`
- `.graphify_extract.json`
- `.graphify_analysis.json`
- `.graphify_labels.json`

Metriche principali:

- AST: 1.583 nodi, 4.479 archi
- Semantico: 45 nodi, 85 archi
- Merge extraction: 1.628 nodi, 4.564 archi
- Grafo finale: 1.608 nodi, 3.978 archi
- Community: 44

Stato Git finale della pipeline:

- nessun file del progetto modificato;
- nessun `git add`;
- nessun commit;
- nessun push;
- unico elemento non tracciato: `graphify-out/`.

## 3. Decisione su `graphify-out/`

`graphify-out/` deve restare output locale di audit/analisi.

Per ora non committare:

- `graph.html`;
- `graph.json`;
- file `.graphify_*.json`;
- output Graphify completi.

Motivo:

- sono file generati;
- possono essere rigenerati;
- includono molta informazione tecnica non necessaria alla repo principale;
- aumenterebbero rumore nel repository.

Questo documento è invece committabile perché contiene solo le conclusioni operative dell'audit.

## 4. Community principali

Le community Graphify sono state etichettate manualmente per renderle leggibili.

Alcune community rilevanti:

- Regole Classificazione Comprovantes
- DRE Gerencial e Report OK
- Pipeline Reprocessamento NF-e
- Import NF-e & Excel Writer
- Import File Bancari
- Motore Decisioni Fase 9
- Candidati e Revisione Classificazione
- CLI & Diagnostica Excel
- Governance, Config e Documentazione
- Matching Bancario Legacy/Audit
- Pacote Contabil Gerencial
- Pacote Revisione Comprovantes

Nota importante: le label sono descrittive del grafo, non decisioni contabili.

## 5. God Nodes principali

Graphify ha evidenziato alcuni nodi ad alta connettività.

I più importanti:

1. `write_lancamentos_to_excel()`
   - punto di uscita centrale dei lançamentos verso Excel.

2. `allocate_nfe_to_lancamentos()`
   - cuore della conversione NF-e → lançamentos.

3. `evaluate_fase9_decisions()`
   - motore decisionale della Fase 9.

4. `build_parser()`
   - registro centrale della CLI.

5. `gerar_relatorios_gerenciais_ok_provisorios()`
   - orchestratore dei report gerenciais provvisori.

6. `gerar_pre_livro_diario_teste()`
   - generatore del pre-livro diario.

7. `importar_planilha_revisao_classificacao()`
   - import del workbook di revisione classificazione.

8. `gerar_dre_gerencial_ok_provisoria()`
   - generatore del DRE gerencial provvisorio.

9. `auditar_relatorios_gerenciais_ok_provisorios()`
   - audit dei report gerenciais.

10. `gerar_proposta_classificacao()`
   - generatore delle proposte di classificazione.

Interpretazione: questi nodi non devono essere modificati superficialmente. Ogni intervento deve avere test mirati.

## 6. Analisi specifica: `build_parser()`

Graphify ha segnalato `build_parser()` come bridge tra più community.

Analisi effettuata:

- definito in `src/fantozzi/cli.py`;
- chiamato da `main()`;
- usato direttamente da circa 26 test;
- collega CLI, banking, Fase 9, comprovantes, diagnostica e test.

Conclusione:

Non refactorare `build_parser()` ora.

Motivi:

- è normale che il parser CLI sia un nodo centrale;
- il suo ruolo è registrare comandi di domini diversi;
- i test coprono bene il contratto del parser;
- un refactor ora avrebbe rischio superiore al beneficio;
- la betweenness alta è compatibile con il ruolo della CLI.

Regola operativa:

`build_parser()` può essere modificato solo con test mirati e con attenzione a:

- nomi dei comandi;
- argomenti obbligatori;
- flag opzionali;
- retrocompatibilità CLI;
- test che importano direttamente `build_parser()`.

Refactor futuro possibile solo se:

- `cli.py` diventa ingestibile;
- si decide una struttura a subparser per dominio;
- esiste una fase dedicata con test di regressione completi.

## 7. Connessioni interessanti

Graphify ha confermato alcune connessioni utili:

### Governance

`CLAUDE.md` e `docs/GOVERNANCA_REGRAS_CONTABEIS.md` esprimono lo stesso principio:

- il sistema applica regole;
- non inventa classificazioni;
- le decisioni contabili appartengono all'utente.

Questa coerenza è positiva.

### Pipeline XML

Tripletta concettuale:

- XML NF-e;
- regras_classificacao_produtos;
- lançamentos.

È il cuore della pipeline NF-e.

### Pipeline bancaria/comprovantes

Tripletta concettuale:

- extrato bancário;
- comprovantes;
- regras_classificacao_comprovantes.

Nota: il matching bancario automatico resta trattato come legacy/audit, non come motore decisionale operativo.

### Pipeline DRE

Tripletta concettuale:

- lançamentos;
- mapeamento_dre;
- DRE gerencial.

Questa relazione conferma che il DRE dipende dalla qualità dei lançamentos e dal mapping DRE.

## 8. Rischi evidenziati

### 8.1 Nodi centrali

Funzioni come `allocate_nfe_to_lancamentos()`, `write_lancamentos_to_excel()` e `evaluate_fase9_decisions()` hanno molte relazioni.

Rischio: cambiamenti piccoli possono avere effetti ampi.

Mitigazione:

- test mirati;
- audit read-only prima del commit;
- diff limitato;
- nessun refactor opportunistico.

### 8.2 Inferenze Graphify

Alcune relazioni del grafo sono `INFERRED`.

Rischio: Graphify può sovrastimare relazioni semantiche o architetturali.

Mitigazione:

- usare Graphify come mappa;
- verificare sempre nel codice reale;
- non usare Graphify come fonte di verità.

### 8.3 Proposte storiche vs configurazione attiva

La community "Proposte Storiche Piano/Regole" può mescolare materiali storici con elementi operativi.

Rischio: confondere proposte vecchie con configurazione corrente.

Mitigazione:

- distinguere sempre `config/` attivo da file proposta;
- trattare file proposta come storico salvo decisione esplicita.

### 8.4 Matching bancario

Graphify distingue "Import File Bancari" da "Matching Bancario Legacy/Audit".

Decisione corrente:

- l'estratto bancario serve come fonte/controllo;
- le classificazioni devono derivare da regole dichiarate;
- il matching automatico non deve diventare motore decisionale senza fase dedicata.

## 9. Uso futuro di Graphify

Graphify è utile per:

- audit architetturale;
- orientamento tra moduli;
- preparare skill Fantozzi;
- identificare punti critici prima di modifiche grandi;
- capire impatto di refactor.

Graphify non va usato per:

- decidere classificazioni contabili;
- creare regole;
- validare fiscalmente dati;
- sostituire test;
- sostituire audit umano.

Frequenza consigliata:

- non eseguire Graphify per ogni microfase;
- rieseguirlo dopo macrofasi importanti;
- utile prima o dopo Fase 14;
- utile prima di creare/aggiornare skill stabili.

## 10. Raccomandazioni operative

1. Non committare `graphify-out/` per ora.

2. Mantenere questo documento come sintesi committabile dell'audit.

3. Non refactorare `build_parser()` ora.

4. Trattare i God Nodes come aree sensibili:
   - `allocate_nfe_to_lancamentos()`;
   - `write_lancamentos_to_excel()`;
   - `evaluate_fase9_decisions()`;
   - `build_parser()`.

5. Ogni modifica a nodi centrali deve avere test mirati.

6. Per skill Fantozzi, usare Graphify come mappa di orientamento, non come fonte normativa.

7. Per la direzione banking/comprovantes, mantenere il principio:
   - regole esplicite dell'utente;
   - nessuna inferenza autonoma;
   - matching automatico solo come audit/controllo se autorizzato.

## 11. Stato finale

Audit Graphify completato e utile.

Conclusione principale:

Il progetto ha una struttura ampia ma coerente.  
I nodi centrali identificati sono attesi per il tipo di sistema.  
Non emergono motivi per refactor immediato.  
Il valore principale di Graphify è come mappa tecnica e supporto di audit per fasi future.
