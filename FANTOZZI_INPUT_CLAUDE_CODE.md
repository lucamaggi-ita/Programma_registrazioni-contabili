# INPUT PER CLAUDE CODE — Progetto “Fantozzi”
## Agente per registrazioni contabili di fatture di acquisto del ristorante

> Repository GitHub: `lucamaggi-ita/Programma_registrazioni-contabili`  
> Cartella locale di lavoro: `C:\Users\Luca Maggi\Documents\Git_Hub\Programma_Registrazioni_Contabili`  
> Ambiente previsto: Windows 10 / Windows 11  
> Nome agente/applicazione: **Fantozzi**

---

## 1. Contesto del progetto

Voglio creare un agente chiamato **Fantozzi** per aiutarmi nella registrazione della contabilità del mio ristorante.

Nella prima fase Fantozzi deve occuparsi della **registrazione automatica delle fatture di acquisto** partendo da file `.csv` contenenti le fatture e scrivendo il risultato in un file Excel strutturato.

Il progetto deve nascere in modo semplice, robusto e migliorabile.  
Questa prima versione non deve essere perfetta, ma deve creare una base tecnica ordinata, testabile e facilmente estendibile.

---

## 2. Obiettivo della prima bozza

Creare una prima versione funzionante di Fantozzi che:

1. Legga uno o più file `.csv` contenenti fatture di acquisto.
2. Estragga i dati principali della fattura.
3. Legga le righe degli articoli/prodotti presenti in fattura.
4. Classifichi ogni articolo nel corretto **conto economico** usando un file di regole configurabile.
5. Raggruppi o divida la fattura in righe contabili in base ai conti economici trovati.
6. Scriva le registrazioni in un file Excel.
7. Crei una scheda di lavoro diversa per ogni mese.
8. Ordini le registrazioni per data fattura.
9. Produca un log/report degli articoli non classificati o classificati con bassa sicurezza.

---

## 3. Vincoli importanti

### 3.1 Non fare una soluzione troppo rigida

Il formato reale dei `.csv` potrebbe cambiare.  
Quindi il parser deve essere progettato in modo configurabile e difensivo.

Claude deve evitare codice “monolitico” difficile da modificare.

### 3.2 Non sovrascrivere dati manuali

Il file Excel finale sarà usato anche per registrazioni manuali non provenienti da fatture.

Quindi Fantozzi deve:

- non cancellare fogli esistenti;
- non cancellare righe manuali;
- possibilmente identificare le righe generate automaticamente;
- evitare duplicati;
- poter lavorare in modalità `dry-run`;
- salvare una copia di backup prima di modificare un Excel esistente.

### 3.3 Progetto pensato per crescere

In futuro Fantozzi potrà registrare anche:

- buste paga dei lavoratori;
- comprovanti di pagamento importati dalla banca;
- movimenti non originati da fatture;
- riconciliazioni tra fatture, pagamenti e cassa.

La fase 1 però deve restare concentrata sulle fatture di acquisto da `.csv`.

---

## 4. Dati da importare dalla fattura

Per ogni fattura Fantozzi deve importare almeno:

- nome fornitore;
- CNPJ fornitore;
- numero fattura;
- data fattura;
- data o date di scadenza pagamento;
- valore totale della fattura;
- articoli/prodotti presenti nella fattura;
- valore di ogni articolo o riga articolo;
- conto economico attribuito a ogni articolo;
- valore imputato a ogni conto economico;
- eventuale osservazione o alerta.

---

## 5. Logica contabile richiesta

Non voglio importare tutti gli articoli come righe contabili finali.

Fantozzi deve leggere gli articoli della fattura, classificarli e registrarli nel conto economico corretto.

Se una stessa fattura contiene prodotti appartenenti a più conti economici, la fattura deve essere divisa in più registrazioni.

### Esempio

Fattura:

- Fornitore: `OESA`
- Numero fattura: `12566`
- Data fattura: `05/05/2026`
- Prodotti:
  - Coca Cola: R$ 10,00
  - Rio Sol: R$ 20,00
  - Esponja: R$ 5,00
- Totale fattura: R$ 35,00
- Scadenze:
  - R$ 17,50 il `20/05/2026`
  - R$ 17,50 il `27/05/2026`

Registrazioni attese:

| Data fattura | Fornitore | CNPJ | Fattura | Conto economico | Descrizione base | Valore conto | Scadenza 1 | Valore scadenza 1 | Scadenza 2 | Valore scadenza 2 | Origine |
|---|---|---|---|---|---|---:|---|---:|---|---:|---|
| 05/05/2026 | OESA |  | 12566 | Refrigerante | Coca Cola | 10,00 | 20/05/2026 | 5,00 | 27/05/2026 | 5,00 | CSV |
| 05/05/2026 | OESA |  | 12566 | Vino | Rio Sol | 20,00 | 20/05/2026 | 10,00 | 27/05/2026 | 10,00 | CSV |
| 05/05/2026 | OESA |  | 12566 | Pulizia | Esponja | 5,00 | 20/05/2026 | 2,50 | 27/05/2026 | 2,50 | CSV |

### Regola di ripartizione delle scadenze

Quando una fattura viene divisa su più conti economici, le scadenze devono essere ripartite **proporzionalmente** al valore attribuito a ogni conto.

Formula:

```text
quota_scadenza_per_conto = valore_scadenza_originale * valore_conto / totale_fattura
```

La somma delle quote ripartite deve tornare al totale originale della scadenza.  
Gestire eventuali differenze di arrotondamento sull’ultima riga della fattura.

---

## 6. Piano dei conti e regole di classificazione

Fantozzi deve usare un file di configurazione esterno per classificare gli articoli.

Creare un file iniziale, ad esempio:

```text
config/plano_contas_rules.yaml
```

oppure:

```text
config/plano_contas_rules.json
```

Preferisco YAML se non crea complicazioni inutili.

Il file deve essere molto dettagliato e facile da modificare manualmente.

### 6.1 Regole iniziali obbligatorie

#### Conto economico: Refrigerante

Classificare come `Refrigerante` tutti i prodotti che contengono parole intere o parti riconoscibili come:

- coca cola
- coca-cola
- coca
- fanta
- sprite
- tonica
- tônica
- schweppes
- sweppes
- sch
- fys

Nota: nel README originale compariva `fantea`, probabilmente errore di digitazione. Inserire comunque `fantea` come alias per sicurezza.

#### Conto economico: Vino

Classificare come `Vino` tutti i prodotti che contengono:

- vinho
- espumante
- rio sol
- zephira
- hote
- sauvignon
- chardonnay
- brut
- moscatel
- demi-sec
- demisec
- cabernet
- merlot
- malbec
- rosé
- rose
- seco
- suave

#### Conto economico: Pulizia

Classificare come `Pulizia` tutti i prodotti che contengono:

- detergente
- sanitaria
- água sanitária
- agua sanitaria
- alcool
- álcool
- esponja
- pano
- vassoura
- sabão
- sabao
- desinfetante
- cloro
- multiuso
- limpador
- papel toalha
- saco lixo
- lixo
- luva
- balde
- rodo

---

## 7. Requisiti per la classificazione

La classificazione deve essere:

- case-insensitive;
- robusta rispetto ad accenti portoghesi;
- robusta rispetto a spazi multipli;
- robusta rispetto a trattini e punteggiatura;
- capace di usare alias parziali quando previsto;
- capace di segnalare prodotti non classificati.

### 7.1 Normalizzazione testo

Implementare una funzione di normalizzazione che:

- converte in minuscolo;
- rimuove o normalizza accenti;
- sostituisce punteggiatura con spazi;
- compatta spazi multipli;
- permette match affidabili su descrizioni sporche.

Esempio:

```text
"Água Sanitária 2L" -> "agua sanitaria 2l"
"COCA-COLA lata" -> "coca cola lata"
```

### 7.2 Conflitti

Se un articolo corrisponde a più conti, Fantozzi non deve scegliere in silenzio.

Deve produrre un alerta:

```text
CLASSIFICACAO_AMBIGUA
```

e mettere la riga in un report di revisione.

Per la prima bozza si può usare la prima regola trovata, ma deve essere tracciato l’alerta.

---

## 8. Struttura Excel desiderata

Il file Excel deve contenere un foglio per ogni mese.

Nome foglio consigliato:

```text
2026-05
2026-06
2026-07
```

Oppure:

```text
Mai_2026
Jun_2026
Jul_2026
```

Preferisco `YYYY-MM` perché è più ordinato e meno ambiguo.

### 8.1 Colonne minime del foglio mensile

Creare almeno queste colonne:

| Colonna | Descrizione |
|---|---|
| `data_fattura` | Data della fattura |
| `fornitore` | Nome fornitore |
| `cnpj` | CNPJ fornitore |
| `numero_fattura` | Numero della fattura |
| `conto_economico` | Conto economico attribuito |
| `descricao_itens` | Descrizione sintetica degli articoli raggruppati |
| `valor_conta` | Valore imputato al conto |
| `data_vencimento_1` | Prima scadenza |
| `valor_vencimento_1` | Valore della prima scadenza per quel conto |
| `data_vencimento_2` | Seconda scadenza, se presente |
| `valor_vencimento_2` | Valore della seconda scadenza per quel conto |
| `data_vencimento_3` | Terza scadenza, se presente |
| `valor_vencimento_3` | Valore della terza scadenza per quel conto |
| `origem` | Per esempio `CSV` o `MANUAL` |
| `arquivo_origem` | Nome del file CSV importato |
| `hash_registro` | Chiave tecnica per evitare duplicati |
| `alertas` | Eventuali alertas |
| `observacoes` | Note libere |

### 8.2 Righe manuali

Le righe manuali potranno essere aggiunte dall’utente.

Per questo motivo:

- Fantozzi deve usare la colonna `origem = CSV` per le righe automatiche;
- l’utente potrà usare `origem = MANUAL`;
- Fantozzi non deve modificare le righe `MANUAL`;
- Fantozzi deve evitare di duplicare righe `CSV` già importate.

---

## 9. Formato CSV

Poiché il formato reale dei CSV può variare, creare un layer di importazione configurabile.

Per la prima bozza:

1. Creare un CSV di esempio in `samples/faturas_exemplo.csv`.
2. Creare una configurazione colonne in `config/csv_columns.yaml`.
3. Implementare un parser che possa essere adattato facilmente.

### 9.1 Colonne esempio del CSV

Esempio minimo:

```csv
fornecedor,cnpj,numero_fatura,data_fatura,item_descricao,item_valor,vencimento_data,vencimento_valor
OESA,,12566,05/05/2026,Coca Cola,10.00,20/05/2026,17.50
OESA,,12566,05/05/2026,Coca Cola,10.00,27/05/2026,17.50
OESA,,12566,05/05/2026,Rio Sol,20.00,20/05/2026,17.50
OESA,,12566,05/05/2026,Rio Sol,20.00,27/05/2026,17.50
OESA,,12566,05/05/2026,Esponja,5.00,20/05/2026,17.50
OESA,,12566,05/05/2026,Esponja,5.00,27/05/2026,17.50
```

Atenzione: questo esempio ripete le scadenze su ogni item.  
Il parser deve evitare di sommare più volte le stesse scadenze della stessa fattura.

Per identificare una fattura usare una chiave tipo:

```text
fornecedor + cnpj + numero_fatura + data_fatura
```

---

## 10. Struttura tecnica proposta

Usare Python.

Creare una struttura simile:

```text
Programma_Registrazioni_Contabili/
├─ README.md
├─ pyproject.toml
├─ .gitignore
├─ config/
│  ├─ plano_contas_rules.yaml
│  └─ csv_columns.yaml
├─ samples/
│  └─ faturas_exemplo.csv
├─ output/
│  └─ .gitkeep
├─ reports/
│  └─ .gitkeep
├─ src/
│  └─ fantozzi/
│     ├─ __init__.py
│     ├─ cli.py
│     ├─ models.py
│     ├─ csv_importer.py
│     ├─ classifier.py
│     ├─ allocation.py
│     ├─ excel_writer.py
│     ├─ normalization.py
│     └─ logging_utils.py
└─ tests/
   ├─ test_normalization.py
   ├─ test_classifier.py
   ├─ test_allocation.py
   └─ test_csv_importer.py
```

---

## 11. CLI desiderata

Creare una CLI semplice.

Esempio:

```powershell
python -m fantozzi.cli import-faturas `
  --csv samples\faturas_exemplo.csv `
  --excel output\registrazioni_contabili.xlsx `
  --rules config\plano_contas_rules.yaml `
  --columns config\csv_columns.yaml
```

Aggiungere modalità prova:

```powershell
python -m fantozzi.cli import-faturas `
  --csv samples\faturas_exemplo.csv `
  --excel output\registrazioni_contabili.xlsx `
  --dry-run
```

La modalità `dry-run` deve:

- leggere il CSV;
- classificare;
- calcolare le registrazioni;
- stampare un riepilogo;
- generare eventualmente un report;
- non modificare il file Excel.

---

## 12. Output di controllo

Dopo ogni importazione Fantozzi deve produrre un riepilogo:

```text
FANTOZZI - Importazione fatture

File CSV: samples/faturas_exemplo.csv
Fatture lette: 1
Righe articolo lette: 3
Registrazioni create: 3
Valore totale fatture: R$ 35,00
Valore totale registrato: R$ 35,00
Articoli non classificati: 0
Classificazioni ambigue: 0
File Excel aggiornato: output/registrazioni_contabili.xlsx
Report: reports/import_report_2026-05-05.txt
```

---

## 13. Report di revisione

Creare un report testuale o CSV con:

- articoli non classificati;
- articoli con classificazione ambigua;
- fatture con totale incoerente;
- scadenze con totale incoerente;
- righe duplicate ignorate;
- errori di formato data;
- errori di formato valore.

Nome consigliato:

```text
reports/import_report_YYYYMMDD_HHMMSS.txt
```

---

## 14. Gestione duplicati

Fantozzi deve generare un `hash_registro` per ogni registrazione automatica.

La chiave può derivare da:

```text
fornitore | cnpj | numero_fattura | data_fattura | conto_economico | valor_conta
```

Prima di scrivere sul file Excel, controllare se lo stesso hash esiste già.

Se esiste:

- non duplicare la riga;
- registrare nel report che la riga è stata ignorata come duplicato.

---

## 15. Test minimi obbligatori

Creare test automatici con `pytest`.

Test richiesti:

1. Normalizzazione:
   - `Água Sanitária` deve diventare `agua sanitaria`;
   - `COCA-COLA` deve matchare `coca cola`.

2. Classificazione:
   - `Coca Cola lata` -> `Refrigerante`;
   - `Rio Sol branco` -> `Vino`;
   - `Esponja limpeza` -> `Pulizia`;
   - prodotto sconosciuto -> alerta `NAO_CLASSIFICADO`.

3. Ripartizione scadenze:
   - fattura totale R$ 35,00;
   - conti: R$ 10,00 / R$ 20,00 / R$ 5,00;
   - scadenze: R$ 17,50 e R$ 17,50;
   - quote attese: 5,00 / 10,00 / 2,50 per ogni scadenza.

4. Anti-duplicato:
   - importare due volte lo stesso CSV non deve duplicare le registrazioni.

5. Scrittura Excel:
   - deve creare il foglio `2026-05`;
   - deve preservare righe `MANUAL`;
   - deve ordinare le righe per `data_fattura`.

---

## 16. Librerie consigliate

Usare preferibilmente:

- `pandas` per lettura CSV e gestione dati tabellari, se utile;
- `openpyxl` per scrittura Excel;
- `PyYAML` per leggere regole YAML;
- `pytest` per test;
- solo librerie semplici e diffuse.

Evitare dipendenze complesse o non necessarie.

---

## 17. Cose da NON fare nella prima bozza

Non implementare ancora:

- importazione da banca;
- buste paga;
- riconciliazione bancaria;
- interfaccia grafica;
- intelligenza artificiale esterna;
- collegamenti diretti al sistema fiscale brasiliano;
- lettura automatica di XML NF-e, salvo creare solo un punto di estensione futuro;
- login, utenti o cloud.

La fase 1 deve essere locale, semplice e verificabile.

---

## 18. Istruzioni operative per Claude Code

Claude, lavora nella cartella:

```text
C:\Users\Luca Maggi\Documents\Git_Hub\Programma_Registrazioni_Contabili
```

Repository GitHub previsto:

```text
lucamaggi-ita/Programma_registrazioni-contabili
```

### Prima di modificare

1. Controlla lo stato del repository:

```powershell
git status
```

2. Verifica quali file esistono già.
3. Se esiste già un `README.md`, leggilo prima.
4. Non cancellare file esistenti senza motivo.
5. Se la cartella è vuota, crea la struttura iniziale.
6. Mantieni modifiche piccole, ordinate e documentate.

### Obiettivo della tua esecuzione

Crea una prima bozza funzionante di Fantozzi:

- struttura progetto;
- README.md migliorato;
- configurazioni iniziali;
- CSV di esempio;
- codice Python;
- CLI;
- test minimi;
- report di importazione;
- scrittura Excel;
- protezione anti-duplicati;
- preservazione righe manuali.

### Alla fine restituisci

1. Elenco dei file creati/modificati.
2. Comandi per installare ed eseguire.
3. Comandi per lanciare i test.
4. Risultato atteso dell’esempio.
5. Eventuali limiti della prima bozza.
6. Prossimi miglioramenti consigliati.

---

## 19. Comandi attesi per test locale

Questi sono comandi indicativi che il progetto dovrebbe supportare.

Da PowerShell:

```powershell
cd "C:\Users\Luca Maggi\Documents\Git_Hub\Programma_Registrazioni_Contabili"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -e ".[dev]"
pytest
python -m fantozzi.cli import-faturas --csv samples\faturas_exemplo.csv --excel output\registrazioni_contabili.xlsx
```

Se `pip install -e ".[dev]"` è troppo complesso per la prima bozza, usare un `requirements.txt`, ma preferisco `pyproject.toml`.

---

## 20. Criterio di accettazione della prima bozza

La prima bozza è accettabile se:

- i test passano;
- il CSV di esempio viene letto;
- il file Excel viene creato;
- il foglio `2026-05` viene creato;
- la fattura OESA viene divisa in tre righe:
  - Refrigerante;
  - Vino;
  - Pulizia;
- le due scadenze vengono ripartite correttamente;
- un secondo import dello stesso CSV non duplica le righe;
- le righe manuali non vengono cancellate;
- gli articoli non classificati vengono segnalati nel report.

---

## 21. Nota sul nome “Fantozzi”

Il nome dell’agente è **Fantozzi**.

Nel codice usare nomi sobri e tecnici, ma nella CLI e nei report può comparire:

```text
FANTOZZI - Importazione fatture
```

Il tono deve essere pratico, chiaro e orientato al controllo contabile, non scherzoso.

---

## 22. Priorità

Se devi scegliere tra velocità e robustezza, scegli robustezza.

Se devi scegliere tra automazione completa e controllo umano, scegli controllo umano.

Fantozzi deve aiutare a evitare errori, non nasconderli.
