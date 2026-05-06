# Fantozzi - Registrazioni Contabili Gerenciais

Fantozzi e un programma Python locale per importare XML NF-e di acquisti,
classificare i produtos secondo un `plano_contas_gerencial`, generare
`Lancamentos` e scrivere un Excel controllabile manualmente.

Principio guida:

```text
controllo > automazione cieca
revisione > classificazione inventata
```

Quando Fantozzi non e sicuro, manda il lancamento in revisione invece di
inventare una classificazione. Le decisioni contabili o fiscali finali vanno
sempre validate con il contador.

## Cosa Fa Oggi

- importa XML NF-e;
- estrae `produtos`, valores e `duplicatas`;
- classifica prodotti tramite `config/regras_classificacao_produtos.yaml`;
- raggruppa i `Lancamentos` per `conta_id`, non per singolo prodotto;
- ripartisce le `duplicatas` proporzionalmente tra i `Lancamentos`;
- gestisce `fretes`, `seguros`, `outras despesas` e `descontos` sulle compras;
- scrive un file Excel `.xlsx`;
- preserva righe manuali;
- evita duplicati;
- crea l'aba `Revisao` per i `Lancamentos` da controllare;
- offre una CLI per importare 1 XML o una cartella di XML.

## Requisiti

- Windows 10/11;
- PowerShell;
- Python 3.14 o versione compatibile gia usata nel progetto;
- virtualenv locale `.venv`;
- dipendenze installate dal `pyproject.toml`;
- Excel o LibreOffice solo per controllo visivo del `.xlsx`, non per i test.

Su questo PC `python` puo non essere nel PATH. Nei comandi usare:

```powershell
.\.venv\Scripts\python.exe
```

## Preparazione Ambiente

Aprire PowerShell nella cartella del progetto:

```powershell
cd "C:\Users\Luca Maggi\Documents\Git_Hub\Programma_Registrazioni_Contabili"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
.\.venv\Scripts\Activate.ps1
```

Se serve installare o riallineare le dipendenze:

```powershell
.\.venv\Scripts\python.exe -m pip install -e .[dev]
```

## Eseguire I Test

```powershell
.\.venv\Scripts\python.exe -m pytest -v
```

Risultato atteso attuale:

```text
53 passed
```

## Importare Un Singolo XML NF-e

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli importar-nfe `
  --xml "samples\nfe_exemplo\nfe_oesa_12566_20260505.xml" `
  --excel "output\lancamentos_fantozzi.xlsx"
```

Parametri:

- `--xml` indica il file XML NF-e da importare;
- `--excel` indica il file Excel da creare o aggiornare.

Il comando stampa log semplici in PowerShell, per esempio quanti `Lancamentos`
sono stati generati e dove e stato aggiornato l'Excel.

## Importare Una Cartella Di XML NF-e

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli importar-pasta-nfe `
  --pasta "samples\nfe_exemplo" `
  --excel "output\lancamentos_fantozzi.xlsx"
```

Il comando:

- legge i file `.xml` e `.XML` nella cartella indicata;
- ordina i file per nome;
- importa gli XML NF-e validi;
- segnala XML non validi senza bloccare gli altri;
- scrive tutti i `Lancamentos` nello stesso Excel.

Se la cartella non esiste, non e una cartella, non contiene XML, o nessun XML
valido viene importato, il comando termina con errore e non crea l'Excel.

## Config Opzionali

La CLI usa automaticamente:

```text
config/plano_contas_gerencial.yaml
config/regras_classificacao_produtos.yaml
```

E possibile indicare file diversi:

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli importar-nfe `
  --xml "samples\nfe_exemplo\nfe_oesa_12566_20260505.xml" `
  --excel "output\lancamentos_fantozzi.xlsx" `
  --plano-contas "config\plano_contas_gerencial.yaml" `
  --regras-produtos "config\regras_classificacao_produtos.yaml"
```

Usare config alternative solo con attenzione: `plano_contas_gerencial` e
`regras_classificacao_produtos` guidano la classificazione contabile.

## Excel Generato

Il file Excel viene creato o aggiornato nel path passato con `--excel`.
Esempio:

```text
output\lancamentos_fantozzi.xlsx
```

Il workbook contiene almeno due abas:

```text
Lancamentos
Revisao
```

### Lancamentos

Contiene i `Lancamentos` generati dagli XML importati.

Campi importanti:

- `data_documento`: data di emissione del documento;
- `competencia`: mese di competenza in formato `YYYY-MM`;
- `fornecedor_ou_favorecido`: nome del fornecedor;
- `documento`: numero della NF-e;
- `conta_id`: codice della conta gerencial;
- `conta_nome`: nome della conta gerencial;
- `grupo_conta`: gruppo della conta;
- `descricao_itens`: prodotti riassunti nel lancamento;
- `valor`: valore del lancamento;
- `duplicatas_json`: duplicatas ripartite sul lancamento, in formato JSON;
- `origem`: origine del lancamento, per esempio `XML_NFE` o `MANUAL`;
- `chave_origem`: chave NF-e usata come origine;
- `hash_registro`: chiave tecnica usata per evitare duplicati;
- `status_revisao`: `OK` oppure `REVISAR`;
- `alertas`: motivi di attenzione o revisione;
- `observacoes`: campo libero per note manuali.

### Revisao

Contiene i `Lancamentos` con:

```text
status_revisao != OK
```

Questa aba serve per controllo manuale. Quando compare un lancamento in
`Revisao`, va controllato prima di usare i dati per analisi gerenciale.

## Anti-Duplicati

Rilanciare lo stesso comando sullo stesso XML o sulla stessa cartella non
duplica i `Lancamentos` gia importati.

La deduplica usa `hash_registro` quando disponibile. Se manca, usa una chiave
stabile basata su:

```text
origem
chave_origem
conta_id
valor
```

## Righe Manuali

Le righe manuali inserite dall'utente nell'Excel non vengono cancellate.

Le righe con:

```text
origem = MANUAL
```

sono preservate e non entrano nella deduplica automatica degli XML.

## Revisione Manuale

Se il comando stampa warning o se l'Excel contiene righe nella aba `Revisao`:

1. aprire l'Excel;
2. controllare l'aba `Revisao`;
3. verificare `conta_id`, `conta_nome`, `valor`, `descricao_itens` e `alertas`;
4. correggere o completare manualmente se necessario;
5. validare eventuali dubbi contabili o fiscali con il contador.

Esempi di motivi di revisione:

- prodotto non classificato;
- classificazione ambigua;
- valore inconsistente;
- possibile investimento;
- altri casi da controllare manualmente.

## Limiti Attuali

Oggi Fantozzi non fa ancora:

- DRE gerencial completa automatica;
- dashboard;
- importazione PDF;
- lettura di email;
- integrazione bancaria;
- invio automatico al contador;
- validazione fiscale ufficiale.

Fantozzi non sostituisce il contador e non rende definitivo un trattamento
contabile o fiscale.

## Comandi Git Utili

```powershell
git status --short
git log --oneline -n 8
```
