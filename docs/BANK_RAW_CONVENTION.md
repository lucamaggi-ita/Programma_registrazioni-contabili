# Convenzione bank_raw

Aggiornato: 2026-05-17
Riallineato al flusso reale: 2026-07-03

**NOTA OPERATIVA 2026-07-03:** questa sezione "Nomi File" è stata riallineata
al flusso decisionale realmente operativo (`importar-revisao-banca`, skill
`importa-revisione-banca`), che è il percorso che alimenta giornale/DRE. La
convenzione di cartelle `bank_raw/<ANNO>/<YYYY-MM>/<CNPJ>/` resta invariata.

## Scopo

`bank_raw/` e la convenzione operativa per organizzare i file bancari reali
prima della normalizzazione, riconciliazione manuale, audit e debug.

La cartella contiene dati bancari reali e non deve essere committata. Per questo `bank_raw/` e inclusa in `.gitignore`.

## Struttura

La struttura ufficiale e:

```text
bank_raw/<ANNO>/<YYYY-MM>/<CNPJ>/
```

Esempio (nomi file reali osservati in `bank_raw/2026/*`, non nomi fissi — vedi
sezione "Nomi File" per i pattern effettivi usati dal codice):

```text
bank_raw/
  2026/
    2026-01/
      08884180000123/
        Extrato_Lancamentos_<conta>_<data>.xlsx
        ConsultaPagamentos.xls
        Extrato_<conta>_<data>.ofx
      55746722000181/
        Extrato_Lancamentos_<conta>_<data>.xlsx
        ConsultaPagamentos.xls
        Extrato_<conta>_<data>.ofx
```

Ogni mese, quando disponibili, sono previste due cartelle CNPJ:

- `08884180000123`
- `55746722000181`

I file devono essere importati e normalizzati separatamente per CNPJ. Eventuali consolidamenti devono avvenire solo dopo la normalizzazione.

## Nomi File

**NOTA OPERATIVA 2026-07-03:** i nomi "standard" descritti in una versione
precedente di questa sezione (`ofx_padrao.ofx`, `ofx_saldos.ofx`,
`lancamentos.xlsx`, `consulta_pagamentos.xls`) non corrispondono ai pattern
realmente riconosciuti dal flusso decisionale operativo
(`importar-revisao-banca`, `src/fantozzi/banking/revisao_banca.py`). Il
codice cerca, per ogni cartella CNPJ, questi pattern (case/prefisso, non nomi
fissi):

- **`ConsultaPagamentos*.xls`** — comprovantes/pagamentos, letti con
  `xlrd>=2.0.1`. **Fonte ausiliaria**: aiuta a leggere beneficiario/referência,
  classificata con `config/regras_classificacao_comprovantes.yaml`.
- **`Extrato_Lan*amentos*.xlsx`** — estratto/lançamenti conto corrente.
  **Fonte primaria operativa**: è l'universo dei movimenti che alimenta
  realmente giornale/DRE nel flusso `importar-revisao-banca`.
- **`Extrato*.ofx`** — file OFX. Il flusso decisionale lo **rileva** ma non lo
  usa per generare i lançamentos (`ofx_usado` è sempre `False` nel risultato
  di `importar_revisao_banca`). L'OFX viene effettivamente parsato solo dal
  comando separato `importar-extratos-bancarios` (Fase 13A); il suo output
  (`bank_transactions_normalized.csv`) è consumato solo dal modulo di
  matching automatico bancario, **oggi disabilitato e non esposto dalla
  CLI operativa** (`src/fantozzi/banking/matching.py`). L'OFX quindi **non
  alimenta oggi il giornale/DRE**, salvo una futura decisione esplicita
  dell'utente di ricollegare quel modulo.

Se in una cartella CNPJ mancano sia `ConsultaPagamentos*.xls` sia
`Extrato_Lan*amentos*.xlsx`, il comando `importar-revisao-banca` si arresta
con errore esplicito.

DA VERIFICARE — non è stato individuato nel codice un nome file "canonico"
unico e obbligatorio per ciascun pattern (i tre esempi nell'albero sopra sono
generati con numero conto e data nel nome, non un nome fisso).

## Uso Con La CLI

La CLI della Fase 13A continua ad accettare path espliciti:

```powershell
python -m fantozzi.cli importar-extratos-bancarios `
  --cnpj <CNPJ> `
  --ofx <CAMINHO_OFX_PADRAO> `
  --ofx-saldos <CAMINHO_OFX_SALDOS> `
  --xlsx-lancamentos <CAMINHO_XLSX_LANCAMENTOS> `
  --consulta-pagamentos <CAMINHO_CONSULTA_PAGAMENTOS_XLS> `
  --saida <DIRETORIO_SAIDA>
```

`bank_raw/` e una convenzione operativa, non un requisito funzionale della
Fase 13A. La Fase 13 non esegue piu matching automatico banco x comprovante:
i comprovantes vengono classificati solo tramite
`config/regras_classificacao_comprovantes.yaml`. Senza regola dichiarata, o con
regole conflittanti verso conti diversi, l'output deve andare in revisione.
