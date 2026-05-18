# Convenzione bank_raw

Aggiornato: 2026-05-17

## Scopo

`bank_raw/` e la convenzione operativa per organizzare i file bancari reali
prima della normalizzazione, riconciliazione manuale, audit e debug.

La cartella contiene dati bancari reali e non deve essere committata. Per questo `bank_raw/` e inclusa in `.gitignore`.

## Struttura

La struttura ufficiale e:

```text
bank_raw/<ANNO>/<YYYY-MM>/<CNPJ>/
```

Esempio:

```text
bank_raw/
  2026/
    2026-01/
      08884180000123/
        ofx_padrao.ofx
        ofx_saldos.ofx
        lancamentos.xlsx
        consulta_pagamentos.xls
      55746722000181/
        ofx_padrao.ofx
        ofx_saldos.ofx
        lancamentos.xlsx
        consulta_pagamentos.xls
```

Ogni mese, quando disponibili, sono previste due cartelle CNPJ:

- `08884180000123`
- `55746722000181`

I file devono essere importati e normalizzati separatamente per CNPJ. Eventuali consolidamenti devono avvenire solo dopo la normalizzazione.

## Nomi File

Dentro ogni cartella CNPJ usare questi nomi standard:

- `ofx_padrao.ofx`: fonte primaria dei movimenti bancari.
- `ofx_saldos.ofx`: fonte solo per audit saldi; righe come `SALDO ANTERIOR` e `SALDO TOTAL DISPONIVEL DIA` non sono movimenti gestionali.
- `lancamentos.xlsx`: fonte di arricchimento e audit, non fonte primaria; puo avere metadati iniziali, header alla riga 10, righe saldo e alias come `Valor`, `Valor (R$)`, `Saldo`, `Saldo (R$)`.
- `consulta_pagamentos.xls`: comprovanti/pagamentos in formato `.xls` reale, letto con `xlrd>=2.0.1`.

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
