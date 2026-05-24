# Importa Revisione XML Prodotti - Operatore

Versione: 1.0 - 2026-05-24

## Scopo

Questo workflow prepara il file Excel ufficiale per revisionare prodotti XML/NF-e
che non sono ancora classificati in modo sicuro.

Il comando non modifica YAML, non crea regole e non inferisce conti. Genera o
aggiorna solo il file decisionale ufficiale:

```text
Database_Dev\_decisoes_regras\xml\Analisi_regole_registrazioni_xml_prodotti.xlsx
```

I CSV sotto `reports\importacao_real\runs\...` sono output tecnici della run.
Non sono il file decisionale primario dell'operatore.

## Step 1 - Generare una run fresca

Eseguire la pipeline reale/provvisoria con il YAML corrente:

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli reprocessar-importacao-real `
  --xml-base Database_Dev\XML_fatture_acquisto `
  --saida-output output\importacao_real `
  --saida-reports reports\importacao_real `
  --modo provisorio
```

Al termine annotare il `RUN_ID`, per esempio:

```text
2026-05-24_001
```

## Step 2 - Generare o aggiornare l'Excel ufficiale

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli importar-revisao-xml --run-id <RUN_ID>
```

Esempio:

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli importar-revisao-xml --run-id 2026-05-24_001
```

Il comando aggiorna:

```text
Database_Dev\_decisoes_regras\xml\Analisi_regole_registrazioni_xml_prodotti.xlsx
Database_Dev\_decisoes_regras\xml\manifest_importa_revisao_xml.json
Database_Dev\_decisoes_regras\xml\manifest_importa_revisao_xml_<RUN_ID>.json
```

Il manifest corrente puo essere sovrascritto. Il manifest con `RUN_ID` conserva
lo storico operativo di quella run.

## Preservazione decisioni

Quando l'Excel viene rigenerato, le decisioni gia compilate vengono preservate
tramite `CHAVE_ESTAVEL`.

Se una chiave non compare piu nella run corrente, non resta nel foglio operativo.
Questo e intenzionale: il file ufficiale mostra la vista operativa della run
corrente, non uno storico di tutte le decisioni passate.

## Colonne da compilare

Compilare solo le colonne decisionali quando esiste una decisione sicura:

```text
STATUS_DECISAO
ACAO_REGRA
CONTA_ESCOLHIDA
PALAVRAS_CHAVE_SUGERIDAS
OBSERVACAO_OPERADOR
```

`CONTA_ID_APROVADA` e `CONTA_NOME_APROVADA` sono derivate da `CONTA_ESCOLHIDA`.
Non devono essere usate per inventare conti.

## BLOQUEADO

Le righe `BLOQUEADO` indicano operazioni speciali o casi ad alto rischio.

Per queste righe non approvare automaticamente alias o regole. Le azioni
ammesse sono:

```text
IGNORAR
BLOQUEADO
REVISAR
```

Non usare `CRIAR_ALIAS` o `CRIAR_REGRA` su righe bloccate senza revisione
contabile esplicita.

## Divieti operativi

Non usare `--allow-stale-run` nel lavoro operativo normale.

`--allow-stale-run` e riservato solo ad audit/debug autorizzato su dati storici.
Se il comando segnala `BLOCCO`, fermarsi e generare una run fresca con
`reprocessar-importacao-real`.

Non modificare manualmente:

```text
config\regras_classificacao_produtos.yaml
config\plano_contas_gerencial.yaml
config\mapeamento_dre.yaml
```

## Comandi legacy

Questi comandi possono esistere per workflow precedenti o tecnici, ma non sono
il workflow primario di questo file ufficiale:

```text
exportar-revisao-classificacao
gerar-planilha-revisao-classificacao
importar-planilha-revisao-classificacao
gerar-proposta-classificacao
aplicar-fase9-revisao
```

## Fase successiva

Dopo la compilazione dell'Excel, la fase successiva e `regole-registrazioni-xml`.

Quella fase applica solo decisioni approvate al YAML prodotti:

```text
config\regras_classificacao_produtos.yaml
```

`importar-revisao-xml` non modifica mai il YAML.
