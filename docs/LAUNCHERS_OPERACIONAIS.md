# Launcher operacionais (.ps1 / .bat)

Questo documento elenca i launcher `.ps1`/`.bat` disponibili in `scripts/` per
eseguire comandi della CLI Fantozzi senza digitare a mano il comando Python
completo.

Standard introdotto in FZ-LAUNCHERS-002B. Copre solo comandi **read-only**.
I comandi che modificano dati (es. `executar-lancamentos-manuais-gerenciais`,
`reprocessar-importacao-real`) non hanno ancora un launcher: saranno oggetto
di fasi dedicate successive.

## Convenzione

- I launcher vivono in `scripts/` (cartella già usata nel repo per script
  operativi Python: `gerar_pdf_regras_banca.py`, `gerar_pdf_regras_xml.py`,
  `gerar_pdfs_contabeis.py`).
- Ogni comando ha una coppia `.ps1` + `.bat`: il `.ps1` contiene la logica,
  il `.bat` è solo un wrapper che richiama il `.ps1` via
  `powershell.exe -NoProfile -ExecutionPolicy Bypass -File ... %*`.
- Ogni `.ps1` risolve da solo il root del repo (`Split-Path -Parent $PSScriptRoot`)
  e vi si posiziona (`Set-Location`) prima di chiamare
  `.\.venv\Scripts\python.exe -m fantozzi.cli ...`. I percorsi passati come
  parametro vengono quindi interpretati **relativi al root del repo**, non
  alla cartella da cui si lancia lo script.
- Ogni launcher propaga l'exit code del comando Python: `0` = successo,
  diverso da `0` = errore o dato mancante da controllare.
- Nessun launcher esegue operazioni Git, non modifica YAML/config e non
  scrive fuori dalla cartella di output esplicitamente indicata.

## Launcher disponibili

### `scripts/auditar_regras_comprovantes.ps1` / `.bat`

- **Comando Python equivalente**:
  `.\.venv\Scripts\python.exe -m fantozzi.cli auditar-regras-comprovantes`
- **Read-only**: sì. Legge `config/regras_classificacao_comprovantes.yaml`,
  non lo modifica.
- **Parametri**: nessuno.
- **Output**: solo a console (nessun file scritto).
- **Uso**:

  ```powershell
  scripts\auditar_regras_comprovantes.ps1
  ```

  o, da `cmd.exe`:

  ```cmd
  scripts\auditar_regras_comprovantes.bat
  ```

- **Nota**: un alert noto `REGRA_9_01_001_TOP_LEVEL_DIGITAL` (regola con
  condizioni multiple obbligatorie) può comparire in output senza che sia un
  errore: l'unico segnale di fallimento affidabile è l'exit code diverso da
  zero.

### `scripts/auditar_comprovantes_xml_cnpj.ps1` / `.bat`

- **Comando Python equivalente**:
  `.\.venv\Scripts\python.exe -m fantozzi.cli auditar-comprovantes-xml-cnpj`
- **Read-only**: sì rispetto ai dati di origine (comprovantes, XML, YAML).
  Scrive **solo** dentro la cartella indicata da `-OutputDir`.
- **Parametri**:
  - `-ComprovantesCsv` (obbligatorio): CSV comprovantes da auditare, es.
    `reports\comprovantes_classificacao\runs\<RUN>\comprovantes_a_revisar.csv`.
    Obbligatorio di proposito: non esiste un default sicuro, il CSV
    "ultima run" cambia ad ogni ciclo di revisione banca.
  - `-XmlRoot` (opzionale, default `Database_Dev\XML_fatture_acquisto`).
  - `-OutputDir` (opzionale): se omesso viene generato automaticamente
    `output\audit_xml_comprovantes_launcher_<yyyyMMdd_HHmmss>`. Se la
    cartella indicata esiste già, il launcher si ferma con errore invece di
    sovrascrivere silenziosamente.
  - `-DiasTolerancia` (opzionale, default `7`).
  - `-SomenteRevisar` (switch opzionale): considera solo comprovantes con
    `status_classificacao` che inizia per `REVISAR`.
- **Uso**:

  ```powershell
  scripts\auditar_comprovantes_xml_cnpj.ps1 `
    -ComprovantesCsv "reports\comprovantes_classificacao\runs\2026-07-07_13L\comprovantes_a_revisar.csv" `
    -SomenteRevisar
  ```

  o, da `cmd.exe`:

  ```cmd
  scripts\auditar_comprovantes_xml_cnpj.bat -ComprovantesCsv "reports\comprovantes_classificacao\runs\2026-07-07_13L\comprovantes_a_revisar.csv" -SomenteRevisar
  ```

## Cosa NON fanno questi launcher

- Non classificano, non escludono, non alterano nessuna regola YAML.
- Non toccano `config/`, `bank_raw/`, XML di origine o Excel/CSV di lancamentos.
- Non fanno `git add`/`commit`/`push`.
- Non eseguono comandi modificativi della pipeline (importazione, run
  multifonte, DRE, pacote gerencial): quei comandi restano fuori scope di
  questa fase e verranno coperti in fasi successive (vedi
  `output/audit_launchers_bat_ps1_20260707/audit_launchers_bat_ps1.md`,
  sezione "Raccomandazioni prossime fasi").
