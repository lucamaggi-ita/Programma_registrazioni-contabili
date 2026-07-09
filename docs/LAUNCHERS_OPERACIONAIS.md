# Launcher operacionais (.ps1 / .bat)

Questo documento elenca i launcher `.ps1`/`.bat` disponibili in `scripts/` per
eseguire comandi della CLI Fantozzi senza digitare a mano il comando Python
completo.

Standard introdotto in FZ-LAUNCHERS-002B per i comandi **read-only** ed esteso
in FZ-LAUNCHERS-002A ai comandi operativi solo quando hanno guardrail espliciti.
I comandi che modificano dati devono richiedere conferma esplicita prima di
avviare Python.

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

Per i launcher operativi/modificativi:

- senza conferma esplicita devono stampare il riepilogo e uscire senza avviare
  Python;
- prima dell'esecuzione confermata devono bloccare se esistono modifiche
  tracked/staged, salvo flag di bypass documentato;
- gli untracked preesistenti possono essere mostrati come avviso, ma non sono
  bloccanti da soli.

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

### `scripts/executar_lancamentos_manuais_gerenciais.ps1` / `.bat`

- **Comando Python equivalente**:
  `.\.venv\Scripts\python.exe -m fantozzi.cli executar-lancamentos-manuais-gerenciais`
- **Operativo/modificativo**: si. Valida il template manuale, genera il CSV
  tecnico e puo produrre output contabili della pipeline gerencial provvisoria.
- **Template ufficiale**:
  `Database_Dev\Registrazioni_manuali\registrazioni_manuali_base.xlsx`.
- **CSV tecnico**:
  `Database_Dev\Registrazioni_manuali\registrazioni_manuali_tecnico.csv`.
- **Template legacy non operativo**:
  `Database_Dev\_templates\lancamentos_manuais_gerenciais_TEMPLATE.xlsx`.
  Il launcher non lo usa e blocca se viene passato come template.
- **Conferma obbligatoria**: senza `-ConfirmarExecucao` il launcher stampa il
  riepilogo, non avvia Python ed esce con codice `2`.
- **Git guardrail**: con `-ConfirmarExecucao`, prima di avviare Python blocca
  se `git diff --name-status` o `git diff --cached --name-status` non sono
  vuoti. `-SkipGitCheck` bypassa solo questo controllo e va usato con cautela.
- **Parametri**:
  - `-ConfirmarExecucao` (obbligatorio per esecuzione reale).
  - `-SkipGitCheck` (opzionale): non esegue il blocco su diff tracked/staged.
  - `-TemplatePath` (opzionale, default template ufficiale).
  - `-CsvTecnico` (opzionale, default CSV tecnico ufficiale).
- **Uso senza esecuzione reale**:

  ```powershell
  scripts\executar_lancamentos_manuais_gerenciais.ps1
  ```

  oppure, da `cmd.exe`:

  ```cmd
  scripts\executar_lancamentos_manuais_gerenciais.bat
  ```

- **Uso con esecuzione reale**:

  ```powershell
  scripts\executar_lancamentos_manuais_gerenciais.ps1 -ConfirmarExecucao
  ```

  oppure, da `cmd.exe`:

  ```cmd
  scripts\executar_lancamentos_manuais_gerenciais.bat -ConfirmarExecucao
  ```

- **Cosa non fa**:
  non modifica YAML/config contabili, non cambia il motore contabile, non usa
  il template legacy, non fa `git add`/`commit`/`push`, non esegue workflow
  incrementali.

### `scripts/reprocessar_importacao_real.ps1` / `.bat`

- **Comando Python equivalente**:
  `.\.venv\Scripts\python.exe -m fantozzi.cli reprocessar-importacao-real`
- **Operativo/modificativo**: si. Puo creare una nuova run sicura, importare
  XML reali dentro la run e generare output/reports contabili.
- **Conferma obbligatoria**: senza `-ConfirmarExecucao` il launcher stampa il
  riepilogo, non avvia Python ed esce con codice `2`.
- **Git guardrail**: con `-ConfirmarExecucao`, prima di avviare Python blocca
  se `git diff --name-status` o `git diff --cached --name-status` non sono
  vuoti. `-SkipGitCheck` e una deroga esplicita e non va usata come default.
- **Parametri principali della CLI**:
  `--xml-base`, `--saida-output`, `--saida-reports`,
  `--modo {provisorio,definitivo}`, `--dry-run`, `--fase9-csv`,
  `--lancamentos-manuais-gerenciais-csv`,
  `--lancamentos-banca-aprovados`, `--gerar-pacote-contabil-gerencial`,
  `--project-root`.
- **Pass-through argomenti**: gli argomenti dopo i parametri del launcher
  vengono passati alla CLI senza inventare parametri business.
- **Uso sicuro senza esecuzione reale**:

  ```powershell
  scripts\reprocessar_importacao_real.ps1
  ```

  oppure, da `cmd.exe`:

  ```cmd
  scripts\reprocessar_importacao_real.bat
  ```

- **Uso reale solo con conferma esplicita**:

  ```powershell
  scripts\reprocessar_importacao_real.ps1 -ConfirmarExecucao `
    --xml-base Database_Dev\XML_fatture_acquisto `
    --saida-output output\importacao_real `
    --saida-reports reports\importacao_real
  ```

### `scripts/importar_revisao_banca.ps1` / `.bat`

- **Comando Python equivalente**:
  `.\.venv\Scripts\python.exe -m fantozzi.cli importar-revisao-banca`
- **Operativo/modificativo**: si. Legge `bank_raw` e genera planilhas
  decisorie banca e manifest operativi.
- **Conferma obbligatoria**: senza `-ConfirmarExecucao` il launcher stampa il
  riepilogo, non avvia Python ed esce con codice `2`.
- **Git guardrail**: con `-ConfirmarExecucao`, prima di avviare Python blocca
  se `git diff --name-status` o `git diff --cached --name-status` non sono
  vuoti. `-SkipGitCheck` e una deroga esplicita e non va usata come default.
- **Parametri principali della CLI**:
  `--bank-raw-dir`, `--ano`, `--mes`, `--saida-dir`, `--regras`,
  `--plano-contas`, `--mapeamento-dre`.
- **Pass-through argomenti**: gli argomenti dopo i parametri del launcher
  vengono passati alla CLI senza inventare parametri business.
- **Uso sicuro senza esecuzione reale**:

  ```powershell
  scripts\importar_revisao_banca.ps1
  ```

  oppure, da `cmd.exe`:

  ```cmd
  scripts\importar_revisao_banca.bat
  ```

- **Uso reale solo con conferma esplicita**:

  ```powershell
  scripts\importar_revisao_banca.ps1 -ConfirmarExecucao --ano 2026 --mes 04
  ```

### `scripts/importar_revisao_xml.ps1` / `.bat`

- **Comando Python equivalente**:
  `.\.venv\Scripts\python.exe -m fantozzi.cli importar-revisao-xml`
- **Operativo/modificativo**: si. Gera planilha decisoria XML e manifest a
  partire da una run di reports gia esistente.
- **Conferma obbligatoria**: senza `-ConfirmarExecucao` il launcher stampa il
  riepilogo, non avvia Python ed esce con codice `2`.
- **Git guardrail**: con `-ConfirmarExecucao`, prima di avviare Python blocca
  se `git diff --name-status` o `git diff --cached --name-status` non sono
  vuoti. `-SkipGitCheck` e una deroga esplicita e non va usata come default.
- **Parametri principali della CLI**:
  `--run-reports-dir`, `--run-id`, `--reports-runs-root`, `--saida-dir`,
  `--regras`, `--plano-contas`, `--allow-stale-run`.
- **Pass-through argomenti**: gli argomenti dopo i parametri del launcher
  vengono passati alla CLI senza inventare parametri business.
- **Uso sicuro senza esecuzione reale**:

  ```powershell
  scripts\importar_revisao_xml.ps1
  ```

  oppure, da `cmd.exe`:

  ```cmd
  scripts\importar_revisao_xml.bat
  ```

- **Uso reale solo con conferma esplicita**:

  ```powershell
  scripts\importar_revisao_xml.ps1 -ConfirmarExecucao --run-id <RUN_ID>
  ```

## Cosa NON fanno questi launcher

- Non classificano, non escludono, non alterano nessuna regola YAML.
- Non toccano `config/`, `bank_raw/`, XML di origine o Excel/CSV di lancamentos.
- Non fanno `git add`/`commit`/`push`.
- Non eseguono comandi modificativi senza un guardrail dedicato e una conferma
  esplicita dell'operatore.
- Non eseguono workflow incrementali: il workflow vivo resta full run da zero
  su `senza-incrementali`.
