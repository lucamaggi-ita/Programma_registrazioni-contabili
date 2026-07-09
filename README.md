# Fantozzi - Registrazioni Contabili Gerenciais

Fantozzi e un programma Python locale per costruire una contabilita
gerenciale auditabile a partire da XML NF-e, banca/comprovantes, registrazioni
manuali e regole dichiarate dall'utente.

Principio guida:

```text
controllo > automazione cieca
revisione > classificazione inventata
```

Quando Fantozzi non e sicuro, lascia il movimento in revisione invece di
inventare una classificazione. Le decisioni contabili appartengono all'utente e
non vengono dedotte automaticamente.

## Stato Operativo Corrente

- Branch operativo: `senza-incrementali`.
- Remote operativo: `origin/senza-incrementali`.
- Workflow vivo: full run da zero sulle fonti disponibili.
- Non usare `master`, `main` o `backup/con-incrementali` come base operativa:
  sono linee storiche/di backup.
- Non usare workflow incrementali o baseline retroattive come percorso
  operativo corrente.

## Cosa Fa Oggi

- importa e classifica XML/NF-e;
- importa revisioni banca/comprovantes;
- applica regole dichiarate in YAML/config;
- lascia in `REVISAR` cio che non ha regola sicura;
- valida registrazioni manuali gerenciais;
- genera CSV tecnico manuali;
- esegue full run provvisorie da zero;
- genera giornale, pre-livro, DRE, pacote contabile gerencial e report;
- genera PDF DRE, dashboard HTML e Excel giornale leggibile da run gia
  prodotte;
- produce audit e output etichettati.

## Launcher Operativi

Il riferimento completo e `docs/LAUNCHERS_OPERACIONAIS.md`.

Usare i launcher `.ps1`/`.bat` quando esistono. La CLI Python diretta resta un
riferimento tecnico avanzato, non il percorso operatore preferito.

### Audit Read-Only

- `scripts/auditar_regras_comprovantes.ps1` / `.bat`
- `scripts/auditar_comprovantes_xml_cnpj.ps1` / `.bat`

Contratto: read-only rispetto a regole e fonti. L'audit XML x comprovantes
scrive solo nella cartella `-OutputDir` indicata e blocca sovrascritture
silenziose.

### Operativi/Modificativi

- `scripts/executar_lancamentos_manuais_gerenciais.ps1` / `.bat`
- `scripts/reprocessar_importacao_real.ps1` / `.bat`
- `scripts/importar_revisao_banca.ps1` / `.bat`
- `scripts/importar_revisao_xml.ps1` / `.bat`

Contratto: richiedono `-ConfirmarExecucao`. Senza conferma stampano il riepilogo,
non avviano Python ed escono con codice `2`. Con conferma applicano guardrail
Git su diff tracked/staged, salvo deroga esplicita `-SkipGitCheck`.

### Generativi/Output

- `scripts/gerar_dre_gerencial_ok_provisoria.ps1` / `.bat`
- `scripts/gerar_pdf_dre.ps1` / `.bat`
- `scripts/gerar_dashboard_html.ps1` / `.bat`
- `scripts/gerar_giornale_leggibile.ps1` / `.bat`

Contratto: richiedono `-ConfirmarExecucao`. Senza conferma non avviano Python,
non generano output ed escono con codice `2`.

Esempio sicuro senza esecuzione reale:

```powershell
scripts\gerar_dashboard_html.ps1
```

Esempio reale solo dopo verifica operativa:

```powershell
scripts\gerar_dashboard_html.ps1 -ConfirmarExecucao --run-dir output\run
```

## Registrazioni Manuali

Template operativo ufficiale:

```text
Database_Dev\Registrazioni_manuali\registrazioni_manuali_base.xlsx
```

CSV tecnico generato:

```text
Database_Dev\Registrazioni_manuali\registrazioni_manuali_tecnico.csv
```

Template legacy da non usare:

```text
Database_Dev\_templates\lancamentos_manuais_gerenciais_TEMPLATE.xlsx
```

Il launcher `executar_lancamentos_manuais_gerenciais` verifica il template
ufficiale e blocca il template legacy.

## Ambiente

Aprire PowerShell nella root del repository:

```powershell
cd "C:\Users\Recanto do Sossego\Documents\Luca\GitHub\Gestione_Ristorante\Programma_registrazioni_contabili"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
.\.venv\Scripts\Activate.ps1
```

Su questo PC `python` puo non essere nel PATH. Nei comandi tecnici usare:

```powershell
.\.venv\Scripts\python.exe
```

Se serve installare o riallineare le dipendenze:

```powershell
.\.venv\Scripts\python.exe -m pip install -e .[dev]
```

## Test

```powershell
.\.venv\Scripts\python.exe -m pytest -v
```

Per modifiche mirate, preferire prima i test dell'area toccata e poi ampliare
solo se il rischio lo richiede.

## Fonti E Config Ufficiali

- `config/plano_contas_gerencial.yaml`
- `config/mapeamento_dre.yaml`
- `config/regras_classificacao_produtos.yaml`
- `config/regras_classificacao_comprovantes.yaml`
- `Database_Dev\XML_fatture_acquisto`
- `bank_raw\`
- `Database_Dev\Registrazioni_manuali\registrazioni_manuali_base.xlsx`

Non modificare YAML/config o fonti operative senza decisione esplicita.

## Output

`output/`, `reports/` e `runs/` contengono output generati, audit o run. Non
committare output reali salvo mandato esplicito e motivato.

## Revisione Manuale

Se un output contiene `REVISAR`, `MANCANTE`, `NON_VERIFICATO`,
`NON_DISPONIBILE` o `N.A.`, il dato va trattato come stato esplicito e non come
zero tecnico. La revisione umana resta parte del contratto operativo.

## Limiti Operativi

Fantozzi non sostituisce il contador, non rende definitivo un trattamento
contabile o fiscale, non classifica senza regola approvata e non deve essere
usato per automazione cieca.

## Comandi Git Utili

```powershell
git status -sb
git diff --name-status
git diff --cached --name-status
git log --oneline -n 8
```
