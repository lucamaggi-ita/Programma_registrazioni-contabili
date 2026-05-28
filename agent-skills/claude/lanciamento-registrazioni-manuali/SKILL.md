---
name: lanciamento-registrazioni-manuali
description: Workflow operatore per registrazioni manuali gerenciais. Apre/genera l'Excel operativo, attende compilazione, valida, applica executar-run-incremental e restituisce un riepilogo sintetico senza full run, audit automatici o conferme tecniche.
---

# Skill: lanciamento-registrazioni-manuali

Skill operativa per la frase naturale:

```text
Fantozzi, voglio fare registrazioni manuali
```

Principio della fase 011B:

```text
Operatore parla naturale
-> Excel operativo
-> operatore compila
-> Fantozzi valida
-> executar-run-incremental
-> risultato sintetico
```

Non fare full run automatica. Non lanciare audit automatici. Non chiedere conferme tecniche intermedie.

---

## Path operativi

```text
Database_Dev\Registrazioni_manuali\registrazioni_manuali_base.xlsx
Database_Dev\Registrazioni_manuali\registrazioni_manuali_tecnico.csv
Database_Dev\Registrazioni_manuali\registrazioni_manuali_validazione_manifest.json
output\baselines\baseline_atual.json
output\incremental_<tag>
```

Per 011B usare, se serve un tag esplicito:

```text
output\incremental_011B_operator_manual_test
```

---

## Trigger

Caso A, apertura/creazione Excel:

```text
Fantozzi, voglio fare registrazioni manuali
iniziare registrazioni manuali
aprire registrazioni manuali
fare registrazioni manuali
devo compilare registrazioni manuali
```

Caso B, Excel compilato:

```text
ho finito
ho compilato e salvato
ho salvato le registrazioni manuali
procedi con le registrazioni manuali
registra le manuali
```

Quando l'operatore dice "ho finito" o equivalente, procedere al Caso B senza chiedere un'altra conferma.

---

## Messaggio Operatore Caso A

Eseguire direttamente apertura/generazione Excel. Mostrare solo:

```text
Compila DATA, CONTA_ESCOLHIDA, DESCRICAO_DA_DESPESA, VALOR, OBSERVACAO.
Poi salva e dimmi "ho finito".
```

Non spiegare CSV, manifest, Git, baseline, DRE o comandi tecnici all'operatore in questa fase.

---

## Caso A - Aprire o Generare Excel

1. Se serve, creare la cartella operativa.
2. Se `registrazioni_manuali_base.xlsx` non esiste, generarlo.
3. Aprire il file se l'ambiente consente GUI.
4. Se la GUI non e disponibile, indicare solo il path del file.
5. Attendere che l'operatore dica "ho finito".

Comandi ammessi, atomici:

```powershell
New-Item -ItemType Directory -Force "Database_Dev\Registrazioni_manuali"
```

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli gerar-template-lancamentos-manuais-gerenciais --saida "Database_Dev\Registrazioni_manuali\registrazioni_manuali_base.xlsx"
```

```powershell
Start-Process "Database_Dev\Registrazioni_manuali\registrazioni_manuali_base.xlsx"
```

Non chiedere conferma per aprire, generare, leggere o validare file.

---

## Caso B - Validare e Applicare Incrementale

Sequenza obbligatoria:

1. Verificare che esista `output\baselines\baseline_atual.json`.
2. Validare l'Excel operativo.
3. Generare CSV tecnico.
4. Generare manifest validazione.
5. Se il manifest e `APROVADO`, chiamare `executar-run-incremental`.
6. Mostrare solo il riepilogo sintetico finale.

Comando validazione:

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli validar-lancamentos-manuais-gerenciais --entrada "Database_Dev\Registrazioni_manuali\registrazioni_manuali_base.xlsx" --saida-csv "Database_Dev\Registrazioni_manuali\registrazioni_manuali_tecnico.csv" --manifest "Database_Dev\Registrazioni_manuali\registrazioni_manuali_validazione_manifest.json"
```

Comando incrementale:

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli executar-run-incremental --baseline "output\baselines\baseline_atual.json" --manuali-csv "Database_Dev\Registrazioni_manuali\registrazioni_manuali_tecnico.csv" --manuali-manifest "Database_Dev\Registrazioni_manuali\registrazioni_manuali_validazione_manifest.json" --saida "output\incremental_011B_operator_manual_test"
```

Adattare il tag `incremental_011B_operator_manual_test` solo per evitare sovrascritture. Non usare `reprocessar-importacao-real` nel workflow manuali ordinario.

---

## Conferme da Non Chiedere

Non chiedere conferme per:

```text
aprire Excel
generare Excel
leggere file
controllare Git
controllare manifest
creare cartella output
validare Excel
generare CSV tecnico
generare manifest
eseguire incremental
generare DRE/pacote
```

Fermarsi solo davanti a blocchi reali.

---

## Blocchi Reali

Baseline assente o non approvata:

```text
BLOCCO: BLOCCO_INCREMENTAL_BASELINE_AUSENTE
Non posso registrare incrementale perche manca una baseline approvata.
Serve prima approvare/registrare una baseline.
```

Manuali non validati:

```text
BLOCCO: BLOCCO_INCREMENTAL_MANUALI_NAO_APROVADO
Le registrazioni manuali non sono approvate dalla validazione. Correggi l'Excel e dimmi di nuovo "ho finito".
```

Delta gia applicato:

```text
BLOCCO: BLOCCO_INCREMENTAL_DELTA_JA_APLICADO
Queste registrazioni manuali risultano gia applicate. Non sono state duplicate.
```

Hash strutturale cambiato:

```text
BLOCCO: BLOCCO_INCREMENTAL_STRUCTURAL_CHANGE
La struttura contabile e cambiata. Non posso applicare l'incrementale in sicurezza.
```

Manuali fuori periodo baseline:

```text
BLOCCO: BLOCCO_INCREMENTAL_BASELINE_INSUFICIENTE
La data delle registrazioni manuali e fuori dal periodo coperto dalla baseline approvata.
```

Excel assente/non leggibile o errore tecnico reale:

```text
BLOCCO: <codice tecnico>
Causa: <descrizione semplice>
Azione: correggere il file o ripetere il comando indicato.
```

---

## Manuali Retroattivi

Sono ammessi manuali retroattivi dentro il periodo coperto dalla baseline.

Esempio ammesso:

```text
baseline: 2026-01-01 -> 2026-04-30
manuale: 2026-02-15
```

La run incrementale deve registrare:

```text
periodo_ini
periodo_fim
competencias_afetadas
```

Non rilanciare XML, banca, comprovantes o conta corrente.

---

## Cosa Non Fare

```text
NON usare reprocessar-importacao-real per manuali operatore ordinari
NON fare full run automatica
NON fare bridge banca
NON reprocessare XML
NON rilanciare comprovantes
NON rilanciare conta corrente
NON lanciare auditar-giornale-multifonte automaticamente
NON lanciare auditar-dre-gerencial automaticamente
NON modificare YAML/config
NON modificare piano dei conti
NON modificare mapeamento DRE
NON usare git add .
NON fare git clean / git reset / git restore
NON committare output reali
```

Audit e full run esistono come comandi separati e richiedono richiesta esplicita dell'operatore.

---

## Output Finale Operatore

Mostrare un riepilogo breve:

```text
Registrazioni manuali completate.

Righe valide: 2
Totale manuali: R$ 125,30
Run: INCREMENTAL
Full run eseguita: no
XML rilanciati: no
Banca rilanciata: no
Comprovantes rilanciati: no
Conta corrente rilanciata: no
DRE: OK_PROVISORIO
Pacote: generato
```

Non mostrare dump JSON, comandi PowerShell, liste lunghe o audit completi salvo errore.

---

## Riferimenti CLI

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli gerar-template-lancamentos-manuais-gerenciais --saida "Database_Dev\Registrazioni_manuali\registrazioni_manuali_base.xlsx"
```

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli validar-lancamentos-manuais-gerenciais --entrada "Database_Dev\Registrazioni_manuali\registrazioni_manuali_base.xlsx" --saida-csv "Database_Dev\Registrazioni_manuali\registrazioni_manuali_tecnico.csv" --manifest "Database_Dev\Registrazioni_manuali\registrazioni_manuali_validazione_manifest.json"
```

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli executar-run-incremental --baseline "output\baselines\baseline_atual.json" --manuali-csv "Database_Dev\Registrazioni_manuali\registrazioni_manuali_tecnico.csv" --manuali-manifest "Database_Dev\Registrazioni_manuali\registrazioni_manuali_validazione_manifest.json" --saida "output\incremental_011B_operator_manual_test"
```
