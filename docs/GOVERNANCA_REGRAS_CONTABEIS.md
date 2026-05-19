# Governance delle regole contabili gestionali

**Progetto:** Fantozzi / Programma_registrazioni_contabili
**Data di adozione:** 2026-05-19
**Priorità:** massima — questo documento prevale su qualsiasi istruzione implicita o comportamento di default degli agenti automatici.

---

## Principio fondamentale

Le regole contabili, di classificazione, esclusione, registrazione, DRE e giornale sono definite **esclusivamente dall'utente**.

Nessun agente automatico — ChatGPT, Claude, Codex o altro — ha l'autorità di definire, modificare o reinterpretare le regole contabili del progetto.

---

## Cosa possono fare gli agenti automatici

Gli agenti possono:

- applicare regole già definite e validate dall'utente;
- controllare la coerenza tecnica delle regole (formato YAML, campi obbligatori, conflitti tra ID);
- segnalare ambiguità o rischi di over-matching;
- chiedere conferma **una sola volta** quando una regola proposta sembra generica o rischiosa;
- produrre report, diff e output di audit senza modificare nulla.

---

## Cosa non possono fare gli agenti automatici

Gli agenti non possono:

- inventare conti contabili;
- inferire classificazioni da beneficiario, valore, data, descrizione o estratto conto;
- creare nuove regole contabili senza decisione esplicita dell'utente;
- reinterpretare o "correggere" decisioni già prese dall'utente;
- continuare a proporre associazioni dopo che l'utente ha dato una risposta esplicita;
- dedurre `conta_id`, `dre_linha_id` o `entra_dre` dall'analisi dei movimenti;
- usare somiglianze tra transazioni per proporre conti;
- usare l'estratto conto come motore di classificazione automatica.

---

## Gerarchia delle decisioni

Le decisioni dell'utente prevalgono sempre, nell'ordine:

1. **Decisione esplicita nell'ultima conversazione** — massima priorità.
2. **Regola scritta in `config/regras_classificacao_comprovantes.yaml`** — contratto operativo.
3. **Documento di governance** — vincolo permanente.
4. **Comportamento di default dell'agente** — ultima priorità, vietato se in conflitto con i punti precedenti.

---

## Azioni vietate anche su richiesta implicita

Se l'utente chiede genericamente "classifica questi comprovanti" senza fornire regole, l'agente deve:

1. rispondere che non può inferire conti senza regole dichiarate;
2. indicare quante righe restano in revisione;
3. chiedere se l'utente vuole definire nuove regole.

Non deve in nessun caso produrre una classificazione automatica.

---

## Applicabilità

Questo documento si applica a:

- `config/regras_classificacao_comprovantes.yaml` — regole di classificazione ed esclusione comprovanti;
- `config/plano_contas_gerencial.yaml` — piano dei conti gestionale;
- `config/mapeamento_dre.yaml` — mapeamento DRE;
- qualsiasi file YAML di regole presente in `config/`;
- qualsiasi logica di lancamentos, giornale o DRE gestionale prodotta dal progetto.

---

## Riferimento tecnico

Il motore che applica le regole è in:

```
src/fantozzi/banking/payment_rules.py
```

Il file operativo delle regole comprovanti è:

```
config/regras_classificacao_comprovantes.yaml
```

L'audit del YAML si esegue con:

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli auditar-regras-comprovantes --regras config\regras_classificacao_comprovantes.yaml
```
