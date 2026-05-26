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

---

## Eccezione contabile puntuale: M A BRITO NF-e 2684

Decisione adottata nella fase `FZ-GIORNALE-DRE-MULTIFONTE-004B`.

Regola generale di governance:

- la fonte XML/NF-e prevale normalmente per i costi documentati;
- i pagamenti bancari servono come fonte DRE solo quando rappresentano uscite non già coperte o quando esiste decisione esplicita;
- il pagamento bancario non deve duplicare automaticamente una voce DRE già rappresentata da XML/NF-e;
- non esiste una regola automatica "preferire banca a XML".

Eccezione nominativa, unica e non replicabile:

| Campo | Valore |
|---|---|
| Fornitore | `M A BRITO COMERCIO LTDA` |
| Documento/NF | `2684` |
| Valore | `12000.00` |
| Descrizione | `GLP P45` |
| Motivo | `EXCECAO_CONTABIL_PONTUAL_M_A_BRITO_USAR_PAGAMENTOS` |

Questa eccezione esclude dal DRE la riga XML/NF-e specifica sopra e mantiene nel DRE i pagamenti bancari M A BRITO approvati/deduplicati.

## Deduplica bancaria cross-fonte con data divergente

Decisione adottata nella fase `FZ-GIORNALE-DRE-MULTIFONTE-004C`.

Questa regola riguarda solo il bridge bancario e non modifica la prevalenza XML/NF-e. Non deduplica XML x banca e non crea eccezioni nominative.

Una riga `CONTA_CORRENTE_CSV` e' esclusa con motivo `DUPLICADO_BANCARIO_DATA_DIVERGENTE` quando esiste un solo match non ambiguo in `COMPROVANTES_CSV` con:

| Campo | Valore |
|---|---|
| Fonte prioritaria | `COMPROVANTES_CSV` |
| Fonte esclusa | `CONTA_CORRENTE_CSV` |
| Data | divergente da 1 a 3 giorni |
| Beneficiario | normalizzato uguale e non generico |
| Valore | valore assoluto identico |
| CNPJ azienda | uguale quando presente in entrambe le fonti |
| Conta | `conta_id` uguale |

Se ci sono piu' candidati concorrenti nella finestra 1-3 giorni, la riga non viene esclusa automaticamente: resta inclusa e viene segnalata come `POSSIVEL_DUPLICADO_DATA_DIVERGENTE`.

La regola non si applica a righe della stessa fonte, beneficiari vuoti/generici, valori diversi, conti diversi, entrate bancarie o fonti diverse da `COMPROVANTES_CSV` x `CONTA_CORRENTE_CSV`.

Divieti:

- non trasformare questo caso in regola automatica per M A BRITO;
- non estendere questa eccezione ad altri documentos, importi, descrizioni o fornitori;
- non usare questa eccezione per derivare una gerarchia generale banca > XML.
