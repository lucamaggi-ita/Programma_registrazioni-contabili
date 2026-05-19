# Direttiva — Fonti bancarie e comprovantes

**Progetto:** Fantozzi / Programma_registrazioni_contabili
**Data di adozione:** 2026-05-19

---

## Principio

La fase banca deve considerare **comprovantes ed estratti conto come fonti complementari**, non alternative.

Il sistema non decide quale formato l'utente deve usare. L'utente sceglie il formato più adatto per leggere, capire e definire regole.

---

## Formati supportati

Devono restare disponibili e funzionanti tutti i formati di input bancario:

| Formato | Contenuto tipico | Uso primario |
|---|---|---|
| OFX | Estratto conto strutturato | Riconciliazione, saldo |
| XLSX / XLS | Estratto conto Excel (es. ConsultaPagamentos) | Comprovantes, dettaglio pagamenti |
| PDF / estratto visuale | Estratto leggibile, se utile | Verifica manuale |
| Comprovantes / Consulta Pagamentos | Dettaglio pagamento per singola riga | Identificazione beneficiario reale |

Nessuno di questi formati deve essere rimosso o reso inaccessibile dal sistema.

---

## Comprovante vs riga di estratto conto

Il comprovante è spesso una rappresentazione **più chiara** della riga bancaria:

- mostra il **beneficiario reale** invece dell'intermediario;
- contiene la **referencia** con descrizione del pagamento;
- ha il `payment_id` tracciabile fino alla fonte;
- non contiene aggregazioni di righe.

Quando la riga di estratto conto è ambigua (beneficiario generico, descrizione assente), il comprovante corrispondente può fornire il criterio di match necessario per definire una regola.

L'utente deve sempre poter confrontare entrambe le viste prima di definire una regola.

---

## Cosa deve fare il sistema

Il sistema deve:

- normalizzare i dati di input (beneficiario, referencia, importo, data);
- mostrare le righe in modo leggibile per l'utente;
- applicare le regole definite dall'utente;
- lasciare in revisione le righe senza regola;
- produrre report di audit con le righe classificate, escluse e in revisione.

---

## Cosa non deve fare il sistema

Il sistema non deve:

- usare l'estratto conto per inferire classificazioni contabili;
- usare la similarità tra transazioni per proporre conti;
- usare il matching bancario automatico per decidere regole;
- modificare `conta_id` o `dre_linha_id` basandosi sull'analisi dei movimenti;
- aggregare righe per dedurre il tipo di spesa.

---

## Matching bancario automatico

Il matching bancario automatico (associazione estratto conto ↔ comprovante per dedurre conti) è **disabilitato**.

Il flag di controllo nel manifest è:

```json
"matching_automatico_executado": false
```

Questo flag deve restare `false` in tutte le run operative.

---

## Priorità delle fonti per definire criteri di match

Quando l'utente deve definire una regola, la fonte preferita per costruire il criterio è:

1. **Comprovante** — `beneficiary_name` e `reference_id` sono i più affidabili;
2. **Estratto conto** — se il comprovante non è disponibile o la descrizione è assente;
3. **Alias/keyword approvato dall'utente** — testo libero confermato dall'utente.

Il criterio deve essere **dichiarato dall'utente**, non derivato automaticamente dalla fonte.

---

## Riferimento tecnico

Parser comprovantes:

```
src/fantozzi/banking/payment_rules.py
src/fantozzi/banking/payment_classification_run.py
```

Normalizzazione input bancario:

```
src/fantozzi/banking/normalizer.py
```

File di regole operativo:

```
config/regras_classificacao_comprovantes.yaml
```
