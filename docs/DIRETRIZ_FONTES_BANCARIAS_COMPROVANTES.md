# Direttiva — Fonti bancarie e comprovantes

**Progetto:** Fantozzi / Programma_registrazioni_contabili
**Data di adozione:** 2026-05-19

---

## Principio

Per la fase banca, il sistema deve usare **congiuntamente estratti conto e comprovantes**.

**L'estratto conto è la fonte primaria** dei movimenti bancari e determina l'universo delle righe da processare.
**I comprovantes sono fonte ausiliaria e arricchimento informativo**, utili per leggere beneficiario/referência e applicare regole definite dall'utente.
Il sistema deve considerare solo movimenti bancari **in uscita** e applicare esclusivamente regole contabili dichiarate dall'utente.

Il sistema non decide quale formato l'utente deve usare. L'utente sceglie il formato più adatto per leggere, capire e definire regole.

---

## Formati supportati

Devono restare disponibili e funzionanti tutti i formati di input bancario:

| Formato | Contenuto tipico | Uso primario |
|---|---|---|
| OFX | Estratto conto strutturato | Riconciliazione, saldo |
| XLSX / XLS | Estratto conto Excel e comprovantes (es. ConsultaPagamentos) | Universo movimenti e dettaglio comprovantes |
| PDF / estratto visuale | Estratto leggibile, se utile | Verifica manuale |
| Comprovantes / Consulta Pagamentos | Dettaglio pagamento per singola riga | Identificazione beneficiario reale |

Nessuno di questi formati deve essere rimosso o reso inaccessibile dal sistema.

---

## Ruoli delle due fonti

### Estratto conto — fonte primaria

L'estratto conto determina l'**universo dei movimenti bancari** da processare.

- definisce quali righe esistono e devono essere trattate;
- contiene data, importo, tipo movimento;
- può avere beneficiario generico o intermediario (es. PJBANK, CAIXA) invece del beneficiario reale;
- deve essere filtrato: si considerano **solo movimenti in uscita** (pagamenti), non entrate/incassi.

### Comprovantes — fonte ausiliaria

Il comprovante è un arricchimento informativo della riga di estratto:

- mostra il **beneficiario reale** invece dell'intermediario;
- contiene la **referencia** con descrizione del pagamento;
- ha il `payment_id` tracciabile fino alla fonte;
- non contiene aggregazioni di righe.

Quando la riga di estratto conto è ambigua, il comprovante corrispondente può fornire il criterio di match per definire una regola.

L'utente deve sempre poter confrontare entrambe le viste prima di definire una regola.

### Collegamento estratto + comprovante

Il sistema deve collegare le righe di estratto con i comprovanti corrispondenti per data, valore, beneficiario o referencia quando possibile.

Il collegamento serve solo come **supporto informativo** per l'operatore. Non deve essere usato per creare classificazioni autonome.

---

## Cosa deve fare il sistema

Il workflow corretto della fase banca è:

1. **Leggere l'estratto conto** — fonte primaria; determina l'universo dei movimenti bancari da processare.
2. **Considerare principalmente movimenti in uscita** — escludere entrate e incassi; il focus è sui pagamenti.
3. **Leggere i comprovantes** — fonte ausiliaria; contengono beneficiario reale e referencia; utili per definire criteri di regola.
4. **Collegare comprovantes all'estratto quando possibile** — per data, valore, beneficiario o referencia; il collegamento è supporto informativo, non classificazione autonoma.
5. **Applicare solo le regole definite dall'utente** — le regole possono usare dati da estratto, da comprovante o da entrambi.
6. **Produrre output strutturato**:
   - movimenti bancari classificati (con `conta_id` e `regra_id`);
   - movimenti esclusi (`XML` o `ESCLUDERE`) con motivo;
   - movimenti a revisar senza regola;
   - comprovantes collegati come evidenza/supporto.

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

## Gerarchia delle fonti

```
estratto conto   →  universo dei movimenti (righe da processare)
comprovantes     →  dettaglio e arricchimento (beneficiario/referencia)
regole utente    →  unica fonte decisionale per classificazione
```

Per costruire il criterio di una regola, l'utente può usare dati da:

1. **Estratto conto** — data, importo, descrizione movimento, beneficiario come appare nel conto;
2. **Comprovante** — `beneficiary_name` e `reference_id` quando più chiari della riga di estratto;
3. **Alias/keyword approvato** — testo libero confermato dall'utente.

Il criterio deve essere **dichiarato dall'utente**, non derivato automaticamente da nessuna fonte.
Le regole possono fare riferimento a campi di estratto, di comprovante o di entrambi.

---

## File operativi ufficiali

I file XLSX utilizzati dall'utente per leggere i dati bancari e definire le regole di classificazione sono:

| File | Contenuto | Uso |
|---|---|---|
| `bank_raw/Analisi_regole_registrazioni_banca_conto_corrente.xlsx` | Estratto conto — universo ufficiale dei movimenti bancari | Fonte primaria: identifica le righe da processare |
| `bank_raw/Analisi_regole_registrazioni_banca_comprovanti.xlsx` | Comprovantes — beneficiario reale e referencia per ogni pagamento | Fonte ausiliaria: aiuta l'utente a leggere e definire criteri di match |

Questi file non vengono modificati dal sistema. L'utente li consulta per decidere le regole; il sistema le applica.

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
