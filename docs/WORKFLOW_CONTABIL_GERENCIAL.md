# Workflow contabile gestionale

**Progetto:** Fantozzi / Programma_registrazioni_contabili
**Data di adozione:** 2026-05-19

---

## Obiettivo

Il progetto genera **registrazioni contabili gestionali**, non fiscali, con obiettivo un **DRE attendibile e tracciabile**.

La Fase 13 (classificazione comprovantes) non è un progetto isolato. È un tassello di un workflow completo che va dalle fonti ai lancamentos, al giornale, al DRE.

---

## Schema generale

```
Fonti
  ↓
Regole definite dall'utente
  ↓
Lançamentos gestionali
  ↓
Pre-livro diário (giornale provvisorio)
  ↓
DRE gestionale
```

---

## Fonti principali

Le fonti di dati contabili del progetto, in ordine di priorità per completezza e affidabilità:

| # | Fonte | Contenuto | Fase |
|---|---|---|---|
| 1 | XML / NF-e | Fatture elettroniche di acquisto e vendita | 9, XML |
| 2 | Estratti conto | Movimenti bancari aggregati | 13 |
| 3 | Comprovantes / Consulta Pagamentos | Pagamenti bancari con beneficiario e referencia | 13 |
| 4 | Lançamentos manuais gestionali | Movimenti inseriti manualmente dall'utente | manuale |
| 5 | Operazioni di cassa / denaro contante | Movimenti non tracciati bancariamente | manuale |
| 6 | Correzioni manuali controllate | Rettifiche su righe automatiche | manuale |

---

## Regole del workflow

### Le fonti non si sostituiscono

Un comprovante non sostituisce un XML. Un estratto conto non sostituisce un comprovante.
Fonti diverse coprono transazioni diverse o forniscono viste complementari della stessa transazione.

### Il costo coperto da XML non va duplicato da comprovante

Se un pagamento è già coperto da XML/NF-e (fornitore con fattura), il relativo comprovante deve essere **escluso** dalla classificazione gestionale con `acao: XML`.

Questo evita la doppia imputazione della stessa spesa.

### Il lançamento gestionale nasce dalla regola, non dall'analisi

Il sistema produce un lançamento solo quando esiste una regola dichiarata dall'utente.
Non produce lançamentos per inferenza automatica.

### Tracciabilità obbligatoria

Ogni lançamento gestionale deve essere tracciabile fino alla fonte:

- XML → numero NF-e, CNPJ emittente, data;
- comprovante → `payment_id`, `beneficiario_normalizado`, `regra_id` applicata;
- lançamento manuale → data inserimento, utente, motivo.

---

## Risultato atteso del progetto completo

Il progetto non termina con "classificare i comprovantes".

Il percorso completo è:

```
fonti → regole → lançamentos → giornale → DRE
```

Il DRE gestionale deve essere:

- **attendibile**: ogni riga DRE ha una fonte tracciabile;
- **controllabile**: l'utente può rivedere, correggere o escludere qualsiasi riga;
- **non duplicato**: ogni costo appare una sola volta, dalla fonte corretta;
- **auditabile**: ogni run produce manifest e log con stato dettagliato.

---

## Diritti permanenti dell'utente

L'utente deve sempre poter:

| Azione | Dove |
|---|---|
| Inserire movimenti manuali | Lançamentos manuais gestionali |
| Correggere righe automatiche | Planilha di revisione |
| Escludere righe (XML, ESCLUDERE) | `exclusoes:` nel YAML comprovantes |
| Approvare nuove regole di classificazione | `regras:` nel YAML comprovantes |
| Rieseguire la classificazione su dati aggiornati | CLI `executar-classificacao-comprovantes-real` |
| Verificare lo stato del DRE provvisorio | Output DRE gestionale |
| Mantenere tracciabilità completa | Manifest e report di ogni run |

---

## Cosa non fare nelle fasi automatiche

Le fasi automatiche del sistema non devono mai:

- produrre lançamentos senza regola dichiarata;
- modificare DRE definitivo senza approvazione;
- correggere automaticamente XML o NF-e;
- rimuovere esclusioni già definite;
- reinterpretare regole esistenti;
- usare l'estratto conto come classificatore.

---

## Fasi del progetto e loro scope

| Fase | Nome | Scope |
|---|---|---|
| 9 | Decisioni contabili Fase 9 | Regole speciali, operazioni straordinarie |
| 13H | Classificazione comprovantes | Applicazione regole YAML su comprovantes |
| 13S | Run classificazione reale | Esecuzione completa su dati reali |
| 13T | Pacchetto revisione | XLSX per revisione manuale operatore |
| 13U | Audit e fix regole | Pulizia YAML, ID stabili, test |
| 13V | Esclusioni controllate | Aggiunta `exclusoes:` XML e ESCLUDERE |
| 13W | Diagnostica pendenze | Analisi `REVISAR_SEM_REGRA` rimasti |

Le fasi future continuano questo percorso verso giornale e DRE attendibile.

---

## File operativi ufficiali

I file XLSX utilizzati dall'utente per leggere i dati bancari e definire le regole sono:

| File | Contenuto | Fonte |
|---|---|---|
| `bank_raw/Analisi_regole_registrazioni_banca_conto_corrente.xlsx` | Estratto conto — universo ufficiale dei movimenti bancari in uscita | Primaria |
| `bank_raw/Analisi_regole_registrazioni_banca_comprovanti.xlsx` | Comprovantes — beneficiario reale e referencia per ogni pagamento | Ausiliaria |

Il sistema non modifica questi file. Sono consultati dall'utente per decidere le regole; il sistema le applica.

---

## Riferimento tecnico

| Componente | Path |
|---|---|
| Regole comprovantes | `config/regras_classificacao_comprovantes.yaml` |
| Motore classificazione | `src/fantozzi/banking/payment_rules.py` |
| Run classificazione | `src/fantozzi/banking/payment_classification_run.py` |
| DRE gestionale provvisorio | `src/fantozzi/dre_gerencial_provisoria.py` |
| Pre-livro diário | `src/fantozzi/pre_livro_diario.py` |
| Lançamentos manuais | `src/fantozzi/lancamentos_manuais_gerenciais.py` |
| CLI principale | `src/fantozzi/cli.py` |
