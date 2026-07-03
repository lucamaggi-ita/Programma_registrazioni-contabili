# ROADMAP_MASTER.md — Programma Registrazioni Contabili / Fantozzi

Aggiornato: 2026-05-19

## Stato generale

Il progetto Fantozzi / Programma_registrazioni_contabili ha completato le fasi di classificazione XML/NF-e (Fase 9) e la classificazione bancaria/comprovantes (Fase 13A–13W). Il sistema dispone di regole dichiarate, motore di classificazione, audit YAML e PDF documentali.

Nota importante di numerazione:

- Il progetto era già arrivato alla Fase 12.
- Siamo tornati alla Fase 9 perché era bloccante.
- La Fase 9 è stata sviluppata e chiusa fino alla Fase 9O.
- La fase bancaria è partita come Fase 13 (non Fase 10).
- La numerazione interna della Fase 13 ha seguito l'alfabeto: 13A → 13W.

---

## Roadmap consolidata

### Fasi chiuse — non riaprire

#### Fase 8 / 8.2 — Revisione assistita classificazione

Stato: **CONCLUSA**

Workbook di revisione assistita, alias e regole di classificazione. Import compatibile.

#### Fase 9 — Applicazione controllata decisioni approvate

Stato: **CONCLUSA** — chiusa e pushata fino a 9O.

- 9A: contratto definitivo documentato.
- 9B: engine decisionale puro da CSV revisionato.
- 9C: integrazione opzionale in pipeline via flag.
- 9F-A: bugfix parsing CSV.
- 9G: nuova run fresca da XML.
- 9H: audit read-only.
- 9K/9L: hardening CFOP 1949, estorno/recusa.
- 9M.1: bonificação recebida in DRE come riduzione costo.
- 9N: lançamentos manuais gerenciais via CSV.
- 9O: pacchetto contabile gerencial provvisorio.

#### Fase 12 — Classificazione / alias / revisione avanzata

Stato: **CONCLUSA** — consolidata prima del ritorno alla Fase 9.

Cautele stabilizzate: PATAXO, investimenti, refrigeradores/cervejeiras, bonificações, comodato.

---

### Fase 13 — Estratti bancari, comprovantes e classificazione gestionale

#### Fase 13A — Import estratti bancari

Stato: **CONCLUSA** — commit `46ff9fc`

Import normalizzato per CNPJ/conta. Convenzione bank_raw documentata (commit `de26d7b`).

#### Fase 13B — Matching bancario

Stato: **CONCLUSA** — commit `fab5202`

#### Fase 13D — Filtro entrate matching

Stato: **CONCLUSA** — commit `b20ea24`

#### Fase 13H — Classificazione comprovantes per regole dichiarate

Stato: **CONCLUSA** — commit `5930e4f`

Il motore applica `regras_classificacao_comprovantes.yaml` senza inferenza automatica.

#### Fase 13R — Neutralizzazione matching automatico

Stato: **CONCLUSA** — commit `c31b109`

Decisione definitiva: le regole dichiarate prevalgono; matching automatico disabilitato.
Non riaprire.

#### Fase 13S — Run classificazione reale

Stato: **CONCLUSA** — commit `e7492d4`

Esecuzione completa su dati reali con audit trail.

#### Fase 13T — Pacchetto revisione (XLSX)

Stato: **CLI disponibile** (`gerar-pacote-revisao-comprovantes`)

Esecuzione on-demand, non obbligatoria prima di 13W.

#### Fase 13U — Audit e fix regole YAML

Stato: **CONCLUSA** — commits `3b699ea` → `9b3571c`

ID stabili, pulizia strutturale, correzioni entra_dre/dre_linha_id.

#### Fase 13V — Esclusioni controllate XML / ESCLUDERE

Stato: **CONCLUSA** — commit `9b3571c`

#### Fase 13W — Diagnostica pendenze REVISAR

Stato: **PARZIALE** — audit read-only eseguito; run `executar-classificacao-comprovantes-real` non ancora materializzata.

Blocco prima della prossima fase.

---

### Correzioni strutturali DRE (sessione 2026-05-19)

- Audit 13W read-only: 4 incoerenze corrette in `regras_classificacao_comprovantes.yaml` (commits precedenti).
- Creazione linea DRE `13.02` (Distribuicao e Antecipacao de Lucros) in `mapeamento_dre.yaml`.
- Remapping `11.01.005` e `11.01.006` da `13.01` a `13.02` — commit `eb8b57d`.
- P1 aperto: fallback `99.01` ha `entra_dre: true` — decisione governance non ancora presa.

---

### Prossima macrofase

#### Chiusura 13W e integrazione DRE completa

Obiettivo:
materializzare la diagnostica 13W con run reale, poi integrare tutte le fonti nel giornale e DRE.

Passo obbligatorio prima di qualsiasi nuova fase:

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli executar-classificacao-comprovantes-real ...
```

Dopo 13W, il percorso naturale è:

```
Comprovantes classificati (13W chiusa)
  ↓
Lançamentos gestionali da comprovantes → pre-livro diário
  ↓
DRE gestionale completo (XML + banca + manuali)
```

La CLI ha già: `gerar-pre-livro-diario-teste`, `gerar-dre-gerencial-ok-provisoria`.

La prossima macrofase numerata sarà **Fase 14** (integrazione giornale e DRE completo).

---

## Regole operative permanenti

1. Mai committare dati reali, XML, estratti, comprovantes, output o reports reali.
2. Evitare `git add .`.
3. Prima di commit:
   - `git status -sb`
   - `git diff --check`
   - `git diff --name-status`
4. Prima di push:
   - verificare `git rev-list --left-right --count origin/<branch-attivo>...HEAD` (branch attivo corrente: `senza-incrementali`, non `master`)
   - verificare che il diff contenga solo file attesi.
5. Quando cambiano regole/parser/logica fiscale/contratto contabile:
   - generare nuova run fresca da XML;
   - usare run vecchie solo per confronto/audit.
6. Non forzare run vecchie con CSV obsoleti.
7. I casi REVISAR non devono entrare in DRE o report OK salvo regole esplicite e testate.
8. Audit obbligatori prima di ogni nuova fase:
   - `git status -sb`
   - `auditar-regras-comprovantes` → 0 erros, 0 alertas
   - validazione YAML mapeamento_dre e plano_contas
   - verifica .gitignore copre output/importacao_real/ e reports/importacao_real/
