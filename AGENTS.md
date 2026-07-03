# AGENTS.md — Programma_registrazioni_contabili (Fantozzi)

Creato: 2026-07-03, dopo ricognizione tecnica completa read-only del progetto.

Questo file è la versione tool-agnostica delle istruzioni operative. Per il
contesto completo (identità, workflow, skill, filosofia del progetto) fare
riferimento a `CLAUDE.md` nella root del repository, che resta la fonte
principale. Qui sono raccolte solo le regole operative essenziali, valide per
qualunque agente AI che lavori su questo repository.

## Identità e principio fondamentale

Le decisioni contabili appartengono esclusivamente all'utente. Un agente che
lavora qui:
- non inferisce conti;
- non crea regole di classificazione autonomamente;
- non modifica YAML/config senza conferma esplicita;
- non fa commit/push senza ordine esplicito;
- se manca una regola, il movimento resta in revisione (mai un fallback silenzioso).

## Precheck obbligatorio

Prima di qualunque lavoro non banale su codice, dati, YAML, output o
documentazione:

```powershell
git status -sb
git branch --show-current
git diff --name-status
git diff --cached --name-status
```

- Mai lavorare sopra uno stato misto staged/unstaged senza prima classificarlo esplicitamente (tabella file → origine → stato → destino).
- Se ci sono file tracked modificati fuori commit, non trattarli come "ok" — vanno classificati.

## Branch

- Branch operativo attuale: **senza-incrementali** (verificare sempre con `git branch --show-current`, non assumere).
- Non usare `master` o `backup/con-incrementali` come base di lavoro: è la vecchia linea "con incrementali" (baseline + run incrementali retroattivi), superata dalla strategia attuale "full run da zero". Usarla solo su decisione esplicita dell'utente.
- `main` e `main-initial-backup` sono rami storici/di backup, non operativi.

## Commit e perimetro

- Commit piccoli, a perimetro stretto: un commit, un tema.
- Nessun `git add .`.
- Nessun output reale (run/report/PDF/HTML/Excel generati) committato — solo codice, config, test, documentazione.
- Distinguere sempre: codice (`src/`, `tests/`), audit/documentazione (`docs/`), output (`output/`, `reports/`, `runs/`).
- Nessun merge/push/reset/switch/checkout distruttivo senza via libera esplicita dell'utente.
- Prima di consolidare modifiche locali importanti o rischiose, considerare una patch di sicurezza (snapshot diff pre-modifica) — pattern già usato nel progetto (`output/_patch_safety_003D_...`).

## Fasi visuali (PDF / HTML / Excel / dashboard)

- Se una fase visuale già pushata viene bocciata, preferire un revert pulito piuttosto che patch correttive sopra codice bocciato (esempio nel progetto: "003E" `4871e12` → revert `2d69d3a`, confermato zero residui di codice dopo il revert).
- Per layout visuali, non "inventare": seguire il modello già approvato, oppure produrre un'anteprima e attendere conferma prima del commit.

## Contratti e invarianti

- Ogni decisione funzionale (regola di classificazione, mapping DRE, gate di validazione) deve essere verificabile in codice/config/test, non restare solo in chat.
- Prima di modificare logica contabile esistente, verificare i contratti già codificati nei test (vedi `docs/DECISIONI_CONTABILI.md` e `docs/GOVERNANCA_REGRAS_CONTABEIS.md` per le regole dichiarate dall'utente).

## Metodo per ogni modifica

1. Precheck Git.
2. Leggere YAML/documentazione ufficiale pertinente.
3. Validare la comprensione dello stato reale (non fidarsi di doc/memoria non verificate).
4. Mostrare la proposta prima di modificare.
5. Attendere conferma esplicita.
6. Modificare solo i file autorizzati dal perimetro concordato.
7. Eseguire audit/test pertinenti.
8. Mostrare il diff.
9. Commit/push solo su ordine esplicito.
