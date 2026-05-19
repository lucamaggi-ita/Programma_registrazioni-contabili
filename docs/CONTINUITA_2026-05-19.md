# Continuità lavoro — 2026-05-19

## Stato alla chiusura della sessione

Branch: `master` — allineato a `origin/master` (commit `352bbff`)
Working tree: pulito — nessun file modificato, nessun untracked rilevante.

## Cosa è stato fatto oggi

| Commit | Contenuto |
|---|---|
| `eb8b57d` | Linea DRE `13.02` creata; `11.01.005` e `11.01.006` rimappate da `13.01` a `13.02` |
| `f0186a2` | `.gitignore` protetto (`output/importacao_real/`, `reports/importacao_real/`, `pendenze_*.txt`); `plano_contas` aggiornato (`Internet` → `Internet e Telefonia`) |
| `c9a8d51` | `CLAUDE.md`, docs aggiornati, 3 script PDF aggiunti |
| `ee5ae52` | `Database_Dev/` aggiunta a `.gitignore` |
| `352bbff` | 10 file obsoleti eliminati (prompt iniziali, `regras_classificacao_despesas.yaml`, contratti Fase 9 chiuse) |

## Punti aperti

### P1 — Fallback `entra_dre: true` (decisione governance)

Il fallback `99.01` (A REVISAR) ha `entra_dre: true` nel `mapeamento_dre.yaml`.
Questo significa che prodotti non classificati entrano nel DRE.
Decisione non ancora presa dall'utente.

Opzioni:
- lasciare `entra_dre: true` → REVISAR visibili nel DRE (attuale)
- impostare `entra_dre: false` → REVISAR fuori dal DRE

### P2 — Fase 13W non materializzata

La diagnostica pendenze (13W) è stata eseguita in modalità read-only.
La run reale con `executar-classificacao-comprovantes-real` non è ancora stata eseguita.

Comando da eseguire:

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli executar-classificacao-comprovantes-real `
    --regras config\regras_classificacao_comprovantes.yaml `
    [altri argomenti della run reale]
```

Dopo la run: verificare quanti `REVISAR_SEM_REGRA` rimangono e decidere se aggiungere nuove regole o chiudere 13W.

## Prossima fase

**Fase 14 — Giornale e DRE completo**

Integrare tutte le fonti (XML + banca + manuali) nel pre-livro diário e DRE gestionale.

CLI già disponibili:
- `gerar-pre-livro-diario-teste`
- `gerar-dre-gerencial-ok-provisoria`
- `gerar-relatorios-gerenciais-ok-provisorios`

Prerequisito: chiudere P2 (Fase 13W materializzata).

## Comandi di avvio sessione

```powershell
cd "c:\Users\Recanto do Sossego\Documents\Luca\GitHub\Gestione_Ristorante\Programma_registrazioni_contabili"
git status -sb
git log --oneline -5
git rev-list --left-right --count origin/master...HEAD
.\.venv\Scripts\python.exe -m fantozzi.cli auditar-regras-comprovantes --regras config\regras_classificacao_comprovantes.yaml
```

Risultati attesi:
- `git status -sb` → pulito
- `rev-list` → `0 0`
- `auditar-regras-comprovantes` → `Regras: 30 | Erros: 0 | Alertas: 0`

## Riferimento rapido file operativi

| Cosa | File |
|---|---|
| Identità Fantozzi | `CLAUDE.md` |
| Regole comprovantes | `config/regras_classificacao_comprovantes.yaml` |
| Regole NF-e / XML | `config/regras_classificacao_produtos.yaml` |
| Piano dei conti | `config/plano_contas_gerencial.yaml` |
| Mapeamento DRE | `config/mapeamento_dre.yaml` |
| Roadmap | `docs/ROADMAP_MASTER.md` |
| Stato attuale | `docs/STATO_ATTUALE.md` |
| Decisioni contabili | `docs/DECISIONI_CONTABILI.md` |
| Governance | `docs/GOVERNANCA_REGRAS_CONTABEIS.md` |
