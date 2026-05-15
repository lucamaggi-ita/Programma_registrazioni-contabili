# FANTOZZI — Fase 9: Contratto Definitivo

**Data redazione:** 2026-05-15
**Stato:** DESIGN / DOCUMENTAZIONE — nessun codice funzionale implementato
**Branch:** master

---

## 0. Stato attuale della Fase 9

La Fase 9 **non è mai stata implementata**.
Esiste esclusivamente come blocco intenzionale nel codice:

```python
# reprocessamento_pipeline.py
if modo == "definitivo":
    raise ReprocessamentoPipelineError(MODO_DEFINITIVO_BLOQUEADO)
```

```python
# manifest al termine di ogni run
"classificacao": {
    "fase9_executada": False,
    "regras_aplicadas": 0,
    "nota": "Revisao contabil pendente. Nenhuma regra aprovada aplicada."
}
```

Il messaggio `"[INFO] Fase 9 nao executada; relatorios permanecem provisorios."` viene
stampato dalla CLI al termine di ogni reprocessamento.

Tutti i relatórios correnti sono **provisórios**.
Nessuna regra approvata è stata applicata automaticamente.
Il modo definitivo è bloccato.

---

## 1. Input obbligatori della Fase 9

| File | Origine | Note |
|------|---------|------|
| `revisao_classificacao_real.csv` | Step `DIAGNOSTICOS_CONCLUIDOS` | CSV con candidati classificazione, uno per CNPJ |
| `pacote_revisao_contabil_fase8_1.xlsx` | Step `PACOTE_REVISAO_CONTABIL_GERADO` | Workbook di revisão contábil con aba `Revisao` compilata manualmente |
| `revisao_classificacao_revisada.csv` | Output di `importar-planilha-revisao-classificacao` | CSV risultante dall'importazione del workbook revisionato |
| `candidatos_classificacao_real.txt` | Step `DIAGNOSTICOS_CONCLUIDOS` | Report testuale candidati |
| `manifest_run.json` | `run_manager` | Stato corrente della run |
| `config/regras_classificacao_produtos.yaml` | Config | Regole prodotti esistenti (solo lettura) |
| `config/plano_contas_gerencial.yaml` | Config | Piano dei conti (riferimento per validazione conta) |

### Trigger operativo

Il file trigger operativo è **`revisao_classificacao_revisada.csv`** — il CSV prodotto
dall'importazione della planilha workbook dopo revisione manuale.
Senza quel file, la Fase 9 non può iniziare.

Flusso di preparazione obbligatorio prima della Fase 9:

```
gerar-planilha-revisao-classificacao   → pacote_revisao_contabil_fase8_1.xlsx
  [revisione manuale del workbook]
importar-planilha-revisao-classificacao → revisao_classificacao_revisada.csv
gerar-proposta-classificacao            → proposta YAML (solo visualizzazione)
  [validazione umana della proposta]
  [SOLO DOPO approvazione] → Fase 9B
```

---

## 2. Azioni manuali ammesse

Definite in `classification_review_validation.py`:

```python
ALLOWED_ACTIONS = {
    "ANALISAR_COM_CONTADOR",
    "CLASSIFICAR",
    "CRIAR_ALIAS",
    "CRIAR_REGRA",
    "IGNORAR",
    "MANTER_REVISAO",
}
```

Descrizione di ciascuna:

| Azione | Significato |
|--------|-------------|
| `CLASSIFICAR` | Classificazione diretta su una conta esistente |
| `CRIAR_ALIAS` | Proporre un alias testuale per una conta esistente (prodotto specifico) |
| `CRIAR_REGRA` | Proporre una regola basata su palavras-chave (famiglia di prodotti) |
| `IGNORAR` | Ignorare il candidato (non classificare, non promuovere) |
| `MANTER_REVISAO` | Mantenere REVISAR per revisione futura |
| `ANALISAR_COM_CONTADOR` | Caso fiscale/contabile dubbio, demandare al commercialista |

---

## 3. Quali azioni possono promuovere REVISAR → OK

Solo le tre azioni seguenti possono portare un lancamento da `REVISAR` a `OK`:

```
CLASSIFICAR
CRIAR_ALIAS
CRIAR_REGRA
```

Queste azioni sono già identificate nel codice come quelle che "propongono classificazione reale":

```python
# classification_review_workbook.py
"sim" if action in {"CLASSIFICAR", "CRIAR_ALIAS", "CRIAR_REGRA"} else "nao"

# classification_review_workbook_importer.py
ACOES_COM_CONTA_OBRIGATORIA = {"CLASSIFICAR", "CRIAR_ALIAS", "CRIAR_REGRA"}
```

### Condizioni minime per la promozione (tutte obbligatorie)

1. **`conta_id_sugerida` deve esistere** nel `plano_contas_gerencial.yaml`
   (verificato da `classification_review_validation.py` → `chart.get_by_id(...)`)
2. **`conta_nome_sugerida` deve essere coerente** con `conta_id_sugerida`
   (verificato da `classification_review_workbook_importer.py`)
3. **Per `CRIAR_REGRA`**: `palavras_chave_sugeridas` obbligatorio, non solo generiche
   (verificato da `_validate_keywords()` in `classification_review_workbook_importer.py`)
4. **Nessun alerta non risolto** sul lancamento candidato
5. **L'azione deve riferirsi al candidato corretto** (corrispondenza descrizione/fornecedor)
6. **Nessun blocco fiscal/contabile** attivo sul lancamento
7. **`descricao`/`alias` non può essere generica** — le GENERIC_KEYWORDS sono bloccanti
   (da `classification_review_workbook.GENERIC_KEYWORDS`)

---

## 4. Azioni che NON possono promuovere a OK — devono restare REVISAR

Le seguenti azioni o condizioni impediscono permanentemente la promozione:

| Caso | Motivo |
|------|--------|
| `MANTER_REVISAO` | Decisione esplicita di non classificare ancora |
| `ANALISAR_COM_CONTADOR` | Richiede validazione contabile/fiscale esterna |
| `IGNORAR` | Nessuna classificazione proposta |
| Qualsiasi alerta non risolto | Incertezza sull'imputazione |
| `conta_id_sugerida` assente o non valida | Manca il presupposto contabile |
| Palavras-chave solo generiche per `CRIAR_REGRA` | Rischio di over-matching futuro |
| `conta_nome_sugerida` incoerente | Errore di compilazione del workbook |
| COMODATO | Vedere §5 |
| CFOP 1949 / NF 359680 | Vedere §5 |

---

## 5. Trattamento dei casi speciali

### 5.1 BONIFICACAO_RECEBIDA normale

- Conta: `2.05.003`
- Categoria gerenciale: `REDUCAO_CUSTO`
- Status: già validata in Fase 12J/12K come operazione normale
- **Non deve essere confusa** con CFOP 1949 atipico (NF 359680)
- Entra in DRE come riduzione di costo **solo se `status_revisao = OK` e senza alertas**
- Qualsiasi bonificação con `status_revisao = REVISAR`, alerta attivo, CFOP atipico
  o dubbio fiscale **resta soggetta ai gate della Fase 9** e non entra in DRE
  finché non è promossa attraverso il processo di revisione ordinario

### 5.2 COMODATO

- Status fisso: `REVISAR` / `FORA_DRE`
- **Non entra mai** in costo, investimento o DRE
- Non può essere promosso a OK automaticamente o manualmente in Fase 9
- Richiede decisione contabile separata fuori dalla pipeline standard

### 5.3 CFOP 1949 / NF 359680

- Caso raro e atipico
- **Non deve essere hardcodato** come rifiuto automatico
- **Non deve essere hardcodato** come vendita
- Resta sotto controllo manuale pre-Fase 9
- Non promuovibile a OK senza decisione manuale contabile esplicita e documentata
- La pipeline deve segnalarlo come `ANALISAR_COM_CONTADOR` o `REVISAR` finché
  non esiste una decisione formale registrata con audit trail

---

## 6. Write-back previsto per la Fase 9B

La Fase 9A non implementa write-back.
Questo paragrafo descrive il meccanismo da progettare in Fase 9B.

### 6.1 File aggiornato

**`revisao_classificacao_revisada.csv` è input immutabile.**
La Fase 9 lo legge ma **non lo modifica mai**.

Il write-back produce un nuovo file di output separato:

```
output/<run_id>/lancamentos_aplicados_fase9.csv
```

I file `config/` (YAML di configurazione) non vengono mai toccati automaticamente:
le regole/alias approvate vengono scritte nella proposta YAML separata
(`output/<run_id>/proposta_regras_fase9.yaml`) e applicate al config reale
solo con autorizzazione manuale esplicita dell'utente.

### 6.2 Campi che cambiano per ogni lancamento promosso

| Campo | Prima | Dopo |
|-------|-------|------|
| `status_revisao` | `REVISAR` | `OK` |
| `acao_aplicada` | vuoto | azione manuale approvata (es. `CRIAR_ALIAS`) |
| `conta_id_applicata` | vuoto/precedente | conta approvata |
| `regra_alias_applicata` | vuoto | nome alias/regra applicata |
| `data_aplicacao_fase9` | vuoto | timestamp ISO 8601 |
| `run_id_fase9` | vuoto | run_id della Fase 9 |

### 6.3 Lancamenti che restano REVISAR

Ogni lancamento che resta `REVISAR` dopo la Fase 9 deve avere registrato:

- `motivo_revisao_persistente`: motivo esplicito (es. `MANTER_REVISAO`, `conta_invalida`, `alerta_nao_resolvido`)

### 6.4 Regole YAML

Le regole/alias approvate vengono scritte nella proposta YAML separata:
`output/<run_id>/proposta_regras_fase9.yaml`

**Il file `config/regras_classificacao_produtos.yaml` non viene mai modificato
automaticamente dalla Fase 9.**
L'applicazione al config reale è un passo manuale successivo alla Fase 9B,
autorizzato esplicitamente dall'utente.

---

## 7. Manifest e `fase9_executada`

### 7.1 Struttura attuale (stato provisório)

```json
{
  "classificacao": {
    "fase9_executada": false,
    "regras_aplicadas": 0,
    "nota": "Revisao contabil pendente. Nenhuma regra aprovada aplicada."
  }
}
```

### 7.2 Struttura attesa dopo Fase 9B completata con successo

```json
{
  "classificacao": {
    "fase9_executada": true,
    "regras_aplicadas": 12,
    "lancamentos_promovidos": 34,
    "lancamentos_revisar_persistentes": 7,
    "run_id_fase9": "20260515_143022",
    "data_fase9": "2026-05-15T14:30:22",
    "audit_trail": "reports/<run_id>/relatorio_fase9_aplicacao_regras.csv"
  }
}
```

### 7.3 Invariante critica

`fase9_executada: true` **da sola non è sufficiente** per sbloccare il modo definitivo.
Devono essere soddisfatte tutte le condizioni del §9.

---

## 8. Idempotenza

La Fase 9 deve essere idempotente. Eseguirla due volte sullo stesso input deve
produrre lo stesso risultato. In pratica:

- Non duplica lancamentos in `Lancamentos_OK`
- Non duplica alias nel YAML di proposta
- Non corrompe il CSV di output
- Non promuove nuovi lancamentos non coperti da decisione manuale registrata
- Se tutti i lancamenti promovibili sono già `OK`, il secondo run riporta
  `regras_aplicadas = 0` (nessuna nuova promozione) senza errore
- Il manifest viene aggiornato solo **dopo** il completamento senza errori bloccanti;
  un run parzialmente fallito non lascia `fase9_executada: true`

---

## 9. Audit trail obbligatorio

Ogni esecuzione della Fase 9 deve produrre due file:

```
reports/<run_id>/relatorio_fase9_aplicacao_regras.txt
reports/<run_id>/relatorio_fase9_aplicacao_regras.csv
```

### 9.1 Contenuto minimo del CSV

| Colonna | Descrizione |
|---------|-------------|
| `run_id` | Identificatore della run Fase 9 |
| `data_hora` | Timestamp ISO 8601 |
| `tipo_candidato` | Tipo candidato (alias, regra, diretto) |
| `descricao_normalizada` | Descrizione normalizzata del lancamento |
| `fornecedor` | Fornitore |
| `acao_manual` | Azione manuale registrata nel workbook |
| `conta_id_proposta` | Conta suggerita nel workbook |
| `conta_id_applicata` | Conta effettivamente applicata |
| `regra_alias_applicata` | Nome alias/regra applicata (se pertinente) |
| `status_antes` | `REVISAR` |
| `status_depois` | `OK` oppure `REVISAR` |
| `motivo_revisar_persistente` | Motivo se status_depois = `REVISAR` |
| `warnings` | Warning non bloccanti |

---

## 10. Condizioni per sbloccare il modo definitivo

Il modo definitivo (`modo = "definitivo"` nel pipeline) può essere sbloccato
**solo quando tutte** le seguenti condizioni sono soddisfatte:

| Condizione | Verifica |
|------------|---------|
| `fase9_executada = True` nel manifest | Manifest JSON |
| `regras_aplicadas >= 0` (valore coerente) | Manifest JSON |
| Audit trail esiste e non è vuoto | File `relatorio_fase9_aplicacao_regras.csv` |
| Nessun errore bloccante nel manifest | Campo `erros` assente o vuoto |
| Nessun `REVISAR` presente in `Lancamentos_OK` | Audit sheet nel pre_livro_diario |
| COMODATO escluso da `Lancamentos_OK` | Verifica operazioni speciali |
| `ANALISAR_COM_CONTADOR` escluso da `Lancamentos_OK` | Status residuo |
| `MANTER_REVISAO` escluso da `Lancamentos_OK` | Status residuo |
| Proposta YAML revisionata manualmente | File `proposta_regras_fase9.yaml` |
| Decisione esplicita dell'utente | Autorizzazione non automatica |

Il codice attuale in `reprocessamento_pipeline.py:140-141` implementa il blocco
e **non deve essere rimosso** finché tutte le condizioni sopra non sono verificabili
automaticamente dalla pipeline.

---

## 11. Piano micro-fasi per la Fase 9B

| Sotto-fase | Descrizione | Output |
|------------|-------------|--------|
| **9B.1** | Funzioni pure: leggere decisioni approvate dal CSV revisado e calcolare set di promozioni | Solo dati in memoria, nessuna scrittura |
| **9B.2** | Write-back controllato su copia/output — non tocca `config/` reale | `output/<run_id>/lancamentos_promovidos.csv`, `proposta_regras_fase9.yaml` |
| **9B.3** | Audit trail Fase 9 (TXT + CSV) | `relatorio_fase9_aplicacao_regras.*` |
| **9B.4** | Integrazione pipeline: nuovo step `FASE9_APLICADA` nel manifest | Step aggiunto a `reprocessamento_pipeline.py` |
| **9B.5** | Gate sblocco modo definitivo: verifica automatica di tutte le condizioni del §9 | Rimozione condizionata del blocco `modo == "definitivo"` |

---

## 12. Test obbligatori per la Fase 9B

Questa sezione definisce la matrice dei test che **devono** essere implementati
prima di qualsiasi merge della Fase 9B in produzione.

### 12.1 Promozione corretta

- `CLASSIFICAR` valido promuove `REVISAR → OK`
- `CRIAR_ALIAS` valido promuove solo descrizione esatta/sicura
- `CRIAR_REGRA` valido promuove solo quando palavras-chave sono sicure e non generiche

### 12.2 Blocchi obbligatori — nessuna promozione possibile

- `ANALISAR_COM_CONTADOR` resta `REVISAR`
- `MANTER_REVISAO` resta `REVISAR`
- `IGNORAR` resta `REVISAR` (nessuna classificazione)
- COMODATO resta `REVISAR` / `FORA_DRE`
- NF 359680 / CFOP 1949 resta `REVISAR` salvo decisione manuale esplicita futura
- conta inesistente nel plano_contas blocca promozione
- alias generico blocca promozione (`GENERIC_KEYWORDS`)
- regra con sole palavras-chave generiche blocca promozione
- alertas non risolti bloccano promozione
- `conta_nome_sugerida` incoerente con `conta_id_sugerida` blocca promozione

### 12.3 Idempotenza

- Fase 9 eseguita due volte produce lo stesso risultato
- Nessun duplicato in `Lancamentos_OK`
- Nessun duplicato nel YAML di proposta

### 12.4 Audit e manifest

- Fase 9 produce audit trail (TXT + CSV) non vuoto
- Manifest aggiornato solo dopo successo completo
- `fase9_executada = True` solo se nessun errore bloccante
- Run parzialmente fallito non lascia `fase9_executada: True`

### 12.5 Gate definitivo

- Modo definitivo resta bloccato se `fase9_executada = False`
- Modo definitivo resta bloccato se audit trail manca
- Nessun `REVISAR` appare in `Lancamentos_OK` dopo Fase 9 completata
- BONIFICACAO_RECEBIDA normale classificata correttamente come `2.05.003 / REDUCAO_CUSTO`
- BONIFICACAO_RECEBIDA con CFOP 1949 non promossa automaticamente

---

## Appendice A — Struttura della pipeline corrente (riferimento)

```
CRIACAO_DA_RUN
IMPORTACAO_XML_CONCLUIDA
DIAGNOSTICOS_CONCLUIDOS           ← produce revisao_classificacao_real.csv
OPERACOES_ESPECIAIS_AUDITADAS
PRE_LIVRO_GERADO
AUDITORIA_PRE_LIVRO_CONCLUIDA
PACOTE_REVISAO_CONTABIL_GERADO    ← produce pacote_revisao_contabil_fase8_1.xlsx
RELATORIOS_OK_GERADOS
AUDITORIA_RELATORIOS_OK_CONCLUIDA
DRE_OK_PROVISORIA_GERADA
FINALIZACAO_PROVISORIA            ← fase9_executada = False, modo definitivo bloccato
```

La Fase 9B aggiunge un checkpoint `FASE9_APLICADA` tra
`PACOTE_REVISAO_CONTABIL_GERADO` e `RELATORIOS_OK_GERADOS`.

---

## Appendice B — File del codice di riferimento per la Fase 9B

| Simbolo / Costante | File | Nota |
|--------------------|------|------|
| `MODO_DEFINITIVO_BLOQUEADO` | `reprocessamento_pipeline.py:43` | Messaggio blocco definitivo |
| `AVISO_FASE9_BLOQUEADA` | `reprocessamento_pipeline.py:46` | Avviso aviso nel manifest |
| `ETAPAS_PREVISTAS` | `reprocessamento_pipeline.py:51` | Aggiungere `FASE9_APLICADA` |
| `ALLOWED_ACTIONS` | `classification_review_validation.py:16` | Azioni valide |
| `ACOES_COM_CONTA_OBRIGATORIA` | `classification_review_workbook_importer.py:20` | Gate conta |
| `ACAO_COM_PALAVRAS_OBRIGATORIAS` | `classification_review_workbook_importer.py:21` | Gate palavras-chave |
| `GENERIC_KEYWORDS` | `classification_review_workbook.py` | Keywords generiche bloccanti |
| `_validate_keywords()` | `classification_review_workbook_importer.py:317` | Validazione palavras |
| `LANCAMENTOS_OK_SHEET` | `pre_livro_diario.py:20` | Sheet Lancamentos_OK |
| `LANCAMENTOS_REVISAR_SHEET` | `pre_livro_diario.py:21` | Sheet Lancamentos_REVISAR |
| `gerar_pacote_revisao_contabil()` | `classification_review_package.py:77` | Genera pacchetto revisão |
| `importar_planilha_revisao_classificacao()` | `classification_review_workbook_importer.py:57` | Importa workbook revisionato |
