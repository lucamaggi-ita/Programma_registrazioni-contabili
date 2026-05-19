# Progetto: Regole YAML autonome per classificazione comprovantes

Data: 2026-05-18

## 1. Obiettivo

L'obiettivo e' rendere il sistema di classificazione dei comprovantes stabile, controllato e indipendente da qualsiasi planilha esterna durante l'esecuzione operativa.

Il programma deve classificare i comprovantes usando solo:

```text
config/regras_classificacao_comprovantes.yaml
```

La planilha manuale deve essere usata solo quando l'operatore decide di aggiornare il YAML. Dopo l'aggiornamento, il YAML deve essere autonomo: se la planilha viene eliminata, riscritta, riordinata o sostituita, le regole gia' esistenti devono continuare a funzionare.

Principio centrale:

```text
La planilha e' una fonte temporanea di decisioni umane.
Il YAML e' il contratto operativo del classificatore.
```

## 2. Decisioni progettuali confermate

### 2.1 Nessun riferimento alla planilha nel YAML operativo

Il YAML non deve contenere campi come:

```text
origem_linha_planilha
linha_planilha
numero_linha
nome_arquivo_planilha
coluna_origem
```

Questi dati possono comparire nei report di importazione, ma non nel YAML operativo.

Motivo:

```text
La riga 2 di una planilha oggi puo' essere diversa dalla riga 2 di una planilha futura.
La posizione nella planilha non e' identita' della regola.
```

### 2.2 Conflitti isolati

Se una nuova planilha propone un criterio gia' esistente ma con una conta diversa, il sistema deve bloccare solo quella candidata.

Il resto dell'importazione deve continuare.

Esempio:

```text
Criterio gia' esistente:
beneficiario = META PLATFORMS
conta = 9.01.001

Nuova candidata:
beneficiario = META PLATFORMS
conta = 8.02.004

Azione:
bloccare solo questa candidata come conflito.
```

### 2.3 Regole vecchie persistenti

Una nuova planilha non deve cancellare o disattivare regole esistenti.

Se una regola gia' presente nel YAML non appare nella nuova planilha, resta attiva.

La rimozione o disattivazione di una regola deve essere un'azione esplicita dell'operatore.

## 3. Architettura proposta

Il sistema viene diviso in tre livelli.

### 3.1 Fonte manuale temporanea

La planilha e' un documento umano di revisione.

Puo':

- essere eliminata;
- essere riscritta;
- essere riordinata;
- avere nuove colonne;
- avere righe diverse;
- essere usata una sola volta;
- essere riutilizzata in futuro per aggiornare il YAML.

Non puo':

- essere letta dal classificatore operativo;
- definire l'identita' stabile di una regola;
- sovrascrivere regole esistenti solo per posizione di riga.

### 3.2 Importatore e validatore di regole

L'importatore viene eseguito solo quando l'operatore vuole aggiornare il YAML.

Comando progettuale:

```text
atualizar-regras-comprovantes --planilha caminho/arquivo.xlsx
```

Compiti:

1. leggere la planilha;
2. estrarre solo campi approvati;
3. generare candidate di regole autonome;
4. validare ogni candidata;
5. identificare duplicati e conflitti;
6. aggiornare il YAML solo con regole valide;
7. produrre un report di importazione.

### 3.3 Classificatore operativo

Il classificatore deve leggere solo:

```text
config/regras_classificacao_comprovantes.yaml
```

Non deve conoscere:

- planilha;
- righe della planilha;
- colonne della planilha;
- file di origine;
- audit di importazione.

Il classificatore deve applicare solo i criteri presenti nel YAML.

## 4. Schema YAML consigliato

Schema proposto:

```yaml
versao_schema: 1

regras:
  - regra_id: REGRA_9_01_001_META_PLATFORMS
    ativa: true

    criterios:
      beneficiario_favorecido_contem_algum:
        - META PLATFORMS
        - FACEBOOK
        - INSTAGRAM

    destino:
      conta_id: 9.01.001
      conta_nome: Publicidade e Propaganda
      dre_linha_id: '11.01'
      entra_dre: true
```

### 4.1 Campi obbligatori

Ogni regola positiva deve avere:

```text
regra_id
ativa
criterios
destino.conta_id
destino.conta_nome
destino.dre_linha_id
destino.entra_dre
```

### 4.2 Criteri validi

Almeno uno di questi criteri deve essere presente:

```text
beneficiario_favorecido_contem_algum
referencia_empresa_contem_algum
alias_ou_keyword_contem_algum
```

Il nome esatto dei campi dovra' essere allineato al motore esistente, ma il principio e' obbligatorio:

```text
nessuna regola positiva puo' esistere senza almeno un criterio reale di match.
```

### 4.3 Campi vietati nel YAML operativo

Nel YAML operativo non devono comparire:

```text
origem_linha_planilha
linha_planilha
numero_linha
arquivo_origem
coluna_origem
status_classificacao_da_planilha
motivo_revisao_da_planilha
```

Questi campi sono ammessi solo nei report, non nelle regole operative.

## 5. Identita' stabile della regola

L'identita' della regola deve derivare dal contenuto approvato, non dalla posizione nella planilha.

Esempi accettabili:

```text
REGRA_9_01_001_META_PLATFORMS
REGRA_4_01_003_SISPAR_8916087
REGRA_9_02_001_EUDES_BRAGA_NASCIMENTO
```

Il sistema dovrebbe costruire un identificatore normalizzato usando:

```text
conta_id
tipo del criterio principale
valore normalizzato del criterio principale
```

Esempio:

```text
conta_id: 9.01.001
criterio: beneficiario_favorecido
valor: META PLATFORMS

regra_id:
REGRA_9_01_001_META_PLATFORMS
```

### 5.1 Chiave logica interna

Per deduplicare e rilevare conflitti, oltre al `regra_id`, il sistema deve calcolare una chiave logica interna.

Esempio:

```text
beneficiario_favorecido:meta_platforms
```

Questa chiave serve a rispondere a tre casi:

```text
stesso criterio + stessa conta = gia' esistente
stesso criterio + conta diversa = conflito isolato
criterio nuovo + conta valida = nuova regola
```

## 6. Regole di importazione da planilha

L'importatore deve leggere solo colonne approvate.

### 6.1 Colonne usabili come destinazione contabile

```text
CONTA_ID_APROVADA
CONTA_NOME_APROVADA
dre_linha_id
entra_dre
```

Se `dre_linha_id` e `entra_dre` non sono affidabili nella planilha, devono essere risolti dal piano/mapeamento contabile esistente a partire da `CONTA_ID_APROVADA`.

### 6.2 Colonne usabili come criteri reali

Usare solo se compilate e approvate:

```text
BENEFICIARIO
beneficiario_normalizado
beneficiario_favorecido
REFERENCIA
referencia_empresa
ALIAS_O_KEYWORD_APROVADO
TIPO_REGRA_SUGERIDA
```

### 6.3 Colonne ignorate per la generazione operativa

Non devono generare criteri, identita' o campi YAML:

```text
origem_linha_planilha
numero_linha
linha_planilha
observacoes
status_classificacao
motivo_revisao
contas_candidatas
regras_candidatas
```

Possono essere usate solo per il report di importazione.

## 7. Comportamento dell'importatore

Per ogni riga della planilha:

1. verificare se esiste una conta approvata;
2. verificare se esiste almeno un criterio reale approvato;
3. se manca la conta, mandare la riga a report come `sem_conta_aprovada`;
4. se manca il criterio, mandare la riga a report come `sem_criterio_real`;
5. se conta e criterio esistono, creare una candidata;
6. calcolare la chiave logica del criterio;
7. confrontare la candidata con il YAML esistente;
8. aggiungere la candidata se e' nuova e valida;
9. ignorarla se e' gia' esistente con stessa destinazione;
10. bloccarla come conflitto se lo stesso criterio punta a conta diversa.

## 8. Casi di aggiornamento

### 8.1 Nuova planilha con nuova riga 2 diversa dalla vecchia riga 2

Comportamento corretto:

```text
La nuova riga 2 non aggiorna nessuna vecchia regola solo perche' e' riga 2.
```

Il sistema legge il contenuto.

Se il contenuto produce un criterio nuovo, crea una nuova regola.

Se il contenuto produce un criterio gia' esistente con stessa conta, non duplica.

Se il contenuto produce un criterio gia' esistente con conta diversa, segnala conflitto.

### 8.2 Regola assente nella nuova planilha

Comportamento corretto:

```text
La regola resta attiva.
```

La planilha non e' una lista completa delle regole valide. E' solo una fonte di nuove decisioni.

### 8.3 Riga con conta approvata ma senza criterio

Comportamento corretto:

```text
Non generare regola.
```

La riga deve andare nel report come:

```text
sem_criterio_real
```

### 8.4 Riga con criterio ma senza conta approvata

Comportamento corretto:

```text
Non generare regola.
```

La riga deve andare nel report come:

```text
sem_conta_aprovada
```

### 8.5 Beneficiario piu' referencia

Se una regola contiene sia beneficiario sia referencia, il motore deve richiedere entrambi.

Quindi l'importatore non deve aggiungere `referencia` automaticamente quando la planilha non fornisce una referencia approvata.

## 9. Report di importazione

Ogni importazione deve produrre un report, preferibilmente in formato `.md` e/o `.csv`.

Contenuto minimo:

```text
regole nuove aggiunte
regole gia' esistenti
righe ignorate per mancanza di criterio reale
righe ignorate per mancanza di conta approvata
conflitti isolati
righe con dati incompleti
```

Per ogni conflitto:

```text
criterio
conta esistente
conta proposta
azione: candidata non importata
```

Per ogni riga senza criterio:

```text
motivo: nessun beneficiario/referencia/alias approvato
azione: nessuna regola generata
```

Il report puo' contenere riferimenti alla riga della planilha per aiutare l'operatore, ma questi riferimenti non devono entrare nel YAML.

## 10. Validazioni obbligatorie

Il validatore deve bloccare ogni regola positiva che non rispetti queste condizioni:

1. `regra_id` presente e non vuoto;
2. `ativa` presente;
3. `destino.conta_id` presente;
4. `destino.conta_nome` presente;
5. `destino.dre_linha_id` coerente;
6. `destino.entra_dre` coerente;
7. almeno un criterio reale presente;
8. nessun campo di origine planilha nel YAML;
9. nessun criterio vuoto;
10. nessun conflitto interno tra regole attive.

Regola fondamentale:

```text
E' vietato generare o accettare una regola positiva senza criteri reali di match.
```

## 11. Strategia di migrazione dal YAML attuale

Il YAML attuale contiene regole come:

```yaml
regras:
- regra_id: COMPROVANTE_LINHA_002
  origem_linha_planilha: 2
  conta_id: 9.01.001
  conta_nome: Publicidade e Propaganda
  dre_linha_id: '11.01'
  entra_dre: true
```

Queste regole non sono operative perche' non hanno criteri reali.

La migrazione deve:

1. leggere il YAML attuale;
2. separare regole valide da regole invalide;
3. rimuovere dal nuovo YAML i riferimenti alla planilha;
4. ricostruire regole valide solo se esistono criteri reali approvati;
5. mandare le regole senza criteri a report;
6. non inventare beneficiario, referencia o alias.

## 12. Piano di realizzazione tecnica

### 12.1 Fase A: audit senza modifiche

Creare o aggiornare uno script di audit che:

- legge `config/regras_classificacao_comprovantes.yaml`;
- identifica regole senza criteri;
- identifica campi vietati;
- identifica conflitti;
- produce un report.

Nessuna modifica automatica al YAML in questa fase.

### 12.2 Fase B: validatore del YAML

Implementare una funzione di validazione centralizzata.

Il classificatore e l'importatore devono usare la stessa validazione.

Se il YAML contiene regole invalide, il sistema deve fallire in modo esplicito oppure segnalarle come non operative, secondo la politica scelta per il comando specifico.

### 12.3 Fase C: nuovo importatore da planilha

Implementare il comando:

```text
atualizar-regras-comprovantes --planilha caminho/arquivo.xlsx
```

Con due modalita':

```text
--dry-run
--apply
```

`--dry-run`:

- legge planilha e YAML;
- calcola candidate;
- produce report;
- non modifica il YAML.

`--apply`:

- esegue le stesse validazioni;
- aggiunge solo candidate valide e non conflittuali;
- salva il YAML aggiornato;
- produce report.

### 12.4 Fase D: migrazione dello schema YAML

Aggiornare il motore per leggere il nuovo schema:

```yaml
criterios:
destino:
```

Se necessario, mantenere temporaneamente compatibilita' con il formato piatto esistente solo per leggere regole vecchie valide.

La direzione finale deve essere un solo formato chiaro.

### 12.5 Fase E: test

Aggiungere test per:

- regola senza criteri viene rifiutata;
- campi di planilha nel YAML vengono rifiutati;
- stessa keyword con stessa conta non duplica;
- stessa keyword con conta diversa genera conflitto isolato;
- regola assente dalla nuova planilha resta attiva;
- nuova riga 2 diversa dalla vecchia riga 2 non sovrascrive una vecchia regola;
- beneficiario piu' referencia richiede entrambi;
- l'importatore non aggiunge referencia automaticamente;
- il classificatore non legge la planilha.

## 13. Comandi operativi finali

Comandi desiderati:

```text
auditar-regras-comprovantes
```

Legge il YAML e produce report di validita'.

```text
atualizar-regras-comprovantes --planilha caminho/arquivo.xlsx --dry-run
```

Simula aggiornamento da planilha.

```text
atualizar-regras-comprovantes --planilha caminho/arquivo.xlsx --apply
```

Applica aggiornamento al YAML.

```text
executar-classificacao-comprovantes-real
```

Classifica comprovantes usando solo YAML.

## 14. Cosa non fare

Non fare in questo progetto:

- non riaprire matching bancario;
- non usare estratto conto;
- non inferire conti da beneficiario, banca, data, valore o similarita';
- non modificare Fase 9;
- non modificare XML/NF;
- non modificare DRE definitivo;
- non cancellare regole vecchie per assenza nella nuova planilha;
- non usare righe della planilha come identita';
- non accettare regole senza match reale.

## 15. Definizione di successo

Il progetto e' riuscito quando:

1. il YAML non contiene riferimenti alla planilha;
2. ogni regola positiva ha almeno un criterio reale;
3. il classificatore funziona senza nessuna planilha;
4. una nuova planilha puo' aggiungere regole senza riscrivere quelle vecchie;
5. conflitti vengono isolati e riportati;
6. regole assenti dalla nuova planilha restano attive;
7. i test impediscono la ricomparsa di `COMPROVANTE_LINHA_002` come identita' operativa basata sulla riga.

