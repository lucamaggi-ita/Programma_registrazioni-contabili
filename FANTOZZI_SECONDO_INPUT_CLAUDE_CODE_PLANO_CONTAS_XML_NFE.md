# FANTOZZI — SECONDO INPUT PER CLAUDE CODE
## Integrazione del Plano de Contas Gerencial e delle regole XML NF-e

> Progetto: **Fantozzi**  
> Repository GitHub: `lucamaggi-ita/Programma_registrazioni-contabili`  
> Cartella locale di lavoro: `C:\Users\Luca Maggi\Documents\Git_Hub\Programma_Registrazioni_Contabili`  
> Ambiente: Windows 10 / Windows 11  
> Lingua operativa del piano dei conti: **portoghese brasiliano**

---

## 1. Contesto

Questo è il **secondo input operativo** per Claude Code.

Il primo input ha definito la creazione iniziale dell’agente **Fantozzi**, incaricato di importare fatture di acquisto del ristorante e registrarle in Excel.

Dopo l’analisi dei file reali, è emerso che le fatture non arrivano principalmente in CSV, ma in formato **XML NF-e brasiliano**.

Sono stati inoltre preparati:

1. un piano dei conti gerencial completo per ristorante in Brasile;
2. una proposta YAML del piano dei conti;
3. una proposta YAML di regole per classificare prodotti delle NF-e;
4. un CSV con le descrizioni reali dei prodotti estratti dalle fatture;
5. un complemento tecnico per l’importazione XML NF-e.

Claude deve ora usare questi materiali per migliorare la prima bozza di Fantozzi.

---

## 2. File da usare come input

I seguenti file devono essere copiati o considerati come documentazione/configurazione del progetto:

```text
FANTOZZI_INPUT_CLAUDE_CODE.md
FANTOZZI_COMPLEMENTO_CLAUDE_CODE_XML_NFE.md
FANTOZZI_ANALISE_NFE_E_CATEGORIAS.md
FANTOZZI_plano_contas_rules_proposta.yaml
FANTOZZI_PRODUTOS_UNICOS_EXTRAIDOS.csv
FANTOZZI_PLANO_DE_CONTAS_GERENCIAL_RESTAURANTE_BRASIL.md
FANTOZZI_plano_contas_gerencial_proposta.yaml
```

Se questi file non sono ancora dentro il repository, Claude deve chiedere/aspettare che vengano copiati oppure creare i file equivalenti nella struttura `docs/` e `config/`.

---

## 3. Obiettivo di questa fase

Integrare in Fantozzi una base più solida per:

1. **Plano de contas gerencial**;
2. **classificação de produtos XML NF-e**;
3. **classificação futura de despesas não originadas de NF-e**;
4. **separação entre custo, despesa, imposto, investimento, retirada de sócio e conta de revisão**;
5. **uso corretto del portoghese brasiliano nei nomi delle contas**.

---

## 4. Modifiche richieste alla struttura del progetto

Creare o aggiornare la struttura seguente:

```text
Programma_Registrazioni_Contabili/
├─ docs/
│  ├─ FANTOZZI_INPUT_CLAUDE_CODE.md
│  ├─ FANTOZZI_COMPLEMENTO_CLAUDE_CODE_XML_NFE.md
│  ├─ FANTOZZI_ANALISE_NFE_E_CATEGORIAS.md
│  └─ FANTOZZI_PLANO_DE_CONTAS_GERENCIAL_RESTAURANTE_BRASIL.md
├─ config/
│  ├─ plano_contas_gerencial.yaml
│  ├─ regras_classificacao_produtos.yaml
│  ├─ regras_classificacao_despesas.yaml
│  └─ csv_columns.yaml
├─ samples/
│  ├─ faturas_exemplo.csv
│  └─ nfe_exemplo/
├─ src/
│  └─ fantozzi/
│     ├─ chart_of_accounts.py
│     ├─ expense_classifier.py
│     ├─ product_classifier.py
│     ├─ nfe_importer.py
│     ├─ models.py
│     └─ cli.py
└─ tests/
   ├─ test_chart_of_accounts.py
   ├─ test_product_classifier.py
   ├─ test_expense_classifier.py
   └─ test_nfe_importer.py
```

Se la prima bozza ha nomi diversi, non riscrivere inutilmente: adattare mantenendo coerenza.

---

## 5. Plano de contas gerencial

Usare come base il file:

```text
FANTOZZI_plano_contas_gerencial_proposta.yaml
```

e copiarlo dentro il progetto come:

```text
config/plano_contas_gerencial.yaml
```

Il piano deve contenere conti in portoghese brasiliano, ad esempio:

```yaml
- id: "3.01.001"
  nome: "Salários e Ordenados CLT"
  grupo: "Pessoal CLT"
  tipo: "despesa"
```

### 5.1 Tipi di conta previsti

Fantozzi deve supportare almeno questi tipi:

```text
receita
deducao_receita
custo
despesa
imposto
investimento
patrimonial
passivo
provisao
ajuste
revisao
despesa_gerencial
```

### 5.2 Requisito importante

Quando Fantozzi classifica qualcosa, deve salvare sia:

```text
conta_id
conta_nome
```

Non basta salvare solo il nome della conta.

---

## 6. Separazione tra contas e regras

Non mischiare il piano dei conti con le regole di classificazione.

Usare tre file separati:

```text
config/plano_contas_gerencial.yaml
config/regras_classificacao_produtos.yaml
config/regras_classificacao_despesas.yaml
```

### 6.1 `plano_contas_gerencial.yaml`

Contiene l’elenco ufficiale delle contas.

### 6.2 `regras_classificacao_produtos.yaml`

Contiene alias e parole chiave per classificare prodotti letti dagli XML NF-e.

Esempio:

```yaml
- conta_id: "2.01.004"
  conta_nome: "Refrigerantes"
  aliases:
    - coca cola
    - coca-cola
    - fanta
    - sprite
    - guaraná
    - guarana
    - tônica
    - tonica
```

### 6.3 `regras_classificacao_despesas.yaml`

Contiene regole future per classificare despesas da banco, boleto, folha, comprovante o registrazioni manuali.

Esempio:

```yaml
- conta_id: "4.01.011"
  conta_nome: "Energia Elétrica"
  aliases:
    - energia eletrica
    - energia elétrica
    - eletricidade
    - neoenergia
    - coelba
```

---

## 7. Gruppi principali del plano de contas

Il piano deve prevedere almeno questi gruppi:

```text
1. Receitas
2. Custos de Mercadorias e Insumos
3. Pessoal, Encargos e Serviços de Mão de Obra
4. Ocupação, Utilidades e Funcionamento
5. Impostos, Taxas e Obrigações Fiscais
6. Despesas Administrativas e Escritório
7. Bancos, Cartões, Juros e Serviços Financeiros
8. Manutenção, Reparos e Conservação
9. Marketing, Comercial e Delivery
10. Investimentos e Imobilizado
11. Sócios, Pró-labore e Distribuição de Lucros
12. Despesas Diversas e Ajustes
13. Contas de Revisão e Classificação Pendente
```

---

## 8. Regole di classificazione prodotti NF-e

Usare come base il file:

```text
FANTOZZI_plano_contas_rules_proposta.yaml
```

e trasformarlo/copiaro come:

```text
config/regras_classificacao_produtos.yaml
```

Le categorie prodotto devono includere almeno:

```text
Cervejas
Vinhos e Espumantes
Destilados, Licores e Aperitivos
Refrigerantes
Águas Minerais
Energéticos
Sucos, Polpas e Bebidas não Alcoólicas
Gelo
Ovos
Hortifruti
Pães e Padaria
Carnes Bovinas
Aves
Suínos
Pescados e Frutos do Mar
Frios e Embutidos
Laticínios e Queijos
Mercearia Seca
Óleos, Azeites e Gorduras
Molhos, Temperos e Condimentos
Conservas e Enlatados
Batatas Congeladas e Pré-prontos
Sobremesas, Sorvetes e Picolés
Descartáveis e Embalagens
Produtos de Limpeza e Higiene
Utensílios de Cozinha e Salão
Gás de Cozinha
Equipamentos e Imobilizado
Manutenção e Obras
Móveis, Decoração e Materiais Promocionais
REVISAR / NÃO CLASSIFICADO
```

### 8.1 Normalizzazione obbligatoria

La classificazione deve essere:

- case-insensitive;
- robusta rispetto agli accenti;
- robusta rispetto a trattini e punteggiatura;
- robusta rispetto a spazi multipli;
- capace di matchare alias parziali quando previsto.

Esempi:

```text
"ÁGUA SANITÁRIA" -> "agua sanitaria"
"COCA-COLA LATA" -> "coca cola lata"
"VINHO RIO SOL" -> "vinho rio sol"
```

---

## 9. Regole di classificazione despesas

Creare una prima bozza di:

```text
config/regras_classificacao_despesas.yaml
```

Questa fase non deve ancora importare automaticamente estratti bancari, ma deve preparare la struttura.

### 9.1 Energia elétrica

```yaml
- conta_id: "4.01.011"
  conta_nome: "Energia Elétrica"
  aliases:
    - energia elétrica
    - energia eletrica
    - conta de luz
    - eletricidade
    - neoenergia
    - coelba
```

### 9.2 Água e esgoto

```yaml
- conta_id: "4.01.012"
  conta_nome: "Água e Esgoto"
  aliases:
    - água
    - agua
    - conta de água
    - conta de agua
    - embasa
    - esgoto
```

### 9.3 Pessoal CLT

```yaml
- conta_id: "3.01.001"
  conta_nome: "Salários e Ordenados CLT"
  aliases:
    - salario
    - salário
    - folha de pagamento
    - pagamento funcionario
    - pagamento funcionário
```

### 9.4 Vale-transporte

```yaml
- conta_id: "3.02.001"
  conta_nome: "Vale-Transporte"
  aliases:
    - vale transporte
    - vale-transporte
    - transporte funcionario
    - transporte funcionário
```

### 9.5 FGTS

```yaml
- conta_id: "3.03.001"
  conta_nome: "FGTS"
  aliases:
    - fgts
    - guia fgts
    - sefip
```

### 9.6 INSS

```yaml
- conta_id: "3.03.002"
  conta_nome: "INSS Patronal"
  aliases:
    - inss patronal
    - gps
    - darf previdenciario
    - darf previdenciário
```

### 9.7 Pró-labore

```yaml
- conta_id: "11.01.001"
  conta_nome: "Pró-labore de Sócios Trabalhadores"
  aliases:
    - pro labore
    - pró-labore
    - pro-labore
    - retirada pro labore
```

### 9.8 Antecipação de lucros

```yaml
- conta_id: "11.01.006"
  conta_nome: "Antecipação de Lucros"
  aliases:
    - antecipacao de lucros
    - antecipação de lucros
    - retirada socio
    - retirada sócio
    - retirada lucro
```

### 9.9 Tarifas bancárias

```yaml
- conta_id: "7.01.001"
  conta_nome: "Tarifas Bancárias"
  aliases:
    - tarifa bancaria
    - tarifa bancária
    - cesta bancaria
    - cesta bancária
    - pacote de serviços
    - pacote de servicos
```

### 9.10 Cartões e máquinas de pagamento

```yaml
- conta_id: "7.02.001"
  conta_nome: "Comissão Cartão de Crédito"
  aliases:
    - taxa cartao credito
    - taxa cartão crédito
    - comissão crédito
    - comissao credito

- conta_id: "7.02.002"
  conta_nome: "Comissão Cartão de Débito"
  aliases:
    - taxa cartao debito
    - taxa cartão débito
    - comissão débito
    - comissao debito

- conta_id: "7.02.004"
  conta_nome: "Aluguel de Máquina de Cartão"
  aliases:
    - aluguel maquina cartao
    - aluguel máquina cartão
    - mensalidade maquininha
    - maquininha
    - rede
    - cielo
    - stone
    - getnet
```

### 9.11 Escritório e papelaria

```yaml
- conta_id: "6.01.005"
  conta_nome: "Material de Escritório"
  aliases:
    - material escritorio
    - material escritório
    - caneta
    - papel a4
    - clips
    - grampo
    - pasta
```

### 9.12 Manutenção

```yaml
- conta_id: "8.01.001"
  conta_nome: "Manutenção Predial"
  aliases:
    - manutencao predial
    - manutenção predial
    - reparo predial
    - conserto predial

- conta_id: "8.02.001"
  conta_nome: "Manutenção de Equipamentos de Cozinha"
  aliases:
    - manutencao fogao
    - manutenção fogão
    - conserto freezer
    - conserto geladeira
    - fritadeira
    - forno
```

### 9.13 Investimentos

```yaml
- conta_id: "10.01.001"
  conta_nome: "Investimento em Obras e Reformas"
  aliases:
    - reforma
    - obra
    - construção
    - construcao
    - benfeitoria

- conta_id: "10.01.003"
  conta_nome: "Equipamentos de Cozinha"
  aliases:
    - freezer novo
    - geladeira nova
    - fogao novo
    - fogão novo
    - forno novo
    - fritadeira nova
    - equipamento cozinha
```

---

## 10. Requisito de segurança: conta de revisão

Quando Fantozzi non è sicuro, deve usare una conta di revisão, non inventare.

Contas di revisione obbligatorie:

```text
13.01.001 Revisar - Produto Não Classificado
13.01.002 Revisar - Classificação Ambígua
13.01.006 Revisar - Valor Inconsistente
13.01.009 Revisar - Possível Investimento
13.01.010 Revisar - Pagamento de Pessoal a Regularizar
13.01.999 Revisar - Outros Casos
```

---

## 11. Pessoal CLT e prestadores

### 11.1 Non usare questa categoria ufficiale

Non creare una conta chiamata:

```text
Pessoal sem carteira assinada
```

### 11.2 Usare invece

```text
Prestadores, diaristas e serviços de terceiros
```

e, se necessario:

```text
Pagamentos Extraordinários a Regularizar
Revisar - Pagamento de Pessoal a Regularizar
```

Motivo: il nome deve essere corretto e meno rischioso.  
Fantozzi deve aiutare a controllare e regolarizzare, non a nascondere classificazioni delicate.

---

## 12. Pró-labore e retiradas dos sócios

Fantozzi deve separare chiaramente:

```text
Pró-labore de Sócios Trabalhadores
INSS sobre Pró-labore
IRRF sobre Pró-labore
Distribuição de Lucros
Antecipação de Lucros
Retiradas Não Classificadas de Sócios
```

Regola:

- `Pró-labore` entra come remuneração do sócio trabalhador.
- `Distribuição de Lucros` non è costo operativo.
- `Antecipação de Lucros` deve restare separata fino al fechamento.
- `Retiradas Não Classificadas de Sócios` deve generare revisão.

---

## 13. Investimento x despesa

Fantozzi deve distinguere:

| Caso | Classificazione |
|---|---|
| Conserto di freezer | Manutenção de Geladeiras e Freezers |
| Acquisto di freezer nuovo | Equipamentos de Cozinha |
| Pintura semplice | Pintura |
| Reforma strutturale | Investimento em Obras e Reformas |
| Acquisto computer | Equipamentos Eletrônicos / Computadores |
| Acquisto imballaggi delivery | Embalagens para Delivery |

Quando c’è dubbio, usare:

```text
13.01.009 Revisar - Possível Investimento
```

---

## 14. XML NF-e

Integrare l’importatore XML NF-e come priorità.

Dagli XML devono essere letti almeno:

```text
chave_nfe
numero_nfe
serie_nfe
data_emissao
cnpj_emitente
nome_emitente
cnpj_destinatario
nome_destinatario
produto_codigo
produto_descricao
ncm
cfop
quantidade
unidade
valor_unitario
valor_total_item
valor_total_nfe
duplicatas / vencimentos, se presenti
```

### 14.1 Attenzione ai diversi CNPJ destinatari

Poiché negli XML reali esistono NF-e destinate a più di un CNPJ, Fantozzi deve sempre salvare:

```text
cnpj_destinatario
nome_destinatario
```

Questo evita di mischiare registrazioni tra aziende diverse.

---

## 15. Excel di output

Aggiungere al file Excel colonne coerenti con il piano dei conti:

```text
cnpj_empresa
nome_empresa
competencia
data_documento
data_pagamento
fornecedor_ou_favorecido
cnpj_cpf_favorecido
documento
tipo_documento
conta_id
conta_nome
grupo_conta
centro_custo
valor
tipo_lancamento
origem
arquivo_origem
chave_origem
hash_registro
status_revisao
alertas
observacoes
```

Per XML NF-e, `origem` deve essere:

```text
XML_NFE
```

Per righe manuali:

```text
MANUAL
```

Fantozzi non deve cancellare o sovrascrivere righe `MANUAL`.

---

## 16. Test minimi richiesti in questa fase

Aggiungere test per:

### 16.1 Plano de contas

- caricare `config/plano_contas_gerencial.yaml`;
- verificare che ogni conta abbia `id`, `nome`, `grupo`, `tipo`;
- verificare che non ci siano `id` duplicati;
- verificare che le contas di revisão esistano.

### 16.2 Classificazione prodotti

- `Coca Cola` -> `2.01.004 Refrigerantes`;
- `Rio Sol` -> `2.01.002 Vinhos e Espumantes`;
- `Esponja` -> `Produtos de Limpeza e Higiene`;
- prodotto sconosciuto -> `13.01.001 Revisar - Produto Não Classificado`.

### 16.3 Classificazione despesas

- `Conta de luz Neoenergia` -> `4.01.011 Energia Elétrica`;
- `Embasa água` -> `4.01.012 Água e Esgoto`;
- `FGTS guia` -> `3.03.001 FGTS`;
- `Tarifa bancária` -> `7.01.001 Tarifas Bancárias`;
- `Retirada sócio` -> `11.01.006 Antecipação de Lucros` oppure revisão se ambiguo.

### 16.4 XML NF-e

- leggere almeno un XML NF-e di esempio;
- estrarre emitente, destinatário, número NF-e, data, produtos e valores;
- classificare itens usando `regras_classificacao_produtos.yaml`;
- salvare `cnpj_destinatario`.

### 16.5 Excel

- creare foglio mensal `YYYY-MM`;
- scrivere `conta_id` e `conta_nome`;
- preservare righe `MANUAL`;
- evitare duplicati tramite `hash_registro`.

---

## 17. CLI aggiornata

Aggiungere o mantenere comandi simili:

```powershell
python -m fantozzi.cli import-nfe `
  --input "samples\nfe_exemplo" `
  --excel "output\registrazioni_contabili.xlsx" `
  --plano-contas "config\plano_contas_gerencial.yaml" `
  --rules-produtos "config\regras_classificacao_produtos.yaml"
```

Modalità prova:

```powershell
python -m fantozzi.cli import-nfe `
  --input "samples\nfe_exemplo" `
  --excel "output\registrazioni_contabili.xlsx" `
  --dry-run
```

Validazione del piano:

```powershell
python -m fantozzi.cli validate-plano-contas `
  --plano-contas "config\plano_contas_gerencial.yaml"
```

---

## 18. Report di revisione

Ogni importazione deve generare report con:

```text
NF-e importate
Itens lidos
Itens classificados
Itens não classificados
Classificações ambíguas
Valores inconsistentes
Possíveis investimentos
CNPJ destinatários encontrados
Duplicados ignorados
Arquivo Excel atualizado
```

Report consigliato:

```text
reports/import_nfe_report_YYYYMMDD_HHMMSS.txt
```

---

## 19. Criterio di accettazione

Questa fase è accettabile se:

1. il piano dei conti YAML viene caricato e validato;
2. le regole prodotti vengono caricate;
3. le regole despesas vengono caricate;
4. XML NF-e di esempio viene letto;
5. prodotti NF-e vengono classificati con `conta_id` e `conta_nome`;
6. Excel viene creato/aggiornato senza cancellare righe manuali;
7. righe duplicate non vengono duplicate;
8. conti di revisione vengono usati quando serve;
9. `pytest` passa;
10. la CLI espone almeno `import-nfe` e `validate-plano-contas`.

---

## 20. Cose da NON fare in questa fase

Non implementare ancora:

- importazione bancaria completa;
- folha de pagamento automatica completa;
- riconciliazione bancaria;
- interfaccia grafica;
- dashboard;
- collegamento a Receita Federal;
- invio automatico al contador;
- automazioni cloud.

Questa fase serve a rendere robusta la base:

```text
XML NF-e -> prodotti -> classificazione -> plano de contas -> Excel gerencial
```

---

## 21. Output finale che Claude deve restituire

Alla fine Claude deve rispondere con:

1. file creati;
2. file modificati;
3. come sono stati copiati/integrati i YAML;
4. comandi PowerShell per installare;
5. comandi PowerShell per testare;
6. comandi PowerShell per importare XML NF-e;
7. esempio del risultato atteso;
8. limiti attuali;
9. prossima fase consigliata.

---

## 22. Regola di fondo

Fantozzi deve privilegiare sempre:

```text
controllo > automazione cieca
revisione > classificazione inventata
dados reais > ipotesi
português brasileiro correto > nomi improvvisati
separação contábil > categorie generiche
```

Se il programma ha dubbio, deve segnalarlo.

Non deve “indovinare” in silenzio.
