# FANTOZZI — TERZO INPUT PER CLAUDE CODE
## Riclassificazione DRE, statistiche e grafici futuri

> Progetto: **Fantozzi**  
> Repository GitHub: `lucamaggi-ita/Programma_registrazioni-contabili`  
> Cartella locale: `C:\Users\Luca Maggi\Documents\Git_Hub\Programma_Registrazioni_Contabili`  
> Lingua operativa: **português brasileiro**  
> Obiettivo di questo input: preparare Fantozzi per generare in futuro **DRE gerencial**, statistiche e grafici, rispettando i raggruppamenti contabili usati in Brasile.

---

## 1. Contesto

Fantozzi nasce per registrare fatture di acquisto, XML NF-e, costi, despesas, imposte, personale, soci, investimenti e movimenti manuali.

Nelle fasi successive dovrà produrre:

1. statistiche;
2. grafici;
3. relatório DRE;
4. riclassificazione contabile gerencial;
5. confronto mensile;
6. analisi per CNPJ;
7. analisi per centro de custo;
8. analisi per competência e per caixa.

Per evitare una riscrittura futura, questa struttura deve essere prevista fin da ora.

---

## 2. Principio fondamentale

Ogni lançamento di Fantozzi deve avere almeno due classificazioni:

```text
1. Conta gerencial detalhada
2. Linha de DRE
```

Esempio:

```text
Conta gerencial:
2.01.004 Refrigerantes

Linha DRE:
03.01 CMV - Bebidas
```

Questo permette a Fantozzi di mantenere il dettaglio operativo, ma generare anche una DRE sintetica e leggibile.

---

## 3. DRE contábil x DRE gerencial

Fantozzi deve distinguere:

### 3.1 DRE contábil

Relatório più vicino alla struttura contabile ufficiale brasiliana, da validare con il contador.

### 3.2 DRE gerencial

Relatório interno per gestione del ristorante.

Può avere più dettaglio, per esempio:

```text
CMV - Alimentos
CMV - Bebidas
CMV - Embalagens
Pessoal Operacional
Pessoal Administrativo
Taxas de Cartão
Delivery
Marketing
Manutenção
```

Ma deve rimanere compatibile con il raggruppamento brasiliano classico:

```text
Receita Bruta
(-) Deduções da Receita
(=) Receita Líquida
(-) Custos
(=) Lucro Bruto
(-) Despesas Operacionais
(+/-) Resultado Financeiro
(+/-) Outras Receitas e Despesas
(=) Resultado antes de IRPJ/CSLL
(-) IRPJ e CSLL
(=) Lucro ou Prejuízo Líquido
```

---

## 4. Competência x Caixa

DRE deve essere prodotta per **regime de competência**, non semplicemente per data di pagamento.

Quindi Fantozzi deve mantenere separati:

```text
competencia
data_documento
data_vencimento
data_pagamento
```

Regola:

- DRE usa `competencia`;
- fluxo de caixa usa `data_pagamento`;
- contas a pagar usa `data_vencimento`;
- audit documentale usa `data_documento`.

Se una NF-e è emessa a maggio ma pagata a giugno, nella DRE deve entrare a maggio, salvo diversa regola contabile definita.

---

## 5. Nuovo file richiesto: mapeamento DRE

Creare un file:

```text
config/mapeamento_dre.yaml
```

Questo file deve mappare ogni `conta_id` del piano gerencial a una linha DRE.

Esempio:

```yaml
- conta_id: "2.01.004"
  conta_nome: "Refrigerantes"
  dre_linha_id: "03.02"
  dre_linha_nome: "CMV - Bebidas"
  dre_grupo: "Custos"
  dre_sinal: "-"
  dre_ordem: 302
  entra_dre: true
```

### 5.1 Campi obbligatori

Ogni mapping deve avere:

```text
conta_id
conta_nome
dre_linha_id
dre_linha_nome
dre_grupo
dre_sinal
dre_ordem
entra_dre
observacoes
```

### 5.2 Sinais

Usare:

```text
"+" per receitas
"-" per deduções, custos e despesas
"+/-" per linhas que podem ser positivas ou negativas
"=" per linhas calculadas
```

---

## 6. Estrutura DRE recomendada para restaurante no Brasil

Criar no `config/mapeamento_dre.yaml` as seguintes linhas DRE.

---

# 01. Receita Bruta

| ID | Linha DRE |
|---|---|
| `01.01` | Receita Bruta de Vendas no Salão |
| `01.02` | Receita Bruta de Delivery |
| `01.03` | Receita Bruta de Balcão / Retirada |
| `01.04` | Receita de Eventos |
| `01.05` | Taxa de Serviço / Gorjeta |
| `01.99` | Outras Receitas Operacionais |

---

# 02. Deduções da Receita Bruta

| ID | Linha DRE |
|---|---|
| `02.01` | Impostos sobre Vendas |
| `02.02` | Descontos Concedidos |
| `02.03` | Cancelamentos e Estornos |
| `02.04` | Devoluções |
| `02.05` | Chargebacks sobre Vendas |
| `02.99` | Outras Deduções da Receita |

---

# 03. Custos das Mercadorias, Insumos e Produção

Para restaurante, usar CMV gerencial separado por tipo de insumo.

| ID | Linha DRE |
|---|---|
| `03.01` | CMV - Alimentos |
| `03.02` | CMV - Bebidas |
| `03.03` | CMV - Sobremesas |
| `03.04` | CMV - Embalagens e Descartáveis |
| `03.05` | CMV - Gás de Cozinha |
| `03.06` | CMV - Perdas de Estoque |
| `03.07` | CMV - Ajustes de Inventário |
| `03.08` | Mão de Obra Operacional da Cozinha |
| `03.09` | Benefícios da Equipe Operacional |
| `03.99` | Outros Custos de Produção |

Observação: se o restaurante quiser DRE mais simples, `03.08` e `03.09` podem ir para despesas com pessoal. Mas para visão gerencial de restaurante, muitas vezes faz sentido separar equipe da cozinha como custo operacional.

---

# 04. Lucro Bruto

Linha calculada:

```text
04.01 Lucro Bruto = Receita Líquida - Custos
```

Não lançar documentos diretamente nesta linha.

---

# 05. Despesas Operacionais

## 05.01 Despesas com Pessoal

| ID | Linha DRE |
|---|---|
| `05.01.01` | Pessoal CLT - Salários |
| `05.01.02` | Pessoal CLT - Horas Extras e Adicionais |
| `05.01.03` | Encargos Trabalhistas |
| `05.01.04` | Benefícios de Funcionários |
| `05.01.05` | Rescisões e Indenizações |
| `05.01.06` | Prestadores, Diaristas e Terceiros |
| `05.01.99` | Outros Custos de Pessoal |

## 05.02 Ocupação e Funcionamento

| ID | Linha DRE |
|---|---|
| `05.02.01` | Aluguel, Condomínio e IPTU |
| `05.02.02` | Energia Elétrica |
| `05.02.03` | Água e Esgoto |
| `05.02.04` | Internet e Telefonia |
| `05.02.05` | Segurança, Monitoramento e Licenças |
| `05.02.06` | Dedetização e Controle de Pragas |
| `05.02.99` | Outras Despesas de Ocupação |

## 05.03 Administrativas e Escritório

| ID | Linha DRE |
|---|---|
| `05.03.01` | Honorários Contábeis e Jurídicos |
| `05.03.02` | Sistemas, Softwares e Assinaturas |
| `05.03.03` | Material de Escritório e Papelaria |
| `05.03.04` | Certificado Digital e Documentação |
| `05.03.99` | Outras Despesas Administrativas |

## 05.04 Manutenção

| ID | Linha DRE |
|---|---|
| `05.04.01` | Manutenção Predial |
| `05.04.02` | Manutenção de Equipamentos de Cozinha |
| `05.04.03` | Manutenção de Equipamentos Eletrônicos |
| `05.04.04` | Manutenção de Veículos |
| `05.04.99` | Outras Manutenções |

## 05.05 Comercial, Marketing e Delivery

| ID | Linha DRE |
|---|---|
| `05.05.01` | Marketing e Propaganda |
| `05.05.02` | Redes Sociais e Tráfego Pago |
| `05.05.03` | Comissões de Aplicativos de Delivery |
| `05.05.04` | Frete, Motoboy e Entregas |
| `05.05.05` | Brindes, Promoções e Cortesias |
| `05.05.99` | Outras Despesas Comerciais |

---

# 06. Resultado Operacional antes do Financeiro

Linha calculada:

```text
06.01 Resultado Operacional antes do Financeiro =
Lucro Bruto - Despesas Operacionais
```

Não lançar documentos diretamente nesta linha.

---

# 07. Resultado Financeiro

| ID | Linha DRE |
|---|---|
| `07.01` | Receitas Financeiras |
| `07.02` | Tarifas Bancárias |
| `07.03` | Juros Bancários |
| `07.04` | IOF |
| `07.05` | Juros de Empréstimos e Financiamentos |
| `07.06` | Comissões de Cartão de Crédito |
| `07.07` | Comissões de Cartão de Débito |
| `07.08` | Taxa de Antecipação de Recebíveis |
| `07.09` | Aluguel de Máquina de Cartão |
| `07.99` | Outros Resultados Financeiros |

Observação: taxas de cartão podem ficar em despesas comerciais ou resultado financeiro. Para restaurante, recomendo manter em `Resultado Financeiro` ou criar indicador separado, mas nunca misturar com CMV.

---

# 08. Outras Receitas e Despesas

| ID | Linha DRE |
|---|---|
| `08.01` | Outras Receitas Operacionais |
| `08.02` | Outras Despesas Operacionais |
| `08.03` | Despesas Não Recorrentes |
| `08.04` | Multas e Juros Fiscais |
| `08.05` | Perdas Extraordinárias |
| `08.06` | Ganhos ou Perdas na Venda de Imobilizado |
| `08.99` | Outros Ajustes de Resultado |

---

# 09. Resultado antes de IRPJ e CSLL

Linha calculada:

```text
09.01 Resultado antes de IRPJ e CSLL =
Resultado Operacional + Resultado Financeiro + Outras Receitas/Despesas
```

---

# 10. IRPJ e CSLL

| ID | Linha DRE |
|---|---|
| `10.01` | IRPJ |
| `10.02` | CSLL |
| `10.03` | Provisão de IRPJ |
| `10.04` | Provisão de CSLL |
| `10.99` | Outros Tributos sobre Lucro |

---

# 11. Lucro ou Prejuízo Líquido

Linha calculada:

```text
11.01 Lucro ou Prejuízo Líquido =
Resultado antes de IRPJ e CSLL - IRPJ - CSLL
```

---

## 7. Contas que NÃO entram diretamente na DRE

Alguns lançamentos são importantes, mas não devem entrar como despesa operacional normal na DRE.

### 7.1 Investimentos e imobilizado

As contas do grupo 10 devem ser tratadas como `investimento` e não como despesa normal.

Exemplos:

```text
Equipamentos de Cozinha
Equipamentos Eletrônicos
Veículos
Obras e Reformas
Benfeitorias
```

Na DRE, futuramente pode entrar apenas:

```text
Depreciação Gerencial
```

### 7.2 Sócios

Separar:

```text
Pró-labore
INSS sobre Pró-labore
IRRF sobre Pró-labore
Distribuição de Lucros
Antecipação de Lucros
Retiradas Não Classificadas
```

Regra:

- Pró-labore pode entrar na DRE como despesa administrativa ou pessoal.
- Distribuição de lucros e antecipação de lucros NÃO devem reduzir o resultado operacional.
- Retiradas não classificadas devem ir para revisão.

### 7.3 Empréstimos e financiamentos

Separar:

- principal da dívida: não entra como despesa DRE;
- juros: entra no resultado financeiro;
- IOF/tarifas: entram no resultado financeiro.

---

## 8. Campos novos recomendados no Excel/base de dados

Adicionar aos lançamentos:

```text
dre_linha_id
dre_linha_nome
dre_grupo
dre_ordem
dre_sinal
entra_dre
regime
competencia_dre
valor_dre
```

### 8.1 `regime`

Valores possíveis:

```text
competencia
caixa
ambos
fora_dre
```

### 8.2 `valor_dre`

Para receitas, positivo.  
Para custos e despesas, pode ser salvo positivo com `dre_sinal = "-"`, oppure diretamente negativo.  
Escolher uma regra e manter consistência.

Minha recomendação:

```text
valor = sempre positivo no lançamento original
dre_sinal = define se soma ou subtrai na DRE
```

---

## 9. Estatísticas futuras

Fantozzi deve preparar os dados para gerar estes indicadores:

### 9.1 Indicadores de receita

```text
Receita bruta por mês
Receita líquida por mês
Receita por CNPJ
Receita por centro de custo
Receita por canal: salão, delivery, balcão
Ticket médio
```

### 9.2 Indicadores de custo

```text
CMV total
CMV alimentos
CMV bebidas
CMV embalagens
CMV sobre receita líquida %
Top 20 fornecedores
Top 20 contas de custo
Evolução mensal por conta
```

### 9.3 Indicadores de pessoal

```text
Custo de pessoal total
Custo de pessoal / receita líquida %
Pessoal CLT
Encargos
Benefícios
Prestadores e diaristas
Custo por centro de custo
```

### 9.4 Indicadores de despesas operacionais

```text
Energia elétrica / receita líquida %
Água e esgoto / receita líquida %
Gás / receita líquida %
Aluguel / receita líquida %
Marketing / receita líquida %
Manutenção mensal
Despesas administrativas
```

### 9.5 Indicadores financeiros

```text
Tarifas bancárias
Comissões de cartão
Taxa média de cartão
Antecipação de recebíveis
Juros bancários
Resultado financeiro mensal
```

### 9.6 Indicadores de investimento

```text
Investimentos por mês
Investimentos por tipo
Investimentos por CNPJ
Investimentos em cozinha
Investimentos em salão
Investimentos em obras
```

---

## 10. Gráficos futuros

Fantozzi deve preparar uma camada de dados para gerar gráficos, mesmo que a primeira versão ainda não desenhe todos.

Gráficos recomendados:

```text
DRE mensal em barras
DRE acumulada no período
Waterfall da DRE
Evolução da receita líquida
Evolução do CMV %
Evolução do custo de pessoal %
Evolução do lucro bruto
Evolução do resultado operacional
Pizza ou barras por grupo de despesa
Top fornecedores
Top categorias de custo
Comparativo CNPJ 1 x CNPJ 2
Fluxo de caixa previsto por vencimento
```

### 10.1 Exportações futuras

Prevedere in futuro:

```text
Excel com abas de relatório
CSV consolidado
PDF gerencial
Dashboard HTML local
Gráficos PNG
```

Non implementare tutto ora se la fase è ancora iniziale. Però il modello dati deve essere pronto.

---

## 11. Nuovi moduli consigliati

Aggiungere in futuro, o predisporre adesso:

```text
src/fantozzi/dre_mapping.py
src/fantozzi/dre_report.py
src/fantozzi/statistics.py
src/fantozzi/charts.py
```

### 11.1 `dre_mapping.py`

Responsabile di:

- leggere `config/mapeamento_dre.yaml`;
- validare che ogni `conta_id` esista nel plano de contas;
- applicare `dre_linha_id` ai lançamentos;
- segnalare contas senza mapeamento.

### 11.2 `dre_report.py`

Responsabile di:

- consolidare lançamentos por competência;
- calcular linhas sintéticas;
- gerar DRE mensal;
- gerar DRE acumulada;
- separar por CNPJ;
- separar por centro de custo.

### 11.3 `statistics.py`

Responsabile di:

- calcular KPIs;
- comparar períodos;
- calcular percentuais sobre receita líquida.

### 11.4 `charts.py`

Responsabile futuro di:

- gerar gráficos;
- exportar imagens;
- preparar dashboard.

---

## 12. Test minimi futuri per DRE

Criar ou preparar test per:

### 12.1 Validação do mapeamento DRE

- ogni `conta_id` del mapping esiste nel plano de contas;
- ogni conta DRE ha `dre_linha_id`;
- non ci sono `dre_linha_id` duplicati non desiderati;
- contas di investimento hanno `entra_dre: false`, salvo depreciazione;
- contas di distribuição de lucros hanno `entra_dre: false`.

### 12.2 Cálculo DRE

Esempio:

```text
Receita Bruta: 10.000
Deduções: 1.000
Receita Líquida: 9.000
CMV: 3.000
Lucro Bruto: 6.000
Despesas Operacionais: 2.000
Resultado Operacional: 4.000
Resultado Financeiro: -500
IRPJ/CSLL: 600
Lucro Líquido: 2.900
```

Il test deve verificare tutti i subtotali.

### 12.3 Competência x caixa

- documento emesso a maggio e pagato a giugno entra nella DRE di maggio;
- lo stesso documento entra nel fluxo de caixa di giugno.

### 12.4 Sócios

- antecipação de lucros não entra na DRE operacional;
- pró-labore entra come despesa.

---

## 13. CLI futura

Prevedere comandi come:

```powershell
python -m fantozzi.cli gerar-dre `
  --excel output\registrazioni_contabili.xlsx `
  --competencia 2026-05 `
  --saida reports\DRE_2026-05.xlsx
```

```powershell
python -m fantozzi.cli gerar-dre `
  --excel output\registrazioni_contabili.xlsx `
  --inicio 2026-01 `
  --fim 2026-12 `
  --por-cnpj
```

```powershell
python -m fantozzi.cli gerar-estatisticas `
  --excel output\registrazioni_contabili.xlsx `
  --inicio 2026-01 `
  --fim 2026-12
```

```powershell
python -m fantozzi.cli gerar-graficos `
  --dre reports\DRE_2026-05.xlsx `
  --saida reports\graficos\
```

---

## 14. Criterio di accettazione per questa preparazione

Anche se DRE e grafici non vengono implementati subito, questa fase è accettabile se:

1. esiste `config/mapeamento_dre.yaml`;
2. esiste una struttura dati per salvare `dre_linha_id`;
3. il piano dei conti distingue `tipo` e `entra_dre`;
4. investimenti e retiradas de sócios non vengono trattati come despesas normali;
5. i lançamentos hanno `competencia`;
6. la futura DRE potrà essere generata senza riclassificare tutto a mano;
7. ci sono test o almeno funzioni di validazione del mapping.

---

## 15. Istruzione pratica per Claude Code

Claude, in questa fase non devi necessariamente implementare tutti i grafici.

Devi però verificare che la struttura creata per Fantozzi sia pronta per:

```text
lançamento detalhado -> plano de contas -> mapeamento DRE -> DRE gerencial -> estatísticas -> gráficos
```

Se manca questa struttura, aggiungila ora in modo semplice e testabile.

Non complicare il progetto con dashboard o UI.

Priorità:

1. `config/mapeamento_dre.yaml`;
2. campi DRE nei lançamentos;
3. validazione del mapping;
4. preservazione della separazione competência/caixa;
5. test minimi;
6. documentazione chiara in português brasileiro.

---

## 16. Regola finale

Fantozzi deve essere capace di rispondere in futuro a domande come:

```text
Quanto foi o CMV de bebidas em maio?
Qual foi o custo de pessoal sobre a receita líquida?
Quanto gastamos com energia elétrica nos últimos 6 meses?
Qual CNPJ teve maior margem?
Quanto foi investimento e quanto foi despesa?
Qual foi o lucro operacional antes do financeiro?
Qual foi o lucro líquido gerencial?
```

Per rispondere bene a queste domande, ogni lançamento deve nascere già con:

```text
conta_id
conta_nome
centro_custo
competencia
tipo_lancamento
dre_linha_id
dre_linha_nome
entra_dre
```

Sem isso, il progetto crescerà male.
