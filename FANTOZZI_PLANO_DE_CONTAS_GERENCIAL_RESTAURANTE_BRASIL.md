# FANTOZZI — Proposta di Plano de Contas Gerencial para Restaurante no Brasil

> Finalidade: criar uma base prática para o agente **Fantozzi** classificar despesas, custos, investimentos e retiradas dos sócios de um restaurante no Brasil.  
> Idioma dos nomes contábeis: **português brasileiro**.  
> Uso previsto: gestão interna, DRE gerencial, controle de caixa e preparação de dados para conferência com o contador.

---

## 1. Premissas importantes

Este plano de contas é **gerencial**. Ele ajuda o restaurante a controlar custos, margens, pagamentos, investimentos e retiradas dos sócios.

Ele **não substitui** o plano contábil oficial do contador, nem as obrigações fiscais, trabalhistas e previdenciárias.

Como os dois CNPJs **não são optantes pelo Simples Nacional**, o plano já prevê contas separadas para:

- IRPJ;
- CSLL;
- PIS;
- COFINS;
- ICMS;
- ISS, se houver serviços sujeitos ao imposto municipal;
- IBS e CBS durante a transição da Reforma Tributária;
- contribuições previdenciárias;
- FGTS;
- retenções na fonte;
- multas e juros fiscais.

Também prevê separação por:

- CNPJ;
- loja/unidade;
- centro de custo;
- competência;
- data de pagamento;
- origem do lançamento;
- tipo de documento.

---

## 2. Campos recomendados para cada lançamento de Fantozzi

Cada registro contábil gerencial deveria ter, no mínimo:

| Campo | Descrição |
|---|---|
| `cnpj_empresa` | CNPJ que recebeu a nota ou realizou o pagamento |
| `nome_empresa` | Nome da empresa destinatária |
| `competencia` | Mês/ano econômico da despesa |
| `data_documento` | Data da nota, folha, boleto, recibo ou comprovante |
| `data_pagamento` | Data real do pagamento |
| `fornecedor_ou_favorecido` | Fornecedor, funcionário, sócio, banco ou órgão público |
| `cnpj_cpf_favorecido` | CNPJ/CPF do favorecido, quando houver |
| `documento` | Número NF-e, NFS-e, recibo, folha, guia, contrato etc. |
| `conta_id` | Código da conta gerencial |
| `conta_nome` | Nome da conta gerencial |
| `centro_custo` | Cozinha, Bar, Salão, Administração, Delivery, Obras etc. |
| `valor` | Valor do lançamento |
| `tipo_lancamento` | Custo, despesa, imposto, investimento, retirada, provisão, ajuste |
| `origem` | XML_NFE, CSV, EXCEL_MANUAL, BANCO, FOLHA, CONTADOR |
| `status_revisao` | OK, REVISAR, NÃO_CLASSIFICADO, AMBÍGUO |
| `observacoes` | Campo livre para explicações |

---

## 3. Centros de custo sugeridos

Usar centros de custo evita misturar despesa de cozinha, salão, administração e investimento.

| Código | Centro de custo |
|---|---|
| `CC-001` | Cozinha |
| `CC-002` | Bar |
| `CC-003` | Salão / Atendimento |
| `CC-004` | Delivery / Retirada |
| `CC-005` | Administração |
| `CC-006` | Limpeza e Higiene |
| `CC-007` | Manutenção |
| `CC-008` | Marketing e Vendas |
| `CC-009` | Veículos |
| `CC-010` | Obras e Investimentos |
| `CC-999` | Geral / A definir |

---

## 4. Estrutura geral do plano de contas

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

# 1. Receitas

Mesmo que Fantozzi comece pelas despesas, recomendo reservar o grupo de receitas para evolução futura.

| Código | Conta |
|---|---|
| `1.01.001` | Receita de Vendas no Salão |
| `1.01.002` | Receita de Delivery |
| `1.01.003` | Receita de Balcão / Retirada |
| `1.01.004` | Receita de Eventos |
| `1.01.005` | Taxa de Serviço / Gorjeta |
| `1.01.006` | Couvert Artístico |
| `1.01.007` | Outras Receitas Operacionais |
| `1.02.001` | Descontos Concedidos |
| `1.02.002` | Cancelamentos e Estornos |
| `1.02.003` | Perdas por Chargeback |

---

# 2. Custos de Mercadorias e Insumos

Este grupo é o coração do custo do restaurante. Ele se conecta diretamente às NF-e de compra.

## 2.01 Bebidas

| Código | Conta |
|---|---|
| `2.01.001` | Cervejas |
| `2.01.002` | Vinhos e Espumantes |
| `2.01.003` | Destilados, Licores e Aperitivos |
| `2.01.004` | Refrigerantes |
| `2.01.005` | Águas Minerais |
| `2.01.006` | Energéticos |
| `2.01.007` | Sucos, Polpas e Bebidas não Alcoólicas |
| `2.01.008` | Gelo |

## 2.02 Alimentos e insumos de cozinha

| Código | Conta |
|---|---|
| `2.02.001` | Carnes Bovinas |
| `2.02.002` | Aves |
| `2.02.003` | Suínos |
| `2.02.004` | Pescados e Frutos do Mar |
| `2.02.005` | Frios e Embutidos |
| `2.02.006` | Laticínios e Queijos |
| `2.02.007` | Ovos |
| `2.02.008` | Hortifruti |
| `2.02.009` | Pães e Padaria |
| `2.02.010` | Mercearia Seca |
| `2.02.011` | Óleos, Azeites e Gorduras |
| `2.02.012` | Molhos, Temperos e Condimentos |
| `2.02.013` | Conservas e Enlatados |
| `2.02.014` | Batatas Congeladas e Pré-prontos |
| `2.02.015` | Sobremesas, Sorvetes e Picolés |
| `2.02.016` | Café, Chá e Achocolatados |
| `2.02.017` | Insumos para Drinks e Coquetelaria |
| `2.02.999` | Outros Insumos de Cozinha |

## 2.03 Materiais diretamente ligados à venda

| Código | Conta |
|---|---|
| `2.03.001` | Descartáveis e Embalagens |
| `2.03.002` | Guardanapos, Canudos e Sachês |
| `2.03.003` | Embalagens para Delivery |
| `2.03.004` | Etiquetas e Lacres |
| `2.03.999` | Outros Materiais de Venda |

---

# 3. Pessoal, Encargos e Serviços de Mão de Obra

## 3.01 Pessoal CLT — Salários e remunerações

| Código | Conta |
|---|---|
| `3.01.001` | Salários e Ordenados CLT |
| `3.01.002` | Horas Extras CLT |
| `3.01.003` | Adicional Noturno |
| `3.01.004` | Adicional de Insalubridade |
| `3.01.005` | Adicional de Periculosidade |
| `3.01.006` | DSR sobre Horas Extras e Adicionais |
| `3.01.007` | Comissões, Prêmios e Bonificações |
| `3.01.008` | Feriados Trabalhados |
| `3.01.009` | Quebra de Caixa |
| `3.01.010` | Pagamento de Folgas / Banco de Horas |
| `3.01.011` | 13º Salário |
| `3.01.012` | Férias |
| `3.01.013` | 1/3 Constitucional de Férias |
| `3.01.014` | Aviso Prévio |
| `3.01.015` | Rescisões Trabalhistas |
| `3.01.016` | Multa Rescisória do FGTS |
| `3.01.999` | Outros Custos de Pessoal CLT |

## 3.02 Benefícios CLT

| Código | Conta |
|---|---|
| `3.02.001` | Vale-Transporte |
| `3.02.002` | Vale-Alimentação |
| `3.02.003` | Vale-Refeição |
| `3.02.004` | Alimentação Interna de Funcionários |
| `3.02.005` | Plano de Saúde |
| `3.02.006` | Plano Odontológico |
| `3.02.007` | Cesta Básica |
| `3.02.008` | Uniformes |
| `3.02.009` | Equipamentos de Proteção Individual - EPI |
| `3.02.010` | Treinamentos de Funcionários |
| `3.02.011` | Exames Admissionais, Periódicos e Demissionais |
| `3.02.999` | Outros Benefícios de Funcionários |

## 3.03 Encargos sobre folha CLT

| Código | Conta |
|---|---|
| `3.03.001` | FGTS |
| `3.03.002` | INSS Patronal |
| `3.03.003` | RAT / SAT |
| `3.03.004` | Terceiros / Sistema S |
| `3.03.005` | INSS sobre 13º Salário |
| `3.03.006` | FGTS sobre 13º Salário |
| `3.03.007` | INSS sobre Férias |
| `3.03.008` | FGTS sobre Férias |
| `3.03.009` | Provisão de 13º Salário |
| `3.03.010` | Provisão de Férias |
| `3.03.011` | Provisão de FGTS |
| `3.03.012` | Provisão de Encargos Sociais |
| `3.03.999` | Outros Encargos Trabalhistas |

## 3.04 Prestadores, diaristas e serviços de terceiros

Usar este grupo para pagamentos de mão de obra **sem vínculo CLT formal**, como prestadores eventuais, freelancers, autônomos ou pessoas jurídicas.

Atenção: se houver subordinação, habitualidade, pessoalidade e onerosidade, pode haver risco trabalhista. Não usar esta categoria para esconder relação de emprego.

| Código | Conta |
|---|---|
| `3.04.001` | Diaristas e Extras de Cozinha |
| `3.04.002` | Diaristas e Extras de Salão |
| `3.04.003` | Diaristas e Extras de Bar |
| `3.04.004` | Freelancers de Eventos |
| `3.04.005` | Prestadores Autônomos - RPA |
| `3.04.006` | Prestadores Pessoa Jurídica - PJ |
| `3.04.007` | Prestadores MEI |
| `3.04.008` | Ajudas de Custo a Prestadores |
| `3.04.009` | Transporte de Prestadores |
| `3.04.010` | Alimentação de Prestadores |
| `3.04.011` | Pagamentos Extraordinários a Regularizar |
| `3.04.012` | Retenções sobre Serviços de Terceiros |
| `3.04.999` | Outros Serviços de Mão de Obra |

---

# 4. Ocupação, Utilidades e Funcionamento

| Código | Conta |
|---|---|
| `4.01.001` | Aluguel do Imóvel |
| `4.01.002` | Condomínio |
| `4.01.003` | IPTU |
| `4.01.004` | Taxa de Coleta de Lixo |
| `4.01.005` | Seguros do Imóvel |
| `4.01.006` | Licenças, Alvarás e Vistorias |
| `4.01.007` | Dedetização e Controle de Pragas |
| `4.01.008` | Limpeza Terceirizada |
| `4.01.009` | Segurança e Monitoramento |
| `4.01.010` | Extintores e Segurança Contra Incêndio |
| `4.01.011` | Energia Elétrica |
| `4.01.012` | Água e Esgoto |
| `4.01.013` | Gás de Cozinha |
| `4.01.014` | Internet |
| `4.01.015` | Telefone |
| `4.01.016` | Sistemas, Softwares e Assinaturas Operacionais |
| `4.01.999` | Outras Despesas de Funcionamento |

---

# 5. Impostos, Taxas e Obrigações Fiscais

## 5.01 Tributos sobre receita e lucro

| Código | Conta |
|---|---|
| `5.01.001` | IRPJ |
| `5.01.002` | CSLL |
| `5.01.003` | PIS |
| `5.01.004` | COFINS |
| `5.01.005` | ICMS |
| `5.01.006` | ISS |
| `5.01.007` | IBS - Período de Transição |
| `5.01.008` | CBS - Período de Transição |
| `5.01.009` | Imposto Seletivo |
| `5.01.010` | Difal / Antecipação de ICMS |
| `5.01.011` | ICMS Substituição Tributária |
| `5.01.999` | Outros Tributos sobre Receita |

## 5.02 Tributos e obrigações sobre folha

| Código | Conta |
|---|---|
| `5.02.001` | INSS a Recolher - Empregados |
| `5.02.002` | INSS a Recolher - Empresa |
| `5.02.003` | FGTS a Recolher |
| `5.02.004` | IRRF sobre Folha |
| `5.02.005` | IRRF sobre Pró-labore |
| `5.02.006` | Contribuição Sindical / Assistencial, se aplicável |
| `5.02.999` | Outras Obrigações sobre Folha |

## 5.03 Retenções sobre serviços tomados

| Código | Conta |
|---|---|
| `5.03.001` | IRRF sobre Serviços Tomados |
| `5.03.002` | PIS Retido |
| `5.03.003` | COFINS Retido |
| `5.03.004` | CSLL Retida |
| `5.03.005` | INSS Retido |
| `5.03.006` | ISS Retido |
| `5.03.999` | Outras Retenções |

## 5.04 Taxas, multas e juros fiscais

| Código | Conta |
|---|---|
| `5.04.001` | Taxas Municipais |
| `5.04.002` | Taxas Estaduais |
| `5.04.003` | Taxas Federais |
| `5.04.004` | Multas Fiscais |
| `5.04.005` | Juros Fiscais |
| `5.04.006` | Parcelamentos Tributários |
| `5.04.999` | Outras Taxas e Penalidades |

---

# 6. Despesas Administrativas e Escritório

| Código | Conta |
|---|---|
| `6.01.001` | Honorários Contábeis |
| `6.01.002` | Honorários Jurídicos |
| `6.01.003` | Consultorias |
| `6.01.004` | Certificado Digital |
| `6.01.005` | Material de Escritório |
| `6.01.006` | Papelaria e Impressos |
| `6.01.007` | Cartuchos, Toner e Suprimentos de Impressão |
| `6.01.008` | Correios e Entregas Administrativas |
| `6.01.009` | Sistemas Administrativos |
| `6.01.010` | Assinaturas Digitais e Armazenamento em Nuvem |
| `6.01.011` | Despesas com Documentação |
| `6.01.012` | Despesas Cartoriais |
| `6.01.013` | Treinamentos Administrativos |
| `6.01.999` | Outras Despesas Administrativas |

---

# 7. Bancos, Cartões, Juros e Serviços Financeiros

## 7.01 Bancos

| Código | Conta |
|---|---|
| `7.01.001` | Tarifas Bancárias |
| `7.01.002` | Pacote de Serviços Bancários |
| `7.01.003` | Tarifas PIX |
| `7.01.004` | Tarifas TED / DOC |
| `7.01.005` | Emissão de Boletos |
| `7.01.006` | Juros Bancários |
| `7.01.007` | IOF |
| `7.01.008` | Juros de Empréstimos |
| `7.01.009` | Juros de Financiamentos |
| `7.01.010` | Multas e Encargos Bancários |
| `7.01.999` | Outros Custos Bancários |

## 7.02 Cartões e meios de pagamento

| Código | Conta |
|---|---|
| `7.02.001` | Comissão Cartão de Crédito |
| `7.02.002` | Comissão Cartão de Débito |
| `7.02.003` | Comissão PIX via Operadora |
| `7.02.004` | Aluguel de Máquina de Cartão |
| `7.02.005` | Taxa de Antecipação de Recebíveis |
| `7.02.006` | Taxa de Gateway / Link de Pagamento |
| `7.02.007` | Chargebacks |
| `7.02.008` | Estornos e Contestações |
| `7.02.009` | Taxas de Vale-Refeição / Voucher |
| `7.02.999` | Outros Custos de Meios de Pagamento |

---

# 8. Manutenção, Reparos e Conservação

## 8.01 Manutenção do imóvel

| Código | Conta |
|---|---|
| `8.01.001` | Manutenção Predial |
| `8.01.002` | Pintura |
| `8.01.003` | Elétrica Predial |
| `8.01.004` | Hidráulica |
| `8.01.005` | Marcenaria |
| `8.01.006` | Serralheria |
| `8.01.007` | Vidraçaria |
| `8.01.008` | Ar-condicionado |
| `8.01.009` | Jardinagem |
| `8.01.999` | Outras Manutenções do Imóvel |

## 8.02 Manutenção de equipamentos

| Código | Conta |
|---|---|
| `8.02.001` | Manutenção de Equipamentos de Cozinha |
| `8.02.002` | Manutenção de Equipamentos de Bar |
| `8.02.003` | Manutenção de Geladeiras e Freezers |
| `8.02.004` | Manutenção de Fogões, Fornos e Fritadeiras |
| `8.02.005` | Manutenção de Equipamentos Eletrônicos |
| `8.02.006` | Manutenção de Computadores e Impressoras |
| `8.02.007` | Manutenção de Sistemas de Som, TV e Câmeras |
| `8.02.008` | Peças de Reposição |
| `8.02.999` | Outras Manutenções de Equipamentos |

## 8.03 Veículos

| Código | Conta |
|---|---|
| `8.03.001` | Combustível de Veículos |
| `8.03.002` | Manutenção de Veículos |
| `8.03.003` | Peças e Pneus |
| `8.03.004` | Seguro de Veículos |
| `8.03.005` | IPVA e Licenciamento |
| `8.03.006` | Multas de Trânsito |
| `8.03.999` | Outras Despesas com Veículos |

---

# 9. Marketing, Comercial e Delivery

| Código | Conta |
|---|---|
| `9.01.001` | Publicidade e Propaganda |
| `9.01.002` | Redes Sociais e Tráfego Pago |
| `9.01.003` | Design Gráfico |
| `9.01.004` | Fotos e Vídeos |
| `9.01.005` | Influenciadores e Parcerias |
| `9.01.006` | Impressos Promocionais |
| `9.01.007` | Brindes e Promoções |
| `9.01.008` | Comissões de Aplicativos de Delivery |
| `9.01.009` | Mensalidade de Plataformas de Delivery |
| `9.01.010` | Frete e Motoboy |
| `9.01.011` | Taxas de Marketplace |
| `9.01.999` | Outras Despesas Comerciais |

---

# 10. Investimentos e Imobilizado

Este grupo não deve ser tratado como despesa operacional normal.  
Serve para controlar compras que aumentam patrimônio, reformas, equipamentos e bens duráveis.

| Código | Conta |
|---|---|
| `10.01.001` | Investimento em Obras e Reformas |
| `10.01.002` | Investimento em Benfeitorias no Imóvel |
| `10.01.003` | Equipamentos de Cozinha |
| `10.01.004` | Equipamentos de Bar |
| `10.01.005` | Equipamentos Eletrônicos |
| `10.01.006` | Computadores e Periféricos |
| `10.01.007` | Móveis e Decoração |
| `10.01.008` | Utensílios Duráveis |
| `10.01.009` | Veículos |
| `10.01.010` | Sistemas e Softwares Implantados |
| `10.01.011` | Projetos de Arquitetura e Engenharia |
| `10.01.012` | Depreciação Gerencial - Imóvel / Benfeitorias |
| `10.01.013` | Depreciação Gerencial - Equipamentos |
| `10.01.014` | Depreciação Gerencial - Veículos |
| `10.01.999` | Outros Investimentos |

---

# 11. Sócios, Pró-labore e Distribuição de Lucros

Separar muito bem remuneração por trabalho e retirada de lucro.

| Código | Conta |
|---|---|
| `11.01.001` | Pró-labore de Sócios Trabalhadores |
| `11.01.002` | INSS sobre Pró-labore |
| `11.01.003` | IRRF sobre Pró-labore |
| `11.01.004` | Reembolso de Despesas dos Sócios |
| `11.01.005` | Distribuição de Lucros |
| `11.01.006` | Antecipação de Lucros |
| `11.01.007` | Empréstimos de Sócios para a Empresa |
| `11.01.008` | Devolução de Empréstimos a Sócios |
| `11.01.009` | Aporte de Capital dos Sócios |
| `11.01.010` | Retiradas Não Classificadas de Sócios |
| `11.01.999` | Outros Movimentos de Sócios |

Observação: `Pró-labore` é custo/despesa de remuneração do sócio trabalhador.  
`Distribuição de lucros` e `antecipação de lucros` não devem ser misturadas com custo operacional do restaurante.

---

# 12. Despesas Diversas e Ajustes

| Código | Conta |
|---|---|
| `12.01.001` | Pequenas Despesas Operacionais |
| `12.01.002` | Diferenças de Caixa |
| `12.01.003` | Quebras e Perdas Operacionais |
| `12.01.004` | Doações e Cortesias |
| `12.01.005` | Reembolsos Diversos |
| `12.01.006` | Ajustes de Inventário |
| `12.01.007` | Perdas de Estoque |
| `12.01.008` | Perdas por Validade Vencida |
| `12.01.009` | Despesas Não Recorrentes |
| `12.01.999` | Outras Despesas Diversas |

---

# 13. Contas de Revisão e Classificação Pendente

Este grupo é obrigatório para Fantozzi.  
O agente não deve inventar classificação quando houver dúvida.

| Código | Conta |
|---|---|
| `13.01.001` | Revisar - Produto Não Classificado |
| `13.01.002` | Revisar - Classificação Ambígua |
| `13.01.003` | Revisar - Documento sem CNPJ |
| `13.01.004` | Revisar - Documento sem Número |
| `13.01.005` | Revisar - Documento sem Data |
| `13.01.006` | Revisar - Valor Inconsistente |
| `13.01.007` | Revisar - Vencimento Inconsistente |
| `13.01.008` | Revisar - CNPJ Destinatário Diferente |
| `13.01.009` | Revisar - Possível Investimento |
| `13.01.010` | Revisar - Pagamento de Pessoal a Regularizar |
| `13.01.999` | Revisar - Outros Casos |

---

## 5. Regras gerenciais importantes para Fantozzi

### 5.1 Compra operacional x investimento

Fantozzi deve diferenciar:

- gasto pequeno e recorrente: despesa/custo;
- bem durável ou reforma relevante: investimento/imobilizado;
- manutenção que apenas conserva: manutenção;
- melhoria que aumenta capacidade ou vida útil: investimento.

Exemplos:

| Descrição | Classificação recomendada |
|---|---|
| Conserto de freezer | Manutenção de Geladeiras e Freezers |
| Compra de freezer novo | Equipamentos de Cozinha |
| Pintura simples de manutenção | Pintura |
| Reforma estrutural da cozinha | Investimento em Obras e Reformas |
| Compra de panelas comuns | Utensílios Duráveis ou Insumo, dependendo do controle |
| Compra de embalagens delivery | Embalagens para Delivery |

### 5.2 Pessoal CLT x prestador

Fantozzi deve separar:

- folha CLT;
- encargos CLT;
- benefícios;
- prestadores eventuais;
- PJ/MEI;
- pagamentos extraordinários a regularizar.

Não recomendo usar o nome “sem carteira assinada” como conta oficial.  
Para controle interno, usar `Prestadores, diaristas e serviços de terceiros` e, quando houver risco, lançar também alerta `PAGAMENTO_A_REGULARIZAR`.

### 5.3 Pró-labore x lucro

- `Pró-labore`: remuneração do sócio que trabalha.
- `INSS sobre Pró-labore`: encargo/contribuição vinculada ao pró-labore.
- `Distribuição de Lucros`: retirada de resultado.
- `Antecipação de Lucros`: retirada antes do fechamento final do lucro.
- `Retiradas Não Classificadas de Sócios`: conta provisória para revisar.

### 5.4 Impostos

Para empresas fora do Simples Nacional, manter contas separadas para cada tributo.  
Não lançar tudo genericamente como “Impostos”.

Separar também:

- imposto principal;
- multa;
- juros;
- parcelamento;
- retenção na fonte;
- antecipação tributária.

---

## 6. Sugestão de implementação no Fantozzi

Na primeira versão do Fantozzi, usar três arquivos separados:

```text
config/plano_contas_gerencial.yaml
config/regras_classificacao_produtos.yaml
config/regras_classificacao_despesas.yaml
```

### 6.1 `plano_contas_gerencial.yaml`

Contém a lista oficial das contas:

```yaml
- id: "3.01.001"
  nome: "Salários e Ordenados CLT"
  grupo: "Pessoal, Encargos e Serviços de Mão de Obra"
  tipo: "despesa"
  exige_revisao: false
```

### 6.2 `regras_classificacao_produtos.yaml`

Usado para XML NF-e de produtos comprados.

Exemplo:

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

Usado para boletos, extratos bancários, comprovantes e lançamentos manuais.

Exemplo:

```yaml
- conta_id: "4.01.011"
  conta_nome: "Energia Elétrica"
  aliases:
    - energia
    - eletricidade
    - coelba
    - neoenergia
```

---

## 7. Próximo passo recomendado

Antes de programar tudo, validar estas decisões:

1. Separar ou unir `Cervejas`, `Vinhos` e `Destilados`?
2. Separar `Energia`, `Água`, `Gás`, `Internet` em contas próprias? Minha recomendação: sim.
3. Separar funcionários CLT por setor? Exemplo: cozinha, salão, bar, limpeza. Minha recomendação: usar centro de custo, não criar contas duplicadas.
4. Tratar retiradas dos sócios fora da DRE operacional? Minha recomendação: sim.
5. Tratar investimentos fora das despesas operacionais? Minha recomendação: sim.
6. Criar alerta para pagamentos de mão de obra não formalizados? Minha recomendação: sim.

---

## 8. Resumo prático

Este plano permite gerar três visões:

1. **DRE gerencial**: receitas, CMV, pessoal, despesas, impostos e lucro operacional.
2. **Fluxo de caixa**: pagamentos reais por data.
3. **Controle patrimonial**: investimentos, equipamentos, obras e veículos.

Essa separação é essencial para o Fantozzi não misturar:

- compra de mercadoria com despesa administrativa;
- salário com retirada de sócio;
- manutenção com investimento;
- imposto normal com multa/juros;
- pagamento de fornecedor com pagamento de pessoal;
- despesa da empresa com gasto pessoal do sócio.
