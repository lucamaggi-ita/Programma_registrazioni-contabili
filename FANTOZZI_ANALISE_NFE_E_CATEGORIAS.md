# Fantozzi — Análise preliminar dos XML NF-e recebidos

## 1. Resultado rápido

Analisei os dois arquivos ZIP enviados. Eles contêm XML de **NF-e recebidas**, não CSV simples. Portanto, a primeira versão de Fantozzi deve prever um importador XML NF-e, além do importador CSV originalmente imaginado.

Resumo extraído:

| Item | Quantidade / Valor |
|---|---:|
| Arquivos XML NF-e lidos | 764 |
| Eventos XML ignorados | 17 |
| NF-e únicas | 764 |
| Itens/produtos lidos | 3341 |
| Descrições únicas de produto | 1093 |
| Valor total das NF-e (`vNF`) | R$ 1.094.246,81 |
| Valor total dos itens (`vProd` dos produtos) | R$ 1.098.904,78 |
| NF-e com duplicatas/vencimentos no XML | 544 |
| NF-e sem duplicatas/vencimentos no XML | 220 |

## 2. Empresas destinatárias encontradas

Os XML misturam NF-e emitidas para mais de um CNPJ. Fantozzi deve gravar o destinatário da NF-e no Excel, para evitar misturar empresas.

| CNPJ destinatário | Nome encontrado no XML | NF-e |
|---|---|---:|
| 55746722000181 | TRATTORIA DO SOSSEGO RESTAURANTE LTDA | 366 |
| 08884180000123 | AMAMI RESTAURANTE LTDA | 277 |
| 08884180000123 | AMAMI RESTAURRANTE LTDA | 32 |
| 08884180000123 | AMAMI RESTAURANTE LTDA ME | 28 |
| 55746722000181 | TRATTORIA DO SOSSEGO RESTAURANTE LT | 24 |
| 55746722000181 | TRATTORIA DO SOSSEGO RESTAURANTE | 19 |
| 08884180000123 | Amami Restaurante Ltda-Me | 7 |
| 08884180000123 | AMAMI RESTAURANTE LTDA - ME | 3 |
| 08884180000123 | CABANA RECANTO DO SOSSEGO | 3 |
| 55746722000181 | (RECANTO DO SOSSEGO) TRATTORIA DO SOSSEGO RESTAURANTE LTDA | 2 |
| 08884180000123 | AMAMI RESTAURANTE LTDA - RECANTO DO SOSSEGO | 1 |
| 08884180000123 | Amami Restaurante Ltda | 1 |
| 08884180000123 | 8788 - AMAMI RESTAURANTE LTDA | 1 |

## 3. Meses encontrados

| Mês emissão | NF-e |
|---|---:|
| 2025-10 | 113 |
| 2025-11 | 99 |
| 2025-12 | 178 |
| 2026-01 | 147 |
| 2026-02 | 103 |
| 2026-03 | 124 |

## 4. Principais fornecedores por quantidade de NF-e

| Fornecedor | NF-e |
|---|---:|
| OESA COMERCIO E REPRESENTACOES S/A | 71 |
| Checon Distribuidora e Transportadora Lt | 50 |
| NORSA REFRIGERANTES S.A. | 47 |
| GP7 NORDESTE DISTRIBUIDOR DE BEBIDAS LTD | 45 |
| FE IMPORTADORA E DISTRA LTDA | 42 |
| OESA COMERCIO E REPRESENTACOES SA | 36 |
| CENTRAL DE ALIMENTOS DA BAHIA - CAB LTDA | 28 |
| JACILENY MAURICIO TEIXEIRA | 25 |
| ARMAZEM TOP ALTO LTDA | 23 |
| QUEIJOS FINOS IND,COM E REPRESENTACAO | 23 |
| KI GELO COMERCIO E FABRICACAO DE GELO LTDA | 22 |
| E SANTOS DE OLIVEIRA CARNES E DERIVADOS | 19 |
| L & C COMERCIO DE PRODUTOS DE LIMPEZA LTDA | 19 |
| DATERRA DISTRIBUIDORA DE ALIMENTOS LTDA | 19 |
| NATURAVES INDUSTRIA E COMERCIO DE ALIMENTOS LTDA | 18 |
| LATICINIO MAIS VIDA LTDA | 16 |
| Marques Fair CD Porto | 15 |
| THIAGO RIBEIRO DE MEDEIROS | 15 |
| PRIME SEAFOOD LTDA - BAHIA | 14 |
| SOST IND E COMERCIO DE ALIMENTOS LTDA | 14 |

## 5. Categorias contábeis propostas

Abaixo estão categorias em português do Brasil, pensadas para restaurante. Elas foram propostas olhando os produtos reais presentes nos XML.

### Bebidas

- Cervejas
- Vinhos e Espumantes
- Destilados, Licores e Aperitivos
- Refrigerantes
- Águas Minerais
- Energéticos
- Sucos, Polpas e Bebidas não Alcoólicas
- Gelo

### Alimentos

- Ovos
- Hortifruti
- Pães e Padaria
- Carnes Bovinas
- Aves
- Pescados e Frutos do Mar
- Frios e Embutidos
- Laticínios e Queijos
- Mercearia Seca
- Óleos, Azeites e Gorduras
- Molhos, Temperos e Condimentos
- Conservas e Enlatados
- Batatas Congeladas e Pré-prontos
- Sobremesas, Sorvetes e Picolés

### Operação, materiais e estrutura

- Descartáveis e Embalagens
- Produtos de Limpeza e Higiene
- Utensílios de Cozinha e Salão
- Gás de Cozinha
- Equipamentos e Imobilizado
- Manutenção e Obras
- Móveis, Decoração e Materiais Promocionais
- REVISAR / NÃO CLASSIFICADO

## 6. Cobertura preliminar das regras propostas

Esta cobertura é apenas uma simulação inicial. A classificação definitiva deve usar regras com prioridade, normalização, palavra inteira para aliases curtos e alerta de ambiguidade.

| Categoria sugerida | Itens | Valor aproximado |
|---|---:|---:|
| Cervejas | 305 | R$ 193.132,56 |
| Laticínios e Queijos | 131 | R$ 106.578,94 |
| Equipamentos e Imobilizado | 25 | R$ 68.043,82 |
| Pescados e Frutos do Mar | 38 | R$ 66.020,13 |
| Vinhos e Espumantes | 119 | R$ 52.882,26 |
| Conservas e Enlatados | 121 | R$ 52.356,01 |
| Refrigerantes | 258 | R$ 50.275,82 |
| Carnes Bovinas | 36 | R$ 46.225,10 |
| Hortifruti | 226 | R$ 40.596,48 |
| Óleos, Azeites e Gorduras | 52 | R$ 36.357,39 |
| Produtos de Limpeza e Higiene | 177 | R$ 33.872,08 |
| Frios e Embutidos | 62 | R$ 32.910,77 |
| Mercearia Seca | 147 | R$ 32.252,85 |
| Manutenção e Obras | 117 | R$ 30.441,88 |
| Móveis, Decoração e Materiais Promocionais | 12 | R$ 28.836,52 |
| Sobremesas, Sorvetes e Picolés | 120 | R$ 28.641,18 |
| Destilados, Licores e Aperitivos | 123 | R$ 26.371,10 |
| Gás de Cozinha | 8 | R$ 25.922,98 |
| Descartáveis e Embalagens | 111 | R$ 24.129,63 |
| Aves | 41 | R$ 20.225,44 |
| Águas Minerais | 37 | R$ 16.549,80 |
| Gelo | 23 | R$ 15.085,00 |
| Pães e Padaria | 473 | R$ 13.810,85 |
| REVISAR / NÃO CLASSIFICADO | 319 | R$ 13.126,39 |
| Batatas Congeladas e Pré-prontos | 14 | R$ 12.188,31 |
| Sucos, Polpas e Bebidas não Alcoólicas | 93 | R$ 9.415,87 |
| Molhos, Temperos e Condimentos | 83 | R$ 7.928,40 |
| Ovos | 24 | R$ 6.203,72 |
| Utensílios de Cozinha e Salão | 18 | R$ 4.450,29 |
| Energéticos | 28 | R$ 4.073,21 |

## 7. Produtos de maior valor ainda para revisar

Estes produtos ficaram sem classificação segura na simulação preliminar. Muitos deles podem virar novas regras ou exceções.

| Produto | Ocorrências | Valor | Fornecedor principal | NCM |
|---|---:|---:|---|---|
| DOSADOR: UNIKIT S | 1 | R$ 850,00 | AREA MARKET COMERCIO LTDA (1) | 84798912 (1) |
| DOSADOR: UNIKIT D | 1 | R$ 850,00 | AREA MARKET COMERCIO LTDA (1) | 84798912 (1) |
| KIT CENTRAL 03 A 15XB190/B125 | 1 | R$ 584,31 | BAHIANA DISTRIBUIDORA DE GAS LTDA (1) | 74122000 (1) |
| GABINETE 1 BAIA - BG-2001 C/FONTE 200W PRETO BLUECASE - USB / MICROATX / MINI-ITX | 1 | R$ 225,00 | Informatica Porto Seguro 2000 Ltda. (1) | 84733011 (1) |
| ESTEIRA TELA SOLDADA 3 X 2 | 2 | R$ 220,00 | PATAXO MATERIAIS DE CONSTRUCAO LTDA EPP (2) | 73142000 (2) |
| ROTHMANS HAND SELECTED BLUE BOX HW25 | 1 | R$ 168,60 | SOUZA CRUZ LTDA. (1) | 24022000 (1) |
| BLUKIT KIT FACIL P/REG 10 EM 1 C/ACAB. C-40 | 1 | R$ 160,00 | PATAXO MATERIAIS DE CONSTRUCAO LTDA EPP (1) | 84819010 (1) |
| GEKRIL ESMALTE BASE D.AGUA 3,6L BRANCO | 1 | R$ 150,00 | PATAXO MATERIAIS DE CONSTRUCAO LTDA EPP (1) | 32091010 (1) |
| RASTELO PLAST PRETA 22D C/ CABO | 1 | R$ 149,50 | GERONIMO FERRAGENS E FERRAMENTAS LTDA (1) | 39269090 (1) |
| TRAMONTINA CAVADEIRA ARTICULADA C/CB 1.20MT FERRO 77562/803 | 1 | R$ 145,00 | PATAXO MATERIAIS DE CONSTRUCAO LTDA EPP (1) | 82019000 (1) |
| TIGRE UNIAO SOLD. 32MM | 1 | R$ 136,80 | PATAXO MATERIAIS DE CONSTRUCAO LTDA EPP (1) | 39174090 (1) |
| DUNHILL CARLTON BLEND 2.0 BOX 10X1 | 1 | R$ 135,32 | SOUZA CRUZ LTDA. (1) | 24022000 (1) |
| MASTER BAC LAVANDA 5 LTS | 1 | R$ 132,29 | BERTHOLDI & SANTOS LTDA (1) | 38089919 (1) |
| CHIC TRIDENT MENTA 1X21 | 3 | R$ 131,97 | DELMAX SUPERMERCADO COMERCIO DE ALIMENTOS LTDA (3) | 21069050 (3) |
| CORANTE | 6 | R$ 122,50 | JMT COMERCIO DE DESCARTAVEIS E ARTIGOS DE LIMPEZA LTDA (4) | JACILENY MAURICIO TEIXEIRA (2) | 21039021 (6) |
| MEXILHAO INTEIRO PCT 500GR VAC CXA 5 KG | 1 | R$ 120,00 | PRIME SEAFOOD LTDA - BAHIA (1) | 03073200 (1) |
| PAF564 ARCO DE BALOES 56 PCS - DOURADO | 1 | R$ 119,80 | CONFETE ARTIGOS DE FESTAS LTDA (1) | 95059000 (1) |
| PAF564 ARCO DE BALOES 56 PCS - PRATA | 1 | R$ 119,80 | CONFETE ARTIGOS DE FESTAS LTDA (1) | 95059000 (1) |
| SAL SACHE D MAIS 1000X08GR | 3 | R$ 118,30 | JMT COMERCIO DE DESCARTAVEIS E ARTIGOS DE LIMPEZA LTDA (3) | 25010020 (3) |
| COOLER INTEL ORIGINAL 1151/1150/1155/1155/1156 | 2 | R$ 118,00 | Informatica Porto Seguro 2000 Ltda. (2) | 84145910 (2) |
| PALITO DENTE COM SACHE PLASTICO 2000UN BOMPACK | 2 | R$ 117,00 | JMT COMERCIO DE DESCARTAVEIS E ARTIGOS DE LIMPEZA LTDA (2) | 44219100 (2) |
| LOMBO CANADENSE GOUMET SEARA KG | 1 | R$ 115,46 | FE IMPORTADORA E DISTRA LTDA (1) | 16010000 (1) |
| ALCOOL NOBRE 70 INPM 76,9GL 1000ML | 1 | R$ 114,00 | JACILENY MAURICIO TEIXEIRA (1) | 38089419 (1) |
| TIGRE FITA ISOLANTE 20MT | 4 | R$ 114,00 | PATAXO MATERIAIS DE CONSTRUCAO LTDA EPP (4) | 39191020 (4) |
| TEKBOND SPRAY UG 350ML BRANCO BRILHANTE | 5 | R$ 113,60 | PATAXO MATERIAIS DE CONSTRUCAO LTDA EPP (5) | 32082019 (5) |
| CARTOM BOTINA S/BIQ PTA 42 | 2 | R$ 112,00 | PATAXO MATERIAIS DE CONSTRUCAO LTDA EPP (2) | 64039990 (2) |
| PAPEL A4 REPORT PREMIUM - SUZANO | 1 | R$ 108,00 | CIA DO PAPEL DISTRIBUIDORA LTDA (1) | 48025610 (1) |
| ADAPTADOR USB WIRELESS - MERCUSYS MW300 | 1 | R$ 105,00 | Informatica Porto Seguro 2000 Ltda. (1) | 85176277 (1) |
| CHUCHU (PLASTICA) | 2 | R$ 101,20 | CENTRAL DE ALIMENTOS DA BAHIA - CAB LTDA (2) | 07099990 (2) |
| TEKBOND SPRAY UG 350ML VERMELHO | 4 | R$ 100,00 | PATAXO MATERIAIS DE CONSTRUCAO LTDA EPP (4) | 32082019 (4) |

## 8. Produtos de maior valor encontrados

| Produto | Ocorrências | Valor | Categoria sugerida |
|---|---:|---:|---|
| CERV HEINEKEN PIL 0,60GFA RT 24UN | 30 | R$ 74.009,76 | Cervejas |
| QUEIJO MOZZARELLA CAPA AZUL 3,6KG | 13 | R$ 37.746,22 | Laticínios e Queijos |
| CERV AMSTEL LAGER 0,60L GFA RT 24UN | 24 | R$ 29.645,00 | Cervejas |
| TOMATE PELADO LATA 2,5KG OPERA I | 6 | R$ 26.709,00 | Conservas e Enlatados |
| QUEIJO MUSSARELA +/- 4 KG DAVACA | 9 | R$ 22.176,12 | Laticínios e Queijos |
| KILOS | 14 | R$ 20.457,50 | Carnes Bovinas |
| ORIGINAL 600ML | 21 | R$ 18.300,82 | Cervejas |
| CAMARAO CINZA  DESCASCADO EVISCERADO 41/50 SACO 2KG FRESCATTO CXA 06 | 8 | R$ 15.646,32 | Pescados e Frutos do Mar |
| GELO 20KG | 22 | R$ 15.000,00 | Gelo |
| GLP EM CILINDRO P45 | 2 | R$ 13.500,00 | Gás de Cozinha |
| GORDURA VEGETAL 14KG FRITADELLA | 12 | R$ 13.359,45 | Óleos, Azeites e Gorduras |
| PALETAS MEXICANAS | 7 | R$ 11.032,50 | Sobremesas, Sorvetes e Picolés |
| PAO DOCE VARIADO | 440 | R$ 10.971,06 | Pães e Padaria |
| FILE DE PEITO IND CONG 18KG | 9 | R$ 10.804,86 | Aves |
| QUEIJO PARMESAO TRES ESTRELA | 2 | R$ 10.545,40 | Laticínios e Queijos |
| TIJOLO GAB REFRATARIO 22.9X114 51MM | 4 | R$ 10.090,00 | Manutenção e Obras |
| STELLA ARTOIS 600 ML | 12 | R$ 9.991,68 | Cervejas |
| COCA-COLA KS 290ML | 16 | R$ 9.945,93 | Refrigerantes |
| REFRIGERADOR,SKOL,NORMAL,220W,9643, | 3 | R$ 9.867,09 | Equipamentos e Imobilizado |
| CARNE FILE MIGNON | 5 | R$ 9.822,00 | Carnes Bovinas |
| BRAHMA CHOPP 600ML | 16 | R$ 9.820,32 | Cervejas |
| AM CRYS 500ML C/GAS | 21 | R$ 9.295,20 | Águas Minerais |
| BACON MANTA +-3,5KG RF SEARA | 11 | R$ 8.960,95 | Frios e Embutidos |
| PAPEL HIG INTERF ELITE FD 48X250 | 14 | R$ 8.803,24 | Produtos de Limpeza e Higiene |
| CC SEM ACUCAR LT 350ML COLORCX12 | 26 | R$ 8.642,65 | Refrigerantes |
| TOMATE S PELE LA BELLA SAN MARZANO 2,5KG | 4 | R$ 8.630,62 | Hortifruti |
| REFR VT VN CC HEINEKEN 8CX BIVOLT METALF | 1 | R$ 8.160,00 | Equipamentos e Imobilizado |
| LAGOSTA | 2 | R$ 8.112,50 | Pescados e Frutos do Mar |
| REFR VT VN CC HEINEKEN 8 CX 220V | 1 | R$ 8.000,00 | Equipamentos e Imobilizado |
| LULA INTEIRA | 2 | R$ 7.735,00 | Pescados e Frutos do Mar |
| COCA-COLA LT 350ML | 21 | R$ 7.734,31 | Refrigerantes |
| BATATA 9MM 2,5KG MAIS BATATA | 9 | R$ 7.683,98 | Batatas Congeladas e Pré-prontos |
| CADEIRA PLAST C/ BRACO DEVASSA LAGER | 1 | R$ 7.200,00 | Móveis, Decoração e Materiais Promocionais |
| AM CRYS 500ML S/GAS | 15 | R$ 7.184,40 | Águas Minerais |
| OLEO COMP AZEITE OLIVA 5,02L DOM GA | 12 | R$ 7.134,36 | Óleos, Azeites e Gorduras |
| AZEITE OLIVA EXT VRG 500ML OPERA N | 1 | R$ 6.801,48 | Óleos, Azeites e Gorduras |
| Embalagem para Pizza TAM 35 | 2 | R$ 6.779,76 | Descartáveis e Embalagens |
| PESC FILE SALMAO VAC 1500/2000 CG PESQUA | 2 | R$ 6.657,11 | Pescados e Frutos do Mar |
| OMBRELONE SOL PREMIUM | 1 | R$ 6.624,00 | Móveis, Decoração e Materiais Promocionais |
| REFR VT VN PC SCHIN 8CX 220V FRICON | 1 | R$ 6.500,00 | Equipamentos e Imobilizado |
| CERVEJ C VIS 8CX 220V BLK PRINC PM | 1 | R$ 6.300,00 | Equipamentos e Imobilizado |
| MASSA ESPAGUETTI 8G DURO 500G DIVELLA | 9 | R$ 6.218,88 | Mercearia Seca |
| CAMARAO DESC. E EVISC. IQF 41\50 | 1 | R$ 6.120,00 | Pescados e Frutos do Mar |
| CARNE DE SOL | 5 | R$ 6.033,15 | Carnes Bovinas |
| ONU 1075 GAS(ES) DE PETROLEO, LIQUEFEITO(S), 2.1, GPL/GLP GRANEL | 3 | R$ 5.962,40 | Gás de Cozinha |
| REFR VT VN PCV DEVSASSA 8CX BIVOLT METAL | 1 | R$ 5.890,00 | Equipamentos e Imobilizado |
| TOA ROLO PLUS FS 6X200 | 13 | R$ 5.722,86 | Produtos de Limpeza e Higiene |
| CC SEM ACUCAR KS 290ML | 10 | R$ 5.613,34 | Refrigerantes |
| REFR VT VN PC AMSTEL 8CX 220V FRICON | 1 | R$ 5.470,00 | Equipamentos e Imobilizado |
| CARNES DE SOL | 6 | R$ 5.442,75 | Carnes Bovinas |

## 9. Pontos técnicos importantes para Claude Code

### 9.1 Importador XML NF-e

Fantozzi deve criar um módulo específico para XML NF-e, por exemplo:

```text
src/fantozzi/nfe_xml_importer.py
```

Campos mínimos a extrair:

| Dado | XPath NF-e aproximado |
|---|---|
| Chave NF-e | atributo `Id` de `infNFe` |
| Número NF-e | `ide/nNF` |
| Série | `ide/serie` |
| Data emissão | `ide/dhEmi` |
| Fornecedor CNPJ | `emit/CNPJ` |
| Fornecedor nome | `emit/xNome` |
| Destinatário CNPJ | `dest/CNPJ` |
| Destinatário nome | `dest/xNome` |
| Valor total NF-e | `total/ICMSTot/vNF` |
| Item número | atributo `nItem` de `det` |
| Código produto | `det/prod/cProd` |
| Descrição produto | `det/prod/xProd` |
| NCM | `det/prod/NCM` |
| CFOP | `det/prod/CFOP` |
| Unidade | `det/prod/uCom` |
| Quantidade | `det/prod/qCom` |
| Valor unitário | `det/prod/vUnCom` |
| Valor item | `det/prod/vProd` |
| Vencimentos, quando existirem | `cobr/dup/dVenc` e `cobr/dup/vDup` |

### 9.2 Vencimentos

Nem todas as NF-e possuem vencimentos no grupo `cobr/dup`. Quando não houver vencimento no XML, Fantozzi deve registrar alerta:

```text
VENCIMENTO_NAO_INFORMADO_NO_XML
```

### 9.3 Empresas diferentes

Como foram encontrados mais de um CNPJ destinatário, o Excel deve ter colunas:

```text
cnpj_destinatario
nome_destinatario
```

Ou então Fantozzi deve criar arquivos separados por empresa. Para a primeira versão, recomendo colunas separadas, não arquivos separados.

### 9.4 Classificação

Usar o arquivo:

```text
config/plano_contas_rules.yaml
```

A versão proposta está no arquivo:

```text
FANTOZZI_plano_contas_rules_proposta.yaml
```

### 9.5 Arquivo de revisão manual

Gerar sempre um relatório com:

- produtos não classificados;
- produtos ambíguos;
- fornecedores com descrições genéricas;
- NF-e sem vencimento;
- diferenças entre soma dos itens e total da NF-e;
- possíveis duplicatas.

## 10. Recomendação

Antes de Claude Code implementar regras definitivas, eu recomendo revisar principalmente:

1. se as categorias propostas batem com seu plano de contas real;
2. se bebidas alcoólicas devem ficar separadas em cervejas, vinhos e destilados ou em uma conta única;
3. se gelo deve ser custo de bar/cozinha ou conta separada;
4. se equipamentos, móveis e materiais promocionais devem entrar no mesmo Excel mensal ou em controle de imobilizado separado;
5. como tratar NF-e sem vencimento no XML.
