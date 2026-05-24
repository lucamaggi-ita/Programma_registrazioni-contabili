# Lancamentos manuais gerenciais — operador

Este fluxo serve para despesas gerenciais que nao vieram de XML e nao devem ser
criadas automaticamente a partir de banco/comprovantes.

Use somente quando houver decisao humana clara. Antes de inserir uma linha,
verifique se a despesa ja nao esta registrada em XML ou se o comprovante bancario
nao e apenas pagamento de uma fatura XML ja contabilizada.

O projeto e DRE/conta economica. O operador informa o que foi comprado ou
contratado e o valor. Ele nao informa caixa, banco, cartao ou qualquer conta
patrimonial.

---

## Gerar template

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli gerar-template-lancamentos-manuais-gerenciais
```

Se a CLI `fantozzi` estiver instalada no PATH, tambem e possivel usar:

```powershell
fantozzi gerar-template-lancamentos-manuais-gerenciais
```

Output default:

```text
Database_Dev\_templates\lancamentos_manuais_gerenciais_template.xlsx
```

O template contem quatro abas:

- `lancamentos_manuais`: aba principal para preenchimento.
- `categorias`: contas operacionais aprovadas com dropdown derivado do plano de contas.
- `plano_contas_operativo`: leitura de apoio do plano de contas.
- `instrucoes`: regras de uso e blockers.

---

## Colunas do operador

O operador deve preencher apenas as colunas visíveis:

| Coluna | Obrigatorio | Nota |
|---|---|---|
| `DATA` | sim | formato `YYYY-MM-DD` |
| `CONTA_ESCOLHIDA` | sim | selecionado via dropdown `ID - NOME` |
| `CONTA_ID_APROVADA` | automatica | derivada de `CONTA_ESCOLHIDA` via VLOOKUP — nao editar |
| `CONTA_NOME_APROVADA` | automatica | derivada de `CONTA_ESCOLHIDA` via VLOOKUP — nao editar |
| `VALOR` | sim | valor positivo |
| `DESCRICAO_DA_DESPESA` | sim | descricao concreta da despesa |
| `OBSERVACAO` | opcional | comprovante, arquivo_origem, payment_id |

As colunas tecnicas (ocultas) sao preenchidas automaticamente pelo sistema:

```text
CONTA_CREDITO_ID       — sempre 99.99.998 (contrapartida tecnica)
CONTA_CREDITO_NOME     — nome da conta tecnica
STATUS_VALIDACAO       — estado da linha (PENDENTE / OK / REVISAR)
MENSAGEM_VALIDACAO     — mensagem de validacao
CHAVE_PREVIEW          — chave hash para deduplicacao
```

**Nao devem comparir no template operativo:**

```text
HISTORICO
DOCUMENTO
FORNECEDOR_OU_CONTRAPARTE
CATEGORIA_GERENCIAL
CNPJ_ORIGEM
CENTRO_CUSTO
FORMA_PAGAMENTO
BANCO
CAIXA
CARTAO
```

---

## Regras de preenchimento

- Informar o que foi comprado/contratado, nao como foi pago.
- Nao inserir despesas ja presentes nos XML.
- Nao inserir comprovantes que sao apenas pagamento de fatura XML ja registrada.
- Usar valor positivo.
- Nao deixar linhas vazias no meio dos dados.
- `OBSERVACAO` pode indicar comprovante, arquivo_origem ou payment_id.

---

## Executar run completa

Quando o template estiver preenchido, usar o comando one-shot:

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli executar-lancamentos-manuais-gerenciais
```

Se a CLI `fantozzi` estiver instalada no PATH, tambem e possivel usar:

```powershell
fantozzi executar-lancamentos-manuais-gerenciais
```

O comando executa em sequencia:

1. valida o template XLSX;
2. gera CSV tecnico intermediario;
3. executa run completa com XML reais + lancamentos manuais;
4. gera DRE, pre-livro/giornale e pacote contabil gerencial;
5. produz manifest e relatorio de auditoria.

Opcoes disponíveis (todas opcionais, os defaults sao corretos para uso normal):

```text
--template         Template XLSX preenchido (default: Database_Dev\_templates\...)
--saida-csv        CSV tecnico intermediario (default: Database_Dev\_templates\...)
--xml-base         Pasta raiz XML NF-e (default: Database_Dev\XML_fatture_acquisto)
--saida-output     Pasta output run (default: output\importacao_real)
--saida-reports    Pasta reports run (default: reports\importacao_real)
--plano-contas     YAML plano de contas (default: config\plano_contas_gerencial.yaml)
--mapeamento-dre   YAML mapeamento DRE (default: config\mapeamento_dre.yaml)
--relatorio-validacao  TXT relatorio de validacao (default: nao gerado)
```

O comando para imediatamente se houver erros de validacao. Nao requer CNPJ,
fornecedor, documento, centro_custo, forma_pagamento nem categoria gerencial.

---

## Outputs produzidos

```text
reports\importacao_real\runs\{run_id}\manifest_run.json                              - manifest da run
reports\importacao_real\runs\{run_id}\dre_gerencial_ok_provisoria.xlsx               - DRE gerencial
reports\importacao_real\runs\{run_id}\pre_livro_diario_gerencial_teste.xlsx          - pre-livro/giornale
reports\importacao_real\runs\{run_id}\pacote_contabil_gerencial_provisorio\          - pacote contabil gerencial
Database_Dev\_templates\..._tecnico.csv                                              - CSV tecnico intermediario
```

---

## Contrapartida tecnica

A conta `99.99.998 - Contrapartida Tecnica Lancamentos Manuais`:

- e conta de natureza tecnica, fora DRE;
- nao representa caixa, banco, cartao nem patrimonio real;
- nao e selecionavel pelo operador no dropdown;
- aparece no pre-livro/giornale como contrapartida — nao no DRE;
- se nao existir no plano de contas, a validacao bloqueia com instrucoes explicitas.

---

## Erros bloqueantes

A validacao bloqueia quando encontra:

- `DATA` ausente ou invalida;
- `CONTA_ESCOLHIDA` ausente ou inexistente no plano;
- `CONTA_ID_APROVADA` inconsistente com `CONTA_ESCOLHIDA`;
- `CONTA_NOME_APROVADA` inconsistente com `CONTA_ESCOLHIDA`;
- `DESCRICAO_DA_DESPESA` vazia;
- `VALOR` menor ou igual a zero;
- conta de contrapartida tecnica ausente do plano;
- linha vazia no meio dos dados;
- duplicata interna.

---

## Regra de ouro

O lancamento manual e excecao controlada. Ele entra no fluxo apenas depois de
revisao humana e nunca substitui XML, comprovante bancario ou regra automatica
quando essas origens ja explicam a despesa.
