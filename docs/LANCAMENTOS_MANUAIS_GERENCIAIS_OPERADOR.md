# Lancamentos manuais gerenciais - operador

Este fluxo serve para despesas gerenciais que nao vieram de XML e nao devem ser
criadas automaticamente a partir de banco/comprovantes.

Use somente quando houver decisao humana clara. Antes de inserir uma linha,
verifique se a despesa ja nao esta registrada em XML ou se o comprovante bancario
nao e apenas pagamento de uma fatura XML ja contabilizada.

O projeto e DRE/conta economica. O operador informa o que foi comprado ou
contratado e o valor. Ele nao informa caixa, banco, cartao ou qualquer conta
patrimonial.

## Gerar template

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli gerar-template-lancamentos-manuais-gerenciais --saida Database_Dev\_templates\lancamentos_manuais_gerenciais_TEMPLATE.xlsx
```

O template contem:

- `lancamentos_manuais`: aba principal para preenchimento.
- `categorias`: categorias operacionais aprovadas e suas contas de debito.
- `plano_contas_operativo`: leitura de apoio do plano de contas.
- `instrucoes`: regras de uso e blockers.

O CSV tecnico usa uma conta de contrapartida apenas para compatibilidade com o
pipeline atual. Essa conta deve existir explicitamente no plano de contas, estar
fora DRE, ser neutra/auditavel e nao representar caixa, banco, cartao ou
patrimonio real.

Se nao existir uma conta tecnica segura, a validacao deve bloquear. A correcao
nesse caso e uma mini-fase separada para criar a conta de contrapartida tecnica
de forma explicita no plano/mapeamento, nao um improviso dentro do template.

## Colunas do operador

Preencha apenas estas colunas:

- `DATA`: data em formato `YYYY-MM-DD`.
- `HISTORICO`: descricao concreta da despesa.
- `DOCUMENTO`: numero ou referencia interna. Para dinheiro, use padrao como
  `CASSA-2026-05-001`.
- `FORNECEDOR_OU_CONTRAPARTE`: fornecedor, pessoa ou contraparte.
- `CATEGORIA_GERENCIAL`: escolha da lista.
- `VALOR`: valor positivo.
- `CNPJ_ORIGEM`: CNPJ da empresa quando aplicavel. Opcional para pagamentos em
  dinheiro ou fornecedores sem CNPJ. Se vazio, o sistema avisa mas nao bloqueia.
- `CENTRO_CUSTO`: opcional.
- `OBSERVACAO`: origem do comprovante, payment_id, arquivo bancario ou nota de
  revisao.

As colunas tecnicas `CONTA_*`, `STATUS_VALIDACAO`, `MENSAGEM_VALIDACAO` e
`CHAVE_PREVIEW` sao calculadas/ocultas e nao devem ser editadas.

## Validar e gerar CSV tecnico

```powershell
.\.venv\Scripts\python.exe -m fantozzi.cli validar-lancamentos-manuais-gerenciais --entrada Database_Dev\_templates\lancamentos_manuais_gerenciais_TEMPLATE.xlsx --saida-csv reports\lancamentos_manuais_gerenciais.csv --relatorio reports\validacao_lancamentos_manuais_gerenciais.txt
```

O CSV tecnico gerado usa UTF-8-SIG, separador `;` e colunas compativeis com o
pipeline existente:

```text
data;historico;documento;fornecedor_ou_contraparte;conta_debito_id;conta_debito_nome;conta_credito_id;conta_credito_nome;valor;cnpj_origem_opcional;centro_custo_opcional;observacao
```

## Erros bloqueantes

A validacao bloqueia o arquivo quando encontra:

- data ausente ou invalida;
- historico vazio;
- fornecedor/contraparte vazio;
- categoria ausente, inexistente ou inativa;
- valor menor ou igual a zero;
- conta de debito inexistente;
- regra DRE invalida para a conta economica;
- linha vazia no meio dos dados;
- duplicata interna.

## Avisos

Avisos nao bloqueiam a geracao do CSV, mas devem ser revisados:

- documento vazio;
- CNPJ vazio;
- observacao vazia;
- historico ou fornecedor generico;
- risco de despesa ja existir em XML.

## Regra de ouro

O lancamento manual e excecao controlada. Ele entra no fluxo apenas depois de
revisao humana e nunca substitui XML, comprovante bancario ou regra automatica
quando essas origens ja explicam a despesa.
