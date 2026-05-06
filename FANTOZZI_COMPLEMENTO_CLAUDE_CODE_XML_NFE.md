# Complemento para Claude Code — Fantozzi deve importar XML NF-e

Este complemento atualiza o escopo inicial de Fantozzi.

Os arquivos enviados pelo usuário não são CSV: são ZIPs contendo XML de NF-e recebidas.

## Arquivos analisados

- `NFEs_Recebidas_01-10-2025-02-04-2026_.zip`
- `NFEs_Recebidas_01-10-2025-02-04-2026_ (1).zip`

## Achados principais

- NF-e lidas: 764
- Itens/produtos lidos: 3341
- Descrições únicas de produto: 1093
- Valor total aproximado das NF-e: R$ 1.094.246,81
- Existem NF-e para mais de um CNPJ destinatário.
- Muitas NF-e possuem vencimento em `cobr/dup`, mas nem todas.

## Nova prioridade técnica

A primeira versão de Fantozzi deve implementar também:

```text
src/fantozzi/nfe_xml_importer.py
```

A CLI deve aceitar ZIP de XML NF-e:

```powershell
python -m fantozzi.cli import-nfe-xml --zip "caminho\NFEs.zip" --excel output\registrazioni_contabili.xlsx --rules config\plano_contas_rules.yaml
```

Também deve aceitar uma pasta de XML:

```powershell
python -m fantozzi.cli import-nfe-xml --xml-dir "caminho\xml" --excel output\registrazioni_contabili.xlsx --rules config\plano_contas_rules.yaml
```

## Arquivos de apoio preparados

Usar como base:

- `FANTOZZI_ANALISE_NFE_E_CATEGORIAS.md`
- `FANTOZZI_plano_contas_rules_proposta.yaml`
- `FANTOZZI_PRODUTOS_UNICOS_EXTRAIDOS.csv`

## Requisitos adicionais

1. Extrair dados da NF-e usando namespace oficial `http://www.portalfiscal.inf.br/nfe`.
2. Ignorar XML de evento, por exemplo `*-evento-*.xml`, na importação de notas.
3. Deduplicar pela chave NF-e.
4. Incluir `chave_nfe` no Excel.
5. Incluir `cnpj_destinatario` e `nome_destinatario` no Excel.
6. Gerar alerta se a NF-e não tiver vencimento em `cobr/dup`.
7. Gerar alerta se a soma dos itens divergir do total da NF-e acima de uma tolerância configurável.
8. Continuar preservando linhas manuais no Excel.
9. Continuar usando `hash_registro` para evitar duplicidade de linhas automáticas.
10. Gerar relatório de revisão para produtos não classificados e ambíguos.
