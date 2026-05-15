from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
import hashlib
from typing import List

from fantozzi.models import ClassificationResult, Duplicata, Lancamento, NFE, ProdutoNFE
from fantozzi.operation_classifier import (
    BONIFICACAO_CFOPS,
    COMODATO_CFOPS,
    TIPO_BONIFICACAO,
    TIPO_COMODATO,
    TIPO_COMPRA_NORMAL,
    ClassificacaoOperacaoEspecial,
    classificar_produto_operacao_especial,
)
from fantozzi.product_classifier import ProductClassifier


CENT = Decimal("0.01")
ORIGEM_XML_NFE = "XML_NFE"
ACCESSORY_ACCOUNT_ID = "2.05.001"
DISCOUNT_ACCOUNT_ID = "2.05.002"
SPECIAL_REVIEW_ACCOUNT_ID = "13.01.001"
ALERTA_VALOR_INCONSISTENTE = "VALOR_INCONSISTENTE"


@dataclass
class _AllocationGroup:
    conta_id: str
    conta_nome: str
    grupo_conta: str
    tipo_lancamento: str
    produtos: List[ProdutoNFE] = field(default_factory=list)
    classificacoes: List[ClassificationResult] = field(default_factory=list)
    operacoes_especiais: List[ClassificacaoOperacaoEspecial] = field(default_factory=list)
    valor: Decimal = Decimal("0")
    alertas: List[str] = field(default_factory=list)
    natureza_operacao: str = ""


def allocate_nfe_to_lancamentos(
    nfe: NFE, product_classifier: ProductClassifier
) -> list[Lancamento]:
    """Create grouped managerial accounting entries from an imported NF-e.

    Products are classified individually, but the final accounting entries are
    grouped by account. Duplicatas are allocated proportionally across those
    grouped entries, preserving both due-date totals and entry totals.
    """
    groups = _group_products_by_account(nfe, product_classifier)
    if not groups:
        return []

    lancamentos = [_build_lancamento(nfe, group) for group in groups]
    _append_total_difference_lancamento(nfe, product_classifier, lancamentos)
    if nfe.duplicatas:
        dup_alertas = _allocate_duplicatas(lancamentos, nfe.duplicatas, nfe.valor_total)
    else:
        dup_alertas = []

    for lancamento in lancamentos:
        alertas = _split_alertas(lancamento.alertas)
        alertas.extend(dup_alertas)
        lancamento.alertas = _join_alertas(alertas)
        if lancamento.alertas or lancamento.status_revisao != "OK":
            lancamento.status_revisao = "REVISAR"
        lancamento.hash_registro = _hash_lancamento(lancamento)

    return lancamentos


def _group_products_by_account(
    nfe: NFE, product_classifier: ProductClassifier
) -> list[_AllocationGroup]:
    groups: OrderedDict[str, _AllocationGroup] = OrderedDict()
    nfe_tem_cfop_especial = _nfe_tem_cfop_especial(nfe)

    for produto in nfe.produtos:
        operacao = classificar_produto_operacao_especial(
            produto,
            nfe.natureza_operacao,
            tem_cobranca=_tem_cobranca(nfe),
            nfe_tem_cfop_especial=nfe_tem_cfop_especial,
        )
        if operacao.tipo == TIPO_COMPRA_NORMAL:
            classification = product_classifier.classify(
                produto.descricao, fornecedor=nfe.nome_emitente
            )
        else:
            conta_revisao = _get_classifier_account(
                product_classifier, SPECIAL_REVIEW_ACCOUNT_ID
            )
            classification = ClassificationResult(
                conta=conta_revisao,
                confianca="NAO_CLASSIFICADO",
                alertas=_operacao_especial_alertas(operacao),
                matches=list(operacao.motivos),
            )
        conta = classification.conta
        group_key = _group_key(conta.id, operacao.tipo)

        if group_key not in groups:
            groups[group_key] = _AllocationGroup(
                conta_id=conta.id,
                conta_nome=conta.nome,
                grupo_conta=conta.grupo,
                tipo_lancamento=conta.tipo,
                natureza_operacao=nfe.natureza_operacao,
            )

        group = groups[group_key]
        group.produtos.append(produto)
        group.classificacoes.append(classification)
        group.operacoes_especiais.append(operacao)
        group.valor += produto.valor_total
        group.alertas.extend(_classification_alertas(classification))

    return list(groups.values())


def _classification_alertas(classification: ClassificationResult) -> list[str]:
    alertas = list(classification.alertas)
    if classification.confianca == "NAO_CLASSIFICADO":
        alertas.append("NAO_CLASSIFICADO")
    elif classification.confianca == "AMBIGUO":
        alertas.append("CLASSIFICACAO_AMBIGUA")
    return alertas


def _operacao_especial_alertas(
    operacao: ClassificacaoOperacaoEspecial,
) -> list[str]:
    if operacao.tipo == TIPO_COMODATO:
        base = ["OPERACAO_ESPECIAL_COMODATO"]
    elif operacao.tipo == TIPO_BONIFICACAO:
        base = ["OPERACAO_ESPECIAL_BONIFICACAO"]
    else:
        base = ["OPERACAO_ESPECIAL_REVISAR"]
    return _unique([*base, *operacao.motivos, *operacao.alertas])


def _group_key(conta_id: str, tipo_operacao: str) -> str:
    suffix = tipo_operacao if tipo_operacao != TIPO_COMPRA_NORMAL else ""
    return f"{conta_id}|{suffix}"


def _build_lancamento(nfe: NFE, group: _AllocationGroup) -> Lancamento:
    status_revisao = (
        "OK"
        if all(c.confianca == "OK" for c in group.classificacoes) and not group.alertas
        else "REVISAR"
    )

    lancamento = Lancamento(
        data_documento=nfe.data_emissao,
        competencia=nfe.data_emissao.strftime("%Y-%m"),
        cnpj_empresa=nfe.cnpj_destinatario,
        nome_empresa=nfe.nome_destinatario,
        fornecedor_ou_favorecido=nfe.nome_emitente,
        cnpj_cpf_favorecido=nfe.cnpj_emitente,
        documento=nfe.numero,
        tipo_documento=ORIGEM_XML_NFE,
        conta_id=group.conta_id,
        conta_nome=group.conta_nome,
        grupo_conta=group.grupo_conta,
        descricao_itens=_describe_items(group.produtos),
        valor=_money(group.valor),
        duplicatas=[],
        origem=ORIGEM_XML_NFE,
        arquivo_origem=nfe.arquivo_origem,
        chave_origem=nfe.chave,
        status_revisao=status_revisao,
        alertas=_join_alertas(group.alertas),
        cfop=_join_produto_cfops(group.produtos),
        natureza_operacao=group.natureza_operacao,
        tipo_operacao_especial=_join_operacao_values(
            op.tipo for op in group.operacoes_especiais
        ),
        confianca_operacao_especial=_join_operacao_values(
            op.confianca for op in group.operacoes_especiais
        ),
        tratamento_operacao_especial=_join_operacao_values(
            op.destino_gerencial for op in group.operacoes_especiais
        ),
        tipo_lancamento=group.tipo_lancamento,
    )
    lancamento.hash_registro = _hash_lancamento(lancamento)
    return lancamento


def _append_total_difference_lancamento(
    nfe: NFE, product_classifier: ProductClassifier, lancamentos: list[Lancamento]
) -> None:
    soma_produtos = sum((produto.valor_total for produto in nfe.produtos), Decimal("0"))
    diferenca = _money(nfe.valor_total - soma_produtos)
    if diferenca == 0:
        return

    if diferenca > 0:
        explained = diferenca == _money(_identified_positive_adjustments(nfe))
        conta = _get_classifier_account(product_classifier, ACCESSORY_ACCOUNT_ID)
        lancamentos.append(
            _build_adjustment_lancamento(
                nfe=nfe,
                conta_id=conta.id,
                conta_nome=conta.nome,
                grupo_conta=conta.grupo,
                tipo_lancamento=conta.tipo,
                descricao_itens="Custos acessorios identificados no total da NF-e",
                valor=diferenca,
                explained=explained,
            )
        )
        return

    explained = abs(diferenca) == _money(nfe.valor_desconto)
    conta = _get_classifier_account(product_classifier, DISCOUNT_ACCOUNT_ID)
    lancamentos.append(
        _build_adjustment_lancamento(
            nfe=nfe,
            conta_id=conta.id,
            conta_nome=conta.nome,
            grupo_conta=conta.grupo,
            tipo_lancamento=conta.tipo,
            descricao_itens="Descontos obtidos identificados no total da NF-e",
            valor=diferenca,
            explained=explained,
        )
    )


def _identified_positive_adjustments(nfe: NFE) -> Decimal:
    return (
        nfe.valor_frete
        + nfe.valor_seguro
        + nfe.valor_outras_despesas
        + nfe.valor_tributos_inclusos
    )


def _get_classifier_account(product_classifier: ProductClassifier, account_id: str):
    return product_classifier._chart.get_by_id(account_id)


def _build_adjustment_lancamento(
    *,
    nfe: NFE,
    conta_id: str,
    conta_nome: str,
    grupo_conta: str,
    tipo_lancamento: str,
    descricao_itens: str,
    valor: Decimal,
    explained: bool,
) -> Lancamento:
    lancamento = Lancamento(
        data_documento=nfe.data_emissao,
        competencia=nfe.data_emissao.strftime("%Y-%m"),
        cnpj_empresa=nfe.cnpj_destinatario,
        nome_empresa=nfe.nome_destinatario,
        fornecedor_ou_favorecido=nfe.nome_emitente,
        cnpj_cpf_favorecido=nfe.cnpj_emitente,
        documento=nfe.numero,
        tipo_documento=ORIGEM_XML_NFE,
        conta_id=conta_id,
        conta_nome=conta_nome,
        grupo_conta=grupo_conta,
        descricao_itens=descricao_itens,
        valor=_money(valor),
        duplicatas=[],
        origem=ORIGEM_XML_NFE,
        arquivo_origem=nfe.arquivo_origem,
        chave_origem=nfe.chave,
        status_revisao="OK" if explained else "REVISAR",
        alertas="" if explained else ALERTA_VALOR_INCONSISTENTE,
        cfop=_join_produto_cfops(nfe.produtos),
        natureza_operacao=nfe.natureza_operacao,
        tipo_operacao_especial=TIPO_COMPRA_NORMAL,
        confianca_operacao_especial="NAO_APLICAVEL",
        tratamento_operacao_especial="DRE",
        tipo_lancamento=tipo_lancamento,
    )
    lancamento.hash_registro = _hash_lancamento(lancamento)
    return lancamento


def _describe_items(produtos: list[ProdutoNFE]) -> str:
    return "; ".join(
        f"{produto.descricao} ({_money(produto.valor_total)})" for produto in produtos
    )


def _join_produto_cfops(produtos: list[ProdutoNFE]) -> str:
    return ", ".join(_unique([produto.cfop.strip() for produto in produtos if produto.cfop]))


def _join_operacao_values(values) -> str:
    return ", ".join(_unique([str(value).strip() for value in values if str(value).strip()]))


def _nfe_tem_cfop_especial(nfe: NFE) -> bool:
    return any(
        produto.cfop.strip() in COMODATO_CFOPS
        or produto.cfop.strip() in BONIFICACAO_CFOPS
        for produto in nfe.produtos
    )


def _tem_cobranca(nfe: NFE) -> bool:
    return nfe.tem_duplicatas or nfe.tem_cobranca or nfe.valor_cobranca > 0


def _allocate_duplicatas(
    lancamentos: list[Lancamento],
    duplicatas: list[Duplicata],
    nfe_total: Decimal,
) -> list[str]:
    expected_total = _money(nfe_total)
    duplicata_total = sum((_money(d.valor) for d in duplicatas), Decimal("0"))
    alertas = [] if duplicata_total == expected_total else ["VALOR_INCONSISTENTE"]

    if expected_total == 0:
        for lancamento in lancamentos:
            lancamento.duplicatas = [
                Duplicata(numero=d.numero, vencimento=d.vencimento, valor=Decimal("0.00"))
                for d in duplicatas
            ]
        return alertas

    matrix = _build_allocation_matrix(
        row_totals=[_money(d.valor) for d in duplicatas],
        column_totals=[l.valor for l in lancamentos],
        total=expected_total,
    )

    for column_index, lancamento in enumerate(lancamentos):
        lancamento.duplicatas = [
            Duplicata(
                numero=duplicata.numero,
                vencimento=duplicata.vencimento,
                valor=matrix[row_index][column_index],
            )
            for row_index, duplicata in enumerate(duplicatas)
        ]

        allocated_total = sum((d.valor for d in lancamento.duplicatas), Decimal("0"))
        if allocated_total != lancamento.valor:
            alertas.append("VALOR_INCONSISTENTE")

    for row_index, duplicata in enumerate(duplicatas):
        allocated_row = sum(matrix[row_index], Decimal("0"))
        if allocated_row != _money(duplicata.valor):
            alertas.append("VALOR_INCONSISTENTE")

    return _unique(alertas)


def _build_allocation_matrix(
    row_totals: list[Decimal], column_totals: list[Decimal], total: Decimal
) -> list[list[Decimal]]:
    row_count = len(row_totals)
    column_count = len(column_totals)
    matrix = [
        [Decimal("0.00") for _ in range(column_count)] for _ in range(row_count)
    ]

    for row_index in range(row_count):
        for column_index in range(column_count):
            is_last_row = row_index == row_count - 1
            is_last_column = column_index == column_count - 1

            if not is_last_row and not is_last_column:
                matrix[row_index][column_index] = _money(
                    row_totals[row_index] * column_totals[column_index] / total
                )
            elif not is_last_column:
                matrix[row_index][column_index] = column_totals[column_index] - sum(
                    matrix[r][column_index] for r in range(row_index)
                )
            else:
                matrix[row_index][column_index] = row_totals[row_index] - sum(
                    matrix[row_index][c] for c in range(column_index)
                )

            matrix[row_index][column_index] = _money(matrix[row_index][column_index])

    return matrix


def _hash_lancamento(lancamento: Lancamento) -> str:
    parts = [
        lancamento.origem,
        lancamento.chave_origem,
        lancamento.conta_id,
        lancamento.tipo_operacao_especial,
        f"{_money(lancamento.valor):.2f}",
    ]
    payload = "|".join(parts)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _money(value: Decimal) -> Decimal:
    return value.quantize(CENT, rounding=ROUND_HALF_UP)


def _split_alertas(alertas: str) -> list[str]:
    if not alertas:
        return []
    return [alerta.strip() for alerta in alertas.split(";") if alerta.strip()]


def _join_alertas(alertas: list[str]) -> str:
    return "; ".join(_unique(alertas))


def _unique(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        if value not in seen:
            seen.add(value)
            result.append(value)
    return result
