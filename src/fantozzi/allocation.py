from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
import hashlib
from typing import List

from fantozzi.models import ClassificationResult, Duplicata, Lancamento, NFE, ProdutoNFE
from fantozzi.product_classifier import ProductClassifier


CENT = Decimal("0.01")
ORIGEM_XML_NFE = "XML_NFE"


@dataclass
class _AllocationGroup:
    conta_id: str
    conta_nome: str
    grupo_conta: str
    tipo_lancamento: str
    produtos: List[ProdutoNFE] = field(default_factory=list)
    classificacoes: List[ClassificationResult] = field(default_factory=list)
    valor: Decimal = Decimal("0")
    alertas: List[str] = field(default_factory=list)


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
    valor_alertas = _adjust_lancamento_totals(lancamentos, nfe.valor_total)
    if nfe.duplicatas:
        dup_alertas = _allocate_duplicatas(lancamentos, nfe.duplicatas, nfe.valor_total)
    else:
        dup_alertas = []

    for lancamento in lancamentos:
        alertas = _split_alertas(lancamento.alertas)
        alertas.extend(valor_alertas)
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

    for produto in nfe.produtos:
        classification = product_classifier.classify(produto.descricao)
        conta = classification.conta

        if conta.id not in groups:
            groups[conta.id] = _AllocationGroup(
                conta_id=conta.id,
                conta_nome=conta.nome,
                grupo_conta=conta.grupo,
                tipo_lancamento=conta.tipo,
            )

        group = groups[conta.id]
        group.produtos.append(produto)
        group.classificacoes.append(classification)
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
        tipo_lancamento=group.tipo_lancamento,
    )
    lancamento.hash_registro = _hash_lancamento(lancamento)
    return lancamento


def _describe_items(produtos: list[ProdutoNFE]) -> str:
    return "; ".join(
        f"{produto.descricao} ({_money(produto.valor_total)})" for produto in produtos
    )


def _adjust_lancamento_totals(
    lancamentos: list[Lancamento], nfe_total: Decimal
) -> list[str]:
    expected_total = _money(nfe_total)
    current_total = sum((l.valor for l in lancamentos), Decimal("0"))
    difference = expected_total - current_total
    if difference == 0:
        return []

    last = lancamentos[-1]
    if last.valor + difference < 0:
        return ["VALOR_INCONSISTENTE"]

    last.valor = _money(last.valor + difference)
    return [] if abs(difference) <= CENT else ["VALOR_INCONSISTENTE"]


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
