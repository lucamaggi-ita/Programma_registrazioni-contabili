from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
import unicodedata

from fantozzi.models import NFE, ProdutoNFE
from fantozzi.nfe_importer import parse_nfe


TIPO_COMPRA_NORMAL = "COMPRA_NORMAL"
TIPO_COMODATO = "COMODATO"
TIPO_BONIFICACAO = "BONIFICACAO_RECEBIDA"
TIPO_OUTRA_REVISAO = "OUTRA_OPERACAO_ESPECIAL_REVISAR"

COMODATO_CFOPS = frozenset({"1908", "2908", "5908", "6908", "1909", "2909", "5909", "6909"})
BONIFICACAO_CFOPS = frozenset({"1910", "2910", "5910", "6910"})
BONIFICACAO_TERMS = ("BONIFICACAO", "DOACAO", "BRINDE")
EQUIPAMENTO_TERMS = (
    "CERVEJEIRA",
    "REFRIGERADOR",
    "FREEZER",
    "GELADEIRA",
    "CHOPEIRA",
    "DISPENSER",
    "DISPLAY",
    "EQUIPAMENTO",
)


@dataclass(frozen=True)
class ClassificacaoOperacaoEspecial:
    tipo: str
    confianca: str
    destino_gerencial: str
    status_revisao: str
    motivos: tuple[str, ...]
    alertas: tuple[str, ...] = ()


@dataclass(frozen=True)
class ProdutoOperacaoEspecial:
    descricao: str
    cfop: str
    valor: Decimal
    motivos: tuple[str, ...]
    alertas: tuple[str, ...]


@dataclass(frozen=True)
class OperacaoEspecialAuditItem:
    nfe: NFE
    classificacao: ClassificacaoOperacaoEspecial
    produtos: tuple[ProdutoOperacaoEspecial, ...]


@dataclass(frozen=True)
class OperacaoEspecialAuditResult:
    pasta: Path
    total_xml: int
    xml_validos: int
    xml_ignorados: int
    itens: tuple[OperacaoEspecialAuditItem, ...]
    erros: tuple[str, ...]


class OperationAuditError(Exception):
    pass


def normalizar_texto_operacao(texto: str) -> str:
    texto = unicodedata.normalize("NFD", texto or "")
    texto = "".join(ch for ch in texto if unicodedata.category(ch) != "Mn")
    texto = texto.upper()
    texto = "".join(ch if ch.isalnum() else " " for ch in texto)
    return " ".join(texto.split())


def classificar_operacao_especial(nfe: NFE) -> ClassificacaoOperacaoEspecial:
    natureza = normalizar_texto_operacao(nfe.natureza_operacao)
    cfops = {produto.cfop.strip() for produto in nfe.produtos if produto.cfop.strip()}

    tem_cfop_comodato = bool(cfops & COMODATO_CFOPS)
    tem_natureza_comodato = "COMODATO" in natureza
    tem_natureza_locacao = "LOCACAO" in natureza
    tem_cfop_bonificacao = bool(cfops & BONIFICACAO_CFOPS)
    tem_natureza_bonificacao = any(term in natureza for term in BONIFICACAO_TERMS)

    motivos: list[str] = []
    alertas: list[str] = []
    if tem_natureza_locacao and not tem_natureza_comodato:
        alertas.append("POSSIVEL_LOCACAO_OU_COMODATO")
    if _tem_alerta_equipamento(nfe.produtos):
        alertas.append("POSSIVEL_EQUIPAMENTO_OU_BEM_DE_TERCEIRO")
    if not _tem_cobranca(nfe) and (tem_cfop_bonificacao or tem_natureza_bonificacao):
        alertas.append("SEM_DUPLICATA_OU_COBRANCA_IDENTIFICADA")

    sinal_comodato = tem_cfop_comodato or tem_natureza_comodato
    sinal_bonificacao = tem_cfop_bonificacao or tem_natureza_bonificacao
    if sinal_comodato and sinal_bonificacao:
        motivos.append("SINAIS_CONFLITANTES_COMODATO_E_BONIFICACAO")
        return ClassificacaoOperacaoEspecial(
            tipo=TIPO_OUTRA_REVISAO,
            confianca="REVISAR",
            destino_gerencial="REVISAR",
            status_revisao="ANALISAR_COM_CONTADOR",
            motivos=tuple(motivos),
            alertas=tuple(_unique(alertas)),
        )

    if tem_cfop_comodato and tem_natureza_comodato:
        motivos.append("CFOP_COMODATO_E_NATUREZA_COMODATO")
        return ClassificacaoOperacaoEspecial(
            tipo=TIPO_COMODATO,
            confianca="FORTE",
            destino_gerencial="FORA_DRE",
            status_revisao="ANALISAR_COM_CONTADOR",
            motivos=tuple(motivos),
            alertas=tuple(_unique(alertas)),
        )
    if tem_cfop_comodato:
        motivos.append("CFOP_COMODATO")
        alertas.append("NATUREZA_OPERACAO_NAO_EXPLICITA_COMODATO")
        return ClassificacaoOperacaoEspecial(
            tipo=TIPO_COMODATO,
            confianca="PROVAVEL",
            destino_gerencial="FORA_DRE",
            status_revisao="ANALISAR_COM_CONTADOR",
            motivos=tuple(motivos),
            alertas=tuple(_unique(alertas)),
        )
    if tem_natureza_comodato:
        motivos.append("NATUREZA_COMODATO_CFOP_INESPERADO")
        alertas.append("CFOP_NAO_COMPATIVEL_COM_COMODATO")
        return ClassificacaoOperacaoEspecial(
            tipo=TIPO_COMODATO,
            confianca="REVISAR",
            destino_gerencial="FORA_DRE",
            status_revisao="ANALISAR_COM_CONTADOR",
            motivos=tuple(motivos),
            alertas=tuple(_unique(alertas)),
        )

    if tem_cfop_bonificacao and tem_natureza_bonificacao:
        motivos.append("CFOP_BONIFICACAO_E_NATUREZA_COMPATIVEL")
        return ClassificacaoOperacaoEspecial(
            tipo=TIPO_BONIFICACAO,
            confianca="FORTE",
            destino_gerencial="FORA_CUSTO_COMPRA",
            status_revisao="ANALISAR_COM_CONTADOR",
            motivos=tuple(motivos),
            alertas=tuple(_unique(alertas)),
        )
    if tem_cfop_bonificacao:
        motivos.append("CFOP_BONIFICACAO")
        alertas.append("NATUREZA_OPERACAO_NAO_EXPLICITA_BONIFICACAO")
        return ClassificacaoOperacaoEspecial(
            tipo=TIPO_BONIFICACAO,
            confianca="PROVAVEL",
            destino_gerencial="FORA_CUSTO_COMPRA",
            status_revisao="ANALISAR_COM_CONTADOR",
            motivos=tuple(motivos),
            alertas=tuple(_unique(alertas)),
        )
    if tem_natureza_bonificacao:
        motivos.append("NATUREZA_BONIFICACAO_CFOP_INESPERADO")
        alertas.append("CFOP_NAO_COMPATIVEL_COM_BONIFICACAO")
        return ClassificacaoOperacaoEspecial(
            tipo=TIPO_BONIFICACAO,
            confianca="REVISAR",
            destino_gerencial="FORA_CUSTO_COMPRA",
            status_revisao="ANALISAR_COM_CONTADOR",
            motivos=tuple(motivos),
            alertas=tuple(_unique(alertas)),
        )

    if alertas:
        return ClassificacaoOperacaoEspecial(
            tipo=TIPO_COMPRA_NORMAL,
            confianca="REVISAR",
            destino_gerencial="DRE",
            status_revisao="ANALISAR_COM_CONTADOR",
            motivos=("ALERTA_SECUNDARIO_SEM_CFOP_NATUREZA_ESPECIAL",),
            alertas=tuple(_unique(alertas)),
        )

    return ClassificacaoOperacaoEspecial(
        tipo=TIPO_COMPRA_NORMAL,
        confianca="NAO_APLICAVEL",
        destino_gerencial="DRE",
        status_revisao="OK",
        motivos=(),
        alertas=(),
    )


def auditar_operacoes_especiais(pasta: Path | str) -> OperacaoEspecialAuditResult:
    pasta = Path(pasta)
    if not pasta.exists():
        raise OperationAuditError(f"Pasta nao encontrada: {pasta}")
    if not pasta.is_dir():
        raise OperationAuditError(f"Caminho nao e uma pasta: {pasta}")

    xml_files = sorted(
        (path for path in pasta.iterdir() if path.is_file() and path.suffix.lower() == ".xml"),
        key=lambda path: path.name.lower(),
    )
    itens: list[OperacaoEspecialAuditItem] = []
    erros: list[str] = []
    validos = 0

    for xml_path in xml_files:
        try:
            nfe = parse_nfe(xml_path)
            validos += 1
        except Exception as exc:
            erros.append(f"{xml_path.name}: {exc}")
            continue

        classificacao = classificar_operacao_especial(nfe)
        if (
            classificacao.tipo != TIPO_COMPRA_NORMAL
            or classificacao.status_revisao != "OK"
        ):
            itens.append(
                OperacaoEspecialAuditItem(
                    nfe=nfe,
                    classificacao=classificacao,
                    produtos=_produtos_relevantes(nfe, classificacao),
                )
            )

    return OperacaoEspecialAuditResult(
        pasta=pasta,
        total_xml=len(xml_files),
        xml_validos=validos,
        xml_ignorados=len(erros),
        itens=tuple(itens),
        erros=tuple(erros),
    )


def relatorio_operacoes_especiais(result: OperacaoEspecialAuditResult) -> str:
    counts = {
        TIPO_COMPRA_NORMAL: result.xml_validos - len(result.itens),
        TIPO_COMODATO: 0,
        TIPO_BONIFICACAO: 0,
        TIPO_OUTRA_REVISAO: 0,
    }
    for item in result.itens:
        counts[item.classificacao.tipo] = counts.get(item.classificacao.tipo, 0) + 1

    lines = [
        "FANTOZZI - OPERACOES ESPECIAIS",
        "================================",
        f"Pasta XML: {result.pasta}",
        "",
        "Resumo",
        "------",
        f"XML encontrados: {result.total_xml}",
        f"XML validos: {result.xml_validos}",
        f"XML ignorados: {result.xml_ignorados}",
        f"Compra normal: {counts.get(TIPO_COMPRA_NORMAL, 0)} NF-e",
        f"Comodato: {counts.get(TIPO_COMODATO, 0)} NF-e",
        f"Bonificacao recebida: {counts.get(TIPO_BONIFICACAO, 0)} NF-e",
        f"Outras operacoes especiais a revisar: {counts.get(TIPO_OUTRA_REVISAO, 0)} NF-e",
        "",
    ]

    if result.erros:
        lines.extend(["XML ignorados", "-------------", *result.erros, ""])

    if result.itens:
        lines.extend(["Detalhes", "--------"])
        for index, item in enumerate(result.itens, start=1):
            nfe = item.nfe
            classificacao = item.classificacao
            cfops = ", ".join(sorted({produto.cfop for produto in nfe.produtos if produto.cfop}))
            produtos_texto = "; ".join(
                f"{produto.descricao} | CFOP {produto.cfop} | R$ {_money_text(produto.valor)}"
                for produto in item.produtos
            )
            lines.extend(
                [
                    f"CASO {index}",
                    "------",
                    f"CNPJ destinatario: {nfe.cnpj_destinatario}",
                    f"Fornecedor: {nfe.nome_emitente}",
                    f"CNPJ fornecedor: {nfe.cnpj_emitente}",
                    f"Numero NF-e: {nfe.numero}",
                    f"Serie: {nfe.serie}",
                    f"Chave NF-e: {nfe.chave}",
                    f"Data emissao: {nfe.data_emissao.isoformat()}",
                    f"Natureza operacao: {nfe.natureza_operacao}",
                    f"CFOPs encontrados: {cfops}",
                    f"Tipo operacao especial: {classificacao.tipo}",
                    f"Confianca: {classificacao.confianca}",
                    f"Destino gerencial sugerido: {classificacao.destino_gerencial}",
                    f"Motivos: {_join_or_dash(classificacao.motivos)}",
                    f"Alertas: {_join_or_dash(classificacao.alertas)}",
                    f"Produtos envolvidos: {produtos_texto}",
                    f"Valor dos itens: R$ {_money_text(sum((p.valor for p in item.produtos), Decimal('0')))}",
                    f"Valor total da NF-e: R$ {_money_text(nfe.valor_total)}",
                    f"Presenca de duplicata/cobranca: {_duplicata_text(nfe)}",
                    "Sugestao: ANALISAR_COM_CONTADOR",
                    "",
                ]
            )

    lines.extend(
        [
            "Seguranca",
            "---------",
            "- Nenhum config/*.yaml foi alterado.",
            "- Nenhum Excel foi alterado.",
            "- Nenhum CSV foi reescrito.",
            "- Nenhuma DRE foi alterada.",
        ]
    )
    return "\n".join(lines) + "\n"


def salvar_relatorio(path: Path | str, text: str) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def console_lines(result: OperacaoEspecialAuditResult) -> list[str]:
    total_especiais = len(result.itens)
    return [
        "[OK] Auditoria de operacoes especiais concluida",
        f"XML encontrados: {result.total_xml}",
        f"XML validos: {result.xml_validos}",
        f"XML ignorados: {result.xml_ignorados}",
        f"Operacoes especiais/revisao: {total_especiais}",
    ]


def _produtos_relevantes(
    nfe: NFE, classificacao: ClassificacaoOperacaoEspecial
) -> tuple[ProdutoOperacaoEspecial, ...]:
    produtos = []
    for produto in nfe.produtos:
        motivos, alertas = _motivos_produto(produto)
        if classificacao.tipo != TIPO_COMPRA_NORMAL or motivos or alertas:
            produtos.append(
                ProdutoOperacaoEspecial(
                    descricao=produto.descricao,
                    cfop=produto.cfop,
                    valor=produto.valor_total,
                    motivos=motivos,
                    alertas=alertas,
                )
            )
    return tuple(produtos)


def _motivos_produto(produto: ProdutoNFE) -> tuple[tuple[str, ...], tuple[str, ...]]:
    motivos: list[str] = []
    alertas: list[str] = []
    if produto.cfop in COMODATO_CFOPS:
        motivos.append("CFOP_COMODATO")
    if produto.cfop in BONIFICACAO_CFOPS:
        motivos.append("CFOP_BONIFICACAO")
    if _descricao_tem_equipamento(produto.descricao):
        alertas.append("POSSIVEL_EQUIPAMENTO_OU_BEM_DE_TERCEIRO")
    return tuple(motivos), tuple(alertas)


def _tem_alerta_equipamento(produtos: list[ProdutoNFE]) -> bool:
    return any(_descricao_tem_equipamento(produto.descricao) for produto in produtos)


def _descricao_tem_equipamento(descricao: str) -> bool:
    normalized = normalizar_texto_operacao(descricao)
    return any(term in normalized for term in EQUIPAMENTO_TERMS) or "REFR VT" in normalized


def _duplicata_text(nfe: NFE) -> str:
    if not nfe.duplicatas and not nfe.tem_cobranca:
        return "duplicatas: nao; cobranca: nao identificada"
    if not nfe.duplicatas:
        return (
            "duplicatas: nao; fatura/liquido: "
            f"R$ {_money_text(nfe.valor_cobranca)}; cobranca: sim"
        )
    total = sum((duplicata.valor for duplicata in nfe.duplicatas), Decimal("0"))
    return f"duplicatas: sim ({len(nfe.duplicatas)}); valor duplicatas: R$ {_money_text(total)}"


def _tem_cobranca(nfe: NFE) -> bool:
    return nfe.tem_duplicatas or nfe.tem_cobranca or nfe.valor_cobranca > 0


def _money_text(value: Decimal) -> str:
    return f"{value:.2f}".replace(".", ",")


def _join_or_dash(values: tuple[str, ...]) -> str:
    return "; ".join(values) if values else "-"


def _unique(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        if value not in seen:
            seen.add(value)
            result.append(value)
    return result
