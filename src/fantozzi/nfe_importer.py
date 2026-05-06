from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path
import xml.etree.ElementTree as ET

from fantozzi.models import Duplicata, NFE, ProdutoNFE

_NS = "http://www.portalfiscal.inf.br/nfe"


def _t(name: str) -> str:
    return f"{{{_NS}}}{name}"


def _text(elem: ET.Element, tag: str, default: str = "") -> str:
    node = elem.find(_t(tag))
    return (node.text or "").strip() if node is not None else default


def _decimal(elem: ET.Element, tag: str) -> Decimal:
    return Decimal(_text(elem, tag) or "0")


def _parse_date(value: str) -> date:
    # handles "2026-05-05T10:00:00-03:00" and bare "2026-05-20"
    return date.fromisoformat(value.split("T")[0])


def parse_nfe(xml_path: Path | str) -> NFE:
    """Parse a Brazilian NF-e XML file and return an NFE dataclass.

    Raises ValueError if the file does not contain a valid <infNFe> element
    (e.g. evento XML, cancellation, etc.).
    """
    xml_path = Path(xml_path)
    root = ET.parse(xml_path).getroot()

    inf_nfe = root.find(f".//{_t('infNFe')}")
    if inf_nfe is None:
        raise ValueError(f"Não é uma NF-e válida: {xml_path.name}")

    # chave: strip the "NFe" prefix from the Id attribute
    raw_id = inf_nfe.get("Id", "")
    chave = raw_id[3:] if raw_id.startswith("NFe") else raw_id

    ide = inf_nfe.find(_t("ide"))
    emit = inf_nfe.find(_t("emit"))
    dest = inf_nfe.find(_t("dest"))
    total_elem = inf_nfe.find(_t("total"))
    cobr = inf_nfe.find(_t("cobr"))

    numero = _text(ide, "nNF")
    serie = _text(ide, "serie")
    data_emissao = _parse_date(_text(ide, "dhEmi"))

    cnpj_emitente = _text(emit, "CNPJ")
    nome_emitente = _text(emit, "xNome")

    # dest uses CNPJ for legal entities or CPF for individuals
    cnpj_destinatario = _text(dest, "CNPJ") or _text(dest, "CPF")
    nome_destinatario = _text(dest, "xNome")

    icms_tot = total_elem.find(_t("ICMSTot")) if total_elem is not None else None
    valor_total = _decimal(icms_tot, "vNF") if icms_tot is not None else Decimal("0")

    produtos = []
    for det in inf_nfe.findall(_t("det")):
        n_item = int(det.get("nItem", "0"))
        prod = det.find(_t("prod"))
        if prod is None:
            continue
        produtos.append(
            ProdutoNFE(
                n_item=n_item,
                codigo=_text(prod, "cProd"),
                descricao=_text(prod, "xProd"),
                ncm=_text(prod, "NCM"),
                cfop=_text(prod, "CFOP"),
                unidade=_text(prod, "uCom"),
                quantidade=_decimal(prod, "qCom"),
                valor_unitario=_decimal(prod, "vUnCom"),
                valor_total=_decimal(prod, "vProd"),
            )
        )

    duplicatas = []
    if cobr is not None:
        for dup in cobr.findall(_t("dup")):
            d_venc_text = _text(dup, "dVenc")
            duplicatas.append(
                Duplicata(
                    numero=_text(dup, "nDup"),
                    vencimento=_parse_date(d_venc_text) if d_venc_text else None,
                    valor=_decimal(dup, "vDup"),
                )
            )

    return NFE(
        chave=chave,
        numero=numero,
        serie=serie,
        data_emissao=data_emissao,
        cnpj_emitente=cnpj_emitente,
        nome_emitente=nome_emitente,
        cnpj_destinatario=cnpj_destinatario,
        nome_destinatario=nome_destinatario,
        valor_total=valor_total,
        produtos=produtos,
        duplicatas=duplicatas,
        tem_duplicatas=len(duplicatas) > 0,
        arquivo_origem=str(xml_path),
    )
