from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
import xml.etree.ElementTree as ET

from fantozzi.normalization import normalize_cnpj


@dataclass(frozen=True)
class NFSE:
    tipo_documento: str
    numero: str
    codigo_verificacao: str
    chave_acesso: str
    data_emissao: date
    competencia: date | None
    cnpj_prestador: str
    razao_social_prestador: str
    cnpj_tomador: str
    razao_social_tomador: str
    valor_servicos: Decimal
    valor_liquido: Decimal
    discriminacao: str
    arquivo_origem: str


def _local_name(elem: ET.Element) -> str:
    return elem.tag.rsplit("}", 1)[-1]


def _children(elem: ET.Element | None, name: str) -> list[ET.Element]:
    if elem is None:
        return []
    return [child for child in list(elem) if _local_name(child) == name]


def _first_child(elem: ET.Element | None, name: str) -> ET.Element | None:
    matches = _children(elem, name)
    return matches[0] if matches else None


def _descendants(elem: ET.Element | None, name: str) -> list[ET.Element]:
    if elem is None:
        return []
    return [child for child in elem.iter() if _local_name(child) == name]


def _first_descendant(elem: ET.Element | None, name: str) -> ET.Element | None:
    matches = _descendants(elem, name)
    return matches[0] if matches else None


def _text(elem: ET.Element | None) -> str:
    return (elem.text or "").strip() if elem is not None else ""


def _child_text(elem: ET.Element | None, name: str) -> str:
    return _text(_first_child(elem, name))


def _descendant_text(elem: ET.Element | None, name: str) -> str:
    return _text(_first_descendant(elem, name))


def _parse_decimal(value: str) -> Decimal:
    text = (value or "").strip()
    if not text:
        return Decimal("0")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal("0")


def _parse_date(value: str) -> date | None:
    text = (value or "").strip()
    if not text:
        return None
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        return date.fromisoformat(text.split("T")[0])
    except ValueError:
        return None


def _required_date(value: str, xml_path: Path) -> date:
    parsed = _parse_date(value)
    if parsed is None:
        raise ValueError(f"NFS-e sem DataEmissao valida: {xml_path.name}")
    return parsed


def _party_cnpj(party: ET.Element | None) -> str:
    return normalize_cnpj(_descendant_text(party, "Cnpj") or _descendant_text(party, "CNPJ"))


def _party_name(party: ET.Element | None) -> str:
    return (
        _descendant_text(party, "RazaoSocial")
        or _descendant_text(party, "NomeFantasia")
        or _descendant_text(party, "xNome")
        or _descendant_text(party, "xFant")
    )


def parse_nfse(xml_path: Path | str) -> NFSE:
    """Parse a supported NFS-e XML file and return an NFSE dataclass.

    Raises ValueError when the XML is not a supported NFS-e document. The parser
    is namespace-agnostic because exporters vary the namespace URI.
    """
    xml_path = Path(xml_path)
    root = ET.parse(xml_path).getroot()
    root_name = _local_name(root)

    if root_name == "CompNfse":
        return _parse_abrasf_nfse(root, xml_path)
    if root_name == "NFSe":
        return _parse_nacional_nfse(root, xml_path)
    raise ValueError(f"Nao e uma NFS-e valida: {xml_path.name}")


def _parse_abrasf_nfse(root: ET.Element, xml_path: Path) -> NFSE:
    inf_nfse = _first_descendant(root, "InfNfse")
    if inf_nfse is None:
        raise ValueError(f"NFS-e sem InfNfse: {xml_path.name}")

    prestador = _first_descendant(inf_nfse, "PrestadorServico")
    if prestador is None:
        prestador = _first_descendant(inf_nfse, "Prestador")
    tomador = _first_descendant(inf_nfse, "TomadorServico")
    if tomador is None:
        tomador = _first_descendant(inf_nfse, "Tomador")

    valor_servicos = _parse_decimal(_descendant_text(inf_nfse, "ValorServicos"))
    valor_liquido = _parse_decimal(_descendant_text(inf_nfse, "ValorLiquidoNfse"))

    return NFSE(
        tipo_documento="NFSE",
        numero=_child_text(inf_nfse, "Numero"),
        codigo_verificacao=_child_text(inf_nfse, "CodigoVerificacao"),
        chave_acesso=_child_text(inf_nfse, "ChaveAcesso"),
        data_emissao=_required_date(_child_text(inf_nfse, "DataEmissao"), xml_path),
        competencia=_parse_date(_descendant_text(inf_nfse, "Competencia")),
        cnpj_prestador=_party_cnpj(prestador),
        razao_social_prestador=_party_name(prestador),
        cnpj_tomador=_party_cnpj(tomador),
        razao_social_tomador=_party_name(tomador),
        valor_servicos=valor_servicos,
        valor_liquido=valor_liquido,
        discriminacao=_descendant_text(inf_nfse, "Discriminacao"),
        arquivo_origem=str(xml_path),
    )


def _parse_nacional_nfse(root: ET.Element, xml_path: Path) -> NFSE:
    inf_nfse = _first_descendant(root, "infNFSe")
    if inf_nfse is None:
        raise ValueError(f"NFS-e nacional sem infNFSe: {xml_path.name}")

    inf_dps = _first_descendant(inf_nfse, "infDPS")
    if inf_dps is None:
        raise ValueError(f"NFS-e nacional sem DPS/infDPS: {xml_path.name}")

    prestador = _first_descendant(inf_dps, "prest")
    emitente = _first_descendant(inf_nfse, "emit")
    tomador = _first_descendant(inf_dps, "toma")

    valor_servicos = _parse_decimal(_descendant_text(inf_dps, "vServ"))
    valor_liquido = _parse_decimal(_descendant_text(inf_nfse, "vLiq")) or valor_servicos

    return NFSE(
        tipo_documento="NFSE",
        numero=_child_text(inf_nfse, "nNFSe") or _child_text(inf_dps, "nDPS"),
        codigo_verificacao="",
        chave_acesso=_nacional_chave_acesso(inf_nfse),
        data_emissao=_required_date(
            _child_text(inf_dps, "dhEmi") or _child_text(inf_nfse, "dhProc"),
            xml_path,
        ),
        competencia=_parse_date(_child_text(inf_dps, "dCompet")),
        cnpj_prestador=_party_cnpj(prestador) or _party_cnpj(emitente),
        razao_social_prestador=_party_name(emitente) or _party_name(prestador),
        cnpj_tomador=_party_cnpj(tomador),
        razao_social_tomador=_party_name(tomador),
        valor_servicos=valor_servicos,
        valor_liquido=valor_liquido,
        discriminacao=_descendant_text(inf_dps, "xDescServ"),
        arquivo_origem=str(xml_path),
    )


def _nacional_chave_acesso(inf_nfse: ET.Element) -> str:
    key = (inf_nfse.attrib.get("Id") or "").strip()
    return key[3:] if key.startswith("NFS") else key
