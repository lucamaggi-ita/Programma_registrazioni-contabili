from __future__ import annotations

import csv
import hashlib
import json
import shutil
import xml.etree.ElementTree as ET
from dataclasses import asdict, dataclass
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from fantozzi.nfe_importer import parse_nfe
from fantozzi.nfse_importer import parse_nfse
from fantozzi.normalization import normalize_cnpj


STATUS_COPIARIA = "COPIARIA"
STATUS_COPIADO = "COPIADO"
STATUS_JA_EXISTE = "JA_EXISTE"
STATUS_DUPLICADO = "DUPLICADO"
STATUS_IGNORADO_CNPJ_TOMADOR_NAO_SUPORTADO = "IGNORADO_CNPJ_TOMADOR_NAO_SUPORTADO"
STATUS_IGNORADO_XML_NAO_SUPORTADO = "IGNORADO_XML_NAO_SUPORTADO"
STATUS_ERRO_PARSE = "ERRO_PARSE"

SUPPORTED_TOMADOR_CNPJS = ("08884180000123", "55746722000181")
CSV_NAME = "manifest_nfstock_ingest.csv"
JSON_NAME = "manifest_nfstock_ingest.json"


@dataclass(frozen=True)
class FiscalXmlCandidate:
    tipo_documento: str
    source_path: Path
    cnpj_tomador: str
    cnpj_emitente_ou_prestador: str
    nome_emitente_ou_prestador: str
    numero: str
    data_emissao: str
    valor_total: Decimal
    chave: str
    dedup_key: str


@dataclass(frozen=True)
class NfstockIngestRow:
    status: str
    tipo_documento: str
    source_path: str
    dest_path: str
    cnpj_tomador: str
    cnpj_emitente_ou_prestador: str
    nome_emitente_ou_prestador: str
    numero: str
    data_emissao: str
    valor_total: str
    chave: str
    motivo: str


@dataclass(frozen=True)
class NfstockIngestResult:
    source_dir: Path
    xml_base: Path
    output_dir: Path
    apply: bool
    rows: tuple[NfstockIngestRow, ...]
    csv_path: Path
    json_path: Path


class NfstockIngestError(ValueError):
    pass


def default_output_dir(base: Path | None = None) -> Path:
    root = base or Path("output")
    return root / f"nfstock_ingest_{datetime.now():%Y%m%d_%H%M%S}"


def ingest_nfstock_xml(
    source_dir: Path,
    xml_base: Path,
    *,
    output_dir: Path | None = None,
    apply: bool = False,
    supported_tomador_cnpjs: tuple[str, ...] = SUPPORTED_TOMADOR_CNPJS,
) -> NfstockIngestResult:
    source_dir = Path(source_dir)
    xml_base = Path(xml_base)
    output_dir = Path(output_dir) if output_dir is not None else default_output_dir()

    if not source_dir.is_dir():
        raise NfstockIngestError(f"Source dir nao encontrada: {source_dir}")
    if not xml_base.is_dir():
        raise NfstockIngestError(f"XML base nao encontrada: {xml_base}")

    output_dir.mkdir(parents=True, exist_ok=True)
    supported = tuple(normalize_cnpj(cnpj) for cnpj in supported_tomador_cnpjs)
    existing_keys = _scan_existing_keys(xml_base, supported)
    seen_source_keys: set[str] = set()
    rows: list[NfstockIngestRow] = []

    for source_path in sorted(source_dir.rglob("*.xml")):
        candidate, error_status, motivo = _parse_candidate(source_path)
        if candidate is None:
            rows.append(
                _row(
                    status=error_status,
                    source_path=source_path,
                    motivo=motivo,
                )
            )
            continue

        dest_path = xml_base / candidate.cnpj_tomador / "xml" / source_path.name
        if candidate.cnpj_tomador not in supported:
            rows.append(
                _row_from_candidate(
                    candidate,
                    dest_path,
                    STATUS_IGNORADO_CNPJ_TOMADOR_NAO_SUPORTADO,
                    "CNPJ tomador fora do perimetro desta fase.",
                )
            )
            continue

        if dest_path.exists():
            rows.append(
                _row_from_candidate(
                    candidate,
                    dest_path,
                    STATUS_JA_EXISTE,
                    "Arquivo de destino ja existe; nenhuma sobrescrita permitida.",
                )
            )
            seen_source_keys.add(candidate.dedup_key)
            continue

        if candidate.dedup_key in existing_keys or candidate.dedup_key in seen_source_keys:
            rows.append(
                _row_from_candidate(
                    candidate,
                    dest_path,
                    STATUS_DUPLICADO,
                    "Duplicado por chave fiscal/conteudo em source ou destino.",
                )
            )
            continue

        seen_source_keys.add(candidate.dedup_key)
        if apply:
            dest_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(candidate.source_path, dest_path)
            status = STATUS_COPIADO
            motivo = "Copiado com --apply."
        else:
            status = STATUS_COPIARIA
            motivo = "Dry-run: nenhuma copia executada."
        rows.append(_row_from_candidate(candidate, dest_path, status, motivo))

    csv_path = output_dir / CSV_NAME
    json_path = output_dir / JSON_NAME
    _write_csv(csv_path, rows)
    _write_json(
        json_path,
        source_dir=source_dir,
        xml_base=xml_base,
        output_dir=output_dir,
        apply=apply,
        rows=rows,
    )
    return NfstockIngestResult(
        source_dir=source_dir,
        xml_base=xml_base,
        output_dir=output_dir,
        apply=apply,
        rows=tuple(rows),
        csv_path=csv_path,
        json_path=json_path,
    )


def console_lines(result: NfstockIngestResult) -> list[str]:
    counts: dict[str, int] = {}
    for row in result.rows:
        counts[row.status] = counts.get(row.status, 0) + 1
    mode = "APPLY" if result.apply else "DRY-RUN"
    lines = [
        f"[OK] Ingest NFSTOCK XML concluido em modo {mode}.",
        f"Source dir: {result.source_dir}",
        f"XML base: {result.xml_base}",
        f"Output dir: {result.output_dir}",
        f"XML avaliados: {len(result.rows)}",
    ]
    for status, count in sorted(counts.items()):
        lines.append(f"  {status}: {count}")
    lines.append(f"Manifest CSV: {result.csv_path}")
    lines.append(f"Manifest JSON: {result.json_path}")
    if not result.apply:
        lines.append("[DRY-RUN] Nenhum XML foi copiado.")
    return lines


def _parse_candidate(
    source_path: Path,
) -> tuple[FiscalXmlCandidate | None, str, str]:
    try:
        root = ET.parse(source_path).getroot()
    except (ET.ParseError, OSError) as exc:
        return None, STATUS_ERRO_PARSE, str(exc)

    local_names = {_local_name(elem) for elem in root.iter()}
    if "infNFe" in local_names or _local_name(root) == "nfeProc":
        try:
            nfe = parse_nfe(source_path)
        except (ValueError, ET.ParseError, OSError) as exc:
            return None, STATUS_ERRO_PARSE, str(exc)
        chave = nfe.chave or _sha256(source_path)
        return (
            FiscalXmlCandidate(
                tipo_documento="NFE",
                source_path=source_path,
                cnpj_tomador=normalize_cnpj(nfe.cnpj_destinatario),
                cnpj_emitente_ou_prestador=normalize_cnpj(nfe.cnpj_emitente),
                nome_emitente_ou_prestador=nfe.nome_emitente,
                numero=nfe.numero,
                data_emissao=nfe.data_emissao.isoformat(),
                valor_total=nfe.valor_total,
                chave=chave,
                dedup_key=f"NFE:{chave}",
            ),
            "",
            "",
        )

    if _local_name(root) == "CompNfse" or "InfNfse" in local_names:
        try:
            nfse = parse_nfse(source_path)
        except (ValueError, ET.ParseError, OSError) as exc:
            return None, STATUS_ERRO_PARSE, str(exc)
        chave = nfse.chave_acesso or _nfse_synthetic_key(nfse)
        return (
            FiscalXmlCandidate(
                tipo_documento="NFSE",
                source_path=source_path,
                cnpj_tomador=nfse.cnpj_tomador,
                cnpj_emitente_ou_prestador=nfse.cnpj_prestador,
                nome_emitente_ou_prestador=nfse.razao_social_prestador,
                numero=nfse.numero,
                data_emissao=nfse.data_emissao.isoformat(),
                valor_total=nfse.valor_liquido or nfse.valor_servicos,
                chave=chave,
                dedup_key=f"NFSE:{chave}",
            ),
            "",
            "",
        )

    return None, STATUS_IGNORADO_XML_NAO_SUPORTADO, "XML fiscal nao suportado."


def _scan_existing_keys(xml_base: Path, supported: tuple[str, ...]) -> set[str]:
    keys: set[str] = set()
    for cnpj in supported:
        dest_dir = xml_base / cnpj / "xml"
        if not dest_dir.is_dir():
            continue
        for xml_path in sorted(dest_dir.rglob("*.xml")):
            candidate, _, _ = _parse_candidate(xml_path)
            if candidate is not None:
                keys.add(candidate.dedup_key)
    return keys


def _row(
    *,
    status: str,
    source_path: Path,
    motivo: str,
) -> NfstockIngestRow:
    return NfstockIngestRow(
        status=status,
        tipo_documento="",
        source_path=str(source_path),
        dest_path="",
        cnpj_tomador="",
        cnpj_emitente_ou_prestador="",
        nome_emitente_ou_prestador="",
        numero="",
        data_emissao="",
        valor_total="",
        chave="",
        motivo=motivo,
    )


def _row_from_candidate(
    candidate: FiscalXmlCandidate,
    dest_path: Path,
    status: str,
    motivo: str,
) -> NfstockIngestRow:
    return NfstockIngestRow(
        status=status,
        tipo_documento=candidate.tipo_documento,
        source_path=str(candidate.source_path),
        dest_path=str(dest_path),
        cnpj_tomador=candidate.cnpj_tomador,
        cnpj_emitente_ou_prestador=candidate.cnpj_emitente_ou_prestador,
        nome_emitente_ou_prestador=candidate.nome_emitente_ou_prestador,
        numero=candidate.numero,
        data_emissao=candidate.data_emissao,
        valor_total=f"{candidate.valor_total:.2f}",
        chave=candidate.chave,
        motivo=motivo,
    )


def _write_csv(path: Path, rows: list[NfstockIngestRow]) -> None:
    headers = tuple(NfstockIngestRow.__dataclass_fields__.keys())
    with path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=headers, delimiter=";")
        writer.writeheader()
        for row in rows:
            writer.writerow(asdict(row))


def _write_json(
    path: Path,
    *,
    source_dir: Path,
    xml_base: Path,
    output_dir: Path,
    apply: bool,
    rows: list[NfstockIngestRow],
) -> None:
    payload = {
        "source_dir": str(source_dir),
        "xml_base": str(xml_base),
        "output_dir": str(output_dir),
        "apply": apply,
        "totais": _status_counts(rows),
        "rows": [asdict(row) for row in rows],
        "zip_support": "NAO_SUPORTADO_NESTA_FASE",
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _status_counts(rows: list[NfstockIngestRow]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for row in rows:
        counts[row.status] = counts.get(row.status, 0) + 1
    return counts


def _nfse_synthetic_key(nfse) -> str:
    return "|".join(
        (
            "NFSE",
            nfse.cnpj_prestador,
            nfse.cnpj_tomador,
            nfse.numero,
            nfse.data_emissao.isoformat(),
            f"{(nfse.valor_liquido or nfse.valor_servicos):.2f}",
        )
    )


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _local_name(elem: ET.Element) -> str:
    return elem.tag.rsplit("}", 1)[-1]
