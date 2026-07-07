"""Audit read-only de segundo nivel: cruza comprovantes de pagamento com XML/NF-e.

Este modulo NAO classifica, NAO exclui e NAO altera nenhuma decisao contabil.
Ele apenas produz evidencia de auditoria (candidatos de match por CNPJ + valor +
proximidade de data) para apoiar decisoes futuras de Luca. Ver
FZ-COMPROVANTES-13I-AUDIT-XML-COMPROVANTES-SEGUNDO-NIVEL.md para o mandato completo.
"""

from __future__ import annotations

import csv
import json
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path

from fantozzi.csv_utils import sniff_delimiter
from fantozzi.nfe_importer import parse_nfe
from fantozzi.normalization import normalize_cnpj

STATUS_MATCH_ALTA = "MATCH_ALTA"
STATUS_MATCH_MEDIA = "MATCH_MEDIA"
STATUS_MATCH_BAIXA = "MATCH_BAIXA"
STATUS_MULTIPLOS_CANDIDATOS = "MULTIPLOS_CANDIDATOS"
STATUS_SEM_XML_NO_DATASET = "SEM_XML_NO_DATASET"
STATUS_DADOS_INSUFICIENTES = "DADOS_INSUFICIENTES"

_TIER_ORDER = {
    STATUS_MATCH_ALTA: 3,
    STATUS_MATCH_MEDIA: 2,
    STATUS_MATCH_BAIXA: 1,
}

_MENSAGEM_SEM_XML = (
    "Nenhuma XML encontrada no dataset escaneado em {timestamp}, com {total} "
    "XML lidas em {pastas}. Isso nao prova ausencia real de NF; apenas "
    "ausencia nos arquivos disponiveis neste audit."
)


class AuditXmlComprovantesError(ValueError):
    pass


@dataclass(frozen=True)
class ComprovanteRecord:
    comprovante_id: str
    data_pagamento: date | None
    valor: Decimal
    cnpj_pagador: str
    cnpj_beneficiario: str
    beneficiario: str
    status_classificacao: str
    acao: str = ""


@dataclass(frozen=True)
class XmlRecord:
    chave: str
    data_emissao: date
    valor_total: Decimal
    cnpj_emitente: str
    nome_emitente: str
    cnpj_destinatario: str
    nome_destinatario: str
    arquivo: str


@dataclass(frozen=True)
class ScanManifestEntry:
    path: str
    tamanho: int
    mtime: str


@dataclass(frozen=True)
class AuditRow:
    comprovante_id: str
    data_pagamento: str
    valor_comprovante: str
    cnpj_pagador: str
    cnpj_beneficiario: str
    beneficiario: str
    status_classificacao: str
    status_audit: str
    confidencia: str
    xml_chave: str = ""
    xml_data_emissao: str = ""
    xml_valor: str = ""
    xml_cnpj_emitente: str = ""
    xml_cnpj_destinatario: str = ""
    xml_arquivo: str = ""
    criterios_match: str = ""
    alerta_segundo_nivel: str = ""
    observacao: str = ""


@dataclass(frozen=True)
class ResumoFornecedor:
    cnpj_beneficiario: str
    beneficiario: str
    qtd_comprovantes: int
    valor_total_comprovantes: Decimal
    qtd_match_alta: int
    valor_match_alta: Decimal
    qtd_sem_xml_dataset: int
    valor_sem_xml_dataset: Decimal
    qtd_multiplos_candidatos: int
    observacoes: str = ""


@dataclass(frozen=True)
class AuditXmlComprovantesResult:
    timestamp: str
    comprovantes_csv: Path
    xml_root: Path
    dias_tolerancia: int
    fornecedor_cnpj_filtro: str
    empresa_cnpj_filtro: str
    somente_revisar: bool
    total_comprovantes_lidos: int
    total_comprovantes_avaliados: int
    total_xml_lidas: int
    total_xml_ignoradas: int
    pastas_escaneadas: tuple[str, ...]
    xml_por_pasta: dict[str, int]
    rows: tuple[AuditRow, ...] = field(default_factory=tuple)
    resumo_por_fornecedor: tuple[ResumoFornecedor, ...] = field(default_factory=tuple)
    manifest: tuple[ScanManifestEntry, ...] = field(default_factory=tuple)


def _parse_decimal(value: object) -> Decimal:
    text = str(value or "").strip()
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


def _parse_data(value: str) -> date | None:
    text = (value or "").strip()
    if not text:
        return None
    try:
        return date.fromisoformat(text.split("T")[0])
    except ValueError:
        return None


def load_comprovantes(csv_path: Path) -> list[ComprovanteRecord]:
    if not csv_path.is_file():
        raise AuditXmlComprovantesError(
            f"Arquivo de comprovantes nao encontrado: {csv_path}"
        )
    sample = csv_path.read_text(encoding="utf-8-sig")[:4096]
    delimiter = sniff_delimiter(sample)
    records: list[ComprovanteRecord] = []
    with csv_path.open(encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file, delimiter=delimiter)
        for row in reader:
            comprovante_id = (row.get("payment_id") or row.get("raw_hash") or "").strip()
            beneficiario = (
                row.get("beneficiary_name") or row.get("beneficiario_normalizado") or ""
            ).strip()
            records.append(
                ComprovanteRecord(
                    comprovante_id=comprovante_id,
                    data_pagamento=_parse_data(row.get("data_pagamento", "")),
                    valor=_parse_decimal(row.get("valor", "")),
                    cnpj_pagador=normalize_cnpj(row.get("cnpj", "")),
                    cnpj_beneficiario=normalize_cnpj(row.get("beneficiary_document", "")),
                    beneficiario=beneficiario,
                    status_classificacao=(row.get("status_classificacao") or "").strip(),
                    acao=(row.get("acao") or "").strip(),
                )
            )
    return records


def scan_xml_root(
    xml_root: Path,
) -> tuple[list[XmlRecord], dict[str, int], list[ScanManifestEntry], int]:
    if not xml_root.is_dir():
        raise AuditXmlComprovantesError(f"Pasta XML nao encontrada: {xml_root}")

    xml_records: list[XmlRecord] = []
    xml_por_pasta: dict[str, int] = {}
    manifest: list[ScanManifestEntry] = []
    total_ignoradas = 0

    for xml_path in sorted(xml_root.rglob("*.xml")):
        stat = xml_path.stat()
        rel_path = xml_path.relative_to(xml_root)
        manifest.append(
            ScanManifestEntry(
                path=str(rel_path).replace("\\", "/"),
                tamanho=stat.st_size,
                mtime=datetime.fromtimestamp(stat.st_mtime, tz=timezone.utc).isoformat(),
            )
        )
        try:
            nfe = parse_nfe(xml_path)
        except (ValueError, ET.ParseError, OSError):
            total_ignoradas += 1
            continue

        pasta = rel_path.parts[0] if rel_path.parts else "."
        xml_por_pasta[pasta] = xml_por_pasta.get(pasta, 0) + 1
        xml_records.append(
            XmlRecord(
                chave=nfe.chave,
                data_emissao=nfe.data_emissao,
                valor_total=nfe.valor_total,
                cnpj_emitente=normalize_cnpj(nfe.cnpj_emitente),
                nome_emitente=nfe.nome_emitente,
                cnpj_destinatario=normalize_cnpj(nfe.cnpj_destinatario),
                nome_destinatario=nfe.nome_destinatario,
                arquivo=str(rel_path).replace("\\", "/"),
            )
        )

    return xml_records, xml_por_pasta, manifest, total_ignoradas


def _tier_for_candidate(
    comp: ComprovanteRecord, xml: XmlRecord, dias_tolerancia: int
) -> tuple[str, tuple[str, ...]]:
    criterios = ["CNPJ_FORNECEDOR", "CNPJ_EMPRESA"]

    diff_valor = abs(xml.valor_total - comp.valor)
    valor_exato = diff_valor <= Decimal("0.01")
    limite_proximo = max(Decimal("0.01"), comp.valor * Decimal("0.02"))
    valor_proximo = (not valor_exato) and diff_valor <= limite_proximo

    date_diff = None
    if comp.data_pagamento is not None:
        date_diff = abs((xml.data_emissao - comp.data_pagamento).days)
    date_compativel = date_diff is not None and date_diff <= dias_tolerancia
    date_plausivel = (
        date_diff is not None
        and not date_compativel
        and date_diff <= dias_tolerancia * 4
    )

    if valor_exato:
        criterios.append("VALOR_EXATO")
    elif valor_proximo:
        criterios.append("VALOR_PROXIMO")
    if date_compativel:
        criterios.append("DATA_COMPATIVEL")
    elif date_plausivel:
        criterios.append("DATA_PLAUSIVEL")

    if valor_exato and date_compativel:
        tier = STATUS_MATCH_ALTA
    elif (valor_exato and date_plausivel) or (valor_proximo and date_compativel):
        tier = STATUS_MATCH_MEDIA
    else:
        tier = STATUS_MATCH_BAIXA

    return tier, tuple(criterios)


def _match_comprovante(
    comp: ComprovanteRecord,
    xml_records: list[XmlRecord],
    dias_tolerancia: int,
    timestamp: str,
    total_xml_lidas: int,
    pastas_escaneadas: tuple[str, ...],
) -> list[AuditRow]:
    base_kwargs = dict(
        comprovante_id=comp.comprovante_id,
        data_pagamento=comp.data_pagamento.isoformat() if comp.data_pagamento else "",
        valor_comprovante=f"{comp.valor:.2f}",
        cnpj_pagador=comp.cnpj_pagador,
        cnpj_beneficiario=comp.cnpj_beneficiario,
        beneficiario=comp.beneficiario,
        status_classificacao=comp.status_classificacao,
    )

    if not comp.cnpj_pagador or not comp.cnpj_beneficiario or not comp.valor:
        return [
            AuditRow(
                **base_kwargs,
                status_audit=STATUS_DADOS_INSUFICIENTES,
                confidencia="N.A.",
                observacao=(
                    "Comprovante sem CNPJ pagador, CNPJ beneficiario ou valor "
                    "suficientes para tentativa de match."
                ),
            )
        ]

    candidatos = [
        xml
        for xml in xml_records
        if xml.cnpj_emitente == comp.cnpj_beneficiario
        and xml.cnpj_destinatario == comp.cnpj_pagador
    ]

    if not candidatos:
        return [
            AuditRow(
                **base_kwargs,
                status_audit=STATUS_SEM_XML_NO_DATASET,
                confidencia="N.A.",
                observacao=_MENSAGEM_SEM_XML.format(
                    timestamp=timestamp,
                    total=total_xml_lidas,
                    pastas=", ".join(pastas_escaneadas) or "(nenhuma)",
                ),
            )
        ]

    scored = [
        (xml, *_tier_for_candidate(comp, xml, dias_tolerancia)) for xml in candidatos
    ]
    melhor_tier = max(_TIER_ORDER[tier] for _, tier, _ in scored)
    melhores = [item for item in scored if _TIER_ORDER[item[1]] == melhor_tier]

    if len(melhores) > 1:
        rows = []
        for xml, tier, criterios in melhores:
            rows.append(
                AuditRow(
                    **base_kwargs,
                    status_audit=STATUS_MULTIPLOS_CANDIDATOS,
                    confidencia=tier,
                    xml_chave=xml.chave,
                    xml_data_emissao=xml.data_emissao.isoformat(),
                    xml_valor=f"{xml.valor_total:.2f}",
                    xml_cnpj_emitente=xml.cnpj_emitente,
                    xml_cnpj_destinatario=xml.cnpj_destinatario,
                    xml_arquivo=xml.arquivo,
                    criterios_match=";".join(criterios),
                    observacao=(
                        f"{len(melhores)} XML candidatas com o mesmo nivel de "
                        "confianca para este comprovante; nenhuma escolhida "
                        "automaticamente."
                    ),
                )
            )
        return rows

    xml, tier, criterios = melhores[0]
    return [
        AuditRow(
            **base_kwargs,
            status_audit=tier,
            confidencia=tier,
            xml_chave=xml.chave,
            xml_data_emissao=xml.data_emissao.isoformat(),
            xml_valor=f"{xml.valor_total:.2f}",
            xml_cnpj_emitente=xml.cnpj_emitente,
            xml_cnpj_destinatario=xml.cnpj_destinatario,
            xml_arquivo=xml.arquivo,
            criterios_match=";".join(criterios),
        )
    ]


def _alerta_segundo_nivel(comp: ComprovanteRecord, status_audit: str) -> str:
    status = comp.status_classificacao.upper()
    acao = comp.acao.upper()
    tem_candidata = status_audit in (
        STATUS_MATCH_ALTA,
        STATUS_MATCH_MEDIA,
        STATUS_MATCH_BAIXA,
        STATUS_MULTIPLOS_CANDIDATOS,
    )

    if status.startswith("REVISAR") and tem_candidata:
        return "REVISAR_COM_XML_CANDIDATA"
    if ("XML" in status or "XML" in acao) and status_audit == STATUS_SEM_XML_NO_DATASET:
        return "XML_SEM_XML_ENCONTRADA"
    if status.startswith("CLASSIFICADO") and tem_candidata:
        return "CLASSIFICADO_COM_XML_CANDIDATA"
    if status_audit == STATUS_SEM_XML_NO_DATASET:
        return "SEM_XML_NO_DATASET"
    return ""


def run_audit(
    comprovantes_csv: Path,
    xml_root: Path,
    *,
    dias_tolerancia: int = 7,
    fornecedor_cnpj: str | None = None,
    empresa_cnpj: str | None = None,
    somente_revisar: bool = False,
) -> AuditXmlComprovantesResult:
    timestamp = datetime.now(timezone.utc).isoformat()

    comprovantes = load_comprovantes(comprovantes_csv)
    total_lidos = len(comprovantes)

    fornecedor_filtro = normalize_cnpj(fornecedor_cnpj) if fornecedor_cnpj else ""
    empresa_filtro = normalize_cnpj(empresa_cnpj) if empresa_cnpj else ""

    avaliados = comprovantes
    if fornecedor_filtro:
        avaliados = [c for c in avaliados if c.cnpj_beneficiario == fornecedor_filtro]
    if empresa_filtro:
        avaliados = [c for c in avaliados if c.cnpj_pagador == empresa_filtro]
    if somente_revisar:
        avaliados = [
            c for c in avaliados if c.status_classificacao.upper().startswith("REVISAR")
        ]

    xml_records, xml_por_pasta, manifest, total_ignoradas = scan_xml_root(xml_root)
    pastas_escaneadas = tuple(sorted(xml_por_pasta.keys()))

    rows: list[AuditRow] = []
    comprovante_status: dict[str, tuple[ComprovanteRecord, str]] = {}
    for comp in avaliados:
        comp_rows = _match_comprovante(
            comp,
            xml_records,
            dias_tolerancia,
            timestamp,
            len(xml_records),
            pastas_escaneadas,
        )
        status_geral = comp_rows[0].status_audit
        alerta = _alerta_segundo_nivel(comp, status_geral)
        comp_rows = [
            AuditRow(**{**row.__dict__, "alerta_segundo_nivel": alerta})
            for row in comp_rows
        ]
        rows.extend(comp_rows)
        comprovante_status[comp.comprovante_id or id(comp)] = (comp, status_geral)

    resumo = build_resumo_por_fornecedor(comprovante_status.values())

    return AuditXmlComprovantesResult(
        timestamp=timestamp,
        comprovantes_csv=comprovantes_csv,
        xml_root=xml_root,
        dias_tolerancia=dias_tolerancia,
        fornecedor_cnpj_filtro=fornecedor_filtro,
        empresa_cnpj_filtro=empresa_filtro,
        somente_revisar=somente_revisar,
        total_comprovantes_lidos=total_lidos,
        total_comprovantes_avaliados=len(avaliados),
        total_xml_lidas=len(xml_records),
        total_xml_ignoradas=total_ignoradas,
        pastas_escaneadas=pastas_escaneadas,
        xml_por_pasta=xml_por_pasta,
        rows=tuple(rows),
        resumo_por_fornecedor=resumo,
        manifest=tuple(manifest),
    )


def build_resumo_por_fornecedor(
    itens,
) -> tuple[ResumoFornecedor, ...]:
    por_cnpj: dict[str, dict[str, object]] = {}
    for comp, status_geral in itens:
        chave = comp.cnpj_beneficiario or "(sem_cnpj)"
        bucket = por_cnpj.setdefault(
            chave,
            {
                "beneficiario": comp.beneficiario,
                "qtd_comprovantes": 0,
                "valor_total_comprovantes": Decimal("0"),
                "qtd_match_alta": 0,
                "valor_match_alta": Decimal("0"),
                "qtd_sem_xml_dataset": 0,
                "valor_sem_xml_dataset": Decimal("0"),
                "qtd_multiplos_candidatos": 0,
            },
        )
        bucket["qtd_comprovantes"] += 1
        bucket["valor_total_comprovantes"] += comp.valor
        if status_geral == STATUS_MATCH_ALTA:
            bucket["qtd_match_alta"] += 1
            bucket["valor_match_alta"] += comp.valor
        elif status_geral == STATUS_SEM_XML_NO_DATASET:
            bucket["qtd_sem_xml_dataset"] += 1
            bucket["valor_sem_xml_dataset"] += comp.valor
        elif status_geral == STATUS_MULTIPLOS_CANDIDATOS:
            bucket["qtd_multiplos_candidatos"] += 1

    resumo = []
    for cnpj_beneficiario, bucket in sorted(por_cnpj.items()):
        resumo.append(
            ResumoFornecedor(
                cnpj_beneficiario=cnpj_beneficiario,
                beneficiario=str(bucket["beneficiario"]),
                qtd_comprovantes=int(bucket["qtd_comprovantes"]),
                valor_total_comprovantes=bucket["valor_total_comprovantes"],
                qtd_match_alta=int(bucket["qtd_match_alta"]),
                valor_match_alta=bucket["valor_match_alta"],
                qtd_sem_xml_dataset=int(bucket["qtd_sem_xml_dataset"]),
                valor_sem_xml_dataset=bucket["valor_sem_xml_dataset"],
                qtd_multiplos_candidatos=int(bucket["qtd_multiplos_candidatos"]),
            )
        )
    return tuple(resumo)


def console_lines(result: AuditXmlComprovantesResult) -> list[str]:
    lines = [
        "[OK] Audit XML x comprovantes (segundo nivel, read-only) concluido.",
        f"Comprovantes CSV: {result.comprovantes_csv}",
        f"XML root: {result.xml_root}",
        f"Timestamp: {result.timestamp}",
        f"Dias tolerancia: {result.dias_tolerancia}",
        f"Comprovantes lidos: {result.total_comprovantes_lidos}",
        f"Comprovantes avaliados: {result.total_comprovantes_avaliados}",
        f"XML lidas: {result.total_xml_lidas} (ignoradas: {result.total_xml_ignoradas})",
        f"Pastas XML escaneadas: {', '.join(result.pastas_escaneadas) or '(nenhuma)'}",
    ]
    contagem: dict[str, int] = {}
    for row in result.rows:
        contagem[row.status_audit] = contagem.get(row.status_audit, 0) + 1
    for status, qtd in sorted(contagem.items()):
        lines.append(f"  {status}: {qtd}")
    lines.append(
        "[INFO] Este audit e evidencia, nao decisao contabil. Nenhuma regra "
        "YAML ou classificacao foi alterada."
    )
    return lines


def report_markdown(result: AuditXmlComprovantesResult) -> str:
    lines = [
        "# Audit XML x Comprovantes - Segundo Nivel (read-only)",
        "",
        f"- Timestamp audit: {result.timestamp}",
        f"- Comprovantes CSV: `{result.comprovantes_csv}`",
        f"- Pastas XML escaneadas: `{result.xml_root}` -> "
        f"{', '.join(result.pastas_escaneadas) or '(nenhuma)'}",
        f"- Total XML lidas: {result.total_xml_lidas} "
        f"(ignoradas/nao NF-e validas: {result.total_xml_ignoradas})",
        f"- Total comprovantes lidos: {result.total_comprovantes_lidos}",
        f"- Total comprovantes avaliados (apos filtros): {result.total_comprovantes_avaliados}",
        f"- Dias de tolerancia de data: {result.dias_tolerancia}",
        f"- Filtro fornecedor CNPJ: {result.fornecedor_cnpj_filtro or '(nenhum)'}",
        f"- Filtro empresa CNPJ: {result.empresa_cnpj_filtro or '(nenhum)'}",
        f"- Somente revisar: {result.somente_revisar}",
        "",
        "## Criterios de matching",
        "",
        "- Chave primaria: CNPJ emitente XML = CNPJ beneficiario comprovante, "
        "e CNPJ destinatario XML = CNPJ pagador comprovante.",
        "- MATCH_ALTA: valor exato + data dentro da tolerancia.",
        "- MATCH_MEDIA: valor exato + data fora da tolerancia mas plausivel, "
        "ou valor proximo + data dentro da tolerancia.",
        "- MATCH_BAIXA: CNPJs batem mas valor/data fora dos criterios acima.",
        "- MULTIPLOS_CANDIDATOS: mais de uma XML no mesmo nivel de confianca; "
        "nenhuma escolhida automaticamente.",
        "- SEM_XML_NO_DATASET: nenhuma XML com CNPJs compativeis encontrada "
        "no dataset escaneado (snapshot, nao prova contabil definitiva).",
        "- DADOS_INSUFICIENTES: comprovante sem CNPJ/valor suficientes.",
        "",
        "## Resumo por fornecedor",
        "",
        "| CNPJ beneficiario | Beneficiario | Qtd comprovantes | Valor total | "
        "Qtd MATCH_ALTA | Valor MATCH_ALTA | Qtd SEM_XML | Valor SEM_XML | "
        "Qtd MULTIPLOS |",
        "|---|---|---|---|---|---|---|---|---|",
    ]
    for r in result.resumo_por_fornecedor:
        lines.append(
            f"| {r.cnpj_beneficiario} | {r.beneficiario} | {r.qtd_comprovantes} | "
            f"{r.valor_total_comprovantes:.2f} | {r.qtd_match_alta} | "
            f"{r.valor_match_alta:.2f} | {r.qtd_sem_xml_dataset} | "
            f"{r.valor_sem_xml_dataset:.2f} | {r.qtd_multiplos_candidatos} |"
        )

    lines += [
        "",
        "## Alertas",
        "",
    ]
    alertas = sorted({row.alerta_segundo_nivel for row in result.rows if row.alerta_segundo_nivel})
    if alertas:
        for alerta in alertas:
            qtd = sum(1 for row in result.rows if row.alerta_segundo_nivel == alerta)
            lines.append(f"- {alerta}: {qtd} linha(s)")
    else:
        lines.append("- Nenhum alerta de segundo nivel gerado nesta execucao.")

    lines += [
        "",
        "## Limitacoes",
        "",
        "- Este controle e read-only: nao classifica, nao exclui e nao altera "
        "YAML/output existentes.",
        "- SEM_XML_NO_DATASET reflete apenas o dataset escaneado no momento do "
        "audit, nao prova ausencia real de nota fiscal.",
        "- Match por CNPJ + valor + data e evidencia de auditoria; a decisao "
        "contabil permanece com Luca.",
    ]
    return "\n".join(lines) + "\n"


_CSV_HEADERS = [
    "comprovante_id",
    "data_pagamento",
    "valor_comprovante",
    "cnpj_pagador",
    "cnpj_beneficiario",
    "beneficiario",
    "status_classificacao",
    "status_audit",
    "confidencia",
    "xml_chave",
    "xml_data_emissao",
    "xml_valor",
    "xml_cnpj_emitente",
    "xml_cnpj_destinatario",
    "xml_arquivo",
    "criterios_match",
    "alerta_segundo_nivel",
    "observacao",
]

_RESUMO_HEADERS = [
    "cnpj_beneficiario",
    "beneficiario",
    "qtd_comprovantes",
    "valor_total_comprovantes",
    "qtd_match_alta",
    "valor_match_alta",
    "qtd_sem_xml_dataset",
    "valor_sem_xml_dataset",
    "qtd_multiplos_candidatos",
    "observacoes",
]


def write_outputs(result: AuditXmlComprovantesResult, output_dir: Path) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)

    md_path = output_dir / "audit_xml_comprovantes.md"
    md_path.write_text(report_markdown(result), encoding="utf-8")

    csv_path = output_dir / "audit_xml_comprovantes.csv"
    with csv_path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=_CSV_HEADERS, delimiter=";")
        writer.writeheader()
        for row in result.rows:
            writer.writerow({header: getattr(row, header) for header in _CSV_HEADERS})

    resumo_path = output_dir / "audit_xml_comprovantes_resumo_por_fornecedor.csv"
    with resumo_path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=_RESUMO_HEADERS, delimiter=";")
        writer.writeheader()
        for r in result.resumo_por_fornecedor:
            writer.writerow(
                {
                    "cnpj_beneficiario": r.cnpj_beneficiario,
                    "beneficiario": r.beneficiario,
                    "qtd_comprovantes": r.qtd_comprovantes,
                    "valor_total_comprovantes": f"{r.valor_total_comprovantes:.2f}",
                    "qtd_match_alta": r.qtd_match_alta,
                    "valor_match_alta": f"{r.valor_match_alta:.2f}",
                    "qtd_sem_xml_dataset": r.qtd_sem_xml_dataset,
                    "valor_sem_xml_dataset": f"{r.valor_sem_xml_dataset:.2f}",
                    "qtd_multiplos_candidatos": r.qtd_multiplos_candidatos,
                    "observacoes": r.observacoes,
                }
            )

    manifest_path = output_dir / "audit_xml_comprovantes_manifest.json"
    manifest_payload = {
        "timestamp": result.timestamp,
        "comprovantes_csv": str(result.comprovantes_csv),
        "xml_root": str(result.xml_root),
        "dias_tolerancia": result.dias_tolerancia,
        "total_comprovantes_lidos": result.total_comprovantes_lidos,
        "total_comprovantes_avaliados": result.total_comprovantes_avaliados,
        "total_xml_lidas": result.total_xml_lidas,
        "total_xml_ignoradas": result.total_xml_ignoradas,
        "pastas_escaneadas": list(result.pastas_escaneadas),
        "xml_por_pasta": result.xml_por_pasta,
        "arquivos_xml": [
            {"path": entry.path, "tamanho": entry.tamanho, "mtime": entry.mtime}
            for entry in result.manifest
        ],
    }
    manifest_path.write_text(
        json.dumps(manifest_payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    return {
        "markdown": md_path,
        "csv": csv_path,
        "resumo": resumo_path,
        "manifest": manifest_path,
    }
