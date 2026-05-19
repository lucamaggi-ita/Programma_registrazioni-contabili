#!/usr/bin/env python3
"""Generator for Regole Registrazioni Bancarie e Comprovantes PDF."""

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import yaml
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

_MATCH_FIELDS = [
    "beneficiario_favorecido_contem_algum",
    "referencia_empresa_contem_algum",
    "tipo_pagamento_contem_algum",
    "descricao_comprovante_contem_algum",
]

HEADER_COLOR_REGRAS = colors.HexColor("#1B4F72")
HEADER_COLOR_EXCL = colors.HexColor("#5D4037")
GRAY_FALSE = colors.HexColor("#C8C8C8")
TEXT_GRAY = colors.HexColor("#555555")


def load_yaml(regras_path):
    with open(regras_path, encoding="utf-8") as f:
        data = yaml.safe_load(f)
    regras = data.get("regras", [])
    all_exclusoes = data.get("exclusoes", [])
    exclusoes_xml = [e for e in all_exclusoes if e.get("acao") == "XML"]
    exclusoes_escludere = [e for e in all_exclusoes if e.get("acao") == "ESCLUDERE"]
    return regras, exclusoes_xml, exclusoes_escludere


def validate_inputs(regras, exclusoes_xml, exclusoes_escludere):
    errors = []
    all_exclusoes = exclusoes_xml + exclusoes_escludere

    # V2: required fields in regras
    regra_required = {"regra_id", "conta_id", "conta_nome", "dre_linha_id", "entra_dre", "match"}
    for r in regras:
        missing = regra_required - set(r.keys())
        if missing:
            errors.append(f"V2 {r.get('regra_id', '?')} campos faltantes: {missing}")

    # V3: required fields and forbidden fields in exclusoes
    excl_required = {"regra_id", "acao", "match"}
    excl_forbidden = {"conta_id", "conta_nome", "dre_linha_id", "entra_dre"}
    for e in all_exclusoes:
        missing = excl_required - set(e.keys())
        if missing:
            errors.append(f"V3 {e.get('regra_id', '?')} campos faltantes: {missing}")
        forbidden = excl_forbidden & set(e.keys())
        if forbidden:
            errors.append(f"V3 {e.get('regra_id', '?')} campos proibidos: {forbidden}")

    # V4: unique regra_ids across the entire YAML
    all_ids = [r.get("regra_id") for r in regras + all_exclusoes]
    seen, dups = set(), set()
    for rid in all_ids:
        if rid in seen:
            dups.add(rid)
        seen.add(rid)
    if dups:
        errors.append(f"V4 regra_ids duplicados: {sorted(dups)}")

    # V5: at least one non-empty match criterion per entry
    for item in regras + all_exclusoes:
        match = item.get("match") or {}
        if not any(match.get(f) for f in _MATCH_FIELDS):
            errors.append(f"V5 {item.get('regra_id', '?')} sem criterio match nao vazio")

    return errors


def _safe(text):
    s = str(text) if text is not None else ""
    return s.encode("latin-1", "replace").decode("latin-1")


def _base_style(header_bg):
    return [
        ("BACKGROUND", (0, 0), (-1, 0), header_bg),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 7),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 6.5),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#AAAAAA")),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]


def _add_alt_rows(style_cmds, n_data_rows):
    for i in range(1, n_data_rows + 1):
        if i % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F2F2F2")))


def _get_pdf_dimensions(pdf_path):
    try:
        with open(pdf_path, "rb") as f:
            content = f.read()
        m = re.search(
            rb"/MediaBox\s*\[\s*([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)", content
        )
        if m:
            w = float(m.group(3)) - float(m.group(1))
            h = float(m.group(4)) - float(m.group(2))
            return w, h
    except Exception:
        pass
    return None, None


def _p(text, style):
    return Paragraph(_safe(str(text) if text is not None else ""), style)


def _fmt_list(lst, style):
    if not lst:
        return Paragraph("[]", style)
    return Paragraph("<br/>".join(_safe(str(v)) for v in lst), style)


def _make_styles():
    base = getSampleStyleSheet()
    title = ParagraphStyle(
        "RBTitle", parent=base["Normal"],
        fontName="Helvetica-Bold", fontSize=12,
        alignment=TA_CENTER, spaceAfter=4 * mm,
    )
    section = ParagraphStyle(
        "RBSection", parent=base["Normal"],
        fontName="Helvetica-Bold", fontSize=9,
        spaceBefore=3 * mm, spaceAfter=2 * mm,
    )
    note = ParagraphStyle(
        "RBNote", parent=base["Normal"],
        fontName="Helvetica-Oblique", fontSize=7,
        textColor=colors.HexColor("#7D3C98"),
        spaceAfter=2 * mm,
    )
    header = ParagraphStyle(
        "RBHeader", parent=base["Normal"],
        fontName="Helvetica-Bold", fontSize=7,
        leading=9, textColor=colors.white, alignment=TA_LEFT,
    )
    cell = ParagraphStyle(
        "RBCell", parent=base["Normal"],
        fontName="Helvetica", fontSize=6.5,
        leading=8, alignment=TA_LEFT,
    )
    return title, section, note, header, cell


def _build_regras_table(regras, usable_w, header_style, cell_style):
    raw = [130, 40, 95, 38, 42, 140, 120, 90, 90]
    scale = usable_w / sum(raw)
    col_widths = [w * scale for w in raw]

    headers = [
        "regra_id", "conta_id", "conta_nome", "dre_linha_id", "entra_dre",
        "beneficiario", "referencia", "tipo_pagamento", "descricao",
    ]
    data = [[_p(h, header_style) for h in headers]]
    gray_rows = []

    for i, r in enumerate(regras, start=1):
        match = r.get("match") or {}
        entra = r.get("entra_dre", True)
        data.append([
            _p(r.get("regra_id", ""), cell_style),
            _p(r.get("conta_id", ""), cell_style),
            _p(r.get("conta_nome", ""), cell_style),
            _p(r.get("dre_linha_id", ""), cell_style),
            _p(entra, cell_style),
            _fmt_list(match.get("beneficiario_favorecido_contem_algum"), cell_style),
            _fmt_list(match.get("referencia_empresa_contem_algum"), cell_style),
            _fmt_list(match.get("tipo_pagamento_contem_algum"), cell_style),
            _fmt_list(match.get("descricao_comprovante_contem_algum"), cell_style),
        ])
        if entra is False:
            gray_rows.append(i)

    style_cmds = _base_style(HEADER_COLOR_REGRAS)
    _add_alt_rows(style_cmds, len(regras))
    for ri in gray_rows:
        style_cmds += [
            ("BACKGROUND", (0, ri), (-1, ri), GRAY_FALSE),
            ("TEXTCOLOR", (0, ri), (-1, ri), TEXT_GRAY),
        ]

    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle(style_cmds))
    return table


def _build_exclusoes_table(items, usable_w, header_style, cell_style):
    raw = [185, 52, 185, 155, 110, 98]
    scale = usable_w / sum(raw)
    col_widths = [w * scale for w in raw]

    headers = ["regra_id", "acao", "beneficiario", "referencia", "tipo_pagamento", "descricao"]
    data = [[_p(h, header_style) for h in headers]]

    for item in items:
        match = item.get("match") or {}
        data.append([
            _p(item.get("regra_id", ""), cell_style),
            _p(item.get("acao", ""), cell_style),
            _fmt_list(match.get("beneficiario_favorecido_contem_algum"), cell_style),
            _fmt_list(match.get("referencia_empresa_contem_algum"), cell_style),
            _fmt_list(match.get("tipo_pagamento_contem_algum"), cell_style),
            _fmt_list(match.get("descricao_comprovante_contem_algum"), cell_style),
        ])

    style_cmds = _base_style(HEADER_COLOR_EXCL)
    _add_alt_rows(style_cmds, len(items))

    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle(style_cmds))
    return table


def build_pdf(regras, exclusoes_xml, exclusoes_escludere, output_path):
    """Build 3-section PDF. Returns (page_count, regra_ids, xml_ids, escludere_ids)."""
    page_size = landscape(A4)
    margin = 10 * mm
    usable_w = page_size[0] - 2 * margin

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=page_size,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=14 * mm,
        bottomMargin=10 * mm,
    )

    title_style, section_style, note_style, header_style, cell_style = _make_styles()

    TITLE = "REGOLE REGISTRAZIONI BANCARIE E COMPROVANTES"
    NOTE_EXCL = "Fuori DRE — non generano lancamentos."

    table_regras = _build_regras_table(regras, usable_w, header_style, cell_style)
    table_xml = _build_exclusoes_table(exclusoes_xml, usable_w, header_style, cell_style)
    table_escl = _build_exclusoes_table(exclusoes_escludere, usable_w, header_style, cell_style)

    story = [
        Paragraph(TITLE, title_style),
        Paragraph(f"Sezione 1 — Regole Positive ({len(regras)} regole)", section_style),
        table_regras,
        PageBreak(),
        Paragraph(TITLE, title_style),
        Paragraph(f"Sezione 2 — Exclusoes XML ({len(exclusoes_xml)} esclusioni)", section_style),
        Paragraph(NOTE_EXCL, note_style),
        table_xml,
        PageBreak(),
        Paragraph(TITLE, title_style),
        Paragraph(
            f"Sezione 3 — Exclusoes ESCLUDERE ({len(exclusoes_escludere)} esclusioni)",
            section_style,
        ),
        Paragraph(NOTE_EXCL, note_style),
        table_escl,
    ]

    page_tracker = [0]

    def on_page(canv, doc):
        page_tracker[0] = doc.page
        canv.saveState()
        canv.setFont("Helvetica", 6)
        canv.drawRightString(page_size[0] - margin, 6 * mm, f"Pagina {doc.page}")
        canv.restoreState()

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)

    regra_ids = [r.get("regra_id", "") for r in regras]
    xml_ids = [e.get("regra_id", "") for e in exclusoes_xml]
    escludere_ids = [e.get("regra_id", "") for e in exclusoes_escludere]
    return page_tracker[0], regra_ids, xml_ids, escludere_ids


def write_manifest(output_dir, regra_ids, xml_ids, escludere_ids):
    manifest = {
        "regra_ids_renderizados": regra_ids,
        "exclusao_xml_ids_renderizados": xml_ids,
        "exclusao_escludere_ids_renderizados": escludere_ids,
        "total_regras": len(regra_ids),
        "total_exclusoes_xml": len(xml_ids),
        "total_exclusoes_escludere": len(escludere_ids),
        "total_exclusoes": len(xml_ids) + len(escludere_ids),
    }
    path = Path(output_dir) / "manifest_regras_banca.json"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    return path


def validate_manifest(manifest_path, regras, exclusoes_xml, exclusoes_escludere):
    with open(manifest_path, encoding="utf-8") as f:
        m = json.load(f)
    divs = []

    exp_regras = [r.get("regra_id", "") for r in regras]
    exp_xml = [e.get("regra_id", "") for e in exclusoes_xml]
    exp_escl = [e.get("regra_id", "") for e in exclusoes_escludere]

    checks = [
        ("total_regras", len(exp_regras), m.get("total_regras")),
        ("total_exclusoes_xml", len(exp_xml), m.get("total_exclusoes_xml")),
        ("total_exclusoes_escludere", len(exp_escl), m.get("total_exclusoes_escludere")),
    ]
    for label, expected, got in checks:
        if expected != got:
            divs.append(f"{label}: manifest={got} vs YAML={expected}")

    id_checks = [
        ("regra_ids", exp_regras, m.get("regra_ids_renderizados", [])),
        ("exclusao_xml_ids", exp_xml, m.get("exclusao_xml_ids_renderizados", [])),
        ("exclusao_escludere_ids", exp_escl, m.get("exclusao_escludere_ids_renderizados", [])),
    ]
    for label, expected, rendered in id_checks:
        miss = set(expected) - set(rendered)
        if miss:
            divs.append(f"{label} ausentes no manifest: {sorted(miss)}")

    return divs


def write_summary(output_dir, stats):
    path = Path(output_dir) / "resumo_geracao_pdf_regras_banca.txt"
    lines = [
        f"Data/ora generazione:        {stats['datetime']}",
        f"YAML regras:                 {stats['regras_path']}",
        f"Totale regole positive:      {stats['total_regras']}",
        f"Totale exclusoes XML:        {stats['total_xml']}",
        f"Totale exclusoes ESCLUDERE:  {stats['total_escl']}",
        f"Totale exclusoes:            {stats['total_xml'] + stats['total_escl']}",
        f"Pagine PDF:                  {stats['pages']}",
        f"Orientamento:                landscape",
        f"Alertas:                     {stats['alertas']}",
    ]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return path


def main():
    parser = argparse.ArgumentParser(
        description="Gera PDF delle Regole Registrazioni Bancarie e Comprovantes"
    )
    parser.add_argument("--regras", required=True, help="regras_classificacao_comprovantes.yaml")
    parser.add_argument("--output", required=True, help="Output directory")
    args = parser.parse_args()

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    regras, exclusoes_xml, exclusoes_escludere = load_yaml(args.regras)
    print(
        f"Loaded: {len(regras)} regras, {len(exclusoes_xml)} exclusoes XML, "
        f"{len(exclusoes_escludere)} exclusoes ESCLUDERE"
    )

    errors = validate_inputs(regras, exclusoes_xml, exclusoes_escludere)
    for e in errors:
        print(f"ERRO: {e}", file=sys.stderr)
    if errors:
        sys.exit(1)
    print("V2-V5: OK")

    pdf_path = output_dir / "regras_registrazioni_banca.pdf"
    print("Gerando PDF...")
    pages, regra_ids, xml_ids, escludere_ids = build_pdf(
        regras, exclusoes_xml, exclusoes_escludere, pdf_path
    )

    manifest_path = write_manifest(output_dir, regra_ids, xml_ids, escludere_ids)

    divs = validate_manifest(manifest_path, regras, exclusoes_xml, exclusoes_escludere)
    if divs:
        for d in divs:
            print(f"MANIFEST DIVERGENCIA: {d}", file=sys.stderr)
        sys.exit(1)
    print("Manifest: OK")

    w, h = _get_pdf_dimensions(str(pdf_path))
    if w is not None:
        orientation = "landscape OK" if w > h else "PORTRAIT (ERRO)"
        print(f"Orientacao: {orientation} ({w:.0f}x{h:.0f}pt)")
        if w < h:
            print("ERRO: PDF nao e landscape!", file=sys.stderr)
            sys.exit(1)

    summary_path = write_summary(
        output_dir,
        {
            "datetime": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "regras_path": args.regras,
            "total_regras": len(regras),
            "total_xml": len(exclusoes_xml),
            "total_escl": len(exclusoes_escludere),
            "pages": pages,
            "alertas": "nenhum",
        },
    )

    print(f"\nPDF: {pdf_path} ({pages} pag.)")
    print(f"Manifest: {manifest_path}")
    print(f"Resumo TXT: {summary_path}")
    print("\nGeracao concluida.")


if __name__ == "__main__":
    main()
