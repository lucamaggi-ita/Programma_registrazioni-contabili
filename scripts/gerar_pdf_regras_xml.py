#!/usr/bin/env python3
"""Generator for Regras e Aliases Classificacao XML/NF-e PDF."""

import argparse
import json
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import yaml
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

EXPECTED_KEYS = {
    "versao", "idioma", "descricao", "normalizacao", "politica_de_match",
    "regras_por_fornecedor", "regras", "conta_fallback_id", "conta_fallback_nome",
}

HDR_ALIASES = colors.HexColor("#1B4F72")
HDR_FORNEC = colors.HexColor("#145A32")
HDR_EVITAR = colors.HexColor("#7D3C98")
HDR_RESUMO = colors.HexColor("#2E4053")
EVITAR_BG = colors.HexColor("#FEF9E7")
EVITAR_TEXT = colors.HexColor("#7D6608")
ALT_ROW = colors.HexColor("#F2F2F2")
GRAY_TEXT = colors.HexColor("#555555")


def load_yaml(regras_path):
    with open(regras_path, encoding="utf-8") as f:
        data = yaml.safe_load(f)
    regras = data.get("regras", [])
    fornecedores = data.get("regras_por_fornecedor", [])
    fallback = {
        "id": data.get("conta_fallback_id", ""),
        "nome": data.get("conta_fallback_nome", ""),
    }
    unmapped = sorted(set(data.keys()) - EXPECTED_KEYS)
    return regras, fornecedores, fallback, unmapped


def validate_inputs(regras, fornecedores):
    errors = []
    warnings = []

    # V2: sezioni presenti
    if not regras:
        warnings.append("V2: sezione 'regras' vuota o assente")
    if not fornecedores:
        warnings.append("V2: sezione 'regras_por_fornecedor' vuota o assente")

    # V3: regras senza conta_id o aliases
    for r in regras:
        if not r.get("conta_id"):
            errors.append(f"V3: regra senza conta_id: {r}")
        if not r.get("aliases"):
            errors.append(f"V3: regra senza aliases: {r.get('conta_id', '?')}")

    # V4: conta_id duplicati (avviso)
    from collections import Counter
    cids = [str(r.get("conta_id", "")) for r in regras]
    dups = [c for c, n in Counter(cids).items() if n > 1]
    if dups:
        warnings.append(f"V4: conta_id duplicati in regras (ammesso): {sorted(dups)}")

    return errors, warnings


def _safe(text):
    s = str(text) if text is not None else ""
    return s.encode("latin-1", "replace").decode("latin-1")


def _p(text, style):
    return Paragraph(_safe(str(text) if text is not None else ""), style)


def _fmt_list(lst, style):
    if not lst:
        return Paragraph("—", style)
    return Paragraph("<br/>".join(_safe(str(v)) for v in lst), style)


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
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), ALT_ROW))


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


def _make_styles():
    base = getSampleStyleSheet()
    title = ParagraphStyle(
        "XMLTitle", parent=base["Normal"],
        fontName="Helvetica-Bold", fontSize=12,
        alignment=TA_CENTER, spaceAfter=4 * mm,
    )
    section = ParagraphStyle(
        "XMLSection", parent=base["Normal"],
        fontName="Helvetica-Bold", fontSize=9,
        spaceBefore=3 * mm, spaceAfter=2 * mm,
    )
    note = ParagraphStyle(
        "XMLNote", parent=base["Normal"],
        fontName="Helvetica-Oblique", fontSize=7,
        textColor=HDR_EVITAR, spaceAfter=2 * mm,
    )
    header = ParagraphStyle(
        "XMLHeader", parent=base["Normal"],
        fontName="Helvetica-Bold", fontSize=7,
        leading=9, textColor=colors.white, alignment=TA_LEFT,
    )
    cell = ParagraphStyle(
        "XMLCell", parent=base["Normal"],
        fontName="Helvetica", fontSize=6.5,
        leading=8, alignment=TA_LEFT,
    )
    cell_r = ParagraphStyle(
        "XMLCellR", parent=base["Normal"],
        fontName="Helvetica", fontSize=6.5,
        leading=8, alignment=TA_RIGHT,
    )
    evitar_cell = ParagraphStyle(
        "XMLEvitar", parent=base["Normal"],
        fontName="Helvetica-Oblique", fontSize=6.5,
        leading=8, textColor=EVITAR_TEXT, alignment=TA_LEFT,
    )
    return title, section, note, header, cell, cell_r, evitar_cell


def build_aliases_table(regras, usable_w, header_style, cell_style):
    raw = [200, 55, 220, 55, 55]
    scale = usable_w / sum(raw)
    col_w = [w * scale for w in raw]

    headers = ["alias", "conta_id", "conta_nome", "prioridade", "ha_evitar"]
    data = [[_p(h, header_style) for h in headers]]

    for r in regras:
        cid = _safe(r.get("conta_id", ""))
        cnome = _safe(r.get("conta_nome", ""))
        prio = _safe(r.get("prioridade", ""))
        ha = "si" if r.get("evitar_aliases") else "no"
        for alias in (r.get("aliases") or []):
            data.append([
                _p(alias, cell_style),
                _p(cid, cell_style),
                _p(cnome, cell_style),
                _p(prio, cell_style),
                _p(ha, cell_style),
            ])

    style_cmds = _base_style(HDR_ALIASES)
    _add_alt_rows(style_cmds, len(data) - 1)

    table = Table(data, colWidths=col_w, repeatRows=1)
    table.setStyle(TableStyle(style_cmds))
    return table


def build_fornecedores_table(fornecedores, usable_w, header_style, cell_style):
    raw = [220, 130, 55, 200, 180]
    scale = usable_w / sum(raw)
    col_w = [w * scale for w in raw]

    headers = ["fornecedor", "produto_contem", "conta_id", "conta_nome", "observacao"]
    data = [[_p(h, header_style) for h in headers]]

    for f in fornecedores:
        if f.get("fornecedor_contem"):
            forn_cell = _p(f["fornecedor_contem"], cell_style)
        else:
            aliases = f.get("fornecedor_aliases") or []
            forn_cell = Paragraph(
                "<br/>".join(_safe(str(a)) for a in aliases), cell_style
            )
        prod = f.get("produto_normalizado_contem")
        data.append([
            forn_cell,
            _p(prod if prod else "—", cell_style),
            _p(f.get("conta_id", ""), cell_style),
            _p(f.get("conta_nome", ""), cell_style),
            _p(f.get("observacao", ""), cell_style),
        ])

    style_cmds = _base_style(HDR_FORNEC)
    _add_alt_rows(style_cmds, len(fornecedores))

    table = Table(data, colWidths=col_w, repeatRows=1)
    table.setStyle(TableStyle(style_cmds))
    return table


def build_evitar_table(regras, usable_w, header_style, evitar_cell_style, cell_style):
    raw = [220, 55, 225, 285]
    scale = usable_w / sum(raw)
    col_w = [w * scale for w in raw]

    headers = ["evitar_alias", "conta_id", "conta_nome", "motivo"]
    data = [[_p(h, header_style) for h in headers]]
    evitar_rows = []

    row_idx = 1
    for r in regras:
        evitar = r.get("evitar_aliases") or []
        cid = _safe(r.get("conta_id", ""))
        cnome = _safe(r.get("conta_nome", ""))
        for ev in evitar:
            data.append([
                _p(ev, evitar_cell_style),
                _p(cid, cell_style),
                _p(cnome, cell_style),
                _p("Protezione anti-falso-positivo", evitar_cell_style),
            ])
            evitar_rows.append(row_idx)
            row_idx += 1

    style_cmds = _base_style(HDR_EVITAR)
    _add_alt_rows(style_cmds, len(data) - 1)
    for ri in evitar_rows:
        style_cmds.append(("BACKGROUND", (0, ri), (-1, ri), EVITAR_BG))

    table = Table(data, colWidths=col_w, repeatRows=1)
    table.setStyle(TableStyle(style_cmds))
    return table


def build_resumo_table(regras, fornecedores, fallback, usable_w, header_style, cell_style, cell_r_style):
    raw = [75, 310, 130, 130, 140]
    scale = usable_w / sum(raw)
    col_w = [w * scale for w in raw]

    headers = ["conta_id", "conta_nome", "tot_aliases", "tot_fornecedores", "tot_protezioni"]
    data = [[_p(h, header_style) for h in headers]]

    aliases_per_conta = defaultdict(int)
    nome_per_conta = {}
    evitar_per_conta = defaultdict(int)

    for r in regras:
        cid = str(r.get("conta_id", ""))
        aliases_per_conta[cid] += len(r.get("aliases") or [])
        nome_per_conta[cid] = r.get("conta_nome", "")
        evitar_per_conta[cid] += len(r.get("evitar_aliases") or [])

    forn_per_conta = defaultdict(int)
    for f in fornecedores:
        cid = str(f.get("conta_id", ""))
        forn_per_conta[cid] += 1
        if cid not in nome_per_conta:
            nome_per_conta[cid] = f.get("conta_nome", "")

    all_ids = sorted(set(list(aliases_per_conta.keys()) + list(forn_per_conta.keys())))

    for cid in all_ids:
        data.append([
            _p(cid, cell_style),
            _p(nome_per_conta.get(cid, ""), cell_style),
            _p(str(aliases_per_conta.get(cid, 0)), cell_r_style),
            _p(str(forn_per_conta.get(cid, 0)), cell_r_style),
            _p(str(evitar_per_conta.get(cid, 0)), cell_r_style),
        ])

    # Fallback row
    fb_id = str(fallback.get("id", ""))
    fb_nome = str(fallback.get("nome", ""))
    if fb_id:
        data.append([
            _p(fb_id, cell_style),
            _p(f"{fb_nome} [fallback]", cell_style),
            _p("—", cell_r_style),
            _p("—", cell_r_style),
            _p("—", cell_r_style),
        ])

    style_cmds = _base_style(HDR_RESUMO)
    _add_alt_rows(style_cmds, len(data) - 1)
    # Highlight fallback row
    if fb_id:
        last = len(data) - 1
        style_cmds += [
            ("BACKGROUND", (0, last), (-1, last), colors.HexColor("#D5E8D4")),
            ("FONTNAME", (0, last), (-1, last), "Helvetica-Oblique"),
        ]

    table = Table(data, colWidths=col_w, repeatRows=1)
    table.setStyle(TableStyle(style_cmds))
    return table


def build_pdf(regras, fornecedores, fallback, output_path):
    page_size = landscape(A4)
    margin = 10 * mm
    usable_w = page_size[0] - 2 * margin

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=page_size,
        leftMargin=margin, rightMargin=margin,
        topMargin=14 * mm, bottomMargin=10 * mm,
    )

    title_s, section_s, note_s, header_s, cell_s, cell_r_s, evitar_s = _make_styles()

    total_aliases = sum(len(r.get("aliases") or []) for r in regras)
    total_evitar = sum(len(r.get("evitar_aliases") or []) for r in regras)
    TITLE = "REGRAS E ALIASES CLASSIFICACAO XML/NF-e"

    t_aliases = build_aliases_table(regras, usable_w, header_s, cell_s)
    t_fornec = build_fornecedores_table(fornecedores, usable_w, header_s, cell_s)
    t_evitar = build_evitar_table(regras, usable_w, header_s, evitar_s, cell_s)
    t_resumo = build_resumo_table(regras, fornecedores, fallback, usable_w, header_s, cell_s, cell_r_s)

    story = [
        Paragraph(TITLE, title_s),
        Paragraph(f"Sezione 1 — Aliases prodotto ({total_aliases} aliases, {len(regras)} conti)", section_s),
        t_aliases,
        PageBreak(),
        Paragraph(TITLE, title_s),
        Paragraph(f"Sezione 2 — Regras por fornecedor ({len(fornecedores)} regole)", section_s),
        t_fornec,
        PageBreak(),
        Paragraph(TITLE, title_s),
        Paragraph(f"Sezione 3 — Evitar aliases / protezioni ({total_evitar} protezioni)", section_s),
        Paragraph("Termini in questa sezione NON vengono classificati nel conto indicato.", note_s),
        t_evitar,
        PageBreak(),
        Paragraph(TITLE, title_s),
        Paragraph("Sezione 4 — Resumo por conta", section_s),
        t_resumo,
    ]

    page_tracker = [0]

    def on_page(canv, doc):
        page_tracker[0] = doc.page
        canv.saveState()
        canv.setFont("Helvetica", 6)
        canv.drawRightString(page_size[0] - margin, 6 * mm, f"Pagina {doc.page}")
        canv.restoreState()

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)

    return page_tracker[0], total_aliases, total_evitar


def write_manifest(output_dir, regras, fornecedores, unmapped):
    total_aliases = sum(len(r.get("aliases") or []) for r in regras)
    total_evitar = sum(len(r.get("evitar_aliases") or []) for r in regras)
    manifest = {
        "total_regras": len(regras),
        "total_aliases": total_aliases,
        "total_fornecedores": len(fornecedores),
        "total_evitar_aliases": total_evitar,
        "conta_ids_regras": [str(r.get("conta_id", "")) for r in regras],
        "conta_ids_fornecedores": [str(f.get("conta_id", "")) for f in fornecedores],
        "sezioni_non_mappate": unmapped,
    }
    path = Path(output_dir) / "manifest_regras_xml.json"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    return path, manifest


def validate_manifest(manifest_path, regras, fornecedores):
    with open(manifest_path, encoding="utf-8") as f:
        m = json.load(f)
    divs = []
    exp_regras = len(regras)
    exp_aliases = sum(len(r.get("aliases") or []) for r in regras)
    exp_forn = len(fornecedores)
    exp_evitar = sum(len(r.get("evitar_aliases") or []) for r in regras)

    checks = [
        ("total_regras", exp_regras, m.get("total_regras")),
        ("total_aliases", exp_aliases, m.get("total_aliases")),
        ("total_fornecedores", exp_forn, m.get("total_fornecedores")),
        ("total_evitar_aliases", exp_evitar, m.get("total_evitar_aliases")),
    ]
    for label, expected, got in checks:
        if expected != got:
            divs.append(f"{label}: manifest={got} vs YAML={expected}")
    return divs


def write_summary(output_dir, stats):
    path = Path(output_dir) / "resumo_regras_xml.txt"
    lines = [
        f"Data/ora generazione:         {stats['datetime']}",
        f"YAML regras:                  {stats['regras_path']}",
        f"Totale regras (conti):        {stats['total_regras']}",
        f"Totale aliases:               {stats['total_aliases']}",
        f"Totale regras_por_fornecedor: {stats['total_forn']}",
        f"Totale evitar_aliases:        {stats['total_evitar']}",
        f"Conta fallback:               {stats['fallback_id']} / {stats['fallback_nome']}",
        f"Pagine PDF:                   {stats['pages']}",
        f"Orientamento:                 landscape",
        f"Sezioni non mappate:          {stats['unmapped'] if stats['unmapped'] else 'nenhuma'}",
        f"Alertas:                      {stats['alertas']}",
    ]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return path


def main():
    parser = argparse.ArgumentParser(
        description="Gera PDF delle Regras e Aliases Classificacao XML/NF-e"
    )
    parser.add_argument("--regras", required=True, help="regras_classificacao_produtos.yaml")
    parser.add_argument("--output", required=True, help="Output directory")
    args = parser.parse_args()

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    regras, fornecedores, fallback, unmapped = load_yaml(args.regras)
    total_aliases = sum(len(r.get("aliases") or []) for r in regras)
    total_evitar = sum(len(r.get("evitar_aliases") or []) for r in regras)
    print(
        f"Loaded: {len(regras)} regras, {total_aliases} aliases, "
        f"{len(fornecedores)} fornecedores, {total_evitar} evitar_aliases"
    )
    if unmapped:
        print(f"Aviso V5 - sezioni non mappate: {unmapped}")

    errors, warnings = validate_inputs(regras, fornecedores)
    for w in warnings:
        print(f"Aviso: {w}")
    for e in errors:
        print(f"ERRO: {e}", file=sys.stderr)
    if errors:
        sys.exit(1)
    print("V2-V5: OK")

    pdf_path = output_dir / "regras_classificacao_xml.pdf"
    print("Gerando PDF...")
    pages, _, _ = build_pdf(regras, fornecedores, fallback, pdf_path)

    manifest_path, _ = write_manifest(output_dir, regras, fornecedores, unmapped)

    divs = validate_manifest(manifest_path, regras, fornecedores)
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

    alertas = []
    if unmapped:
        alertas.append(f"sezioni non mappate: {unmapped}")

    summary_path = write_summary(
        output_dir,
        {
            "datetime": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "regras_path": args.regras,
            "total_regras": len(regras),
            "total_aliases": total_aliases,
            "total_forn": len(fornecedores),
            "total_evitar": total_evitar,
            "fallback_id": fallback.get("id", ""),
            "fallback_nome": fallback.get("nome", ""),
            "pages": pages,
            "unmapped": unmapped,
            "alertas": "; ".join(alertas) if alertas else "nenhum",
        },
    )

    print(f"\nPDF: {pdf_path} ({pages} pag.)")
    print(f"Manifest: {manifest_path}")
    print(f"Resumo TXT: {summary_path}")
    print("\nGeracao concluida.")


if __name__ == "__main__":
    main()
