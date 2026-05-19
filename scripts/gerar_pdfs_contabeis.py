#!/usr/bin/env python3
"""Generator for Plano de Contas and Legenda DRE PDFs."""

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import yaml
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)


def load_yaml(plano_path, mapeamento_path):
    with open(plano_path, encoding="utf-8") as f:
        plano = yaml.safe_load(f)
    with open(mapeamento_path, encoding="utf-8") as f:
        mapeamento_data = yaml.safe_load(f)
    contas = plano.get("contas", [])
    linhas_dre = mapeamento_data.get("linhas_dre", [])
    mapeamento = mapeamento_data.get("mapeamento", {})
    return contas, linhas_dre, mapeamento


def validate_inputs(contas, linhas_dre, mapeamento):
    errors = []
    warnings = []

    conta_ids_plano = {str(c["id"]) for c in contas}
    conta_ids_map = {str(k) for k in mapeamento.keys()}
    linha_ids = {str(l["id"]) for l in linhas_dre}
    linha_entra_dre = {str(l["id"]): l.get("entra_dre", True) for l in linhas_dre}

    v1 = conta_ids_plano - conta_ids_map
    if v1:
        errors.append(f"V1 conta_id no plano sem mapeamento: {sorted(v1)}")

    v2 = conta_ids_map - conta_ids_plano
    if v2:
        errors.append(f"V2 conta_id no mapeamento sem plano: {sorted(v2)}")

    v3 = set()
    for cid, entry in mapeamento.items():
        dlid = str(entry.get("dre_linha_id", ""))
        if dlid and dlid not in linha_ids:
            v3.add(dlid)
    if v3:
        errors.append(f"V3 dre_linha_id invalidos: {sorted(v3)}")

    for cid, entry in mapeamento.items():
        dlid = str(entry.get("dre_linha_id", ""))
        map_entra = entry.get("entra_dre", True)
        if dlid in linha_entra_dre and map_entra != linha_entra_dre[dlid]:
            warnings.append(
                f"V4 {cid}: mapeamento.entra_dre={map_entra} "
                f"vs linha {dlid}.entra_dre={linha_entra_dre[dlid]}"
            )

    return errors, warnings


def _safe(text):
    """Encode to Latin-1 for reportlab standard fonts, replacing unmappable chars."""
    s = str(text) if text is not None else ""
    return s.encode("latin-1", "replace").decode("latin-1")


def _base_style(header_bg):
    return [
        ("BACKGROUND", (0, 0), (-1, 0), header_bg),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 7),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 6.5),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#AAAAAA")),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]


def _add_alt_rows(style_cmds, n_data_rows):
    for i in range(1, n_data_rows + 1):
        if i % 2 == 0:
            style_cmds.append(
                ("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F2F2F2"))
            )


def build_plano_pdf(contas, mapeamento, output_path):
    """Generate PDF 1 - Plano de Contas Gerencial.

    Returns (page_count, conta_ids_rendered).
    """
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

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "PDFTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=11,
        alignment=TA_CENTER,
        spaceAfter=4 * mm,
    )

    raw_widths = [55, 215, 195, 65, 65, 48]
    scale = usable_w / sum(raw_widths)
    col_widths = [w * scale for w in raw_widths]

    headers = ["conta_id", "nome", "grupo", "tipo", "dre_linha_id", "entra_dre"]
    data = [headers]
    conta_ids_rendered = []
    gray_rows = []

    for i, conta in enumerate(contas, start=1):
        cid = str(conta["id"])
        entry = mapeamento.get(cid, {})
        entra = entry.get("entra_dre", "")
        dre = entry.get("dre_linha_id", "")
        data.append(
            [
                _safe(cid),
                _safe(conta.get("nome", "")),
                _safe(conta.get("grupo", "")),
                _safe(conta.get("tipo", "")),
                _safe(dre),
                _safe(entra),
            ]
        )
        conta_ids_rendered.append(cid)
        if entra is False:
            gray_rows.append(i)

    style_cmds = _base_style(colors.HexColor("#2C3E50"))
    _add_alt_rows(style_cmds, len(contas))
    for ri in gray_rows:
        style_cmds += [
            ("BACKGROUND", (0, ri), (-1, ri), colors.HexColor("#C8C8C8")),
            ("TEXTCOLOR", (0, ri), (-1, ri), colors.HexColor("#555555")),
        ]

    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle(style_cmds))

    page_tracker = [0]

    def on_page(canv, doc):
        page_tracker[0] = doc.page
        canv.saveState()
        canv.setFont("Helvetica", 6)
        canv.drawRightString(
            page_size[0] - margin, 6 * mm, f"Pagina {doc.page}"
        )
        canv.restoreState()

    doc.build(
        [Paragraph("PLANO DE CONTAS GERENCIAL", title_style), table],
        onFirstPage=on_page,
        onLaterPages=on_page,
    )

    return page_tracker[0], conta_ids_rendered


def build_legenda_pdf(linhas_dre, mapeamento, contas, output_path):
    """Generate PDF 2 - Legenda DRE e Mapeamento.

    Returns (page_count, linha_ids_rendered, map_ids_rendered).
    """
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

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "PDFTitle2",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=11,
        alignment=TA_CENTER,
        spaceAfter=3 * mm,
    )
    section_style = ParagraphStyle(
        "Section",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        spaceBefore=3 * mm,
        spaceAfter=2 * mm,
    )

    conta_index = {str(c["id"]): c for c in contas}
    linha_index = {str(l["id"]): l for l in linhas_dre}

    # --- Section 1: linhas_dre ---
    raw_w1 = [60, 210, 175, 45, 60, 60]
    col_w1 = [w * usable_w / sum(raw_w1) for w in raw_w1]

    heads1 = ["dre_linha_id", "nome linha DRE", "grupo", "sinal", "calculada", "entra_dre"]
    data1 = [heads1]
    linha_ids_rendered = []
    calc_rows = []
    false_rows1 = []

    for i, linha in enumerate(linhas_dre, start=1):
        lid = str(linha["id"])
        calculada = linha.get("calculada", False)
        entra = linha.get("entra_dre", True)
        data1.append(
            [
                _safe(lid),
                _safe(linha.get("nome", "")),
                _safe(linha.get("grupo", "")),
                _safe(linha.get("sinal", "")),
                _safe(calculada),
                _safe(entra),
            ]
        )
        linha_ids_rendered.append(lid)
        if calculada:
            calc_rows.append(i)
        if not entra:
            false_rows1.append(i)

    style1 = _base_style(colors.HexColor("#1A5276"))
    _add_alt_rows(style1, len(linhas_dre))
    for ri in calc_rows:
        style1 += [
            ("BACKGROUND", (0, ri), (-1, ri), colors.HexColor("#D5E8D4")),
            ("FONTNAME", (0, ri), (-1, ri), "Helvetica-Bold"),
        ]
    for ri in false_rows1:
        style1 += [
            ("BACKGROUND", (0, ri), (-1, ri), colors.HexColor("#C8C8C8")),
            ("TEXTCOLOR", (0, ri), (-1, ri), colors.HexColor("#555555")),
        ]

    table1 = Table(data1, colWidths=col_w1, repeatRows=1)
    table1.setStyle(TableStyle(style1))

    # --- Section 2: mapeamento sorted by dre_linha_id then conta_id ---
    raw_w2 = [60, 200, 65, 240, 55]
    col_w2 = [w * usable_w / sum(raw_w2) for w in raw_w2]

    heads2 = ["dre_linha_id", "nome linha DRE", "conta_id", "nome conta", "entra_dre"]
    data2 = [heads2]
    map_ids_rendered = []
    false_rows2 = []

    sorted_entries = sorted(
        mapeamento.items(),
        key=lambda x: (str(x[1].get("dre_linha_id", "")), str(x[0])),
    )

    for i, (cid, entry) in enumerate(sorted_entries, start=1):
        dlid = str(entry.get("dre_linha_id", ""))
        entra = entry.get("entra_dre", True)
        nome_linha = linha_index.get(dlid, {}).get("nome", "")
        nome_conta = conta_index.get(str(cid), {}).get("nome", "")
        data2.append(
            [
                _safe(dlid),
                _safe(nome_linha),
                _safe(cid),
                _safe(nome_conta),
                _safe(entra),
            ]
        )
        map_ids_rendered.append(str(cid))
        if entra is False:
            false_rows2.append(i)

    style2 = _base_style(colors.HexColor("#1A5276"))
    _add_alt_rows(style2, len(mapeamento))
    for ri in false_rows2:
        style2 += [
            ("BACKGROUND", (0, ri), (-1, ri), colors.HexColor("#C8C8C8")),
            ("TEXTCOLOR", (0, ri), (-1, ri), colors.HexColor("#555555")),
        ]

    table2 = Table(data2, colWidths=col_w2, repeatRows=1)
    table2.setStyle(TableStyle(style2))

    page_tracker = [0]

    def on_page(canv, doc):
        page_tracker[0] = doc.page
        canv.saveState()
        canv.setFont("Helvetica", 6)
        canv.drawRightString(
            page_size[0] - margin, 6 * mm, f"Pagina {doc.page}"
        )
        canv.restoreState()

    story = [
        Paragraph("LEGENDA NUMERAZIONE DRE E MAPEAMENTO DEI CONTI", title_style),
        Paragraph("Secao 1 - Linhas DRE", section_style),
        table1,
        PageBreak(),
        Paragraph("Secao 2 - Mapeamento Contas para Linhas DRE", section_style),
        table2,
    ]

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)

    return page_tracker[0], linha_ids_rendered, map_ids_rendered


def _get_pdf_dimensions(pdf_path):
    """Parse MediaBox from PDF bytes to verify landscape orientation."""
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


def write_manifest(output_dir, conta_ids, linha_ids, map_ids):
    manifest = {
        "conta_ids_renderizados": conta_ids,
        "linha_dre_ids_renderizados": linha_ids,
        "mapeamentos_renderizados": map_ids,
        "total_contas": len(conta_ids),
        "total_linhas_dre": len(linha_ids),
        "total_mapeamentos": len(map_ids),
    }
    path = Path(output_dir) / "manifest_geracao.json"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    return path


def validate_manifest(manifest_path, contas, linhas_dre, mapeamento):
    with open(manifest_path, encoding="utf-8") as f:
        m = json.load(f)
    divs = []
    exp_c = [str(c["id"]) for c in contas]
    exp_l = [str(l["id"]) for l in linhas_dre]
    exp_m = [str(k) for k in mapeamento.keys()]

    if m["total_contas"] != len(exp_c):
        divs.append(f"total_contas: manifest={m['total_contas']} vs YAML={len(exp_c)}")
    if m["total_linhas_dre"] != len(exp_l):
        divs.append(f"total_linhas_dre: manifest={m['total_linhas_dre']} vs YAML={len(exp_l)}")
    if m["total_mapeamentos"] != len(exp_m):
        divs.append(
            f"total_mapeamentos: manifest={m['total_mapeamentos']} vs YAML={len(exp_m)}"
        )

    miss_c = set(exp_c) - set(m["conta_ids_renderizados"])
    if miss_c:
        divs.append(f"conta_ids ausentes no manifest: {sorted(miss_c)}")
    miss_l = set(exp_l) - set(m["linha_dre_ids_renderizados"])
    if miss_l:
        divs.append(f"linha_ids ausentes no manifest: {sorted(miss_l)}")
    miss_m = set(exp_m) - set(m["mapeamentos_renderizados"])
    if miss_m:
        divs.append(f"mapeamentos ausentes no manifest: {sorted(miss_m)}")

    return divs


def write_summary(output_dir, stats):
    path = Path(output_dir) / "resumo_geracao_pdfs_contabeis.txt"
    lines = [
        f"Data/hora geracao:        {stats['datetime']}",
        f"YAML plano_contas:        {stats['plano_path']}",
        f"YAML mapeamento_dre:      {stats['mapeamento_path']}",
        f"Total de contas:          {stats['total_contas']}",
        f"Total linhas DRE:         {stats['total_linhas_dre']}",
        f"Contas entra_dre true:    {stats['entra_dre_true']}",
        f"Contas entra_dre false:   {stats['entra_dre_false']}",
        f"Paginas plano_contas PDF: {stats['pages_plano']}",
        f"Paginas legenda_dre PDF:  {stats['pages_legenda']}",
        f"Alertas:                  {stats['alertas']}",
    ]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return path


def main():
    parser = argparse.ArgumentParser(
        description="Gera PDFs do Plano de Contas e Legenda DRE"
    )
    parser.add_argument("--plano", required=True, help="plano_contas_gerencial.yaml")
    parser.add_argument("--mapeamento", required=True, help="mapeamento_dre.yaml")
    parser.add_argument("--output", required=True, help="Output directory")
    args = parser.parse_args()

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    contas, linhas_dre, mapeamento = load_yaml(args.plano, args.mapeamento)
    print(
        f"Loaded: {len(contas)} contas, {len(linhas_dre)} linhas DRE, "
        f"{len(mapeamento)} mapeamentos"
    )

    errors, warnings = validate_inputs(contas, linhas_dre, mapeamento)
    for e in errors:
        print(f"ERRO: {e}", file=sys.stderr)
    if errors:
        sys.exit(1)
    for w in warnings:
        print(f"Aviso: {w}")

    plano_pdf = output_dir / "plano_contas_gerencial.pdf"
    legenda_pdf = output_dir / "legenda_dre_mapeamento_contas.pdf"

    print("Gerando PDF 1 (Plano de Contas)...")
    pages_plano, conta_ids = build_plano_pdf(contas, mapeamento, plano_pdf)

    print("Gerando PDF 2 (Legenda DRE)...")
    pages_legenda, linha_ids, map_ids = build_legenda_pdf(
        linhas_dre, mapeamento, contas, legenda_pdf
    )

    manifest_path = write_manifest(output_dir, conta_ids, linha_ids, map_ids)

    divs = validate_manifest(manifest_path, contas, linhas_dre, mapeamento)
    if divs:
        for d in divs:
            print(f"MANIFEST DIVERGENCIA: {d}", file=sys.stderr)
        sys.exit(1)

    for pdf_p, label in [(plano_pdf, "plano"), (legenda_pdf, "legenda")]:
        w, h = _get_pdf_dimensions(str(pdf_p))
        if w is not None:
            orientation = "landscape OK" if w > h else "PORTRAIT (ERRO)"
            print(f"Orientacao {label}: {orientation} ({w:.0f}x{h:.0f}pt)")
            if w < h:
                print(f"ERRO: {label} nao e landscape!", file=sys.stderr)
                sys.exit(1)

    entra_true = sum(
        1 for cid in conta_ids if mapeamento.get(cid, {}).get("entra_dre", True)
    )
    entra_false = len(conta_ids) - entra_true

    alertas = []
    if pages_plano > 2:
        alertas.append(f"plano ultrapassou 2 paginas ({pages_plano})")
    for w in warnings:
        alertas.append(w)

    summary_path = write_summary(
        output_dir,
        {
            "datetime": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "plano_path": args.plano,
            "mapeamento_path": args.mapeamento,
            "total_contas": len(contas),
            "total_linhas_dre": len(linhas_dre),
            "entra_dre_true": entra_true,
            "entra_dre_false": entra_false,
            "pages_plano": pages_plano,
            "pages_legenda": pages_legenda,
            "alertas": "; ".join(alertas) if alertas else "nenhum",
        },
    )

    print(f"\nPDF 1: {plano_pdf} ({pages_plano} pag.)")
    print(f"PDF 2: {legenda_pdf} ({pages_legenda} pag.)")
    print(f"Manifest: {manifest_path}")
    print(f"Resumo TXT: {summary_path}")
    if alertas:
        print(f"Alertas: {'; '.join(alertas)}")
    print("\nGeracao concluida.")


if __name__ == "__main__":
    main()
