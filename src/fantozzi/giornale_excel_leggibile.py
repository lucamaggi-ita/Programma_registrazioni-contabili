from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from openpyxl.formatting.rule import DataBarRule as _DataBarRule
    _HAS_DATABAR = True
except Exception:
    _HAS_DATABAR = False


STATUS = "TESTE_TECNICO_NAO_OFICIAL"

_PRELIVRO_CANDIDATES = [
    Path("pacote_contabil_gerencial_provisorio") / "02_PRE_LIVRO_DIARIO_GERENCIAL.xlsx",
    Path("pre_livro_diario_gerencial_teste.xlsx"),
]
_DRE_CANDIDATES = [
    Path("pacote_contabil_gerencial_provisorio") / "01_DRE_GERENCIAL_CONSOLIDADO.xlsx",
    Path("dre_gerencial_ok_provisoria.xlsx"),
]

# Colors
_C_HEADER_BG = "1F4E78"
_C_HEADER_FG = "FFFFFF"
_C_CARD_BG = "EEF2F9"
_C_CARD_ACCENT = "2E75B6"
_C_ALT_ROW = "F4F6FA"
_C_REVISAR_BG = "FFF2CC"
_C_STATUS_BADGE = "C84B2A"
_C_TOP_ACCOUNT = "E2EFDA"  # light green for top accounts in 03_ANALISE_POR_CONTA
_C_LEITURA_BG = "F0F4FB"   # light blue for leituras rápidas block

_FMT_BRL = "#,##0.00"
_FMT_DATE = "DD/MM/YYYY"

# Human-readable column map for 04_GIORNALE: (tech_key, human_label, col_width)
_GIORNALE_COLS: list[tuple[str, str, int]] = [
    ("data_documento", "Data", 14),
    ("competencia", "Mês", 10),
    ("origem", "Origem", 16),
    ("fornecedor_ou_favorecido", "Fornecedor / Favorecido", 38),
    ("documento", "Documento", 16),
    ("descricao_itens", "Histórico", 48),
    ("conta_id", "Conta", 14),
    ("conta_nome", "Nome da Conta", 34),
    ("grupo_conta", "Grupo DRE", 28),
    ("dre_linha_nome", "Linha DRE", 34),
    ("valor", "Valor", 16),
    ("status_revisao", "Status", 14),
    ("observacoes", "Observação", 36),
]

# Technical names that MUST NOT appear as column headers in sheets 01–04
FORBIDDEN_TECH_HEADERS: frozenset[str] = frozenset({
    "arquivo_excel_origem",
    "cnpj_empresa",
    "nome_empresa",
    "cnpj_cpf_favorecido",
    "tipo_documento",
    "arquivo_origem",
    "chave_origem",
    "hash_registro",
    "centro_custo",
    "cfop",
    "natureza_operacao",
    "tipo_operacao_especial",
    "confianca_operacao_especial",
    "tratamento_operacao_especial",
    "tipo_lancamento",
    "dre_linha_id",
    "data_pagamento",
    "duplicatas_json",
    "alertas",
    "fornecedor_ou_favorecido",
    "status_revisao",
    "conta_debito_id",
    "conta_credito_id",
    "linha_origem",
    "snake_case",
})


class GiornaleExcelLeggibileError(Exception):
    pass


@dataclass(frozen=True)
class GiornaleResult:
    xlsx_path: Path
    prelivro_source: Path
    dre_source: Path | None
    run_id: str | None
    periodos: tuple[str, ...]
    lancamentos_ok: int
    lancamentos_revisar: int
    status: str


def gerar_giornale_leggibile(
    run_dir: Path | str,
    output_dir: Path | str | None = None,
) -> GiornaleResult:
    run_dir = Path(run_dir)
    if not run_dir.exists() or not run_dir.is_dir():
        raise GiornaleExcelLeggibileError(f"run-dir nao encontrado: {run_dir}")

    prelivro_xlsx = _find_prelivro(run_dir)

    dre_xlsx: Path | None = None
    for c in _DRE_CANDIDATES:
        p = run_dir / c
        if p.exists():
            dre_xlsx = p
            break

    if output_dir is None:
        output_dir = run_dir / "output_visual"
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    manifest = _load_manifest(run_dir)
    run_id = manifest.get("run_id") or None

    lancamentos_ok = _load_sheet(prelivro_xlsx, "Lancamentos_OK")
    lancamentos_revisar = _load_sheet(prelivro_xlsx, "Lancamentos_REVISAR")
    resumo_mensal = _load_sheet(prelivro_xlsx, "Resumo_Mensal")

    dre_geral: list[dict] = []
    dre_mensal: list[dict] = []
    base_ok_dre: list[dict] = []
    if dre_xlsx:
        dre_geral = _load_sheet(dre_xlsx, "DRE_Geral")
        dre_mensal = _load_sheet(dre_xlsx, "DRE_Mensal")
        base_ok_dre = _load_sheet(dre_xlsx, "Base_OK_DRE")

    periodos = _collect_periodos(lancamentos_ok, dre_mensal)

    wb = Workbook()

    dashboard_ws = wb.active
    dashboard_ws.title = "01_DASHBOARD"
    dashboard_ws.sheet_properties.tabColor = _C_HEADER_BG
    _write_dashboard(
        dashboard_ws, run_id, periodos, lancamentos_ok, lancamentos_revisar,
        dre_geral, resumo_mensal,
    )

    dre_det_ws = wb.create_sheet("02_DRE_DETTAGLIO")
    dre_det_ws.sheet_properties.tabColor = "2E75B6"
    _write_dre_dettaglio(dre_det_ws, dre_mensal, base_ok_dre)

    analise_ws = wb.create_sheet("03_ANALISE_POR_CONTA")
    analise_ws.sheet_properties.tabColor = "2A9D8F"
    _write_analise_por_conta(analise_ws, base_ok_dre, lancamentos_ok, periodos)

    conta_lookup = _build_conta_dre_lookup(base_ok_dre)

    giornale_ws = wb.create_sheet("04_GIORNALE")
    giornale_ws.sheet_properties.tabColor = "F4A261"
    _write_giornale(giornale_ws, lancamentos_ok, lancamentos_revisar, conta_lookup)

    audit_ws = wb.create_sheet("05_BASE_TECNICA_AUDIT")
    audit_ws.sheet_properties.tabColor = "8EA9C1"
    _write_base_tecnica(audit_ws, lancamentos_ok, lancamentos_revisar)

    xlsx_path = output_dir / "03_GIORNALE_LEGGIBILE.xlsx"
    wb.save(str(xlsx_path))

    return GiornaleResult(
        xlsx_path=xlsx_path,
        prelivro_source=prelivro_xlsx,
        dre_source=dre_xlsx,
        run_id=run_id,
        periodos=tuple(periodos),
        lancamentos_ok=len(lancamentos_ok),
        lancamentos_revisar=len(lancamentos_revisar),
        status=STATUS,
    )


def console_lines(result: GiornaleResult) -> list[str]:
    lines = [
        "[OK] Giornale Excel leggibile gerado.",
        f"     Status: {result.status}",
        f"     Arquivo: {result.xlsx_path}",
        f"     Lancamentos OK: {result.lancamentos_ok}",
        f"     Lancamentos REVISAR: {result.lancamentos_revisar}",
    ]
    if result.run_id:
        lines.append(f"     Run: {result.run_id}")
    if result.periodos:
        lines.append(f"     Periodos: {' | '.join(result.periodos)}")
    return lines


# ── Data loading ───────────────────────────────────────────────────────────────

def _find_prelivro(run_dir: Path) -> Path:
    for c in _PRELIVRO_CANDIDATES:
        p = run_dir / c
        if p.exists():
            return p
    raise GiornaleExcelLeggibileError(
        f"Pre-livro nao encontrado em: {run_dir}. "
        f"Candidatos: {[str(c) for c in _PRELIVRO_CANDIDATES]}"
    )


def _load_manifest(run_dir: Path) -> dict:
    for name in (
        "manifest_run.json",
        "pacote_contabil_gerencial_provisorio/manifest_pacote_contabil_gerencial.json",
    ):
        p = run_dir / name
        if p.exists():
            try:
                return json.loads(p.read_text(encoding="utf-8"))
            except Exception:
                pass
    return {}


def _load_sheet(xlsx: Path, sheet_name: str) -> list[dict]:
    try:
        wb = load_workbook(str(xlsx), read_only=True, data_only=True)
    except Exception as exc:
        raise GiornaleExcelLeggibileError(
            f"Nao foi possivel abrir: {xlsx}"
        ) from exc
    if sheet_name not in wb.sheetnames:
        wb.close()
        return []
    ws = wb[sheet_name]
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not all_rows:
        return []
    headers = [str(h).strip() if h is not None else "" for h in all_rows[0]]
    result: list[dict] = []
    for row in all_rows[1:]:
        d = {headers[i]: (row[i] if i < len(row) else None) for i in range(len(headers))}
        if any(v is not None for v in d.values()):
            result.append(d)
    return result


def _collect_periodos(lancamentos_ok: list[dict], dre_mensal: list[dict]) -> list[str]:
    periodos: set[str] = set()
    for r in lancamentos_ok:
        v = str(r.get("competencia") or "").strip()
        if v:
            periodos.add(v)
    for r in dre_mensal:
        v = str(r.get("ano_mes") or "").strip()
        if v:
            periodos.add(v)
    return sorted(periodos)


def _build_conta_dre_lookup(base_ok_dre: list[dict]) -> dict[str, dict]:
    """Build {conta_id -> {linha_dre, grupo_dre}} from Base_OK_DRE for GIORNALE enrichment."""
    lookup: dict[str, dict] = {}
    for r in base_ok_dre:
        cid = str(r.get("conta_id") or "").strip()
        if cid and cid not in lookup:
            lookup[cid] = {
                "linha_dre": r.get("linha_dre"),
                "grupo_dre": r.get("grupo_dre"),
            }
    return lookup


# ── Helpers ────────────────────────────────────────────────────────────────────

def _to_decimal(value: Any) -> Decimal:
    if value is None or value == "":
        return Decimal("0.00")
    try:
        return Decimal(str(value)).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError):
        return Decimal("0.00")


def _fmt_brl_value(value: Decimal) -> str:
    abs_val = abs(value)
    s = f"{abs_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-" if value < 0 else "") + s


def _float_safe(val: Any) -> float | None:
    if val is None:
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


_THIN = Side(style="thin", color="CCCCCC")
_THIN_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


def _style_header_cells(ws, row: int, n_cols: int, *, bg: str = _C_HEADER_BG, fg: str = _C_HEADER_FG) -> None:
    fill = _fill(bg)
    font = Font(bold=True, color=fg)
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cell.border = _THIN_BORDER


def _apply_table(ws, header_row: int, last_data_row: int, n_cols: int) -> None:
    """Freeze panes + autofilter on an already-written table."""
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    if last_data_row >= header_row:
        end_letter = get_column_letter(n_cols)
        ws.auto_filter.ref = f"A{header_row}:{end_letter}{last_data_row}"


def _row_bg(i: int) -> str:
    return _C_ALT_ROW if i % 2 else "FFFFFF"


def _apply_databar(ws, col_letter: str, data_start: int, data_end: int, color: str = "2E75B6") -> None:
    """Apply data bars conditional formatting to a column range. No-op if unavailable."""
    if not _HAS_DATABAR or data_end < data_start:
        return
    try:
        rule = _DataBarRule(
            start_type="num", start_value=0,
            end_type="max", end_value=None,
            color=color,
            showValue=True,
            minLength=0,
            maxLength=100,
        )
        ws.conditional_formatting.add(f"{col_letter}{data_start}:{col_letter}{data_end}", rule)
    except Exception:
        pass


# ── Sheet writers ──────────────────────────────────────────────────────────────

def _write_leituras_rapidas(
    ws,
    start_row: int,
    total_revisar: int,
    resumo_mensal: list[dict],
    dre_geral: list[dict],
    periodos: list[str],
) -> int:
    """Write the 'Leituras rápidas' auto-insights block. Returns next available row."""
    row = start_row

    ws.merge_cells(f"A{row}:J{row}")
    c = ws[f"A{row}"]
    c.value = "LEITURAS RÁPIDAS"
    c.font = Font(bold=True, size=12, color=_C_HEADER_BG)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 22
    row += 1

    insights: list[str] = []

    if resumo_mensal:
        mes_valores = [
            (str(r.get("competencia") or ""), _to_decimal(r.get("valor_ok")))
            for r in resumo_mensal
            if r.get("competencia")
        ]
        if mes_valores:
            mes_max = max(mes_valores, key=lambda x: x[1])
            mes_min = min(mes_valores, key=lambda x: x[1])
            insights.append(
                f"Período com maior volume classificado: {mes_max[0]}"
                f"  —  R$ {_fmt_brl_value(mes_max[1])}"
            )
            if len(mes_valores) > 1 and mes_min[0] != mes_max[0]:
                insights.append(
                    f"Período com menor volume classificado: {mes_min[0]}"
                    f"  —  R$ {_fmt_brl_value(mes_min[1])}"
                )

    if dre_geral:
        grupos: dict[str, Decimal] = {}
        for dr in dre_geral:
            g = str(dr.get("grupo_dre") or "—")
            grupos[g] = grupos.get(g, Decimal("0")) + _to_decimal(dr.get("valor"))
        if grupos:
            maior_grupo = max(grupos, key=lambda g: abs(grupos[g]))
            insights.append(
                f"Maior grupo DRE: {maior_grupo}"
                f"  —  R$ {_fmt_brl_value(abs(grupos[maior_grupo]))}"
            )

    if total_revisar > 0:
        insights.append(
            f"Lançamentos em revisão: {total_revisar}"
            f"  —  verificar antes de emitir DRE definitiva"
        )
    else:
        insights.append(
            "Todos os lançamentos estão classificados  —  nenhuma linha em revisão"
        )

    insights.append(
        f"Status: {STATUS}  —  este arquivo não substitui escrituração contábil oficial"
    )

    for insight in insights:
        ws.merge_cells(f"A{row}:J{row}")
        cell = ws[f"A{row}"]
        cell.value = f"  •  {insight}"
        cell.fill = _fill(_C_LEITURA_BG)
        cell.font = Font(size=10, color="333333")
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cell.border = _THIN_BORDER
        ws.row_dimensions[row].height = 16
        row += 1

    return row


def _write_dashboard(
    ws,
    run_id: str | None,
    periodos: list[str],
    lancamentos_ok: list[dict],
    lancamentos_revisar: list[dict],
    dre_geral: list[dict],
    resumo_mensal: list[dict],
) -> None:
    period_str = " · ".join(periodos) if periodos else "—"
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    total_ok = len(lancamentos_ok)
    total_revisar = len(lancamentos_revisar)
    valor_ok = sum(_to_decimal(r.get("valor")) for r in lancamentos_ok)
    n_contas = len({str(r.get("conta_id") or "") for r in lancamentos_ok if r.get("conta_id")})

    # Row 1: status badge (10 cols)
    ws.merge_cells("A1:J1")
    c = ws["A1"]
    c.value = (
        f"TESTE TÉCNICO / NÃO OFICIAL  —  "
        f"STATUS: {STATUS}  —  "
        "Este relatório não substitui escrituração contábil oficial sem revisão do contador."
    )
    c.fill = _fill(_C_STATUS_BADGE)
    c.font = Font(bold=True, color="FFFFFF", size=10)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # Rows 2–3: main title
    ws.merge_cells("A2:J3")
    c = ws["A2"]
    c.value = "DRE Gerencial Provisória"
    c.fill = _fill(_C_HEADER_BG)
    c.font = Font(bold=True, color="FFFFFF", size=22)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=2)
    ws.row_dimensions[2].height = 40
    ws.row_dimensions[3].height = 6

    # Row 4: meta info
    ws.merge_cells("A4:J4")
    c = ws["A4"]
    run_str = f"  ·  Run: {run_id}" if run_id else ""
    c.value = f"Período: {period_str}  ·  Gerado em: {now_str}{run_str}"
    c.fill = _fill(_C_CARD_ACCENT)
    c.font = Font(color="FFFFFF", size=11)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=2)
    ws.row_dimensions[4].height = 22

    # Row 5: spacer
    ws.row_dimensions[5].height = 12

    # Cards: rows 6–8 (label / value / accent bar), 2 cols per card × 5 cards
    cards_data = [
        ("DESPESAS OK (R$)", f"R$ {_fmt_brl_value(valor_ok)}", _C_CARD_ACCENT),
        ("LANÇAMENTOS OK", str(total_ok), "2A9D8F"),
        ("EM REVISÃO", str(total_revisar), "E76F51" if total_revisar > 0 else "AAAAAA"),
        ("CONTAS ATIVAS", str(n_contas), _C_CARD_ACCENT),
        ("MESES COBERTOS", str(len(periodos)), _C_CARD_ACCENT),
    ]
    ws.row_dimensions[6].height = 16
    ws.row_dimensions[7].height = 34
    ws.row_dimensions[8].height = 5

    for idx, (label, value, accent) in enumerate(cards_data):
        cs = idx * 2 + 1
        ce = cs + 1
        ls = get_column_letter(cs)
        le = get_column_letter(ce)

        ws.merge_cells(f"{ls}6:{le}6")
        ws.merge_cells(f"{ls}7:{le}7")
        ws.merge_cells(f"{ls}8:{le}8")

        lc = ws[f"{ls}6"]
        lc.value = label
        lc.fill = _fill(_C_CARD_BG)
        lc.font = Font(size=9, color="666666")
        lc.alignment = Alignment(horizontal="center", vertical="bottom")

        vc = ws[f"{ls}7"]
        vc.value = value
        vc.fill = _fill(_C_CARD_BG)
        vc.font = Font(bold=True, size=15, color=_C_HEADER_BG)
        vc.alignment = Alignment(horizontal="center", vertical="center")

        ac = ws[f"{ls}8"]
        ac.fill = _fill(accent)

    # Row 9: spacer
    ws.row_dimensions[9].height = 14

    # DRE groups summary
    row = 10
    ws.merge_cells(f"A{row}:J{row}")
    c = ws[f"A{row}"]
    c.value = "RESUMO POR GRUPO DRE"
    c.font = Font(bold=True, size=12, color=_C_HEADER_BG)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 22
    row += 1

    grupo_header_row = row  # row 11 — table header for DRE groups
    ws[f"A{row}"].value = "Grupo DRE"
    ws[f"B{row}"].value = "Valor Total (R$)"
    _style_header_cells(ws, row, 2)
    ws.row_dimensions[row].height = 20
    row += 1

    grupo_data_start = row
    grupo_data_end = row - 1  # will advance in loop

    if dre_geral:
        grupos: dict[str, Decimal] = {}
        for dr in dre_geral:
            g = str(dr.get("grupo_dre") or "—")
            grupos[g] = grupos.get(g, Decimal("0")) + _to_decimal(dr.get("valor"))
        grand = sum(grupos.values(), Decimal("0"))
        for i, (g, v) in enumerate(grupos.items()):
            bg = _row_bg(i)
            ws[f"A{row}"].value = g
            ws[f"B{row}"].value = float(v)
            ws[f"B{row}"].number_format = _FMT_BRL
            for cl in ("A", "B"):
                ws[f"{cl}{row}"].fill = _fill(bg)
                ws[f"{cl}{row}"].border = _THIN_BORDER
            ws[f"B{row}"].alignment = Alignment(horizontal="right")
            row += 1
        grupo_data_end = row - 1

        # Data bars on value column for quick visual comparison
        _apply_databar(ws, "B", grupo_data_start, grupo_data_end, _C_CARD_ACCENT)

        for cl in ("A", "B"):
            ws[f"{cl}{row}"].fill = _fill(_C_HEADER_BG)
            ws[f"{cl}{row}"].font = Font(bold=True, color="FFFFFF")
            ws[f"{cl}{row}"].border = _THIN_BORDER
        ws[f"A{row}"].value = "TOTAL GERAL"
        ws[f"B{row}"].value = float(grand)
        ws[f"B{row}"].number_format = _FMT_BRL
        ws[f"B{row}"].alignment = Alignment(horizontal="right")
        row += 1

    row += 1  # spacer
    ws.row_dimensions[row].height = 10
    row += 1

    # Monthly summary
    ws.merge_cells(f"A{row}:J{row}")
    c = ws[f"A{row}"]
    c.value = "RESUMO MENSAL"
    c.font = Font(bold=True, size=12, color=_C_HEADER_BG)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 22
    row += 1

    ws[f"A{row}"].value = "Mês"
    ws[f"B{row}"].value = "Lançamentos OK"
    ws[f"C{row}"].value = "Valor OK (R$)"
    ws[f"D{row}"].value = "Em Revisão"
    ws[f"E{row}"].value = "Valor Revisão (R$)"
    _style_header_cells(ws, row, 5)
    ws.row_dimensions[row].height = 20
    row += 1

    for i, mr in enumerate(resumo_mensal):
        bg = _row_bg(i)
        ws[f"A{row}"].value = mr.get("competencia", "")
        ws[f"B{row}"].value = mr.get("qtd_lancamentos_ok", 0)
        ws[f"C{row}"].value = _float_safe(mr.get("valor_ok"))
        ws[f"D{row}"].value = mr.get("qtd_lancamentos_revisar", 0)
        ws[f"E{row}"].value = _float_safe(mr.get("valor_revisar"))
        for cl in ("A", "B", "C", "D", "E"):
            ws[f"{cl}{row}"].fill = _fill(bg)
            ws[f"{cl}{row}"].border = _THIN_BORDER
        ws[f"C{row}"].number_format = _FMT_BRL
        ws[f"E{row}"].number_format = _FMT_BRL
        ws[f"C{row}"].alignment = Alignment(horizontal="right")
        ws[f"E{row}"].alignment = Alignment(horizontal="right")
        row += 1

    row += 1

    # Leituras rápidas — auto-insight block
    row = _write_leituras_rapidas(ws, row, total_revisar, resumo_mensal, dre_geral, periodos)

    row += 1

    # Instructions
    ws.merge_cells(f"A{row}:J{row}")
    ws[f"A{row}"].value = "Como usar este workbook:"
    ws[f"A{row}"].font = Font(bold=True, size=10, color=_C_HEADER_BG)
    ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 20
    row += 1

    instructions = [
        "• Use os filtros das abas seguintes para analisar por mês, conta, fornecedor ou grupo DRE",
        "• 02_DRE_DETTAGLIO — DRE detalhada por grupo, linha DRE, conta e mês (filtrável)",
        "• 03_ANALISE_POR_CONTA — cada conta mês a mês com total e % do total (ordenado por valor)",
        "• 04_GIORNALE — todos os lançamentos com filtros (REVISÃO destacada em amarelo)",
        "• 05_BASE_TECNICA_AUDIT — base técnica completa para auditoria interna",
        f"• Lançamentos em REVISÃO = {total_revisar} — verificar antes de emitir DRE definitiva",
    ]
    for instr in instructions:
        ws.merge_cells(f"A{row}:J{row}")
        ws[f"A{row}"].value = instr
        ws[f"A{row}"].font = Font(size=10, color="444444")
        ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[row].height = 16
        row += 1

    # Column widths
    for cl, w in [("A", 34), ("B", 18), ("C", 18), ("D", 16), ("E", 20),
                  ("F", 14), ("G", 14), ("H", 14), ("I", 14), ("J", 14)]:
        ws.column_dimensions[cl].width = w


def _write_dre_dettaglio(
    ws,
    dre_mensal: list[dict],
    base_ok_dre: list[dict],
) -> None:
    if base_ok_dre:
        agg: dict[tuple, Decimal] = {}
        for r in base_ok_dre:
            key = (
                str(r.get("grupo_dre") or "—"),
                str(r.get("linha_dre") or "—"),
                str(r.get("conta_id") or "—"),
                str(r.get("conta_nome") or "—"),
                str(r.get("ano_mes") or "—"),
            )
            agg[key] = agg.get(key, Decimal("0")) + _to_decimal(r.get("valor"))

        headers = ["Grupo DRE", "Linha DRE", "Conta", "Nome da Conta", "Mês", "Valor"]
        widths = [32, 36, 14, 34, 10, 16]
        rows_data = [
            {
                "Grupo DRE": k[0],
                "Linha DRE": k[1],
                "Conta": k[2],
                "Nome da Conta": k[3],
                "Mês": k[4],
                "Valor": float(v),
            }
            for k, v in sorted(agg.items(), key=lambda x: (x[0][0], x[0][1], x[0][2], x[0][4]))
        ]
    else:
        headers = ["Mês", "Grupo DRE", "Linha DRE", "Valor"]
        widths = [10, 32, 36, 16]
        rows_data = [
            {
                "Mês": str(r.get("ano_mes") or "—"),
                "Grupo DRE": str(r.get("grupo_dre") or "—"),
                "Linha DRE": str(r.get("linha_dre") or "—"),
                "Valor": _float_safe(r.get("valor")),
            }
            for r in dre_mensal
        ]

    n_cols = len(headers)
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i).value = h
    _style_header_cells(ws, 1, n_cols)
    ws.row_dimensions[1].height = 22

    for ri, r in enumerate(rows_data, 2):
        bg = _row_bg(ri)
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(row=ri, column=ci)
            val = r.get(h)
            cell.value = val
            cell.fill = _fill(bg)
            cell.border = _THIN_BORDER
            if h == "Valor":
                cell.number_format = _FMT_BRL
                cell.alignment = Alignment(horizontal="right", vertical="top")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top")

    _apply_table(ws, 1, max(len(rows_data) + 1, 1), n_cols)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_analise_por_conta(
    ws,
    base_ok_dre: list[dict],
    lancamentos_ok: list[dict],
    periodos: list[str],
) -> None:
    if base_ok_dre:
        pivot: dict[tuple, dict[str, Decimal]] = {}
        for r in base_ok_dre:
            key = (
                str(r.get("conta_id") or "—"),
                str(r.get("conta_nome") or "—"),
                str(r.get("grupo_dre") or "—"),
                str(r.get("linha_dre") or "—"),
            )
            mes = str(r.get("ano_mes") or "").strip()
            if mes:
                bucket = pivot.setdefault(key, {})
                bucket[mes] = bucket.get(mes, Decimal("0")) + _to_decimal(r.get("valor"))
    else:
        pivot = {}
        for r in lancamentos_ok:
            key = (
                str(r.get("conta_id") or "—"),
                str(r.get("conta_nome") or "—"),
                str(r.get("grupo_conta") or "—"),
                str(r.get("dre_linha_nome") or "—"),
            )
            mes = str(r.get("competencia") or "").strip()
            if mes:
                bucket = pivot.setdefault(key, {})
                bucket[mes] = bucket.get(mes, Decimal("0")) + _to_decimal(r.get("valor"))

    all_mes = sorted({mes for bucket in pivot.values() for mes in bucket})
    periodos_sorted = periodos if periodos else all_mes

    fixed_headers = ["Conta", "Nome da Conta", "Grupo DRE", "Linha DRE"]
    n_fixed = len(fixed_headers)
    total_col = n_fixed + len(periodos_sorted) + 1
    pct_col = total_col + 1
    all_headers = fixed_headers + periodos_sorted + ["Total", "% do Total"]
    n_cols = len(all_headers)

    for i, h in enumerate(all_headers, 1):
        ws.cell(row=1, column=i).value = h
    _style_header_cells(ws, 1, n_cols)
    ws.row_dimensions[1].height = 22

    # Pre-calculate grand total for % column
    grand = sum(sum(b.values(), Decimal("0")) for b in pivot.values())

    # Sort by absolute Total descending (largest accounts by magnitude first)
    sorted_keys = sorted(
        pivot.keys(),
        key=lambda k: -abs(sum(pivot[k].values(), Decimal("0"))),
    )

    row = 2
    n_accounts = len(sorted_keys)
    for i, key in enumerate(sorted_keys):
        # Top 3 accounts get light green highlight (only when there are more than 3)
        bg = _C_TOP_ACCOUNT if i < 3 and n_accounts > 3 else _row_bg(i)
        bucket = pivot[key]
        total = sum(bucket.values(), Decimal("0"))

        ws.cell(row=row, column=1).value = key[0]
        ws.cell(row=row, column=2).value = key[1]
        ws.cell(row=row, column=3).value = key[2]
        ws.cell(row=row, column=4).value = key[3]

        for j, mes in enumerate(periodos_sorted, n_fixed + 1):
            val = bucket.get(mes, Decimal("0"))
            cell = ws.cell(row=row, column=j)
            cell.value = float(val) if val != Decimal("0") else None
            cell.number_format = _FMT_BRL
            cell.alignment = Alignment(horizontal="right")

        total_cell = ws.cell(row=row, column=total_col)
        total_cell.value = float(total)
        total_cell.number_format = _FMT_BRL
        total_cell.alignment = Alignment(horizontal="right")
        total_cell.font = Font(bold=True)

        pct_cell = ws.cell(row=row, column=pct_col)
        if grand != Decimal("0"):
            pct_cell.value = float(total / grand)
            pct_cell.number_format = "0.00%"
        pct_cell.alignment = Alignment(horizontal="right")

        for col in range(1, n_cols + 1):
            c = ws.cell(row=row, column=col)
            c.fill = _fill(bg)
            c.border = _THIN_BORDER
            if col <= n_fixed:
                c.alignment = Alignment(horizontal="left", vertical="top")
        row += 1

    # Grand total row
    for col in range(1, n_cols + 1):
        c = ws.cell(row=row, column=col)
        c.fill = _fill(_C_HEADER_BG)
        c.font = Font(bold=True, color="FFFFFF")
        c.border = _THIN_BORDER
    ws.cell(row=row, column=1).value = "TOTAL GERAL"
    ws.cell(row=row, column=1).alignment = Alignment(horizontal="left")
    grand_cell = ws.cell(row=row, column=total_col)
    grand_cell.value = float(grand)
    grand_cell.number_format = _FMT_BRL
    grand_cell.alignment = Alignment(horizontal="right")
    for j, mes in enumerate(periodos_sorted, n_fixed + 1):
        mes_total = sum(b.get(mes, Decimal("0")) for b in pivot.values())
        mc = ws.cell(row=row, column=j)
        mc.value = float(mes_total)
        mc.number_format = _FMT_BRL
        mc.alignment = Alignment(horizontal="right")
    pct_grand_cell = ws.cell(row=row, column=pct_col)
    pct_grand_cell.value = 1.0
    pct_grand_cell.number_format = "0.00%"
    pct_grand_cell.alignment = Alignment(horizontal="right")

    # Data bars on % do Total column
    if row > 2:
        _apply_databar(ws, get_column_letter(pct_col), 2, row - 1, "2A9D8F")

    # Freeze: cols A–D (fixed) + row 1
    ws.freeze_panes = ws.cell(row=2, column=n_fixed + 1)

    end_letter = get_column_letter(n_cols)
    ws.auto_filter.ref = f"A1:{end_letter}{max(row, 1)}"

    fixed_widths = [14, 34, 28, 34]
    for i, w in enumerate(fixed_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for j in range(n_fixed + 1, total_col):
        ws.column_dimensions[get_column_letter(j)].width = 12
    ws.column_dimensions[get_column_letter(total_col)].width = 14
    ws.column_dimensions[get_column_letter(pct_col)].width = 12


def _write_giornale(
    ws,
    lancamentos_ok: list[dict],
    lancamentos_revisar: list[dict],
    conta_lookup: dict[str, dict] | None = None,
) -> None:
    n_cols = len(_GIORNALE_COLS)
    human_headers = [c[1] for c in _GIORNALE_COLS]
    widths = [c[2] for c in _GIORNALE_COLS]
    _lookup = conta_lookup or {}

    for i, h in enumerate(human_headers, 1):
        ws.cell(row=1, column=i).value = h
    _style_header_cells(ws, 1, n_cols)
    ws.row_dimensions[1].height = 22

    row = 2
    all_rows = [("OK", r) for r in lancamentos_ok] + [("REVISAR", r) for r in lancamentos_revisar]
    for i, (status_type, r) in enumerate(all_rows):
        bg = _C_REVISAR_BG if status_type == "REVISAR" else _row_bg(i)
        conta_id = str(r.get("conta_id") or "").strip()
        dre_info = _lookup.get(conta_id, {})

        for j, (tech_key, _, _) in enumerate(_GIORNALE_COLS, 1):
            cell = ws.cell(row=row, column=j)
            # Enrich DRE columns: pre-livro dre_linha_nome/grupo_conta are unreliable;
            # use Base_OK_DRE lookup as primary source.
            if tech_key == "dre_linha_nome":
                raw = r.get("dre_linha_nome") or dre_info.get("linha_dre")
            elif tech_key == "grupo_conta":
                raw = dre_info.get("grupo_dre") or r.get("grupo_conta")
            else:
                raw = r.get(tech_key)
            cell.fill = _fill(bg)
            cell.border = _THIN_BORDER
            if tech_key == "valor":
                cell.value = float(_to_decimal(raw)) if raw is not None else None
                cell.number_format = _FMT_BRL
                cell.alignment = Alignment(horizontal="right", vertical="top")
            elif tech_key == "data_documento" and raw is not None and not isinstance(raw, str):
                cell.value = raw
                cell.number_format = _FMT_DATE
                cell.alignment = Alignment(horizontal="center", vertical="top")
            else:
                cell.value = raw
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=False)
        row += 1

    _apply_table(ws, 1, max(row - 1, 1), n_cols)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_base_tecnica(
    ws,
    lancamentos_ok: list[dict],
    lancamentos_revisar: list[dict],
) -> None:
    all_rows = [*lancamentos_ok, *lancamentos_revisar]
    if not all_rows:
        ws.cell(row=1, column=1).value = "Nenhum lancamento disponivel."
        return

    all_keys: list[str] = []
    for r in all_rows:
        for k in r:
            if k not in all_keys:
                all_keys.append(k)

    n_cols = len(all_keys)
    for i, h in enumerate(all_keys, 1):
        ws.cell(row=1, column=i).value = h
    _style_header_cells(ws, 1, n_cols)
    ws.row_dimensions[1].height = 18

    for row_i, r in enumerate(all_rows, 2):
        bg = _row_bg(row_i)
        for col_i, key in enumerate(all_keys, 1):
            cell = ws.cell(row=row_i, column=col_i)
            raw = r.get(key)
            cell.value = raw
            cell.fill = _fill(bg)
            cell.border = _THIN_BORDER
            cell.alignment = Alignment(horizontal="left", vertical="top")
            if key == "valor" and raw is not None:
                cell.number_format = _FMT_BRL

    _apply_table(ws, 1, len(all_rows) + 1, n_cols)
    for i, key in enumerate(all_keys, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(12, min(len(key) + 4, 40))
