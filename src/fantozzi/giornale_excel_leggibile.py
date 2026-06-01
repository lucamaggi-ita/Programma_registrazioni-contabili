from __future__ import annotations

import json
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
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
        dre_geral, dre_mensal, base_ok_dre, resumo_mensal,
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


def _normalize_label(value: Any) -> str:
    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def _is_receita_group(label: Any) -> bool:
    return "receita" in _normalize_label(label)


def _is_custo_despesa_group(label: Any) -> bool:
    normalized = _normalize_label(label)
    return any(part in normalized for part in ("custo", "despesa", "deducao", "deducoes"))


def _dre_group_totals(dre_geral: list[dict]) -> dict[str, Decimal]:
    grupos: dict[str, Decimal] = {}
    for dr in dre_geral:
        group_name = str(dr.get("grupo_dre") or "Sem grupo DRE")
        grupos[group_name] = grupos.get(group_name, Decimal("0")) + _to_decimal(dr.get("valor"))
    return grupos


def _dre_monthly_summary(
    dre_mensal: list[dict],
    resumo_mensal: list[dict],
    periodos: list[str],
) -> list[tuple[str, Decimal, Decimal, Decimal]]:
    monthly: dict[str, dict[str, Decimal]] = {}
    for row in dre_mensal:
        mes = str(row.get("ano_mes") or "").strip()
        if not mes:
            continue
        bucket = monthly.setdefault(
            mes,
            {"receitas": Decimal("0"), "custos": Decimal("0"), "resultado": Decimal("0")},
        )
        value = _to_decimal(row.get("valor"))
        group_name = row.get("grupo_dre")
        bucket["resultado"] += value
        if _is_receita_group(group_name):
            bucket["receitas"] += value
        elif _is_custo_despesa_group(group_name):
            bucket["custos"] += abs(value)

    if not monthly:
        for row in resumo_mensal:
            mes = str(row.get("competencia") or "").strip()
            if not mes:
                continue
            value = _to_decimal(row.get("valor_ok"))
            monthly[mes] = {
                "receitas": Decimal("0"),
                "custos": abs(value),
                "resultado": value,
            }

    ordered_months = periodos or sorted(monthly)
    return [
        (
            mes,
            monthly.get(mes, {}).get("receitas", Decimal("0")),
            monthly.get(mes, {}).get("custos", Decimal("0")),
            monthly.get(mes, {}).get("resultado", Decimal("0")),
        )
        for mes in ordered_months
    ]


def _write_section_title(ws, row: int, title: str, start_col: str = "A", end_col: str = "J") -> int:
    ws.merge_cells(f"{start_col}{row}:{end_col}{row}")
    cell = ws[f"{start_col}{row}"]
    cell.value = title
    cell.fill = _fill("FFFFFF")
    cell.font = Font(bold=True, size=12, color=_C_HEADER_BG)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 22
    return row + 1


def _write_leituras_rapidas(
    ws,
    start_row: int,
    lancamentos_revisar: list[dict],
    resumo_mensal: list[dict],
    dre_geral: list[dict],
    base_ok_dre: list[dict],
    periodos: list[str],
) -> int:
    row = _write_section_title(ws, start_row, "LEITURAS RÁPIDAS")
    insights: list[str] = []

    grupos = _dre_group_totals(dre_geral)
    custo_groups = {g: v for g, v in grupos.items() if _is_custo_despesa_group(g)}
    if custo_groups:
        maior_grupo = max(custo_groups, key=lambda g: abs(custo_groups[g]))
        insights.append(
            f"Maior grupo de custo/despesa: {maior_grupo}"
            f"  —  R$ {_fmt_brl_value(abs(custo_groups[maior_grupo]))}"
        )
    elif grupos:
        maior_grupo = max(grupos, key=lambda g: abs(grupos[g]))
        insights.append(
            f"Grupo DRE mais relevante: {maior_grupo}"
            f"  —  R$ {_fmt_brl_value(abs(grupos[maior_grupo]))}"
        )

    mes_valores = [
        (str(r.get("competencia") or ""), abs(_to_decimal(r.get("valor_ok"))))
        for r in resumo_mensal
        if r.get("competencia")
    ]
    if mes_valores:
        mes_max = max(mes_valores, key=lambda x: x[1])
        insights.append(
            f"Mês com maior custo/despesa: {mes_max[0]}"
            f"  —  R$ {_fmt_brl_value(mes_max[1])}"
        )

    conta_totals: dict[str, Decimal] = {}
    for row_data in base_ok_dre:
        conta = str(row_data.get("conta_nome") or row_data.get("conta_id") or "").strip()
        if conta:
            conta_totals[conta] = conta_totals.get(conta, Decimal("0")) + _to_decimal(row_data.get("valor"))
    if conta_totals:
        conta_top = max(conta_totals, key=lambda c: abs(conta_totals[c]))
        insights.append(
            f"Conta mais relevante: {conta_top}"
            f"  —  R$ {_fmt_brl_value(abs(conta_totals[conta_top]))}"
        )

    total_revisar = len(lancamentos_revisar)
    valor_revisar = sum(_to_decimal(r.get("valor")) for r in lancamentos_revisar)
    if total_revisar > 0:
        insights.append(
            f"Lançamentos em revisão: {total_revisar}"
            f"  —  R$ {_fmt_brl_value(abs(valor_revisar))}"
        )
    elif not insights:
        insights.append("Não disponível")
    else:
        insights.append("Lançamentos em revisão: 0")

    for insight in insights:
        ws.merge_cells(f"A{row}:J{row}")
        cell = ws[f"A{row}"]
        cell.value = f"• {insight}"
        cell.fill = _fill(_C_LEITURA_BG)
        cell.font = Font(size=10, color="333333")
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cell.border = _THIN_BORDER
        ws.row_dimensions[row].height = 19
        row += 1

    return row


def _write_pendencias(ws, start_row: int, lancamentos_revisar: list[dict]) -> int:
    row = _write_section_title(ws, start_row, "PENDÊNCIAS / REVISAR")
    total_revisar = len(lancamentos_revisar)
    valor_revisar = sum(_to_decimal(r.get("valor")) for r in lancamentos_revisar)

    ws.merge_cells(f"A{row}:J{row}")
    cell = ws[f"A{row}"]
    if total_revisar == 0:
        cell.value = "Nenhuma pendência crítica identificada no dashboard."
    else:
        cell.value = (
            f"{total_revisar} lançamento(s) em revisão"
            f"  —  R$ {_fmt_brl_value(abs(valor_revisar))}"
        )
    cell.fill = _fill("FFF2CC" if total_revisar else "E2EFDA")
    cell.font = Font(bold=True, size=10, color="7A4A00" if total_revisar else "375623")
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = _THIN_BORDER
    ws.row_dimensions[row].height = 20
    row += 1

    if total_revisar:
        buckets: dict[str, tuple[int, Decimal]] = {}
        for item in lancamentos_revisar:
            label = str(
                item.get("conta_nome")
                or item.get("fornecedor_ou_favorecido")
                or item.get("conta_id")
                or "Sem identificação"
            )
            count, value = buckets.get(label, (0, Decimal("0")))
            buckets[label] = (count + 1, value + _to_decimal(item.get("valor")))

        for label, (count, value) in sorted(
            buckets.items(),
            key=lambda kv: (-kv[1][0], -abs(kv[1][1]), kv[0]),
        )[:3]:
            ws.merge_cells(f"A{row}:J{row}")
            cell = ws[f"A{row}"]
            cell.value = f"• {label}: {count} lançamento(s) — R$ {_fmt_brl_value(abs(value))}"
            cell.fill = _fill("FFFFFF")
            cell.font = Font(size=10, color="444444")
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = _THIN_BORDER
            ws.row_dimensions[row].height = 18
            row += 1

    return row


def _write_dashboard_links(ws, start_row: int) -> int:
    row = _write_section_title(ws, start_row, "PARA DETALHES")
    links = [
        ("02_DRE_DETTAGLIO", "DRE detalhada"),
        ("03_ANALISE_POR_CONTA", "Análise por conta"),
        ("04_GIORNALE", "Giornale filtrável"),
        ("05_BASE_TECNICA_AUDIT", "Base técnica"),
    ]
    for col, (sheet_name, label) in enumerate(links, 1):
        cell = ws.cell(row=row, column=col)
        cell.value = f"{sheet_name} — {label}"
        cell.hyperlink = f"#'{sheet_name}'!A1"
        cell.style = "Hyperlink"
        cell.font = Font(color=_C_CARD_ACCENT, underline="single", size=9)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 32
    return row + 1


def _write_dashboard_helper_tables(
    ws,
    monthly_rows: list[tuple[str, Decimal, Decimal, Decimal]],
    group_totals: dict[str, Decimal],
) -> tuple[tuple[int, int], tuple[int, int]]:
    helper_col = 12  # L
    row = 2
    monthly_start = row
    monthly_headers = ["Mês", "Receitas", "Custos / Despesas", "Resultado"]
    for offset, header in enumerate(monthly_headers):
        ws.cell(row=row, column=helper_col + offset).value = header
    row += 1
    data_rows = monthly_rows or [("N/D", Decimal("0"), Decimal("0"), Decimal("0"))]
    for mes, receitas, custos, resultado in data_rows:
        ws.cell(row=row, column=helper_col).value = mes
        ws.cell(row=row, column=helper_col + 1).value = float(receitas)
        ws.cell(row=row, column=helper_col + 2).value = float(custos)
        ws.cell(row=row, column=helper_col + 3).value = float(resultado)
        row += 1
    monthly_end = row - 1

    row += 2
    group_start = row
    ws.cell(row=row, column=helper_col).value = "Grupo DRE"
    ws.cell(row=row, column=helper_col + 1).value = "Valor total"
    row += 1
    group_items = sorted(group_totals.items(), key=lambda kv: abs(kv[1]), reverse=True)
    if not group_items:
        group_items = [("Não disponível", Decimal("0"))]
    for group_name, value in group_items[:8]:
        ws.cell(row=row, column=helper_col).value = group_name
        ws.cell(row=row, column=helper_col + 1).value = float(abs(value))
        row += 1
    group_end = row - 1

    for col in range(helper_col, helper_col + 4):
        ws.column_dimensions[get_column_letter(col)].width = 18
        for r in range(2, group_end + 1):
            cell = ws.cell(row=r, column=col)
            cell.font = Font(size=8, color="666666")
            if col > helper_col and isinstance(cell.value, (int, float)):
                cell.number_format = _FMT_BRL

    return (monthly_start, monthly_end), (group_start, group_end)


def _add_dashboard_charts(
    ws,
    monthly_range: tuple[int, int],
    group_range: tuple[int, int],
) -> None:
    monthly_start, monthly_end = monthly_range
    group_start, group_end = group_range

    line = LineChart()
    line.title = "Evolução mensal DRE"
    line.style = 13
    line.height = 8
    line.width = 16
    line.y_axis.title = "R$"
    line.x_axis.title = "Mês"
    line_data = Reference(ws, min_col=13, max_col=15, min_row=monthly_start, max_row=monthly_end)
    line_cats = Reference(ws, min_col=12, min_row=monthly_start + 1, max_row=monthly_end)
    line.add_data(line_data, titles_from_data=True)
    line.set_categories(line_cats)
    line.legend.position = "b"
    ws.add_chart(line, "D11")

    bar = BarChart()
    bar.type = "bar"
    bar.title = "Composição por Grupo DRE"
    bar.style = 10
    bar.height = 8
    bar.width = 16
    bar.y_axis.title = "Grupo DRE"
    bar.x_axis.title = "R$"
    bar_data = Reference(ws, min_col=13, min_row=group_start, max_row=group_end)
    bar_cats = Reference(ws, min_col=12, min_row=group_start + 1, max_row=group_end)
    bar.add_data(bar_data, titles_from_data=True)
    bar.set_categories(bar_cats)
    bar.legend = None
    ws.add_chart(bar, "D27")


def _write_dashboard(
    ws,
    run_id: str | None,
    periodos: list[str],
    lancamentos_ok: list[dict],
    lancamentos_revisar: list[dict],
    dre_geral: list[dict],
    dre_mensal: list[dict],
    base_ok_dre: list[dict],
    resumo_mensal: list[dict],
) -> None:
    ws.sheet_view.showGridLines = False
    period_str = " · ".join(periodos) if periodos else "—"
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    total_ok = len(lancamentos_ok)
    total_revisar = len(lancamentos_revisar)
    valor_ok = sum(_to_decimal(r.get("valor")) for r in lancamentos_ok)
    valor_revisar = sum(_to_decimal(r.get("valor")) for r in lancamentos_revisar)
    grupos = _dre_group_totals(dre_geral)
    receitas = sum(v for g, v in grupos.items() if _is_receita_group(g))
    custos_despesas = sum(abs(v) for g, v in grupos.items() if _is_custo_despesa_group(g))
    resultado = sum(grupos.values(), Decimal("0")) if grupos else valor_ok
    monthly_rows = _dre_monthly_summary(dre_mensal, resumo_mensal, periodos)
    monthly_range, group_range = _write_dashboard_helper_tables(ws, monthly_rows, grupos)

    # Row 1: status badge (10 cols)
    ws.merge_cells("A1:J1")
    c = ws["A1"]
    c.value = f"TESTE TÉCNICO / NÃO OFICIAL  —  STATUS: {STATUS}"
    c.fill = _fill(_C_STATUS_BADGE)
    c.font = Font(bold=True, color="FFFFFF", size=10)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # Rows 2–3: main title
    ws.merge_cells("A2:J3")
    c = ws["A2"]
    c.value = "PAINEL ECONÔMICO GERENCIAL"
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

    # Cards: rows 6–9 (label / value / hint / accent bar), 2 cols per card × 5 cards
    cards_data = [
        ("RECEITAS", f"R$ {_fmt_brl_value(receitas)}", "DRE", "2A9D8F"),
        ("CUSTOS / DESPESAS", f"R$ {_fmt_brl_value(custos_despesas)}", "DRE", "E76F51"),
        ("RESULTADO", f"R$ {_fmt_brl_value(resultado)}", "DRE", _C_CARD_ACCENT),
        ("LANÇAMENTOS OK", str(total_ok), "classificados", "2A9D8F"),
        ("EM REVISÃO", str(total_revisar), f"R$ {_fmt_brl_value(abs(valor_revisar))}", "E76F51" if total_revisar > 0 else "AAAAAA"),
    ]
    ws.row_dimensions[6].height = 16
    ws.row_dimensions[7].height = 42
    ws.row_dimensions[8].height = 15
    ws.row_dimensions[9].height = 5

    for idx, (label, value, hint, accent) in enumerate(cards_data):
        cs = idx * 2 + 1
        ce = cs + 1
        ls = get_column_letter(cs)
        le = get_column_letter(ce)

        ws.merge_cells(f"{ls}6:{le}6")
        ws.merge_cells(f"{ls}7:{le}7")
        ws.merge_cells(f"{ls}8:{le}8")
        ws.merge_cells(f"{ls}9:{le}9")

        lc = ws[f"{ls}6"]
        lc.value = label
        lc.fill = _fill(_C_CARD_BG)
        lc.font = Font(size=9, color="666666")
        lc.alignment = Alignment(horizontal="center", vertical="bottom")

        vc = ws[f"{ls}7"]
        vc.value = value
        vc.fill = _fill(_C_CARD_BG)
        vc.font = Font(bold=True, size=24, color=_C_HEADER_BG)
        vc.alignment = Alignment(horizontal="center", vertical="center")

        hc = ws[f"{ls}8"]
        hc.value = hint
        hc.fill = _fill(_C_CARD_BG)
        hc.font = Font(size=9, color="666666")
        hc.alignment = Alignment(horizontal="center", vertical="top")

        ac = ws[f"{ls}9"]
        ac.fill = _fill(accent)

    # Row 10: spacer
    ws.row_dimensions[10].height = 12

    _add_dashboard_charts(ws, monthly_range, group_range)

    row = 44
    row = _write_leituras_rapidas(
        ws, row, lancamentos_revisar, resumo_mensal, dre_geral, base_ok_dre, periodos,
    )

    row += 1
    row = _write_pendencias(ws, row, lancamentos_revisar)

    row += 1
    _write_dashboard_links(ws, row)

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
