from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import Iterable

import yaml
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from fantozzi.relatorios_gerenciais_ok import (
    IDENTIFIER_HEADER,
    SOURCE_HEADER,
    _derive_identifier,
    _format_currency,
    _is_bonificacao_recebida_decidida,
    _linha_bonificacao_recebida_ok,
    _lancamento_matches_special,
    _read_lancamentos,
    _read_operacoes_especiais,
    _resolve_excel_paths,
    _special_keys,
    _sum_values,
    _text,
)


DRE_GERAL_SHEET = "DRE_Geral"
DRE_MENSAL_SHEET = "DRE_Mensal"
DRE_CNPJ_SHEET = "DRE_por_CNPJ"
BASE_OK_DRE_SHEET = "Base_OK_DRE"
CONTAS_SEM_MAPEAMENTO_SHEET = "Contas_sem_Mapeamento"
AUDITORIA_DRE_SHEET = "Auditoria_DRE"
INSTRUCOES_SHEET = "Instrucoes"

DRE_GERAL_HEADERS = ["ordem", "grupo_dre", "linha_dre", "valor", "observacao"]
DRE_MENSAL_HEADERS = ["ano_mes", "ordem", "grupo_dre", "linha_dre", "valor"]
DRE_CNPJ_HEADERS = ["identificador_operacional", "grupo_dre", "linha_dre", "valor"]
BASE_HEADERS = [
    "data",
    "ano_mes",
    "identificador_operacional",
    "fornecedor",
    "documento",
    "descricao",
    "conta_id",
    "conta_nome",
    "grupo_dre",
    "linha_dre",
    "valor",
    "status_revisao",
    "tipo_operacao_especial",
    "tratamento_operacao_especial",
    "origem_arquivo",
]
CONTAS_SEM_MAPEAMENTO_HEADERS = [
    "conta_id",
    "conta_nome",
    "quantidade_lancamentos",
    "valor_total",
    "observacao",
]
AUDITORIA_HEADERS = ["indicador", "valor", "status", "observacao"]

DEFAULT_MAPEAMENTO_DRE = (
    Path(__file__).resolve().parents[2] / "config" / "mapeamento_dre.yaml"
)


class DreGerencialProvisoriaError(Exception):
    pass


@dataclass(frozen=True)
class DreGerencialProvisoriaResult:
    entradas: tuple[Path, ...]
    excel_paths: tuple[Path, ...]
    mapeamento_path: Path
    xlsx_path: Path
    report_path: Path | None
    total_lancamentos_origem: int
    lancamentos_base_ok: int
    lancamentos_revisar_excluidos: int
    operacoes_especiais_excluidas: int
    valor_total_base_ok: Decimal
    valor_total_mapeado: Decimal
    valor_total_sem_mapeamento: Decimal
    quantidade_contas_sem_mapeamento: int
    status_dre: str


def gerar_dre_gerencial_ok_provisoria(
    entradas: Iterable[Path | str],
    mapeamento_path: Path | str,
    xlsx_path: Path | str,
    report_path: Path | str | None = None,
    pastas_operacoes_especiais: Iterable[Path | str] | None = None,
) -> DreGerencialProvisoriaResult:
    entradas = tuple(Path(path) for path in entradas)
    mapeamento_path = Path(mapeamento_path)
    xlsx_path = Path(xlsx_path)
    report_path = Path(report_path) if report_path else None
    pastas_operacoes = tuple(Path(path) for path in (pastas_operacoes_especiais or ()))
    excel_paths = _resolve_excel_paths(entradas)
    mapping = _load_mapeamento(mapeamento_path)

    all_rows: list[dict[str, object]] = []
    for excel_path in excel_paths:
        identifier = _derive_identifier(excel_path)
        for row in _read_lancamentos(excel_path):
            all_rows.append({IDENTIFIER_HEADER: identifier, SOURCE_HEADER: str(excel_path), **row})

    operacoes = _read_operacoes_especiais(pastas_operacoes)
    special_keys = _special_keys(operacoes)
    base_ok: list[dict[str, object]] = []
    excluded_revisar: list[dict[str, object]] = []
    for row in all_rows:
        if _is_bonificacao_recebida_decidida(row):
            base_ok.append(_linha_bonificacao_recebida_ok(row))
            continue
        if _status(row) == "OK" and not _lancamento_matches_special(row, special_keys):
            base_ok.append(row)
            continue
        if _status(row) != "OK":
            excluded_revisar.append(row)
    base_rows = [_dre_base_row(row, mapping) for row in base_ok]
    sem_mapeamento = [row for row in base_rows if not row["linha_dre_id"]]
    mapped = [row for row in base_rows if row["linha_dre_id"] and row["entra_dre"]]
    dre_geral = _dre_geral_rows(mapped)
    dre_mensal = _dre_mensal_rows(mapped)
    dre_cnpj = _dre_cnpj_rows(mapped)
    contas_sem = _contas_sem_mapeamento_rows(sem_mapeamento)
    status = "ATENCAO_MAPEAMENTO_INCOMPLETO" if contas_sem else "OK_PROVISORIO"

    result = DreGerencialProvisoriaResult(
        entradas=entradas,
        excel_paths=excel_paths,
        mapeamento_path=mapeamento_path,
        xlsx_path=xlsx_path,
        report_path=report_path,
        total_lancamentos_origem=len(all_rows),
        lancamentos_base_ok=len(base_ok),
        lancamentos_revisar_excluidos=len(excluded_revisar),
        operacoes_especiais_excluidas=len(operacoes),
        valor_total_base_ok=_sum_values(base_ok, "valor"),
        valor_total_mapeado=sum((_abs_decimal(row["valor"]) for row in mapped), Decimal("0.00")).quantize(Decimal("0.01")),
        valor_total_sem_mapeamento=_sum_values(sem_mapeamento, "valor_original"),
        quantidade_contas_sem_mapeamento=len(contas_sem),
        status_dre=status,
    )

    workbook = Workbook()
    _write_table(workbook.active, DRE_GERAL_SHEET, DRE_GERAL_HEADERS, dre_geral)
    _write_table(workbook.create_sheet(DRE_MENSAL_SHEET), DRE_MENSAL_SHEET, DRE_MENSAL_HEADERS, dre_mensal)
    _write_table(workbook.create_sheet(DRE_CNPJ_SHEET), DRE_CNPJ_SHEET, DRE_CNPJ_HEADERS, dre_cnpj)
    _write_table(workbook.create_sheet(BASE_OK_DRE_SHEET), BASE_OK_DRE_SHEET, BASE_HEADERS, base_rows)
    _write_table(workbook.create_sheet(CONTAS_SEM_MAPEAMENTO_SHEET), CONTAS_SEM_MAPEAMENTO_SHEET, CONTAS_SEM_MAPEAMENTO_HEADERS, contas_sem or [_no_unmapped_row()])
    _write_table(workbook.create_sheet(AUDITORIA_DRE_SHEET), AUDITORIA_DRE_SHEET, AUDITORIA_HEADERS, _auditoria_rows(result))
    _write_instrucoes(workbook.create_sheet(INSTRUCOES_SHEET))

    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(xlsx_path)
    if report_path:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(relatorio_texto(result, dre_geral, dre_mensal), encoding="utf-8", newline="\n")
    return result


def console_lines(result: DreGerencialProvisoriaResult) -> list[str]:
    lines = [
        "[OK] DRE gerencial OK provisoria gerada.",
        f"     Total origem: {result.total_lancamentos_origem}",
        f"     Lancamentos gerenciais usados: {result.lancamentos_base_ok}",
        f"     Valor total base OK: {_format_currency(result.valor_total_base_ok)}",
        f"     Contas sem mapeamento: {result.quantidade_contas_sem_mapeamento}",
        f"     Status DRE: {result.status_dre}",
        f"     XLSX gerado: {result.xlsx_path}",
    ]
    if result.report_path:
        lines.append(f"     Relatorio: {result.report_path}")
    return lines


def relatorio_texto(
    result: DreGerencialProvisoriaResult,
    dre_geral: list[dict[str, object]],
    dre_mensal: list[dict[str, object]],
) -> str:
    lines = [
        "DRE GERENCIAL PROVISORIA - LANCAMENTOS GERENCIAIS OK",
        "=================================================",
        "",
        "Status: PROVISORIO",
        "",
        "Base:",
        f"- Total lancamentos origem: {result.total_lancamentos_origem}",
        f"- Lancamentos gerenciais usados: {result.lancamentos_base_ok}",
        f"- Lancamentos REVISAR excluidos: {result.lancamentos_revisar_excluidos}",
        f"- Operacoes especiais excluidas: {result.operacoes_especiais_excluidas}",
        f"- Valor total base OK: {_format_currency(result.valor_total_base_ok)}",
        "",
        "Mapeamento:",
        f"- Arquivo mapeamento: {result.mapeamento_path}",
        f"- Contas mapeadas: {max(0, len(dre_geral))}",
        f"- Contas sem mapeamento: {result.quantidade_contas_sem_mapeamento}",
        f"- Valor sem mapeamento: {_format_currency(result.valor_total_sem_mapeamento)}",
        "",
        "DRE Geral:",
        *_dre_lines(dre_geral),
        "",
        "DRE Mensal:",
        *_dre_lines(dre_mensal[:24], include_month=True),
        "",
        "Auditoria:",
        f"- Status DRE: {result.status_dre}",
        "- Alertas:",
        "- Revisao contabil real pendente.",
        "- Nao e DRE definitiva.",
        "- Nao e livro diario.",
        "- Nao usar para obrigacoes fiscais.",
        "",
        "Limitacoes:",
        "- Revisao contabil real pendente.",
        "- Nao e DRE definitiva.",
        "- Nao e livro diario.",
        "- Nao usar para obrigacoes fiscais.",
    ]
    return "\n".join(lines) + "\n"


def _load_mapeamento(path: Path) -> dict[str, object]:
    if not path.exists():
        raise DreGerencialProvisoriaError(f"Mapeamento DRE nao encontrado: {path}")
    try:
        with path.open("r", encoding="utf-8") as file:
            data = yaml.safe_load(file) or {}
    except Exception as exc:
        raise DreGerencialProvisoriaError(f"Mapeamento DRE nao pode ser lido: {path}") from exc
    linhas = {}
    for row in data.get("linhas_dre", []) or []:
        if isinstance(row, dict) and row.get("id"):
            linhas[_text(row["id"])] = row
    mapeamento = data.get("mapeamento", {}) or {}
    if not isinstance(mapeamento, dict):
        raise DreGerencialProvisoriaError("Mapeamento DRE invalido: campo mapeamento.")
    return {"linhas": linhas, "mapeamento": mapeamento}


def _dre_base_row(row: dict[str, object], mapping: dict[str, object]) -> dict[str, object]:
    conta_id = _text(row.get("conta_id"))
    map_row = mapping["mapeamento"].get(conta_id) or {}
    linha_id = _text(map_row.get("dre_linha_id")) if isinstance(map_row, dict) else ""
    linha = mapping["linhas"].get(linha_id, {})
    entra_dre = bool(map_row.get("entra_dre", True)) if isinstance(map_row, dict) else False
    valor_original = _abs_decimal(row.get("valor"))
    valor_signed = _signed_value(valor_original, _text(linha.get("sinal")))
    return {
        "data": row.get("data_documento"),
        "ano_mes": _text(row.get("competencia")) or _ano_mes(row.get("data_documento")),
        "identificador_operacional": row.get(IDENTIFIER_HEADER),
        "fornecedor": row.get("fornecedor_ou_favorecido"),
        "documento": row.get("documento"),
        "descricao": row.get("descricao_itens"),
        "conta_id": conta_id,
        "conta_nome": row.get("conta_nome"),
        "grupo_dre": _text(linha.get("grupo")),
        "linha_dre": _text(linha.get("nome")),
        "linha_dre_id": linha_id,
        "ordem": int(linha.get("ordem") or 999999),
        "sinal": _text(linha.get("sinal")),
        "entra_dre": entra_dre,
        "valor": valor_signed,
        "valor_original": valor_original,
        "status_revisao": row.get("status_revisao"),
        "tipo_operacao_especial": row.get("tipo_operacao_especial"),
        "tratamento_operacao_especial": row.get("tratamento_operacao_especial"),
        "origem_arquivo": row.get(SOURCE_HEADER),
    }


def _dre_geral_rows(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    grouped = _group_dre(rows, ("ordem", "grupo_dre", "linha_dre"))
    return [
        {"ordem": key[0], "grupo_dre": key[1], "linha_dre": key[2], "valor": value, "observacao": "PROVISORIO - OK + bonificacao decidida"}
        for key, value in sorted(grouped.items(), key=lambda item: item[0])
    ]


def _dre_mensal_rows(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    grouped = _group_dre(rows, ("ano_mes", "ordem", "grupo_dre", "linha_dre"))
    return [
        {"ano_mes": key[0], "ordem": key[1], "grupo_dre": key[2], "linha_dre": key[3], "valor": value}
        for key, value in sorted(grouped.items(), key=lambda item: item[0])
    ]


def _dre_cnpj_rows(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    grouped = _group_dre(rows, ("identificador_operacional", "grupo_dre", "linha_dre"))
    return [
        {"identificador_operacional": key[0], "grupo_dre": key[1], "linha_dre": key[2], "valor": value}
        for key, value in sorted(grouped.items(), key=lambda item: item[0])
    ]


def _group_dre(rows: list[dict[str, object]], fields: tuple[str, ...]) -> dict[tuple[object, ...], Decimal]:
    grouped: dict[tuple[object, ...], Decimal] = defaultdict(lambda: Decimal("0.00"))
    for row in rows:
        grouped[tuple(row.get(field) for field in fields)] += _abs_decimal(row.get("valor"))
    return {key: value.quantize(Decimal("0.01")) for key, value in grouped.items()}


def _contas_sem_mapeamento_rows(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    grouped: dict[tuple[str, str], dict[str, object]] = {}
    for row in rows:
        key = (_text(row.get("conta_id")), _text(row.get("conta_nome")))
        bucket = grouped.setdefault(
            key,
            {
                "conta_id": key[0],
                "conta_nome": key[1],
                "quantidade_lancamentos": 0,
                "valor_total": Decimal("0.00"),
                "observacao": "Conta gerencial sem mapeamento DRE.",
            },
        )
        bucket["quantidade_lancamentos"] = int(bucket["quantidade_lancamentos"]) + 1
        bucket["valor_total"] = bucket["valor_total"] + _abs_decimal(row.get("valor_original"))
    return sorted(grouped.values(), key=lambda row: _text(row["conta_id"]))


def _auditoria_rows(result: DreGerencialProvisoriaResult) -> list[dict[str, object]]:
    return [
        {"indicador": "total_lancamentos_base_ok", "valor": result.lancamentos_base_ok, "status": "OK", "observacao": "Status OK mais BONIFICACAO_RECEBIDA decidida em 2.05.003."},
        {"indicador": "valor_total_base_ok", "valor": result.valor_total_base_ok, "status": "OK", "observacao": "Valor da base gerencial."},
        {"indicador": "valor_total_mapeado", "valor": result.valor_total_mapeado, "status": "OK", "observacao": "Valor bruto das contas mapeadas."},
        {"indicador": "valor_total_sem_mapeamento", "valor": result.valor_total_sem_mapeamento, "status": "ATENCAO" if result.valor_total_sem_mapeamento else "OK", "observacao": "Contas sem mapeamento devem ser revisadas."},
        {"indicador": "quantidade_contas_sem_mapeamento", "valor": result.quantidade_contas_sem_mapeamento, "status": "ATENCAO" if result.quantidade_contas_sem_mapeamento else "OK", "observacao": "Quantidade de contas sem linha DRE."},
        {"indicador": "status_dre", "valor": result.status_dre, "status": result.status_dre, "observacao": "Status consolidado da DRE provisoria."},
    ]


def _write_table(sheet, title: str, headers: list[str], rows: list[dict[str, object]]) -> None:
    sheet.title = title
    sheet.append(headers)
    for row in rows:
        sheet.append([_cell_value(row.get(header)) for header in headers])
    _style_table(sheet)
    _set_widths(sheet, headers)


def _write_instrucoes(sheet) -> None:
    lines = [
        "DRE GERENCIAL PROVISORIA - LANCAMENTOS GERENCIAIS OK",
        "",
        "Esta DRE e gerencial, provisoria e parcial.",
        "Inclui lancamentos com status OK.",
        "Inclui BONIFICACAO_RECEBIDA decidida em 2.05.003 como reducao de custo.",
        "Nao promove genericamente lancamentos REVISAR.",
        "Nao inclui COMODATO nem operacoes especiais sem decisao contabil segura.",
        "Nao substitui DRE contabil definitiva.",
        "Nao substitui revisao do contador.",
        "Contas sem mapeamento devem ser revisadas antes de uma DRE final.",
    ]
    for line in lines:
        sheet.append([line])
    sheet["A1"].font = Font(bold=True, size=14)
    sheet.column_dimensions["A"].width = 120
    for row in sheet.iter_rows():
        row[0].alignment = Alignment(wrap_text=True, vertical="top")


def _no_unmapped_row() -> dict[str, object]:
    return {"conta_id": "-", "conta_nome": "Nenhuma conta sem mapeamento encontrada", "quantidade_lancamentos": 0, "valor_total": Decimal("0.00"), "observacao": "OK"}


def _dre_lines(rows: list[dict[str, object]], *, include_month: bool = False) -> list[str]:
    if not rows:
        return ["-"]
    lines = []
    for row in rows:
        prefix = f"{row.get('ano_mes')} | " if include_month else ""
        lines.append(f"- {prefix}{row.get('grupo_dre')} | {row.get('linha_dre')} | {_format_currency(row.get('valor'))}")
    return lines


def _signed_value(value: Decimal, sinal: str) -> Decimal:
    if sinal == "-":
        return -value
    return value


def _status(row: dict[str, object]) -> str:
    return _text(row.get("status_revisao")).upper()


def _ano_mes(value: object) -> str:
    text = _text(value)
    if len(text) >= 7:
        return text[:7]
    return "SEM_COMPETENCIA"


def _abs_decimal(value: object) -> Decimal:
    if isinstance(value, Decimal):
        return abs(value).quantize(Decimal("0.01"))
    return abs(Decimal(str(value or "0"))).quantize(Decimal("0.01"))


def _style_table(sheet) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(bold=True, color="FFFFFF")
    for cell in sheet[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if isinstance(cell.value, (int, float)) and "valor" in _text(sheet.cell(row=1, column=cell.column).value):
                cell.number_format = '#,##0.00'


def _set_widths(sheet, headers: list[str]) -> None:
    for index, header in enumerate(headers, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = max(14, min(len(header) + 8, 52))


def _cell_value(value: object) -> object:
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, bool):
        return "SIM" if value else "NAO"
    return value
