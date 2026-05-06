from __future__ import annotations

import csv
from pathlib import Path
from typing import Iterable

import yaml
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from fantozzi.classification_review_export import CSV_HEADERS


DEFAULT_PLANO_CONTAS = (
    Path(__file__).resolve().parents[2] / "config" / "plano_contas_gerencial.yaml"
)

REVISAO_SHEET = "Revisao"
PLANO_CONTAS_SHEET = "Plano_de_Contas"
ACOES_VALIDAS_SHEET = "Acoes_Validas"
INSTRUCOES_SHEET = "Instrucoes"

EDITABLE_COLUMNS = {
    "acao_manual",
    "conta_id_sugerida",
    "conta_nome_sugerida",
    "palavras_chave_sugeridas",
    "observacao_manual",
}

ALLOWED_ACTIONS_DESCRIPTIONS = {
    "ANALISAR_COM_CONTADOR": (
        "Usar quando a decisao depende de validacao contabil/fiscal."
    ),
    "CLASSIFICAR": (
        "Usar quando o produto deve ser classificado diretamente em uma conta existente."
    ),
    "CRIAR_ALIAS": (
        "Usar quando a conta atual esta correta e queremos ensinar uma variacao de nome."
    ),
    "CRIAR_REGRA": (
        "Usar quando queremos criar uma regra mais geral de classificacao."
    ),
    "IGNORAR": "Usar quando o candidato nao deve gerar regra ou classificacao.",
    "MANTER_REVISAO": (
        "Usar quando o item deve permanecer pendente para analise futura."
    ),
}


class ClassificationReviewWorkbookError(Exception):
    pass


def gerar_planilha_revisao_classificacao(
    csv_path: Path | str,
    output_path: Path | str,
    plano_contas_path: Path | str | None = None,
) -> Path:
    csv_path = Path(csv_path)
    output_path = Path(output_path)
    plano_contas_path = Path(plano_contas_path or DEFAULT_PLANO_CONTAS)

    headers, rows = ler_csv_revisao(csv_path)
    contas = carregar_plano_contas(plano_contas_path)

    workbook = Workbook()
    criar_aba_revisao(workbook, headers, rows)
    criar_aba_plano_de_contas(workbook, contas)
    criar_aba_acoes_validas(workbook)
    criar_aba_instrucoes(workbook)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return output_path


def ler_csv_revisao(csv_path: Path | str) -> tuple[list[str], list[dict[str, str]]]:
    csv_path = Path(csv_path)
    if not csv_path.exists():
        raise ClassificationReviewWorkbookError(
            f"CSV de revisao nao encontrado: {csv_path}"
        )
    if not csv_path.is_file():
        raise ClassificationReviewWorkbookError(
            f"CSV de revisao nao e um arquivo: {csv_path}"
        )
    if csv_path.stat().st_size == 0:
        raise ClassificationReviewWorkbookError("CSV de revisao esta vazio.")

    try:
        with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
            reader = csv.DictReader(file, delimiter=";")
            if not reader.fieldnames:
                raise ClassificationReviewWorkbookError(
                    "CSV de revisao nao possui cabecalho valido."
                )
            headers = [header.strip() for header in reader.fieldnames if header]
            missing_headers = [header for header in CSV_HEADERS if header not in headers]
            if missing_headers:
                raise ClassificationReviewWorkbookError(
                    "Colunas obrigatorias ausentes: " + ", ".join(missing_headers)
                )
            return headers, [dict(row) for row in reader]
    except UnicodeDecodeError as exc:
        raise ClassificationReviewWorkbookError(
            f"CSV de revisao nao pode ser lido em UTF-8: {csv_path}"
        ) from exc


def carregar_plano_contas(plano_path: Path | str) -> list[dict[str, str]]:
    plano_path = Path(plano_path)
    if not plano_path.exists():
        raise ClassificationReviewWorkbookError(
            f"Plano de contas nao encontrado: {plano_path}"
        )
    if not plano_path.is_file():
        raise ClassificationReviewWorkbookError(
            f"Plano de contas nao e um arquivo: {plano_path}"
        )

    try:
        with plano_path.open("r", encoding="utf-8") as file:
            data = yaml.safe_load(file) or {}
    except Exception as exc:
        raise ClassificationReviewWorkbookError(
            f"Plano de contas nao pode ser lido: {plano_path}"
        ) from exc

    contas = data.get("contas", [])
    if not isinstance(contas, list):
        raise ClassificationReviewWorkbookError(
            "Plano de contas nao possui lista valida em 'contas'."
        )

    rows: list[dict[str, str]] = []
    for item in contas:
        if not isinstance(item, dict):
            continue
        rows.append(
            {
                "conta_id": _text(item.get("id") or item.get("conta_id")),
                "conta_nome": _text(item.get("nome") or item.get("conta_nome")),
                "grupo": _text(item.get("grupo")),
                "observacao": _text(item.get("observacao")),
            }
        )
    return rows


def criar_aba_revisao(
    workbook: Workbook, headers: list[str], rows: list[dict[str, str]]
) -> None:
    sheet = workbook.active
    sheet.title = REVISAO_SHEET
    sheet.append(headers)

    for row in rows:
        sheet.append([row.get(header, "") for header in headers])

    _style_header(sheet)
    _apply_revisao_styles(sheet, headers)
    _set_filter_and_freeze(sheet)
    _apply_action_validation(sheet, headers, len(rows))


def criar_aba_plano_de_contas(
    workbook: Workbook, contas: Iterable[dict[str, str]]
) -> None:
    sheet = workbook.create_sheet(PLANO_CONTAS_SHEET)
    headers = ["conta_id", "conta_nome", "grupo", "observacao"]
    sheet.append(headers)
    for conta in contas:
        sheet.append([conta.get(header, "") for header in headers])

    _style_header(sheet)
    _set_filter_and_freeze(sheet)
    _set_widths(sheet, {"A": 16, "B": 42, "C": 36, "D": 50})


def criar_aba_acoes_validas(workbook: Workbook) -> None:
    sheet = workbook.create_sheet(ACOES_VALIDAS_SHEET)
    sheet.append(["acao_manual", "descricao"])
    for action, description in ALLOWED_ACTIONS_DESCRIPTIONS.items():
        sheet.append([action, description])

    _style_header(sheet)
    _set_filter_and_freeze(sheet)
    _set_widths(sheet, {"A": 28, "B": 90})
    for row in sheet.iter_rows(min_row=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")


def criar_aba_instrucoes(workbook: Workbook) -> None:
    sheet = workbook.create_sheet(INSTRUCOES_SHEET)
    instructions = [
        "Planilha assistida de revisao de classificacao",
        "",
        "Esta planilha foi gerada a partir do CSV de revisao de classificacao.",
        "",
        "Preencha principalmente as colunas:",
        "- acao_manual",
        "- conta_id_sugerida",
        "- conta_nome_sugerida",
        "- palavras_chave_sugeridas",
        "- observacao_manual",
        "",
        "Use a aba Plano_de_Contas para consultar os codigos de contas disponiveis.",
        "",
        (
            "Nao altere as colunas geradas automaticamente, como descricao, "
            "fornecedores, documentos, valores, conta atual, status e alertas."
        ),
        "",
        (
            "Use ANALISAR_COM_CONTADOR quando a decisao depender de validacao "
            "contabil/fiscal."
        ),
        (
            "Use ANALISAR_COM_CONTADOR para casos de comodato, bonificacao, "
            "doacao, brinde, equipamentos, materiais de obra, itens fiscais "
            "duvidosos ou qualquer caso sem seguranca."
        ),
        "Use MANTER_REVISAO quando ainda nao houver seguranca para classificar.",
        (
            "Use CLASSIFICAR quando o produto deve ser classificado diretamente "
            "em uma conta existente."
        ),
        (
            "Use CRIAR_ALIAS quando a conta atual esta correta e a intencao e "
            "ensinar uma variacao de nome."
        ),
        (
            "Use CRIAR_REGRA quando a intencao e criar uma regra geral de "
            "classificacao."
        ),
        "Use IGNORAR quando o candidato nao deve gerar regra ou classificacao.",
        "",
        "Salve o arquivo XLSX apos a revisao.",
        (
            "Uma fase futura podera importar esta planilha assistida e gerar "
            "um CSV validado."
        ),
        "Nesta fase, nenhuma regra YAML e modificada automaticamente.",
    ]
    for line in instructions:
        sheet.append([line])

    sheet["A1"].font = Font(bold=True, size=14)
    sheet.column_dimensions["A"].width = 110
    for row in sheet.iter_rows():
        row[0].alignment = Alignment(wrap_text=True, vertical="top")


def _style_header(sheet) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(bold=True, color="FFFFFF")
    for cell in sheet[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(wrap_text=True, vertical="center")


def _apply_revisao_styles(sheet, headers: list[str]) -> None:
    editable_fill = PatternFill("solid", fgColor="FFF2CC")
    generated_fill = PatternFill("solid", fgColor="D9EAF7")
    valor_total_index = _column_index(headers, "valor_total")

    for index, header in enumerate(headers, start=1):
        letter = get_column_letter(index)
        width = _suggested_width(header)
        sheet.column_dimensions[letter].width = width
        fill = editable_fill if header in EDITABLE_COLUMNS else generated_fill
        sheet.cell(row=1, column=index).fill = fill
        sheet.cell(row=1, column=index).font = Font(bold=True, color="000000")

        for cell in sheet.iter_cols(
            min_col=index,
            max_col=index,
            min_row=2,
            max_row=max(sheet.max_row, 2),
        ):
            for item in cell:
                item.alignment = Alignment(wrap_text=True, vertical="top")
                if header in EDITABLE_COLUMNS:
                    item.fill = editable_fill
                if index == valor_total_index:
                    item.alignment = Alignment(
                        horizontal="right", wrap_text=True, vertical="top"
                    )


def _apply_action_validation(sheet, headers: list[str], row_count: int) -> None:
    if "acao_manual" not in headers:
        return
    column_letter = get_column_letter(_column_index(headers, "acao_manual"))
    last_row = max(row_count + 1, 2)
    validation = DataValidation(
        type="list",
        formula1=f"={ACOES_VALIDAS_SHEET}!$A$2:$A${len(ALLOWED_ACTIONS_DESCRIPTIONS) + 1}",
        allow_blank=True,
    )
    validation.error = "Escolha uma acao manual valida."
    validation.errorTitle = "Acao manual invalida"
    validation.prompt = "Selecione uma acao manual valida."
    validation.promptTitle = "acao_manual"
    sheet.add_data_validation(validation)
    validation.add(f"{column_letter}2:{column_letter}{last_row}")


def _set_filter_and_freeze(sheet) -> None:
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions


def _set_widths(sheet, widths: dict[str, int]) -> None:
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width


def _suggested_width(header: str) -> int:
    widths = {
        "tipo_candidato": 24,
        "descricao_normalizada": 34,
        "descricao_exemplo": 42,
        "ocorrencias": 12,
        "fornecedores": 38,
        "documentos": 26,
        "valor_total": 16,
        "conta_atual_id": 16,
        "conta_atual_nome": 34,
        "status_revisao": 18,
        "alertas": 34,
        "acao_manual": 28,
        "conta_id_sugerida": 20,
        "conta_nome_sugerida": 34,
        "palavras_chave_sugeridas": 34,
        "observacao_manual": 44,
    }
    return widths.get(header, max(14, min(len(header) + 4, 40)))


def _column_index(headers: list[str], name: str) -> int:
    return headers.index(name) + 1


def _text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()
