from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from fantozzi.classification_review_export import CSV_HEADERS
from fantozzi.classification_review_validation import ALLOWED_ACTIONS
from fantozzi.classification_review_workbook import (
    ClassificationReviewWorkbookError,
    DEFAULT_PLANO_CONTAS,
    REVISAO_SHEET,
    carregar_plano_contas,
)


ACOES_COM_CONTA_OBRIGATORIA = {"CLASSIFICAR", "CRIAR_REGRA"}


class ClassificationReviewWorkbookImportError(Exception):
    def __init__(
        self,
        message: str,
        result: ResultadoImportacaoPlanilhaRevisao | None = None,
    ) -> None:
        super().__init__(message)
        self.result = result


@dataclass(frozen=True)
class ResultadoImportacaoPlanilhaRevisao:
    xlsx_path: Path
    csv_path: Path | None
    relatorio_path: Path | None
    linhas_lidas: int
    linhas_exportadas: int
    erros: tuple[str, ...]
    avisos: tuple[str, ...]
    contagem_acoes: dict[str, int]

    @property
    def is_ok(self) -> bool:
        return not self.erros


def importar_planilha_revisao_classificacao(
    xlsx_path: Path | str,
    output_csv_path: Path | str,
    plano_contas_path: Path | str | None = None,
    relatorio_path: Path | str | None = None,
) -> ResultadoImportacaoPlanilhaRevisao:
    xlsx_path = Path(xlsx_path)
    output_csv_path = Path(output_csv_path)
    plano_contas_path = Path(plano_contas_path or DEFAULT_PLANO_CONTAS)
    relatorio_path = Path(relatorio_path) if relatorio_path else None

    linhas = ler_aba_revisao(xlsx_path)
    contas_validas = carregar_contas_validas(plano_contas_path)
    rows, erros, avisos, contagem_acoes = validar_linhas_revisao(
        linhas, contas_validas
    )

    result = ResultadoImportacaoPlanilhaRevisao(
        xlsx_path=xlsx_path,
        csv_path=None if erros else output_csv_path,
        relatorio_path=relatorio_path,
        linhas_lidas=len(linhas),
        linhas_exportadas=0 if erros else len(rows),
        erros=tuple(erros),
        avisos=tuple(avisos),
        contagem_acoes=contagem_acoes,
    )

    if relatorio_path:
        escrever_relatorio_validacao(result, relatorio_path)

    if erros:
        raise ClassificationReviewWorkbookImportError(erros[0], result)

    escrever_csv_revisado(rows, output_csv_path)
    return result


def ler_aba_revisao(xlsx_path: Path | str) -> list[tuple[int, dict[str, str]]]:
    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        raise ClassificationReviewWorkbookImportError(
            f"Planilha assistida nao encontrada: {xlsx_path}"
        )
    if not xlsx_path.is_file():
        raise ClassificationReviewWorkbookImportError(
            f"Planilha assistida nao e um arquivo: {xlsx_path}"
        )

    try:
        workbook = load_workbook(xlsx_path, data_only=True)
    except Exception as exc:
        raise ClassificationReviewWorkbookImportError(
            f"Planilha assistida nao pode ser lida: {xlsx_path}"
        ) from exc

    if REVISAO_SHEET not in workbook.sheetnames:
        raise ClassificationReviewWorkbookImportError(
            f"Aba obrigatoria nao encontrada: {REVISAO_SHEET}"
        )

    sheet = workbook[REVISAO_SHEET]
    raw_headers = [_text(cell.value) for cell in sheet[1]]
    headers = [header.strip() for header in raw_headers]
    _validar_headers_revisao(headers)

    rows: list[tuple[int, dict[str, str]]] = []
    for row_index, values in enumerate(
        sheet.iter_rows(min_row=2, values_only=True), start=2
    ):
        row_by_header = {
            header: _cell_to_text(values[headers.index(header)])
            for header in CSV_HEADERS
        }
        if _linha_totalmente_vazia(row_by_header):
            continue
        rows.append((row_index, row_by_header))
    return rows


def validar_linhas_revisao(
    linhas: list[tuple[int, dict[str, str]]],
    contas_validas: set[str],
) -> tuple[list[dict[str, str]], list[str], list[str], dict[str, int]]:
    rows: list[dict[str, str]] = []
    erros: list[str] = []
    avisos: list[str] = []
    contagem_acoes = {acao: 0 for acao in sorted(ALLOWED_ACTIONS)}

    for linha_xlsx, row in linhas:
        normalized_row = {header: _text(row.get(header)) for header in CSV_HEADERS}
        acao_original = normalized_row["acao_manual"]
        acao = acao_original.strip().upper()
        normalized_row["acao_manual"] = acao

        if not acao:
            erros.append(f"Linha {linha_xlsx}: acao_manual vazia.")
        elif acao not in ALLOWED_ACTIONS:
            erros.append(f"Linha {linha_xlsx}: acao_manual invalida: {acao_original}")
        else:
            contagem_acoes[acao] += 1

        conta_id = _normalize_account_id(normalized_row["conta_id_sugerida"])
        normalized_row["conta_id_sugerida"] = conta_id

        if acao in ACOES_COM_CONTA_OBRIGATORIA and not conta_id:
            erros.append(
                f"Linha {linha_xlsx}: conta_id_sugerida obrigatoria "
                f"para acao_manual {acao}."
            )

        if conta_id and conta_id not in contas_validas:
            erros.append(
                f"Linha {linha_xlsx}: conta_id_sugerida nao existe "
                f"no plano de contas: {conta_id}"
            )

        conta_atual_id = _normalize_account_id(normalized_row["conta_atual_id"])
        normalized_row["conta_atual_id"] = conta_atual_id
        if acao == "CRIAR_ALIAS" and not conta_id and not conta_atual_id:
            avisos.append(
                f"Linha {linha_xlsx}: CRIAR_ALIAS sem conta_id_sugerida "
                "e sem conta_atual_id."
            )

        rows.append(normalized_row)

    return rows, erros, avisos, contagem_acoes


def escrever_csv_revisado(rows: list[dict[str, str]], output_path: Path | str) -> None:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=CSV_HEADERS, delimiter=";")
        writer.writeheader()
        writer.writerows(
            [{header: row.get(header, "") for header in CSV_HEADERS} for row in rows]
        )


def carregar_contas_validas(plano_contas_path: Path | str) -> set[str]:
    try:
        contas = carregar_plano_contas(plano_contas_path)
    except ClassificationReviewWorkbookError as exc:
        raise ClassificationReviewWorkbookImportError(str(exc)) from exc
    return {conta["conta_id"] for conta in contas if conta.get("conta_id")}


def escrever_relatorio_validacao(
    result: ResultadoImportacaoPlanilhaRevisao, relatorio_path: Path | str
) -> None:
    relatorio_path = Path(relatorio_path)
    relatorio_path.parent.mkdir(parents=True, exist_ok=True)
    relatorio_path.write_text(relatorio_validacao_text(result), encoding="utf-8")


def relatorio_validacao_text(result: ResultadoImportacaoPlanilhaRevisao) -> str:
    csv_path = result.csv_path if result.csv_path else "-"
    lines = [
        "VALIDACAO DA PLANILHA ASSISTIDA DE REVISAO DE CLASSIFICACAO",
        "============================================================",
        "",
        f"Arquivo XLSX: {result.xlsx_path}",
        f"Arquivo CSV gerado: {csv_path}",
        "",
        f"Linhas lidas: {result.linhas_lidas}",
        f"Linhas exportadas: {result.linhas_exportadas}",
        "",
        "Acoes:",
    ]
    for action in sorted(ALLOWED_ACTIONS):
        lines.append(f"- {action}: {result.contagem_acoes.get(action, 0)}")

    lines.extend(
        [
            "",
            f"Erros: {len(result.erros)}",
            f"Avisos: {len(result.avisos)}",
            "",
            "ERROS",
            "-----",
        ]
    )
    lines.extend(result.erros or ["-"])
    lines.extend(["", "AVISOS", "------"])
    lines.extend(result.avisos or ["-"])
    lines.extend(["", f"Resultado: {'OK' if result.is_ok else 'ERRO'}"])
    return "\n".join(lines) + "\n"


def console_lines(result: ResultadoImportacaoPlanilhaRevisao) -> list[str]:
    if result.is_ok:
        return [
            "[OK] Planilha importada com sucesso.",
            f"[OK] CSV revisado gerado: {result.csv_path}",
            f"Linhas exportadas: {result.linhas_exportadas}",
            f"Erros: {len(result.erros)}",
            f"Avisos: {len(result.avisos)}",
        ]
    return [
        "[ERRO] Planilha assistida contem erros bloqueantes.",
        f"Linhas lidas: {result.linhas_lidas}",
        f"Erros: {len(result.erros)}",
        f"Avisos: {len(result.avisos)}",
        *[f"[ERRO] {erro}" for erro in result.erros],
    ]


def _validar_headers_revisao(headers: list[str]) -> None:
    for required_header in CSV_HEADERS:
        if required_header not in headers:
            raise ClassificationReviewWorkbookImportError(
                f"Coluna obrigatoria ausente na aba Revisao: {required_header}"
            )


def _linha_totalmente_vazia(row: dict[str, str]) -> bool:
    return all(not _text(value) for value in row.values())


def _normalize_account_id(value: str) -> str:
    return "".join(value.split())


def _cell_to_text(value: object) -> str:
    return _text(value)


def _text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()
