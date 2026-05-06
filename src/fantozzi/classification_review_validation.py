from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from fantozzi.chart_of_accounts import ChartOfAccounts
from fantozzi.classification_review_export import CSV_HEADERS


DEFAULT_PLANO_CONTAS = (
    Path(__file__).resolve().parents[2] / "config" / "plano_contas_gerencial.yaml"
)

ALLOWED_ACTIONS = {
    "ANALISAR_COM_CONTADOR",
    "CLASSIFICAR",
    "CRIAR_ALIAS",
    "CRIAR_REGRA",
    "IGNORAR",
    "MANTER_REVISAO",
}


class ClassificationReviewValidationError(Exception):
    pass


@dataclass(frozen=True)
class ValidationProblem:
    csv_line: int
    description: str
    supplier: str
    problem: str
    value: str


@dataclass(frozen=True)
class ClassificationReviewValidationResult:
    csv_path: Path
    plano_contas_path: Path
    total_rows: int
    pending_rows: int
    valid_rows: int
    invalid_rows: int
    unknown_accounts: int
    unknown_actions: int
    status: str
    problems: tuple[ValidationProblem, ...]


def validate_classification_review_csv(
    csv_path: Path | str,
    plano_contas_path: Path | str | None = None,
) -> ClassificationReviewValidationResult:
    csv_path = Path(csv_path)
    plano_contas_path = Path(plano_contas_path or DEFAULT_PLANO_CONTAS)
    _validate_csv_file(csv_path)

    chart = _load_chart(plano_contas_path)
    rows = _read_csv_rows(csv_path)

    pending_rows = 0
    valid_rows = 0
    invalid_rows = 0
    unknown_accounts = 0
    unknown_actions = 0
    problems: list[ValidationProblem] = []

    for index, row in enumerate(rows, start=2):
        row_problems: list[ValidationProblem] = []
        row_pending = False
        row_invalid = False

        raw_account_id = _text(row.get("conta_id_sugerida"))
        normalized_account_id = _normalize_account_id(raw_account_id)
        if not raw_account_id:
            row_pending = True
            row_problems.append(
                _problem(index, row, "conta_id_sugerida pendente", raw_account_id)
            )
        else:
            if raw_account_id != normalized_account_id:
                row_problems.append(
                    _problem(
                        index,
                        row,
                        "conta_id_sugerida contem espacos ou formato suspeito",
                        raw_account_id,
                    )
                )
            try:
                chart.get_by_id(normalized_account_id)
            except KeyError:
                row_invalid = True
                unknown_accounts += 1
                row_problems.append(
                    _problem(
                        index,
                        row,
                        "conta_id_sugerida desconhecida",
                        raw_account_id,
                    )
                )

        raw_action = _text(row.get("acao_manual"))
        normalized_action = raw_action.strip().upper()
        if not raw_action:
            row_pending = True
            row_problems.append(_problem(index, row, "acao_manual pendente", raw_action))
        elif normalized_action not in ALLOWED_ACTIONS:
            row_invalid = True
            unknown_actions += 1
            row_problems.append(
                _problem(index, row, "acao_manual desconhecida", raw_action)
            )

        if row_invalid:
            invalid_rows += 1
        elif row_pending:
            pending_rows += 1
        else:
            valid_rows += 1

        problems.extend(row_problems)

    status = (
        "PRONTO PARA PROXIMA FASE"
        if pending_rows == 0 and invalid_rows == 0
        else "REVISAO MANUAL NECESSARIA"
    )

    return ClassificationReviewValidationResult(
        csv_path=csv_path,
        plano_contas_path=plano_contas_path,
        total_rows=len(rows),
        pending_rows=pending_rows,
        valid_rows=valid_rows,
        invalid_rows=invalid_rows,
        unknown_accounts=unknown_accounts,
        unknown_actions=unknown_actions,
        status=status,
        problems=tuple(problems),
    )


def report_text(
    result: ClassificationReviewValidationResult,
    *,
    max_details: int = 100,
) -> str:
    lines = [
        "FANTOZZI - VALIDACAO DA REVISAO MANUAL DE CLASSIFICACAO",
        "=======================================================",
        f"Arquivo CSV: {result.csv_path}",
        f"Plano de contas: {result.plano_contas_path}",
        "",
        "Resumo",
        "----------------------------------------",
        f"Total de linhas: {result.total_rows}",
        f"Linhas pendentes: {result.pending_rows}",
        f"Linhas validas: {result.valid_rows}",
        f"Linhas invalidas: {result.invalid_rows}",
        f"Contas sugeridas desconhecidas: {result.unknown_accounts}",
        f"Acoes manuais desconhecidas: {result.unknown_actions}",
        "",
        f"Status: {result.status}",
    ]

    if result.total_rows == 0:
        lines.extend(["", "Nenhuma linha encontrada no CSV de revisao."])

    if result.problems:
        lines.extend(["", "LINHAS COM PROBLEMAS", "----------------------------------------"])
        for problem in result.problems[:max_details]:
            lines.extend(
                [
                    f"Linha CSV: {problem.csv_line}",
                    f"Produto/descricao: {_or_dash(problem.description)}",
                    f"Fornecedor: {_or_dash(problem.supplier)}",
                    f"Problema: {problem.problem}",
                    f"Valor informado: {_or_dash(problem.value)}",
                    "",
                ]
            )
        if len(result.problems) > max_details:
            lines.append(
                f"Detalhes omitidos: {len(result.problems) - max_details} problema(s)."
            )

    lines.extend(
        [
            "",
            "Observacoes",
            "----------------------------------------",
            "- Validacao somente leitura.",
            "- Nenhum CSV foi alterado.",
            "- Nenhuma regra YAML foi alterada.",
            "- Nenhuma classificacao foi aplicada automaticamente.",
        ]
    )
    return "\n".join(lines) + "\n"


def save_report(report_path: Path | str, text: str) -> None:
    report_path = Path(report_path)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(text, encoding="utf-8", newline="\n")


def _validate_csv_file(csv_path: Path) -> None:
    if not csv_path.exists():
        raise ClassificationReviewValidationError(
            f"Arquivo CSV nao encontrado: {csv_path}"
        )
    if not csv_path.is_file():
        raise ClassificationReviewValidationError(
            f"Arquivo CSV nao e um arquivo: {csv_path}"
        )
    if csv_path.stat().st_size == 0:
        raise ClassificationReviewValidationError("CSV de revisao esta vazio.")


def _load_chart(plano_contas_path: Path) -> ChartOfAccounts:
    if not plano_contas_path.exists():
        raise ClassificationReviewValidationError(
            f"Plano de contas nao encontrado: {plano_contas_path}"
        )
    try:
        chart = ChartOfAccounts.load(plano_contas_path)
    except Exception as exc:
        raise ClassificationReviewValidationError(
            f"Plano de contas nao pode ser lido: {plano_contas_path}"
        ) from exc

    errors = chart.validate()
    if errors:
        raise ClassificationReviewValidationError(
            "Plano de contas invalido: " + "; ".join(errors)
        )
    return chart


def _read_csv_rows(csv_path: Path) -> list[dict[str, str]]:
    try:
        with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
            reader = csv.DictReader(file, delimiter=";")
            if not reader.fieldnames:
                raise ClassificationReviewValidationError(
                    "CSV de revisao nao possui cabecalho valido."
                )
            fieldnames = [field.strip() for field in reader.fieldnames if field]
            _validate_headers(fieldnames)
            return [dict(row) for row in reader]
    except UnicodeDecodeError as exc:
        raise ClassificationReviewValidationError(
            f"CSV de revisao nao pode ser lido em UTF-8: {csv_path}"
        ) from exc


def _validate_headers(fieldnames: list[str]) -> None:
    missing_headers = [header for header in CSV_HEADERS if header not in fieldnames]
    if missing_headers:
        raise ClassificationReviewValidationError(
            "Colunas obrigatorias ausentes: " + ", ".join(missing_headers)
        )


def _problem(
    csv_line: int, row: dict[str, str], problem: str, value: str
) -> ValidationProblem:
    return ValidationProblem(
        csv_line=csv_line,
        description=_text(row.get("descricao_exemplo"))
        or _text(row.get("descricao_normalizada")),
        supplier=_text(row.get("fornecedores")),
        problem=problem,
        value=value,
    )


def _normalize_account_id(value: str) -> str:
    return "".join(value.split())


def _text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _or_dash(value: str) -> str:
    return value if value else "-"
