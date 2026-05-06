from __future__ import annotations

import csv
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

from fantozzi.classification_candidates import (
    ClassificationCandidatesError,
    ClassificationCandidatesResult,
    analyze_classification_candidates,
)


CSV_HEADERS = [
    "tipo_candidato",
    "descricao_normalizada",
    "descricao_exemplo",
    "ocorrencias",
    "fornecedores",
    "documentos",
    "valor_total",
    "conta_atual_id",
    "conta_atual_nome",
    "status_revisao",
    "alertas",
    "acao_manual",
    "conta_id_sugerida",
    "conta_nome_sugerida",
    "palavras_chave_sugeridas",
    "observacao_manual",
]

MANUAL_FIELDS = {
    "acao_manual": "",
    "conta_id_sugerida": "",
    "conta_nome_sugerida": "",
    "palavras_chave_sugeridas": "",
    "observacao_manual": "",
}


class ClassificationReviewExportError(Exception):
    pass


@dataclass(frozen=True)
class ClassificationReviewExportResult:
    excel_path: Path
    csv_path: Path
    exported_candidates: int


def export_classification_review_csv(
    excel_path: Path | str,
    csv_path: Path | str,
) -> ClassificationReviewExportResult:
    excel_path = Path(excel_path)
    csv_path = Path(csv_path)

    try:
        candidates_result = analyze_classification_candidates(excel_path)
    except ClassificationCandidatesError as exc:
        raise ClassificationReviewExportError(str(exc)) from exc

    rows = _csv_rows(candidates_result)
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with csv_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=CSV_HEADERS, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)

    return ClassificationReviewExportResult(
        excel_path=excel_path,
        csv_path=csv_path,
        exported_candidates=len(rows),
    )


def console_lines(result: ClassificationReviewExportResult) -> list[str]:
    return [
        f"[OK] CSV de revisao de classificacao gerado: {result.csv_path}",
        f"     Candidatos exportados: {result.exported_candidates}",
        f"     Origem: {result.excel_path}",
    ]


def _csv_rows(result: ClassificationCandidatesResult) -> list[dict[str, object]]:
    rows = []
    for candidate in result.candidates:
        account_ids, account_names = _split_accounts(candidate.accounts)
        rows.append(
            {
                "tipo_candidato": "CLASSIFICACAO_PRODUTO",
                "descricao_normalizada": candidate.normalized_description,
                "descricao_exemplo": candidate.description,
                "ocorrencias": candidate.occurrences,
                "fornecedores": _join_values(candidate.suppliers),
                "documentos": _join_values(candidate.documents),
                "valor_total": _format_currency_br(candidate.total_value),
                "conta_atual_id": _join_values(account_ids),
                "conta_atual_nome": _join_values(account_names),
                "status_revisao": _join_values(candidate.statuses),
                "alertas": _join_values(candidate.alerts),
                **MANUAL_FIELDS,
            }
        )
    return rows


def _split_accounts(accounts: tuple[str, ...]) -> tuple[tuple[str, ...], tuple[str, ...]]:
    ids = []
    names = []
    for account in accounts:
        if " - " in account:
            account_id, account_name = account.split(" - ", 1)
            ids.append(account_id)
            names.append(account_name)
        elif account:
            ids.append(account)
    return tuple(ids), tuple(names)


def _join_values(values: tuple[str, ...]) -> str:
    return ", ".join(values)


def _format_currency_br(value: Decimal) -> str:
    text = f"{value.quantize(Decimal('0.01')):.2f}"
    integer_part, decimal_part = text.split(".")
    groups = []
    while integer_part:
        groups.append(integer_part[-3:])
        integer_part = integer_part[:-3]
    return "R$ " + ".".join(reversed(groups)) + "," + decimal_part
