from __future__ import annotations

import csv
from dataclasses import dataclass
from io import StringIO
from pathlib import Path
from typing import Mapping

from fantozzi.fase9_decision_engine import (
    Fase9BlockedDecision,
    Fase9DecisionReason,
    Fase9DecisionResult,
    Fase9PromotionCandidate,
)


STATUS_PROMOVIDO_OK = "PROMOVIDO_OK"
STATUS_BLOQUEADO_FASE9 = "BLOQUEADO_FASE9"
STATUS_JA_OK_OU_FORA_ESCOPO = "JA_OK_OU_FORA_ESCOPO"

AUDIT_HEADERS = (
    "fase9_status",
    "acao_fase9",
    "conta_id_aplicada",
    "conta_nome_aplicada",
    "motivo_promocao",
    "fonte_decisao",
    "fase9_aplicada",
    "motivo_bloqueio_fase9",
)

OUT_OF_SCOPE_REASONS = frozenset(
    {
        Fase9DecisionReason.JA_OK,
        Fase9DecisionReason.STATUS_NAO_REVISAR,
    }
)


@dataclass(frozen=True)
class Fase9WritebackResult:
    rows: tuple[dict[str, str], ...]
    headers: tuple[str, ...]
    total_rows: int
    promoted_rows: int
    blocked_rows: int
    out_of_scope_rows: int
    output_path: Path | None = None


def apply_fase9_writeback(
    rows: list[Mapping[str, object]] | tuple[Mapping[str, object], ...],
    decision_result: Fase9DecisionResult,
    *,
    headers: tuple[str, ...] | list[str] | None = None,
) -> Fase9WritebackResult:
    """Gera linhas novas com aplicacao Fase 9, sem modificar as linhas originais."""
    original_headers = tuple(headers or _headers_from_rows(rows))
    output_headers = _output_headers(original_headers)
    promotions = {
        promotion.decision.line_number: promotion
        for promotion in decision_result.promotions
    }
    blocked = {
        blocked_decision.decision.line_number: blocked_decision
        for blocked_decision in (
            *decision_result.blocked,
            *decision_result.ignored,
        )
    }

    output_rows: list[dict[str, str]] = []
    promoted_count = 0
    blocked_count = 0
    out_of_scope_count = 0

    for index, row in enumerate(rows, start=2):
        line_number = _row_line_number(row, index)
        output_row = _string_row(row, output_headers)
        promotion = promotions.get(line_number)
        blocked_decision = blocked.get(line_number)

        if promotion is not None:
            _apply_promotion(output_row, promotion)
            promoted_count += 1
        elif blocked_decision is not None:
            if blocked_decision.reason in OUT_OF_SCOPE_REASONS:
                _apply_out_of_scope(output_row, blocked_decision)
                out_of_scope_count += 1
            else:
                _apply_block(output_row, blocked_decision)
                blocked_count += 1
        else:
            _apply_no_decision(output_row)
            out_of_scope_count += 1

        output_rows.append(output_row)

    return Fase9WritebackResult(
        rows=tuple(output_rows),
        headers=output_headers,
        total_rows=len(output_rows),
        promoted_rows=promoted_count,
        blocked_rows=blocked_count,
        out_of_scope_rows=out_of_scope_count,
    )


def render_fase9_writeback_csv(
    rows: list[Mapping[str, object]] | tuple[Mapping[str, object], ...],
    decision_result: Fase9DecisionResult,
    *,
    headers: tuple[str, ...] | list[str] | None = None,
) -> str:
    result = apply_fase9_writeback(rows, decision_result, headers=headers)
    output = StringIO()
    writer = csv.DictWriter(
        output,
        fieldnames=result.headers,
        delimiter=";",
        lineterminator="\n",
    )
    writer.writeheader()
    writer.writerows(result.rows)
    return output.getvalue()


def write_fase9_applied_csv(
    input_csv_path: Path | str,
    output_csv_path: Path | str,
    decision_result: Fase9DecisionResult,
) -> Fase9WritebackResult:
    """Le um CSV e grava um novo CSV em path explicito, preservando o original."""
    input_csv_path = Path(input_csv_path)
    output_csv_path = Path(output_csv_path)
    if input_csv_path.resolve() == output_csv_path.resolve():
        raise ValueError("output_csv_path deve ser diferente do CSV original.")

    headers, rows, dialect = _read_csv(input_csv_path)
    result = apply_fase9_writeback(rows, decision_result, headers=headers)
    output_csv_path.parent.mkdir(parents=True, exist_ok=True)
    with output_csv_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(
            file,
            fieldnames=result.headers,
            delimiter=getattr(dialect, "delimiter", ";") or ";",
            lineterminator="\n",
        )
        writer.writeheader()
        writer.writerows(result.rows)

    return Fase9WritebackResult(
        rows=result.rows,
        headers=result.headers,
        total_rows=result.total_rows,
        promoted_rows=result.promoted_rows,
        blocked_rows=result.blocked_rows,
        out_of_scope_rows=result.out_of_scope_rows,
        output_path=output_csv_path,
    )


def _apply_promotion(
    row: dict[str, str],
    promotion: Fase9PromotionCandidate,
) -> None:
    decision = promotion.decision
    row["status_revisao"] = "OK"
    _set_if_present(row, "conta_id", promotion.account_id)
    _set_if_present(row, "conta_nome", promotion.account_name)
    _set_if_present(row, "conta_atual_id", promotion.account_id)
    _set_if_present(row, "conta_atual_nome", promotion.account_name)
    row["fase9_status"] = STATUS_PROMOVIDO_OK
    row["acao_fase9"] = promotion.action
    row["conta_id_aplicada"] = promotion.account_id
    row["conta_nome_aplicada"] = promotion.account_name
    row["motivo_promocao"] = promotion.reason.value
    row["fonte_decisao"] = _decision_source(decision.line_number)
    row["fase9_aplicada"] = "true"
    row["motivo_bloqueio_fase9"] = ""


def _apply_block(
    row: dict[str, str],
    blocked_decision: Fase9BlockedDecision,
) -> None:
    row["fase9_status"] = STATUS_BLOQUEADO_FASE9
    row["acao_fase9"] = blocked_decision.action
    row["conta_id_aplicada"] = ""
    row["conta_nome_aplicada"] = ""
    row["motivo_promocao"] = ""
    row["fonte_decisao"] = _decision_source(blocked_decision.decision.line_number)
    row["fase9_aplicada"] = "false"
    row["motivo_bloqueio_fase9"] = blocked_decision.reason.value


def _apply_out_of_scope(
    row: dict[str, str],
    blocked_decision: Fase9BlockedDecision,
) -> None:
    row["fase9_status"] = STATUS_JA_OK_OU_FORA_ESCOPO
    row["acao_fase9"] = blocked_decision.action
    row["conta_id_aplicada"] = ""
    row["conta_nome_aplicada"] = ""
    row["motivo_promocao"] = ""
    row["fonte_decisao"] = _decision_source(blocked_decision.decision.line_number)
    row["fase9_aplicada"] = "false"
    row["motivo_bloqueio_fase9"] = blocked_decision.reason.value


def _apply_no_decision(row: dict[str, str]) -> None:
    row["fase9_status"] = STATUS_JA_OK_OU_FORA_ESCOPO
    row["acao_fase9"] = ""
    row["conta_id_aplicada"] = ""
    row["conta_nome_aplicada"] = ""
    row["motivo_promocao"] = ""
    row["fonte_decisao"] = ""
    row["fase9_aplicada"] = "false"
    row["motivo_bloqueio_fase9"] = "SEM_DECISAO_FASE9"


def _read_csv(csv_path: Path) -> tuple[tuple[str, ...], tuple[dict[str, str], ...], csv.Dialect]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
        sample = file.read(4096)
        file.seek(0)
        dialect = csv.Sniffer().sniff(sample, delimiters=";,") if sample else csv.excel
        reader = csv.DictReader(file, dialect=dialect)
        headers = tuple(reader.fieldnames or ())
        rows = tuple({key: _text(value) for key, value in row.items()} for row in reader)
    return headers, rows, dialect


def _headers_from_rows(
    rows: list[Mapping[str, object]] | tuple[Mapping[str, object], ...],
) -> tuple[str, ...]:
    headers: list[str] = []
    for row in rows:
        for key in row.keys():
            key_text = _text(key)
            if key_text and key_text not in headers and key_text != "_fase9_line_number":
                headers.append(key_text)
    return tuple(headers)


def _output_headers(headers: tuple[str, ...]) -> tuple[str, ...]:
    output = list(headers)
    for header in AUDIT_HEADERS:
        if header not in output:
            output.append(header)
    return tuple(output)


def _string_row(row: Mapping[str, object], headers: tuple[str, ...]) -> dict[str, str]:
    return {header: _text(row.get(header)) for header in headers}


def _row_line_number(row: Mapping[str, object], fallback: int) -> int:
    value = _text(row.get("_fase9_line_number"))
    if not value:
        return fallback
    try:
        return int(value)
    except ValueError:
        return fallback


def _set_if_present(row: dict[str, str], key: str, value: str) -> None:
    if key in row:
        row[key] = value


def _decision_source(line_number: int) -> str:
    return f"csv_line:{line_number}"


def _text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()
