from __future__ import annotations

"""Historical bank/payment matching helpers.

This module is intentionally not exposed by the operational CLI. Fase 13 now
classifies payment receipts only through declared rules in
config/regras_classificacao_comprovantes.yaml.
"""

import csv
import json
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path

from fantozzi.banking.common import normalize_key, only_digits, parse_date, parse_decimal
from fantozzi.csv_utils import sniff_delimiter


MATCH_HEADERS = [
    "cnpj",
    "periodo",
    "bank_transaction_id",
    "receipt_id",
    "match_status",
    "match_confidence",
    "match_rule",
    "bank_date",
    "receipt_date",
    "bank_amount",
    "receipt_amount",
    "amount_delta",
    "description_bank",
    "description_receipt",
    "notes",
]

AUDIT_HEADERS = [
    "tipo",
    "entity_id",
    "cnpj",
    "periodo",
    "status",
    "rule",
    "confidence",
    "details",
]

DEFAULT_TRANSACTION_FILE = "bank_transactions_normalized.csv"
DEFAULT_RECEIPT_FILE = "payment_receipts_normalized.csv"


@dataclass(frozen=True)
class MatchingManifest:
    cnpj: str
    periodo: str
    input_files: dict[str, str]
    output_dir: Path
    transaction_count: int
    outbound_transaction_count: int
    out_of_scope_entry_count: int
    receipt_count: int
    exact_match_count: int
    probable_match_count: int
    ambiguous_count: int
    unmatched_transaction_count: int
    unmatched_receipt_count: int
    status: str
    window_days: int
    generated_files: dict[str, str]


@dataclass(frozen=True)
class _Transaction:
    transaction_id: str
    cnpj: str
    transaction_date: date
    amount: Decimal
    description_bank: str
    technical_category: str
    source_row: dict[str, str]


@dataclass(frozen=True)
class _Receipt:
    receipt_id: str
    cnpj: str
    payment_date: date | None
    amount: Decimal
    description_receipt: str
    beneficiary_name: str
    beneficiary_document: str
    reference_id: str
    source_row: dict[str, str]


@dataclass(frozen=True)
class _Candidate:
    receipt: _Receipt
    status: str
    confidence: Decimal
    rule: str
    notes: str


@dataclass(frozen=True)
class _Decision:
    transaction: _Transaction
    receipts: tuple[_Receipt, ...]
    status: str
    confidence: Decimal
    rule: str
    notes: str


def match_bank_files(
    *,
    input_dir: Path,
    output_dir: Path,
    periodo: str | None = None,
    window_days: int = 3,
) -> MatchingManifest:
    if window_days < 0:
        raise ValueError("window_days deve ser maior ou igual a zero.")

    transaction_path = input_dir / DEFAULT_TRANSACTION_FILE
    receipt_path = input_dir / DEFAULT_RECEIPT_FILE
    _ensure_file(transaction_path, "CSV de transacoes bancarias normalizadas")
    _ensure_file(receipt_path, "CSV de comprovantes normalizados")

    transaction_headers, transaction_rows = _read_csv(transaction_path)
    receipt_headers, receipt_rows = _read_csv(receipt_path)
    transactions = [_transaction_from_row(row) for row in transaction_rows]
    receipts = [_receipt_from_row(row) for row in receipt_rows]
    effective_period = periodo or _infer_period(transactions, receipts)
    effective_cnpj = _infer_cnpj(transactions, receipts)
    outbound_transactions = [
        transaction
        for transaction in transactions
        if _is_matching_scope_transaction(transaction)
    ]
    out_of_scope_entries = [
        transaction
        for transaction in transactions
        if transaction.technical_category == "ENTRADA_FORA_ESCOPO"
    ]

    decisions = _decide_matches(
        outbound_transactions,
        receipts,
        periodo=effective_period,
        window_days=window_days,
    )
    assigned_receipt_ids = {
        decision.receipts[0].receipt_id
        for decision in decisions
        if decision.status in {"MATCH_EXATO", "MATCH_PROVAVEL"} and decision.receipts
    }
    unmatched_transactions = [
        decision.transaction.source_row
        for decision in decisions
        if decision.status in {"MATCH_AMBIGUO", "NAO_ENCONTRADO"}
    ]
    unmatched_receipts = [
        receipt.source_row
        for receipt in receipts
        if receipt.receipt_id not in assigned_receipt_ids
    ]

    output_dir.mkdir(parents=True, exist_ok=True)
    matches_path = output_dir / "bank_matches.csv"
    unmatched_transactions_path = output_dir / "bank_unmatched_transactions.csv"
    unmatched_receipts_path = output_dir / "bank_unmatched_receipts.csv"
    out_of_scope_entries_path = output_dir / "bank_entradas_fora_escopo.csv"
    audit_path = output_dir / "bank_matching_audit.csv"
    manifest_path = output_dir / "manifest_bank_matching.json"

    _write_csv(matches_path, MATCH_HEADERS, _match_rows(decisions, effective_period))
    _write_csv(
        unmatched_transactions_path,
        transaction_headers,
        unmatched_transactions,
    )
    _write_csv(unmatched_receipts_path, receipt_headers, unmatched_receipts)
    _write_csv(
        out_of_scope_entries_path,
        [*transaction_headers, "technical_category", "exclusion_reason"],
        _out_of_scope_entry_rows(out_of_scope_entries),
    )
    _write_csv(
        audit_path,
        AUDIT_HEADERS,
        _audit_rows(decisions, receipts, out_of_scope_entries, effective_period),
    )

    exact_count = sum(1 for decision in decisions if decision.status == "MATCH_EXATO")
    probable_count = sum(
        1 for decision in decisions if decision.status == "MATCH_PROVAVEL"
    )
    ambiguous_count = sum(
        1 for decision in decisions if decision.status == "MATCH_AMBIGUO"
    )
    unmatched_count = sum(
        1 for decision in decisions if decision.status == "NAO_ENCONTRADO"
    )
    status = "OK" if not ambiguous_count and not unmatched_count else "ATENCAO"

    generated_files = {
        "bank_matches": str(matches_path),
        "bank_unmatched_transactions": str(unmatched_transactions_path),
        "bank_unmatched_receipts": str(unmatched_receipts_path),
        "bank_entradas_fora_escopo": str(out_of_scope_entries_path),
        "bank_matching_audit": str(audit_path),
        "manifest": str(manifest_path),
    }
    manifest = MatchingManifest(
        cnpj=effective_cnpj,
        periodo=effective_period,
        input_files={
            "bank_transactions_normalized": str(transaction_path),
            "payment_receipts_normalized": str(receipt_path),
        },
        output_dir=output_dir,
        transaction_count=len(transactions),
        outbound_transaction_count=len(outbound_transactions),
        out_of_scope_entry_count=len(out_of_scope_entries),
        receipt_count=len(receipts),
        exact_match_count=exact_count,
        probable_match_count=probable_count,
        ambiguous_count=ambiguous_count,
        unmatched_transaction_count=unmatched_count + ambiguous_count,
        unmatched_receipt_count=len(unmatched_receipts),
        status=status,
        window_days=window_days,
        generated_files=generated_files,
    )
    _write_manifest(manifest)
    return manifest


def _decide_matches(
    transactions: list[_Transaction],
    receipts: list[_Receipt],
    *,
    periodo: str,
    window_days: int,
) -> list[_Decision]:
    proposals: list[_Decision] = []
    for transaction in transactions:
        candidates = _candidates_for(transaction, receipts, window_days)
        if not candidates:
            proposals.append(
                _Decision(
                    transaction=transaction,
                    receipts=(),
                    status="NAO_ENCONTRADO",
                    confidence=Decimal("0.00"),
                    rule="SEM_CANDIDATO",
                    notes="Nenhum comprovante com mesmo CNPJ e valor absoluto.",
                )
            )
            continue

        best_confidence = max(candidate.confidence for candidate in candidates)
        best_candidates = [
            candidate for candidate in candidates if candidate.confidence == best_confidence
        ]
        if len(best_candidates) > 1:
            proposals.append(
                _Decision(
                    transaction=transaction,
                    receipts=tuple(candidate.receipt for candidate in best_candidates),
                    status="MATCH_AMBIGUO",
                    confidence=best_confidence,
                    rule=best_candidates[0].rule,
                    notes=(
                        "Mais de um comprovante candidato para a mesma transacao "
                        f"no periodo {periodo}."
                    ),
                )
            )
            continue

        candidate = best_candidates[0]
        proposals.append(
            _Decision(
                transaction=transaction,
                receipts=(candidate.receipt,),
                status=candidate.status,
                confidence=candidate.confidence,
                rule=candidate.rule,
                notes=candidate.notes,
            )
        )

    unique_receipt_targets: dict[str, list[int]] = {}
    for index, proposal in enumerate(proposals):
        if proposal.status not in {"MATCH_EXATO", "MATCH_PROVAVEL"} or not proposal.receipts:
            continue
        unique_receipt_targets.setdefault(proposal.receipts[0].receipt_id, []).append(index)

    conflicted_indexes = {
        index
        for indexes in unique_receipt_targets.values()
        if len(indexes) > 1
        for index in indexes
    }
    if not conflicted_indexes:
        return proposals

    decisions = list(proposals)
    for index in conflicted_indexes:
        proposal = proposals[index]
        decisions[index] = _Decision(
            transaction=proposal.transaction,
            receipts=proposal.receipts,
            status="MATCH_AMBIGUO",
            confidence=proposal.confidence,
            rule="CONFLITO_COMPROVANTE",
            notes="Mesmo comprovante candidato para mais de uma transacao bancaria.",
        )
    return decisions


def _candidates_for(
    transaction: _Transaction,
    receipts: list[_Receipt],
    window_days: int,
) -> list[_Candidate]:
    candidates: list[_Candidate] = []
    for receipt in receipts:
        if transaction.cnpj and receipt.cnpj and transaction.cnpj != receipt.cnpj:
            continue
        if _money(transaction.amount.copy_abs()) != _money(receipt.amount):
            continue

        delta = _date_delta(transaction.transaction_date, receipt.payment_date)
        if delta == 0:
            candidates.append(
                _Candidate(
                    receipt=receipt,
                    status="MATCH_EXATO",
                    confidence=Decimal("1.00"),
                    rule="VALOR_DATA_EXATA",
                    notes="Valor absoluto e data iguais.",
                )
            )
        elif delta is not None and delta <= window_days:
            confidence = Decimal("0.90") - (Decimal(delta) * Decimal("0.01"))
            candidates.append(
                _Candidate(
                    receipt=receipt,
                    status="MATCH_PROVAVEL",
                    confidence=confidence,
                    rule="VALOR_JANELA_DATA",
                    notes=f"Valor absoluto igual e data dentro de {window_days} dias.",
                )
            )
        elif _has_strong_text_identifier(transaction, receipt):
            candidates.append(
                _Candidate(
                    receipt=receipt,
                    status="MATCH_PROVAVEL",
                    confidence=Decimal("0.80"),
                    rule="IDENTIFICADOR_TEXTUAL_FORTE",
                    notes="Valor absoluto igual e identificador textual forte encontrado.",
                )
            )
    return candidates


def _transaction_from_row(row: dict[str, str]) -> _Transaction:
    transaction_date = parse_date(row.get("transaction_date"))
    if transaction_date is None:
        raise ValueError("Transacao bancaria sem transaction_date.")
    description_bank = " | ".join(
        part
        for part in (
            row.get("description_enriched", ""),
            row.get("description_raw", ""),
            row.get("counterparty_name", ""),
            row.get("counterparty_document", ""),
        )
        if part
    )
    return _Transaction(
        transaction_id=row.get("transaction_id", ""),
        cnpj=only_digits(row.get("cnpj", "")),
        transaction_date=transaction_date,
        amount=parse_decimal(row.get("amount")),
        description_bank=description_bank,
        technical_category=_classify_bank_transaction(row, description_bank),
        source_row=row,
    )


def _receipt_from_row(row: dict[str, str]) -> _Receipt:
    description_receipt = " | ".join(
        part
        for part in (
            row.get("beneficiary_name", ""),
            row.get("beneficiary_document", ""),
            row.get("payment_type", ""),
            row.get("reference_id", ""),
            row.get("status", ""),
        )
        if part
    )
    return _Receipt(
        receipt_id=row.get("payment_id", ""),
        cnpj=only_digits(row.get("cnpj", "")),
        payment_date=parse_date(row.get("payment_date")),
        amount=parse_decimal(row.get("amount")),
        description_receipt=description_receipt,
        beneficiary_name=row.get("beneficiary_name", ""),
        beneficiary_document=row.get("beneficiary_document", ""),
        reference_id=row.get("reference_id", ""),
        source_row=row,
    )


def _match_rows(decisions: list[_Decision], periodo: str) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for decision in decisions:
        transaction = decision.transaction
        receipts = decision.receipts
        receipt = receipts[0] if len(receipts) == 1 else None
        receipt_amount = receipt.amount if receipt else None
        rows.append(
            {
                "cnpj": transaction.cnpj,
                "periodo": periodo,
                "bank_transaction_id": transaction.transaction_id,
                "receipt_id": "|".join(item.receipt_id for item in receipts),
                "match_status": decision.status,
                "match_confidence": _decimal_text(decision.confidence),
                "match_rule": _rule_with_category(
                    decision.rule,
                    transaction.technical_category,
                ),
                "bank_date": transaction.transaction_date.isoformat(),
                "receipt_date": (
                    receipt.payment_date.isoformat()
                    if receipt and receipt.payment_date
                    else ""
                ),
                "bank_amount": _decimal_text(transaction.amount),
                "receipt_amount": _decimal_text(receipt_amount) if receipt_amount else "",
                "amount_delta": (
                    _decimal_text(_money(transaction.amount.copy_abs()) - _money(receipt_amount))
                    if receipt_amount is not None
                    else ""
                ),
                "description_bank": transaction.description_bank,
                "description_receipt": (
                    receipt.description_receipt
                    if receipt
                    else " | ".join(item.description_receipt for item in receipts)
                ),
                "notes": decision.notes,
            }
        )
    return rows


def _audit_rows(
    decisions: list[_Decision],
    receipts: list[_Receipt],
    out_of_scope_entries: list[_Transaction],
    periodo: str,
) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    assigned_receipt_ids = {
        decision.receipts[0].receipt_id
        for decision in decisions
        if decision.status in {"MATCH_EXATO", "MATCH_PROVAVEL"} and decision.receipts
    }
    for decision in decisions:
        rows.append(
            {
                "tipo": "BANK_TRANSACTION",
                "entity_id": decision.transaction.transaction_id,
                "cnpj": decision.transaction.cnpj,
                "periodo": periodo,
                "status": decision.status,
                "rule": _rule_with_category(
                    decision.rule,
                    decision.transaction.technical_category,
                ),
                "confidence": _decimal_text(decision.confidence),
                "details": decision.notes,
            }
        )
    for transaction in out_of_scope_entries:
        rows.append(
            {
                "tipo": "BANK_TRANSACTION",
                "entity_id": transaction.transaction_id,
                "cnpj": transaction.cnpj,
                "periodo": periodo,
                "status": "FORA_ESCOPO",
                "rule": "FORA_ESCOPO_ENTRADA",
                "confidence": "0.00",
                "details": (
                    "Entrada bancaria excluida do matching com comprovantes. "
                    f"Categoria tecnica: {transaction.technical_category}."
                ),
            }
        )
    for receipt in receipts:
        if receipt.receipt_id in assigned_receipt_ids:
            continue
        rows.append(
            {
                "tipo": "PAYMENT_RECEIPT",
                "entity_id": receipt.receipt_id,
                "cnpj": receipt.cnpj,
                "periodo": periodo,
                "status": "NAO_ENCONTRADO",
                "rule": "COMPROVANTE_SEM_TRANSACAO",
                "confidence": "0.00",
                "details": "Comprovante nao associado a transacao bancaria unica.",
            }
        )
    return rows


def _read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        sample = file.read(4096)
        file.seek(0)
        reader = csv.DictReader(file, delimiter=sniff_delimiter(sample))
        headers = list(reader.fieldnames or [])
        rows = [
            {header: (row.get(header) or "").strip() for header in headers}
            for row in reader
        ]
    return headers, rows


def _write_csv(path: Path, headers: list[str], rows: list[dict[str, object]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=headers, delimiter=";")
        writer.writeheader()
        for row in rows:
            writer.writerow({header: _csv_value(row.get(header, "")) for header in headers})


def _write_manifest(manifest: MatchingManifest) -> Path:
    path = manifest.output_dir / "manifest_bank_matching.json"
    payload = {
        "cnpj": manifest.cnpj,
        "periodo": manifest.periodo,
        "status": manifest.status,
        "input_files": manifest.input_files,
        "transaction_count": manifest.transaction_count,
        "outbound_transaction_count": manifest.outbound_transaction_count,
        "out_of_scope_entry_count": manifest.out_of_scope_entry_count,
        "receipt_count": manifest.receipt_count,
        "exact_match_count": manifest.exact_match_count,
        "probable_match_count": manifest.probable_match_count,
        "ambiguous_count": manifest.ambiguous_count,
        "unmatched_transaction_count": manifest.unmatched_transaction_count,
        "unmatched_receipt_count": manifest.unmatched_receipt_count,
        "window_days": manifest.window_days,
        "generated_files": manifest.generated_files,
        "matching_executed": True,
        "classification_executed": False,
    }
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return path


def _infer_period(transactions: list[_Transaction], receipts: list[_Receipt]) -> str:
    dates = [transaction.transaction_date for transaction in transactions]
    dates.extend(receipt.payment_date for receipt in receipts if receipt.payment_date)
    if not dates:
        return ""
    months = {item.strftime("%Y-%m") for item in dates}
    if len(months) == 1:
        return months.pop()
    return f"{min(dates).isoformat()}..{max(dates).isoformat()}"


def _infer_cnpj(transactions: list[_Transaction], receipts: list[_Receipt]) -> str:
    cnpjs = {item.cnpj for item in [*transactions, *receipts] if item.cnpj}
    if not cnpjs:
        return ""
    if len(cnpjs) == 1:
        return cnpjs.pop()
    return "MULTI"


def _date_delta(bank_date: date, receipt_date: date | None) -> int | None:
    if receipt_date is None:
        return None
    return abs((bank_date - receipt_date).days)


def _has_strong_text_identifier(transaction: _Transaction, receipt: _Receipt) -> bool:
    bank_text = normalize_key(transaction.description_bank)
    bank_digits = only_digits(transaction.description_bank)
    receipt_document = only_digits(receipt.beneficiary_document)
    if len(receipt_document) >= 8 and receipt_document in bank_digits:
        return True

    reference = normalize_key(receipt.reference_id)
    if len(reference) >= 4 and reference in bank_text:
        return True

    beneficiary = normalize_key(receipt.beneficiary_name)
    return len(beneficiary) >= 8 and beneficiary in bank_text


def _classify_bank_transaction(row: dict[str, str], description_bank: str) -> str:
    amount = parse_decimal(row.get("amount"))
    text_key = normalize_key(description_bank)

    if amount > 0:
        return "ENTRADA_FORA_ESCOPO"

    if amount < 0:
        if "sispag" in text_key:
            return "SISPAG_AGREGADO"
        if "tarifa" in text_key or "taxa" in text_key:
            return "TARIFA_BANCARIA"
        if "pix" in text_key:
            return "PIX_TRANSFERENCIA_SAIDA"
        if any(
            token in text_key
            for token in ("boleto", "ted", "transferencia", "transf")
        ):
            return "BOLETO_TED_TRANSFERENCIA_SAIDA"
        return "SAIDA_PAGAMENTO"

    return "OUTRO"


def _is_matching_scope_transaction(transaction: _Transaction) -> bool:
    return transaction.technical_category in {
        "SAIDA_PAGAMENTO",
        "SISPAG_AGREGADO",
        "PIX_TRANSFERENCIA_SAIDA",
        "BOLETO_TED_TRANSFERENCIA_SAIDA",
        "TARIFA_BANCARIA",
    }


def _rule_with_category(rule: str, technical_category: str) -> str:
    if technical_category == "SISPAG_AGREGADO":
        return f"{rule}|SISPAG_AGREGADO"
    return rule


def _out_of_scope_entry_rows(transactions: list[_Transaction]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for transaction in transactions:
        rows.append(
            {
                **transaction.source_row,
                "technical_category": transaction.technical_category,
                "exclusion_reason": "FORA_ESCOPO_ENTRADA",
            }
        )
    return rows


def _money(value: Decimal) -> Decimal:
    return value.quantize(Decimal("0.01"))


def _decimal_text(value: Decimal) -> str:
    return f"{_money(value):.2f}"


def _csv_value(value: object) -> str:
    if value is None:
        return ""
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)


def _ensure_file(path: Path, label: str) -> None:
    if not path.exists():
        raise FileNotFoundError(f"{label} nao encontrado: {path}")
    if not path.is_file():
        raise ValueError(f"{label} nao e um arquivo: {path}")
