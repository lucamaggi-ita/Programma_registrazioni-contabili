from __future__ import annotations

import csv
from dataclasses import asdict
from pathlib import Path

from fantozzi.banking.common import is_balance_description, only_digits, stable_hash
from fantozzi.banking.manifest import write_manifest
from fantozzi.banking.models import (
    BankBalanceAudit,
    BankImportManifest,
    BankLaunchEnrichment,
    BankTransactionNormalized,
    BankTransactionRaw,
    PaymentReceiptNormalized,
    PaymentReceiptRaw,
)
from fantozzi.banking.ofx_parser import parse_ofx_balance_audit, parse_ofx_transactions
from fantozzi.banking.payment_receipt_parser import parse_payment_receipts
from fantozzi.banking.xlsx_bank_parser import parse_bank_launches, summarize_bank_launch_rows


TRANSACTION_HEADERS = [
    "transaction_id",
    "bank_name",
    "account_id",
    "cnpj",
    "transaction_date",
    "amount",
    "transaction_type",
    "description_raw",
    "description_enriched",
    "counterparty_name",
    "counterparty_document",
    "fitid",
    "source_file",
    "source_type",
    "raw_hash",
]

PAYMENT_HEADERS = [
    "payment_id",
    "cnpj",
    "payment_date",
    "amount",
    "beneficiary_name",
    "beneficiary_document",
    "payment_type",
    "reference_id",
    "status",
    "source_file",
    "raw_hash",
]

AUDIT_HEADERS = [
    "cnpj",
    "account_id",
    "period",
    "opening_balance",
    "closing_balance",
    "source_file",
]


def import_bank_files(
    *,
    cnpj: str,
    ofx: Path,
    output_root: Path,
    xlsx_lancamentos: Path | None = None,
    consulta_pagamentos: Path | None = None,
    ofx_saldos: Path | None = None,
) -> BankImportManifest:
    cnpj_digits = only_digits(cnpj)
    if not cnpj_digits:
        raise ValueError("CNPJ obrigatorio.")
    _ensure_file(ofx, "OFX primario")

    output_dir = output_root / cnpj_digits
    output_dir.mkdir(parents=True, exist_ok=True)

    all_primary_ofx_transactions = parse_ofx_transactions(ofx, include_balance_lines=True)
    raw_transactions = [
        transaction
        for transaction in all_primary_ofx_transactions
        if not is_balance_description(transaction.description_raw)
    ]
    enrichments: list[BankLaunchEnrichment] = []
    audits: list[BankBalanceAudit] = []
    avisos: list[str] = []
    optional_inputs_absent: list[str] = []
    filtered_rows: dict[str, int] = {
        "primary_ofx_balance_rows_filtered": (
            len(all_primary_ofx_transactions) - len(raw_transactions)
        )
    }

    if filtered_rows["primary_ofx_balance_rows_filtered"]:
        avisos.append("Linhas de saldo filtradas do OFX primario.")

    if xlsx_lancamentos:
        _ensure_file(xlsx_lancamentos, "XLSX lancamentos")
        enrichments, xlsx_audits = parse_bank_launches(xlsx_lancamentos)
        audits.extend(_fill_audit(audit, cnpj_digits, raw_transactions) for audit in xlsx_audits)
        filtered_rows.update(summarize_bank_launch_rows(xlsx_lancamentos))
        if filtered_rows.get("xlsx_lancamentos_balance_rows_filtered", 0):
            avisos.append("Linhas de saldo filtradas do XLSX de lancamentos.")
    else:
        optional_inputs_absent.append("xlsx_lancamentos")
        avisos.append("Arquivo opcional ausente: xlsx_lancamentos.")

    if ofx_saldos:
        _ensure_file(ofx_saldos, "OFX saldos")
        audits.extend(parse_ofx_balance_audit(ofx_saldos, cnpj=cnpj_digits))
    else:
        optional_inputs_absent.append("ofx_saldos")
        avisos.append("Arquivo opcional ausente: ofx_saldos.")

    transactions = normalize_transactions(raw_transactions, cnpj_digits, enrichments)

    receipts: list[PaymentReceiptNormalized] = []
    if consulta_pagamentos:
        _ensure_file(consulta_pagamentos, "ConsultaPagamentos")
        receipts = normalize_payment_receipts(
            parse_payment_receipts(consulta_pagamentos),
            cnpj_digits,
        )
    else:
        optional_inputs_absent.append("consulta_pagamentos")
        avisos.append("Arquivo opcional ausente: consulta_pagamentos.")

    transaction_path = output_dir / "bank_transactions_normalized.csv"
    receipts_path = output_dir / "payment_receipts_normalized.csv"
    audit_path = output_dir / "bank_balance_audit.csv"

    _write_csv(transaction_path, TRANSACTION_HEADERS, transactions)
    _write_csv(receipts_path, PAYMENT_HEADERS, receipts)
    _write_csv(audit_path, AUDIT_HEADERS, audits)

    input_files = {"ofx": str(ofx)}
    if xlsx_lancamentos:
        input_files["xlsx_lancamentos"] = str(xlsx_lancamentos)
    if consulta_pagamentos:
        input_files["consulta_pagamentos"] = str(consulta_pagamentos)
    if ofx_saldos:
        input_files["ofx_saldos"] = str(ofx_saldos)

    manifest = BankImportManifest(
        cnpj=cnpj_digits,
        output_dir=output_dir,
        input_files=input_files,
        transaction_count=len(transactions),
        payment_receipt_count=len(receipts),
        balance_audit_count=len(audits),
        status="OK" if transactions else "ATENCAO",
        avisos=avisos,
        optional_inputs_absent=optional_inputs_absent,
        filtered_rows=filtered_rows,
        generated_files={
            "bank_transactions_normalized": str(transaction_path),
            "payment_receipts_normalized": str(receipts_path),
            "bank_balance_audit": str(audit_path),
            "manifest": str(output_dir / "manifest_bank_import.json"),
        },
    )
    write_manifest(manifest)
    return manifest


def normalize_transactions(
    transactions: list[BankTransactionRaw],
    cnpj: str,
    enrichments: list[BankLaunchEnrichment] | None = None,
) -> list[BankTransactionNormalized]:
    enrichment_index = _enrichment_index(enrichments or [])
    normalized: list[BankTransactionNormalized] = []
    for tx in transactions:
        key = (tx.transaction_date.isoformat(), str(tx.amount))
        enrichment = enrichment_index.get(key, []).pop(0) if enrichment_index.get(key) else None
        transaction_id = (
            f"{cnpj}:{tx.account_id}:{tx.fitid}"
            if tx.fitid
            else stable_hash(cnpj, tx.account_id, tx.transaction_date, tx.amount, tx.raw_hash)
        )
        normalized.append(
            BankTransactionNormalized(
                transaction_id=transaction_id,
                bank_name=tx.bank_name,
                account_id=tx.account_id,
                cnpj=cnpj,
                transaction_date=tx.transaction_date,
                amount=tx.amount,
                transaction_type=tx.transaction_type,
                description_raw=tx.description_raw,
                description_enriched=enrichment.description if enrichment else "",
                counterparty_name=enrichment.counterparty_name if enrichment else "",
                counterparty_document=enrichment.counterparty_document if enrichment else "",
                fitid=tx.fitid,
                source_file=tx.source_file,
                source_type=tx.source_type,
                raw_hash=tx.raw_hash,
            )
        )
    return normalized


def normalize_payment_receipts(
    receipts: list[PaymentReceiptRaw],
    cnpj: str,
) -> list[PaymentReceiptNormalized]:
    normalized: list[PaymentReceiptNormalized] = []
    for receipt in receipts:
        payment_id = stable_hash(
            cnpj,
            receipt.payment_date,
            receipt.amount,
            receipt.beneficiary_document,
            receipt.reference_id,
            receipt.raw_hash,
        )
        normalized.append(
            PaymentReceiptNormalized(
                payment_id=payment_id,
                cnpj=cnpj,
                payment_date=receipt.payment_date,
                amount=receipt.amount,
                beneficiary_name=receipt.beneficiary_name,
                beneficiary_document=receipt.beneficiary_document,
                payment_type=receipt.payment_type,
                reference_id=receipt.reference_id,
                status=receipt.status,
                source_file=receipt.source_file,
                raw_hash=receipt.raw_hash,
            )
        )
    return normalized


def _enrichment_index(
    enrichments: list[BankLaunchEnrichment],
) -> dict[tuple[str, str], list[BankLaunchEnrichment]]:
    result: dict[tuple[str, str], list[BankLaunchEnrichment]] = {}
    for enrichment in enrichments:
        if enrichment.amount is None:
            continue
        key = (enrichment.transaction_date.isoformat(), str(enrichment.amount))
        result.setdefault(key, []).append(enrichment)
    return result


def _fill_audit(
    audit: BankBalanceAudit,
    cnpj: str,
    transactions: list[BankTransactionRaw],
) -> BankBalanceAudit:
    account_id = transactions[0].account_id if transactions else ""
    return BankBalanceAudit(
        cnpj=cnpj,
        account_id=audit.account_id or account_id,
        period=audit.period,
        opening_balance=audit.opening_balance,
        closing_balance=audit.closing_balance,
        source_file=audit.source_file,
    )


def _write_csv(path: Path, headers: list[str], rows: list[object]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=headers, delimiter=";")
        writer.writeheader()
        for row in rows:
            data = asdict(row)
            writer.writerow({header: _csv_value(data.get(header)) for header in headers})


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
