from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from pathlib import Path


@dataclass(frozen=True)
class BankTransactionRaw:
    transaction_date: date
    amount: Decimal
    transaction_type: str
    description_raw: str
    fitid: str
    bank_name: str
    account_id: str
    source_file: str
    source_type: str
    raw_hash: str


@dataclass(frozen=True)
class BankTransactionNormalized:
    transaction_id: str
    bank_name: str
    account_id: str
    cnpj: str
    transaction_date: date
    amount: Decimal
    transaction_type: str
    description_raw: str
    description_enriched: str
    counterparty_name: str
    counterparty_document: str
    fitid: str
    source_file: str
    source_type: str
    raw_hash: str


@dataclass(frozen=True)
class BankLaunchEnrichment:
    transaction_date: date
    amount: Decimal | None
    description: str
    counterparty_name: str
    counterparty_document: str
    row_balance: Decimal | None
    source_file: str
    raw_hash: str


@dataclass(frozen=True)
class PaymentReceiptRaw:
    payment_date: date | None
    amount: Decimal
    beneficiary_name: str
    beneficiary_document: str
    payment_type: str
    reference_id: str
    status: str
    source_file: str
    raw_hash: str


@dataclass(frozen=True)
class PaymentReceiptNormalized:
    payment_id: str
    cnpj: str
    payment_date: date | None
    amount: Decimal
    beneficiary_name: str
    beneficiary_document: str
    payment_type: str
    reference_id: str
    status: str
    source_file: str
    raw_hash: str


@dataclass(frozen=True)
class BankBalanceAudit:
    cnpj: str
    account_id: str
    period: str
    opening_balance: Decimal | None
    closing_balance: Decimal | None
    source_file: str


@dataclass(frozen=True)
class BankImportManifest:
    cnpj: str
    output_dir: Path
    input_files: dict[str, str]
    transaction_count: int
    payment_receipt_count: int
    balance_audit_count: int
    status: str = "OK"
    avisos: list[str] = field(default_factory=list)
    optional_inputs_absent: list[str] = field(default_factory=list)
    filtered_rows: dict[str, int] = field(default_factory=dict)
    generated_files: dict[str, str] = field(default_factory=dict)
