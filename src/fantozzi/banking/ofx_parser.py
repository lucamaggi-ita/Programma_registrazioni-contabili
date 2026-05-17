from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path
import re

from fantozzi.banking.common import is_balance_description, parse_date, stable_hash, text
from fantozzi.banking.models import BankBalanceAudit, BankTransactionRaw


def parse_ofx_transactions(
    path: Path,
    *,
    include_balance_lines: bool = False,
) -> list[BankTransactionRaw]:
    content = _read_text(path)
    bank_name = _tag(content, "ORG") or _tag(content, "BANKID")
    account_id = _tag(content, "ACCTID")
    transactions: list[BankTransactionRaw] = []
    for block in _transaction_blocks(content):
        memo = _tag(block, "MEMO")
        if not include_balance_lines and is_balance_description(memo):
            continue
        tx_date = parse_date((_tag(block, "DTPOSTED") or "")[:8])
        if tx_date is None:
            raise ValueError(f"Transacao OFX sem data: {path}")
        amount = Decimal((_tag(block, "TRNAMT") or "0").replace(",", "."))
        trntype = _tag(block, "TRNTYPE")
        fitid = _tag(block, "FITID")
        raw_hash = stable_hash(
            path.name,
            _tag(block, "DTPOSTED"),
            amount,
            trntype,
            memo,
            fitid,
        )
        transactions.append(
            BankTransactionRaw(
                transaction_date=tx_date,
                amount=amount,
                transaction_type=trntype,
                description_raw=memo,
                fitid=fitid,
                bank_name=bank_name,
                account_id=account_id,
                source_file=path.name,
                source_type="ofx",
                raw_hash=raw_hash,
            )
        )
    return transactions


def parse_ofx_balance_audit(path: Path, *, cnpj: str) -> list[BankBalanceAudit]:
    content = _read_text(path)
    balance_transactions = [
        transaction
        for transaction in parse_ofx_transactions(path, include_balance_lines=True)
        if is_balance_description(transaction.description_raw)
    ]
    if not balance_transactions:
        balance = _tag(content, "BALAMT")
        if not balance:
            return []
        return [
            BankBalanceAudit(
                cnpj=cnpj,
                account_id=_tag(content, "ACCTID"),
                period=_period(content),
                opening_balance=None,
                closing_balance=Decimal(balance.replace(",", ".")),
                source_file=path.name,
            )
        ]

    opening: Decimal | None = None
    closing: Decimal | None = None
    for transaction in balance_transactions:
        memo = transaction.description_raw.upper()
        if "SALDO ANTERIOR" in memo and opening is None:
            opening = transaction.amount
        if "SALDO TOTAL" in memo or "DISPON" in memo or "SALDO DO DIA" in memo:
            closing = transaction.amount
    return [
        BankBalanceAudit(
            cnpj=cnpj,
            account_id=_tag(content, "ACCTID"),
            period=_period(content),
            opening_balance=opening,
            closing_balance=closing,
            source_file=path.name,
        )
    ]


def _read_text(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", errors="replace")


def _transaction_blocks(content: str) -> list[str]:
    return re.findall(
        r"<STMTTRN>(.*?)(?=<STMTTRN>|</BANKTRANLIST>)",
        content,
        flags=re.IGNORECASE | re.DOTALL,
    )


def _tag(content: str, tag: str) -> str:
    match = re.search(rf"<{tag}>([^<\r\n]*)", content, flags=re.IGNORECASE)
    return text(match.group(1)) if match else ""


def _period(content: str) -> str:
    start = _date_text(_tag(content, "DTSTART"))
    end = _date_text(_tag(content, "DTEND"))
    if start and end:
        return f"{start}/{end}"
    return start or end


def _date_text(value: str) -> str:
    parsed = parse_date(value[:8]) if value else None
    return parsed.isoformat() if isinstance(parsed, date) else ""
