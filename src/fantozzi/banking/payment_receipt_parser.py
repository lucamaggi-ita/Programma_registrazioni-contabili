from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from fantozzi.banking.common import normalize_key, parse_date, parse_decimal, stable_hash, text
from fantozzi.banking.models import PaymentReceiptRaw


COLUMN_ALIASES = {
    "beneficiary_name": {
        "favorecidobeneficiario",
        "favorecidobeneficiário",
        "beneficiario",
    },
    "beneficiary_document": {"cpfcnpj"},
    "payment_type": {"tipodepagamento"},
    "reference_id": {"referenciadaempresa", "referenciaempresa"},
    "payment_date": {"datadopagamento"},
    "amount": {"valorr", "valor"},
    "status": {"status"},
}


def parse_payment_receipts(path: Path) -> list[PaymentReceiptRaw]:
    rows = _read_xls_rows(path)
    return parse_payment_receipts_rows(rows, source_file=path.name)


def parse_payment_receipts_rows(
    rows: list[list[object]],
    *,
    source_file: str = "ConsultaPagamentos.xls",
) -> list[PaymentReceiptRaw]:
    header_index, columns = _detect_header(rows)
    receipts: list[PaymentReceiptRaw] = []
    for row in rows[header_index + 1 :]:
        record = {
            name: row[index] if index < len(row) else None
            for name, index in columns.items()
        }
        if not text(record.get("beneficiary_name")):
            continue
        amount = parse_decimal(record.get("amount"))
        payment_date = parse_date(record.get("payment_date"))
        raw_hash = stable_hash(
            source_file,
            record.get("beneficiary_name"),
            record.get("beneficiary_document"),
            record.get("payment_type"),
            record.get("reference_id"),
            payment_date,
            amount,
            record.get("status"),
        )
        receipts.append(
            PaymentReceiptRaw(
                payment_date=payment_date,
                amount=amount,
                beneficiary_name=text(record.get("beneficiary_name")),
                beneficiary_document=text(record.get("beneficiary_document")),
                payment_type=text(record.get("payment_type")),
                reference_id=text(record.get("reference_id")),
                status=text(record.get("status")),
                source_file=source_file,
                raw_hash=raw_hash,
            )
        )
    return receipts


def _read_xls_rows(path: Path) -> list[list[object]]:
    try:
        import xlrd
    except ImportError as exc:
        raise RuntimeError(
            "Leitura de .xls requer a dependencia xlrd>=2.0.1."
        ) from exc

    workbook = xlrd.open_workbook(str(path))
    sheet = workbook.sheet_by_index(0)
    rows: list[list[object]] = []
    for row_index in range(sheet.nrows):
        values = []
        for col_index in range(sheet.ncols):
            cell = sheet.cell(row_index, col_index)
            values.append(_xlrd_value(cell, workbook.datemode, xlrd))
        rows.append(values)
    return rows


def _xlrd_value(cell: object, datemode: int, xlrd_module: object) -> object:
    ctype = getattr(cell, "ctype", None)
    value = getattr(cell, "value", None)
    xldate_type = getattr(xlrd_module, "XL_CELL_DATE", object())
    if ctype == xldate_type:
        parsed = xlrd_module.xldate.xldate_as_datetime(value, datemode)
        return parsed.date() if isinstance(parsed, datetime) else parsed
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def _detect_header(rows: list[list[object]]) -> tuple[int, dict[str, int]]:
    for index, row in enumerate(rows):
        columns = _columns_from_header(row)
        if {"beneficiary_name", "payment_date", "amount", "status"}.issubset(columns):
            return index, columns
    raise ValueError("Header ConsultaPagamentos nao encontrado.")


def _columns_from_header(row: list[object]) -> dict[str, int]:
    columns: dict[str, int] = {}
    for index, header in enumerate(row):
        key = normalize_key(header)
        for canonical, aliases in COLUMN_ALIASES.items():
            if key in aliases:
                columns[canonical] = index
    return columns
