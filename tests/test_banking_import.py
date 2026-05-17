from __future__ import annotations

import csv
import json
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
import pytest

from fantozzi.banking.normalizer import (
    import_bank_files,
    normalize_payment_receipts,
    normalize_transactions,
)
from fantozzi.banking.ofx_parser import parse_ofx_balance_audit, parse_ofx_transactions
from fantozzi.banking.payment_receipt_parser import (
    parse_payment_receipts,
    parse_payment_receipts_rows,
)
from fantozzi.banking.xlsx_bank_parser import detect_header_row, parse_bank_launches
from fantozzi.cli import build_parser, main


def test_parse_ofx_standard_preserves_fitid_and_amount_sign(tmp_path):
    ofx_path = _write_ofx(
        tmp_path / "extrato.ofx",
        [
            _ofx_tx("20260201", "-10.50", "DEBIT", "PIX FORNECEDOR", "FIT-1"),
            _ofx_tx("20260202", "25.00", "CREDIT", "RECEBIMENTO", "FIT-2"),
        ],
    )

    transactions = parse_ofx_transactions(ofx_path)

    assert len(transactions) == 2
    assert transactions[0].fitid == "FIT-1"
    assert transactions[0].amount == Decimal("-10.50")
    assert transactions[1].amount == Decimal("25.00")


def test_ofx_balance_lines_are_isolated_from_transactions(tmp_path):
    ofx_path = _write_ofx(
        tmp_path / "extrato_saldos.ofx",
        [
            _ofx_tx("20260201", "100.00", "CREDIT", "SALDO ANTERIOR", "SALDO-1"),
            _ofx_tx("20260201", "-20.00", "DEBIT", "PAGAMENTO", "FIT-1"),
            _ofx_tx(
                "20260228",
                "80.00",
                "CREDIT",
                "SALDO TOTAL DISPONIVEL DIA",
                "SALDO-2",
            ),
        ],
    )

    transactions = parse_ofx_transactions(ofx_path)
    audit = parse_ofx_balance_audit(ofx_path, cnpj="08884180000123")

    assert [tx.fitid for tx in transactions] == ["FIT-1"]
    assert audit[0].opening_balance == Decimal("100.00")
    assert audit[0].closing_balance == Decimal("80.00")


def test_xlsx_lancamentos_detects_header_row_10_aliases_and_balance_rows(tmp_path):
    xlsx_path = tmp_path / "Extrato_Lancamentos.xlsx"
    _write_lancamentos_xlsx(
        xlsx_path,
        headers=["Data", "Lançamento", "Razão Social", "CPF/CNPJ", "Valor (R$)", "Saldo (R$)"],
    )

    header_index = detect_header_row(xlsx_path)
    enrichments, audits = parse_bank_launches(xlsx_path)

    assert header_index == 9
    assert len(enrichments) == 1
    assert enrichments[0].amount == Decimal("-20.00")
    assert enrichments[0].counterparty_name == "Fornecedor Teste"
    assert audits[0].opening_balance == Decimal("100.00")
    assert audits[0].closing_balance == Decimal("80.00")


def test_xlsx_lancamentos_accepts_plain_valor_saldo_aliases(tmp_path):
    xlsx_path = tmp_path / "Extrato_Lancamentos_alias.xlsx"
    _write_lancamentos_xlsx(
        xlsx_path,
        headers=["Data", "Lançamento", "Razão Social", "CPF CNPJ", "Valor", "Saldo"],
    )

    enrichments, audits = parse_bank_launches(xlsx_path)

    assert enrichments[0].amount == Decimal("-20.00")
    assert enrichments[0].counterparty_document == "12.345.678/0001-90"
    assert audits[0].closing_balance == Decimal("80.00")


def test_parse_consulta_pagamentos_xls_via_parser_rows(monkeypatch, tmp_path):
    rows = _consulta_pagamentos_rows()

    def fake_read_xls_rows(path):
        assert path.suffix == ".xls"
        return rows

    monkeypatch.setattr(
        "fantozzi.banking.payment_receipt_parser._read_xls_rows",
        fake_read_xls_rows,
    )

    receipts = parse_payment_receipts(tmp_path / "ConsultaPagamentos.xls")

    assert len(receipts) == 1
    assert receipts[0].beneficiary_name == "Fornecedor Teste"
    assert receipts[0].amount == Decimal("20.00")
    assert receipts[0].status == "Efetuado"


def test_normalizes_payment_receipts_and_hash_is_deterministic():
    raw = parse_payment_receipts_rows(_consulta_pagamentos_rows())

    first = normalize_payment_receipts(raw, "08884180000123")
    second = normalize_payment_receipts(raw, "08884180000123")

    assert first[0].payment_id == second[0].payment_id
    assert first[0].cnpj == "08884180000123"


def test_normalizes_transactions_with_stable_fitid_id_and_enrichment(tmp_path):
    ofx_path = _write_ofx(
        tmp_path / "extrato.ofx",
        [_ofx_tx("20260201", "-20.00", "DEBIT", "PIX RAW", "FIT-1")],
    )
    xlsx_path = tmp_path / "lancamentos.xlsx"
    _write_lancamentos_xlsx(xlsx_path)
    raw_transactions = parse_ofx_transactions(ofx_path)
    enrichments, _ = parse_bank_launches(xlsx_path)

    normalized = normalize_transactions(
        raw_transactions,
        "08884180000123",
        enrichments,
    )

    assert normalized[0].transaction_id == "08884180000123:23130-4:FIT-1"
    assert normalized[0].description_enriched == "PIX Fornecedor"
    assert normalized[0].counterparty_name == "Fornecedor Teste"


def test_import_bank_files_writes_manifest_and_keeps_cnpj_separated(monkeypatch, tmp_path):
    ofx_path = _write_ofx(
        tmp_path / "extrato.ofx",
        [_ofx_tx("20260201", "-20.00", "DEBIT", "PIX RAW", "FIT-1")],
    )
    xlsx_path = tmp_path / "lancamentos.xlsx"
    _write_lancamentos_xlsx(xlsx_path)

    monkeypatch.setattr(
        "fantozzi.banking.normalizer.parse_payment_receipts",
        lambda path: parse_payment_receipts_rows(_consulta_pagamentos_rows()),
    )
    consulta_path = tmp_path / "ConsultaPagamentos.xls"
    consulta_path.write_bytes(b"synthetic-xls-placeholder")

    manifest = import_bank_files(
        cnpj="08.884.180/0001-23",
        ofx=ofx_path,
        xlsx_lancamentos=xlsx_path,
        consulta_pagamentos=consulta_path,
        output_root=tmp_path / "output",
    )

    assert manifest.cnpj == "08884180000123"
    output_dir = tmp_path / "output" / "08884180000123"
    assert (output_dir / "bank_transactions_normalized.csv").exists()
    assert (output_dir / "payment_receipts_normalized.csv").exists()
    assert (output_dir / "bank_balance_audit.csv").exists()

    manifest_json = json.loads((output_dir / "manifest_bank_import.json").read_text())
    assert manifest_json["status"] == "OK"
    assert isinstance(manifest_json["avisos"], list)
    assert "ofx_saldos" in manifest_json["optional_inputs_absent"]
    assert manifest_json["matching_executed"] is False
    assert manifest_json["classification_executed"] is False
    assert manifest_json["transaction_count"] == 1
    assert manifest_json["payment_receipt_count"] == 1
    assert manifest_json["filtered_rows"]["primary_ofx_balance_rows_filtered"] == 0
    assert manifest_json["filtered_rows"]["xlsx_lancamentos_balance_rows_filtered"] == 2
    assert "Linhas de saldo filtradas do XLSX de lancamentos." in manifest_json["avisos"]

    with (output_dir / "bank_transactions_normalized.csv").open(
        encoding="utf-8-sig", newline=""
    ) as file:
        rows = list(csv.DictReader(file, delimiter=";"))
    assert rows[0]["cnpj"] == "08884180000123"


def test_import_bank_files_requires_cnpj_and_primary_ofx(tmp_path):
    with pytest.raises(ValueError):
        import_bank_files(cnpj="", ofx=tmp_path / "missing.ofx", output_root=tmp_path)

    with pytest.raises(FileNotFoundError):
        import_bank_files(
            cnpj="08884180000123",
            ofx=tmp_path / "missing.ofx",
            output_root=tmp_path,
        )


def test_parser_accepts_importar_extratos_bancarios_command():
    parser = build_parser()

    args = parser.parse_args(
        [
            "importar-extratos-bancarios",
            "--cnpj",
            "08884180000123",
            "--ofx",
            "extrato.ofx",
            "--xlsx-lancamentos",
            "lancamentos.xlsx",
            "--consulta-pagamentos",
            "ConsultaPagamentos.xls",
            "--ofx-saldos",
            "saldos.ofx",
            "--saida",
            "output/banking_import",
        ]
    )

    assert args.command == "importar-extratos-bancarios"
    assert args.cnpj == "08884180000123"
    assert args.ofx == "extrato.ofx"
    assert args.xlsx_lancamentos == "lancamentos.xlsx"
    assert args.consulta_pagamentos == "ConsultaPagamentos.xls"
    assert args.ofx_saldos == "saldos.ofx"


def test_cli_importar_extratos_returns_error_without_primary_ofx(tmp_path):
    result = main(
        [
            "importar-extratos-bancarios",
            "--cnpj",
            "08884180000123",
            "--ofx",
            str(tmp_path / "missing.ofx"),
            "--saida",
            str(tmp_path / "output"),
        ]
    )

    assert result == 1


def _write_ofx(path: Path, transactions: list[str]) -> Path:
    content = """OFXHEADER:100
DATA:OFXSGML

<OFX>
<BANKMSGSRSV1>
<STMTTRNRS>
<STMTRS>
<BANKACCTFROM>
<BANKID>341
<ACCTID>23130-4
</BANKACCTFROM>
<BANKTRANLIST>
<DTSTART>20260201000000[-3:BRT]
<DTEND>20260228235959[-3:BRT]
{transactions}
</BANKTRANLIST>
</STMTRS>
</STMTTRNRS>
</BANKMSGSRSV1>
</OFX>
""".format(transactions="\n".join(transactions))
    path.write_text(content, encoding="utf-8")
    return path


def _ofx_tx(date: str, amount: str, trntype: str, memo: str, fitid: str) -> str:
    return f"""<STMTTRN>
<TRNTYPE>{trntype}
<DTPOSTED>{date}000000[-3:BRT]
<TRNAMT>{amount}
<FITID>{fitid}
<MEMO>{memo}
"""


def _write_lancamentos_xlsx(
    path: Path,
    *,
    headers: list[str] | None = None,
) -> None:
    headers = headers or [
        "Data",
        "Lançamento",
        "Razão Social",
        "CPF/CNPJ",
        "Valor (R$)",
        "Saldo (R$)",
    ]
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Lançamentos"
    for row_index in range(1, 10):
        sheet.append(["Metadado", f"linha {row_index}"])
    sheet.append(headers)
    sheet.append(["31/01/2026", "SALDO ANTERIOR", "", "", "", "100.00"])
    sheet.append(
        [
            "01/02/2026",
            "PIX Fornecedor",
            "Fornecedor Teste",
            "12.345.678/0001-90",
            "-20.00",
            "",
        ]
    )
    sheet.append(["28/02/2026", "SALDO TOTAL DISPONÍVEL DIA", "", "", "", "80.00"])
    workbook.save(path)


def _consulta_pagamentos_rows() -> list[list[object]]:
    return [
        ["Dados da conta"],
        ["Agência/conta", "1648 / 23130-4"],
        ["CNPJ:", "08.884.180/0001-23"],
        [],
        [
            "favorecido / beneficiário",
            "CPF/CNPJ",
            "tipo de pagamento",
            "referência da empresa",
            "data do pagamento",
            "valor (R$)",
            "status",
        ],
        [
            "Fornecedor Teste",
            "12.345.678/0001-90",
            "PIX Transferências",
            "REF-1",
            "05/02/2026",
            "20,00",
            "Efetuado",
        ],
    ]
