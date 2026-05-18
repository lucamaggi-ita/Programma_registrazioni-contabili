from __future__ import annotations

import csv
import json
from pathlib import Path

import pytest

from fantozzi.banking.matching import match_bank_files
from fantozzi.cli import build_parser, main


def test_matching_exact_amount_and_date(tmp_path):
    input_dir = _write_inputs(
        tmp_path,
        transactions=[
            _tx("TX-1", "2026-02-05", "-20.00", "PIX FORNECEDOR TESTE"),
        ],
        receipts=[
            _receipt("RC-1", "2026-02-05", "20.00", "Fornecedor Teste"),
        ],
    )

    manifest = match_bank_files(
        input_dir=input_dir,
        output_dir=tmp_path / "matching",
        periodo="2026-02",
    )

    rows = _read_csv(tmp_path / "matching" / "bank_matches.csv")
    assert manifest.exact_match_count == 1
    assert rows[0]["match_status"] == "MATCH_EXATO"
    assert rows[0]["match_rule"] == "VALOR_DATA_EXATA"
    assert rows[0]["amount_delta"] == "0.00"


def test_matching_ignores_credit_with_same_amount_and_date(tmp_path):
    input_dir = _write_inputs(
        tmp_path,
        transactions=[
            _tx("TX-CREDIT", "2026-02-05", "20.00", "PIX QR CODE RECEBIDO"),
            _tx("TX-DEBIT", "2026-02-05", "-20.00", "PIX FORNECEDOR TESTE"),
        ],
        receipts=[
            _receipt("RC-1", "2026-02-05", "20.00", "Fornecedor Teste"),
        ],
    )

    manifest = match_bank_files(
        input_dir=input_dir,
        output_dir=tmp_path / "matching",
        periodo="2026-02",
    )

    rows = _read_csv(tmp_path / "matching" / "bank_matches.csv")
    out_of_scope = _read_csv(tmp_path / "matching" / "bank_entradas_fora_escopo.csv")
    audit_rows = _read_csv(tmp_path / "matching" / "bank_matching_audit.csv")
    assert manifest.exact_match_count == 1
    assert manifest.out_of_scope_entry_count == 1
    assert rows[0]["bank_transaction_id"] == "TX-DEBIT"
    assert out_of_scope[0]["transaction_id"] == "TX-CREDIT"
    assert out_of_scope[0]["exclusion_reason"] == "FORA_ESCOPO_ENTRADA"
    assert any(row["rule"] == "FORA_ESCOPO_ENTRADA" for row in audit_rows)


def test_matching_does_not_match_receipt_against_credit_only(tmp_path):
    input_dir = _write_inputs(
        tmp_path,
        transactions=[
            _tx("TX-CREDIT", "2026-02-05", "20.00", "RECEBIMENTO REDE VISA"),
        ],
        receipts=[
            _receipt("RC-1", "2026-02-05", "20.00", "Fornecedor Teste"),
        ],
    )

    manifest = match_bank_files(
        input_dir=input_dir,
        output_dir=tmp_path / "matching",
        periodo="2026-02",
    )

    matches = _read_csv(tmp_path / "matching" / "bank_matches.csv")
    unmatched_receipts = _read_csv(tmp_path / "matching" / "bank_unmatched_receipts.csv")
    out_of_scope = _read_csv(tmp_path / "matching" / "bank_entradas_fora_escopo.csv")
    assert manifest.exact_match_count == 0
    assert manifest.outbound_transaction_count == 0
    assert matches == []
    assert unmatched_receipts[0]["payment_id"] == "RC-1"
    assert out_of_scope[0]["transaction_id"] == "TX-CREDIT"


def test_matching_probable_inside_date_window(tmp_path):
    input_dir = _write_inputs(
        tmp_path,
        transactions=[_tx("TX-1", "2026-02-07", "-20.00", "PIX")],
        receipts=[_receipt("RC-1", "2026-02-05", "20.00", "Fornecedor Teste")],
    )

    match_bank_files(
        input_dir=input_dir,
        output_dir=tmp_path / "matching",
        periodo="2026-02",
        window_days=3,
    )

    rows = _read_csv(tmp_path / "matching" / "bank_matches.csv")
    assert rows[0]["match_status"] == "MATCH_PROVAVEL"
    assert rows[0]["match_rule"] == "VALOR_JANELA_DATA"


def test_matching_marks_unmatched_transaction_and_receipt(tmp_path):
    input_dir = _write_inputs(
        tmp_path,
        transactions=[_tx("TX-1", "2026-02-07", "-30.00", "PIX SEM COMPROVANTE")],
        receipts=[_receipt("RC-1", "2026-02-05", "20.00", "Fornecedor Teste")],
    )

    manifest = match_bank_files(
        input_dir=input_dir,
        output_dir=tmp_path / "matching",
        periodo="2026-02",
    )

    matches = _read_csv(tmp_path / "matching" / "bank_matches.csv")
    unmatched_transactions = _read_csv(
        tmp_path / "matching" / "bank_unmatched_transactions.csv"
    )
    unmatched_receipts = _read_csv(tmp_path / "matching" / "bank_unmatched_receipts.csv")
    assert manifest.status == "ATENCAO"
    assert matches[0]["match_status"] == "NAO_ENCONTRADO"
    assert unmatched_transactions[0]["transaction_id"] == "TX-1"
    assert unmatched_receipts[0]["payment_id"] == "RC-1"


def test_matching_marks_ambiguous_when_two_receipts_are_candidates(tmp_path):
    input_dir = _write_inputs(
        tmp_path,
        transactions=[_tx("TX-1", "2026-02-05", "-20.00", "PIX")],
        receipts=[
            _receipt("RC-1", "2026-02-05", "20.00", "Fornecedor Um"),
            _receipt("RC-2", "2026-02-05", "20.00", "Fornecedor Dois"),
        ],
    )

    manifest = match_bank_files(
        input_dir=input_dir,
        output_dir=tmp_path / "matching",
        periodo="2026-02",
    )

    rows = _read_csv(tmp_path / "matching" / "bank_matches.csv")
    assert manifest.ambiguous_count == 1
    assert rows[0]["match_status"] == "MATCH_AMBIGUO"
    assert rows[0]["receipt_id"] == "RC-1|RC-2"


def test_matching_uses_strong_text_identifier_when_date_is_outside_window(tmp_path):
    input_dir = _write_inputs(
        tmp_path,
        transactions=[
            _tx("TX-1", "2026-02-20", "-20.00", "PAGAMENTO REF-FORTE"),
        ],
        receipts=[
            _receipt(
                "RC-1",
                "2026-02-05",
                "20.00",
                "Fornecedor Teste",
                reference_id="REF-FORTE",
            ),
        ],
    )

    match_bank_files(
        input_dir=input_dir,
        output_dir=tmp_path / "matching",
        periodo="2026-02",
        window_days=3,
    )

    rows = _read_csv(tmp_path / "matching" / "bank_matches.csv")
    assert rows[0]["match_status"] == "MATCH_PROVAVEL"
    assert rows[0]["match_rule"] == "IDENTIFICADOR_TEXTUAL_FORTE"


def test_matching_keeps_sispag_as_technical_conservative_category(tmp_path):
    input_dir = _write_inputs(
        tmp_path,
        transactions=[
            _tx("TX-1", "2026-02-05", "-20.00", "SISPAG FORNECEDORES"),
        ],
        receipts=[
            _receipt("RC-1", "2026-02-05", "20.00", "Fornecedor Teste"),
        ],
    )

    match_bank_files(
        input_dir=input_dir,
        output_dir=tmp_path / "matching",
        periodo="2026-02",
    )

    rows = _read_csv(tmp_path / "matching" / "bank_matches.csv")
    audit_rows = _read_csv(tmp_path / "matching" / "bank_matching_audit.csv")
    assert rows[0]["match_status"] == "MATCH_EXATO"
    assert rows[0]["match_rule"] == "VALOR_DATA_EXATA|SISPAG_AGREGADO"
    assert audit_rows[0]["rule"] == "VALOR_DATA_EXATA|SISPAG_AGREGADO"


def test_matching_manifest_has_status_counts_and_generated_files(tmp_path):
    input_dir = _write_inputs(
        tmp_path,
        transactions=[
            _tx("TX-1", "2026-02-05", "-20.00", "PIX FORNECEDOR TESTE"),
            _tx("TX-2", "2026-02-08", "-99.00", "PIX SEM COMPROVANTE"),
        ],
        receipts=[
            _receipt("RC-1", "2026-02-05", "20.00", "Fornecedor Teste"),
        ],
    )

    manifest = match_bank_files(
        input_dir=input_dir,
        output_dir=tmp_path / "matching",
        periodo="2026-02",
    )

    manifest_json = json.loads(
        (tmp_path / "matching" / "manifest_bank_matching.json").read_text(
            encoding="utf-8"
        )
    )
    assert manifest.unmatched_transaction_count == 1
    assert manifest_json["status"] == "ATENCAO"
    assert manifest_json["matching_executed"] is True
    assert manifest_json["classification_executed"] is False
    assert manifest_json["exact_match_count"] == 1
    assert manifest_json["outbound_transaction_count"] == 2
    assert manifest_json["out_of_scope_entry_count"] == 0
    assert "bank_matching_audit" in manifest_json["generated_files"]
    assert "bank_entradas_fora_escopo" in manifest_json["generated_files"]


def test_parser_rejects_executar_matching_bancario_command():
    parser = build_parser()

    with pytest.raises(SystemExit) as exc_info:
        parser.parse_args(
            [
                "executar-matching-bancario",
                "--entrada",
                "output/banking_import/08884180000123",
                "--saida",
                "output/banking_matching/08884180000123",
                "--periodo",
                "2026-02",
                "--janela-dias",
                "2",
            ]
        )

    assert exc_info.value.code == 2


def test_cli_matching_command_is_not_operational(tmp_path):
    with pytest.raises(SystemExit) as exc_info:
        main(
            [
                "executar-matching-bancario",
                "--entrada",
                str(tmp_path / "missing"),
                "--saida",
                str(tmp_path / "matching"),
            ]
        )

    assert exc_info.value.code == 2


def _write_inputs(
    tmp_path: Path,
    *,
    transactions: list[dict[str, str]],
    receipts: list[dict[str, str]],
) -> Path:
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    _write_csv(
        input_dir / "bank_transactions_normalized.csv",
        [
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
        ],
        transactions,
    )
    _write_csv(
        input_dir / "payment_receipts_normalized.csv",
        [
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
        ],
        receipts,
    )
    return input_dir


def _tx(
    transaction_id: str,
    transaction_date: str,
    amount: str,
    description: str,
) -> dict[str, str]:
    return {
        "transaction_id": transaction_id,
        "bank_name": "Itau",
        "account_id": "23130-4",
        "cnpj": "08884180000123",
        "transaction_date": transaction_date,
        "amount": amount,
        "transaction_type": "DEBIT",
        "description_raw": description,
        "description_enriched": "",
        "counterparty_name": "",
        "counterparty_document": "",
        "fitid": transaction_id,
        "source_file": "synthetic.ofx",
        "source_type": "ofx",
        "raw_hash": transaction_id,
    }


def _receipt(
    payment_id: str,
    payment_date: str,
    amount: str,
    beneficiary_name: str,
    *,
    reference_id: str | None = None,
) -> dict[str, str]:
    return {
        "payment_id": payment_id,
        "cnpj": "08884180000123",
        "payment_date": payment_date,
        "amount": amount,
        "beneficiary_name": beneficiary_name,
        "beneficiary_document": "12.345.678/0001-90",
        "payment_type": "PIX",
        "reference_id": reference_id or payment_id,
        "status": "Efetuado",
        "source_file": "ConsultaPagamentos.xls",
        "raw_hash": payment_id,
    }


def _write_csv(path: Path, headers: list[str], rows: list[dict[str, str]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=headers, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


def _read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file, delimiter=";"))
