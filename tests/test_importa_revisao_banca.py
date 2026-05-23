from __future__ import annotations

import json
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook

from fantozzi.banking.models import BankLaunchEnrichment, PaymentReceiptRaw
from fantozzi.banking.revisao_banca import (
    COMPROVANTES_SHEET,
    COMPROVANTES_XLSX,
    CONTA_CORRENTE_SHEET,
    CONTA_CORRENTE_XLSX,
    ImportaRevisaoBancaError,
    importar_revisao_banca,
    resolve_bank_raw_dir,
)
from fantozzi.cli import build_parser, main


def test_resolve_bank_raw_dir_from_year_and_month(tmp_path):
    path = resolve_bank_raw_dir(
        ano="2026",
        mes="04",
        project_root=tmp_path,
    )

    assert path == tmp_path / "bank_raw" / "2026" / "2026-04"


def test_importa_revisao_banca_preserves_existing_decisions_and_ignores_missing_ofx(
    monkeypatch,
    tmp_path,
):
    bank_raw_dir = tmp_path / "bank_raw" / "2026" / "2026-04"
    cnpj_dir = bank_raw_dir / "08884180000123"
    cnpj_dir.mkdir(parents=True)
    (cnpj_dir / "ConsultaPagamentos.xls").write_bytes(b"placeholder")
    (cnpj_dir / "Extrato_Lancamentos.xlsx").write_bytes(b"placeholder")
    output_dir = tmp_path / "bank_raw" / "_decisoes_regras"
    output_dir.mkdir(parents=True)
    _write_existing_comprovantes(output_dir / COMPROVANTES_XLSX)

    monkeypatch.setattr(
        "fantozzi.banking.revisao_banca.parse_payment_receipts",
        lambda path: [
            PaymentReceiptRaw(
                payment_date=date(2026, 4, 5),
                amount=Decimal("20.00"),
                beneficiary_name="Fornecedor Teste",
                beneficiary_document="12.345.678/0001-90",
                payment_type="PIX Transferencias",
                reference_id="-",
                status="Efetuado",
                source_file=path.name,
                raw_hash="receipt-1",
            )
        ],
    )
    monkeypatch.setattr(
        "fantozzi.banking.revisao_banca.parse_bank_launches",
        lambda path: (
            [
                BankLaunchEnrichment(
                    transaction_date=date(2026, 4, 5),
                    amount=Decimal("-20.00"),
                    description="PIX Fornecedor Teste",
                    counterparty_name="Fornecedor Teste",
                    counterparty_document="12.345.678/0001-90",
                    row_balance=None,
                    source_file=path.name,
                    raw_hash="launch-1",
                )
            ],
            [],
        ),
    )

    result = importar_revisao_banca(
        bank_raw_dir=bank_raw_dir,
        output_dir=output_dir,
        regras_path=_write_rules(tmp_path),
        plano_contas_path=_write_plan(tmp_path),
        mapeamento_dre_path=_write_mapping(tmp_path),
    )

    assert result.ofx_encontrado is False
    assert result.ofx_usado is False
    assert result.comprovantes_lidos == 1
    assert result.lancamentos_lidos == 1
    assert result.comprovantes_candidatos == 1
    assert result.conta_corrente_candidatos == 1
    assert result.decisoes_preservadas == 1
    assert result.novas_chaves_adicionadas == 1

    workbook = load_workbook(output_dir / COMPROVANTES_XLSX, data_only=False)
    sheet = workbook[COMPROVANTES_SHEET]
    headers = [cell.value for cell in sheet[1]]
    status_col = headers.index("STATUS_DECISAO") + 1
    conta_col = headers.index("CONTA_ID_APROVADA") + 1
    qtd_col = headers.index("QTD_OCORRENCIAS") + 1
    assert sheet.cell(row=2, column=status_col).value == "APROVADO"
    assert sheet.cell(row=2, column=conta_col).value == "4.01.012"
    assert sheet.cell(row=2, column=qtd_col).value == 1

    conta_workbook = load_workbook(output_dir / CONTA_CORRENTE_XLSX, data_only=False)
    assert CONTA_CORRENTE_SHEET in conta_workbook.sheetnames
    manifest = json.loads((output_dir / "manifest_importa_revisao_banca.json").read_text())
    assert manifest["ofx_usado"] is False
    assert manifest["yaml_modificado"] is False


def test_importa_revisao_banca_filters_total_receipt_from_decision_sheets(
    monkeypatch,
    tmp_path,
):
    bank_raw_dir = tmp_path / "bank_raw" / "2026" / "2026-04"
    cnpj_dir = bank_raw_dir / "08884180000123"
    cnpj_dir.mkdir(parents=True)
    (cnpj_dir / "ConsultaPagamentos.xls").write_bytes(b"placeholder")
    output_dir = tmp_path / "bank_raw" / "_decisoes_regras"

    monkeypatch.setattr(
        "fantozzi.banking.revisao_banca.parse_payment_receipts",
        lambda path: [
            PaymentReceiptRaw(
                payment_date=date(2026, 4, 5),
                amount=Decimal("20.00"),
                beneficiary_name="Fornecedor Teste",
                beneficiary_document="12.345.678/0001-90",
                payment_type="PIX Transferencias",
                reference_id="-",
                status="Efetuado",
                source_file=path.name,
                raw_hash="receipt-1",
            ),
            PaymentReceiptRaw(
                payment_date=None,
                amount=Decimal("1000.00"),
                beneficiary_name="TOTAL",
                beneficiary_document="",
                payment_type="",
                reference_id="",
                status="",
                source_file=path.name,
                raw_hash="receipt-total",
            ),
        ],
    )

    result = importar_revisao_banca(
        bank_raw_dir=bank_raw_dir,
        output_dir=output_dir,
        regras_path=_write_rules(tmp_path),
        plano_contas_path=_write_plan(tmp_path),
        mapeamento_dre_path=_write_mapping(tmp_path),
    )

    assert result.comprovantes_lidos == 2
    assert result.comprovantes_candidatos == 1

    workbook = load_workbook(output_dir / COMPROVANTES_XLSX, data_only=False)
    decisoes = workbook[COMPROVANTES_SHEET]
    decisoes_rows = list(decisoes.iter_rows(values_only=True))
    headers = [cell or "" for cell in decisoes_rows[0]]
    beneficiary_col = headers.index("BENEFICIARIO_FAVORECIDO")
    assert len(decisoes_rows) == 2
    assert all(row[beneficiary_col] != "TOTAL" for row in decisoes_rows[1:])

    tecnico = workbook["tecnico_comprovantes"]
    tecnico_rows = list(tecnico.iter_rows(values_only=True))
    tecnico_headers = [cell or "" for cell in tecnico_rows[0]]
    key_col = tecnico_headers.index("CHAVE_ESTAVEL")
    assert len(tecnico_rows) == 2
    assert all("BENEFICIARIO=TOTAL" not in row[key_col] for row in tecnico_rows[1:])


def test_cli_accepts_importar_revisao_banca_modes():
    parser = build_parser()

    args = parser.parse_args(
        [
            "importar-revisao-banca",
            "--bank-raw-dir",
            "bank_raw/2026/2026-04",
        ]
    )

    assert args.command == "importar-revisao-banca"
    assert args.bank_raw_dir == "bank_raw/2026/2026-04"

    args = parser.parse_args(
        [
            "importar-revisao-banca",
            "--ano",
            "2026",
            "--mes",
            "04",
        ]
    )

    assert args.ano == "2026"
    assert args.mes == "04"


def test_cli_importar_revisao_banca_requires_period():
    assert main(["importar-revisao-banca"]) == 1


def test_resolve_bank_raw_dir_rejects_ambiguous_inputs(tmp_path):
    try:
        resolve_bank_raw_dir(
            bank_raw_dir=tmp_path / "bank_raw" / "2026" / "2026-04",
            ano="2026",
            mes="04",
        )
    except ImportaRevisaoBancaError as exc:
        assert "nao ambos" in str(exc)
    else:
        raise AssertionError("expected ImportaRevisaoBancaError")


def _write_existing_comprovantes(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = COMPROVANTES_SHEET
    headers = [
        "STATUS_DECISAO",
        "ACAO_REGRA",
        "PERIODO",
        "CNPJ",
        "DATA_PAGAMENTO",
        "VALOR",
        "BENEFICIARIO_FAVORECIDO",
        "REFERENCIA_EMPRESA",
        "TIPO_PAGAMENTO",
        "CONTA_ESCOLHIDA",
        "CONTA_ID_APROVADA",
        "CONTA_NOME_APROVADA",
        "CHAVE_REGRA_BENEFICIARIO",
        "CHAVE_REGRA_REFERENCIA",
        "RISCO_DUPLICIDADE_XML",
        "ESCOPO_REGRA",
        "QTD_OCORRENCIAS",
        "VALOR_TOTAL",
        "EXEMPLOS",
        "OBSERVACAO_OPERADOR",
    ]
    sheet.append(headers)
    sheet.append(
        [
            "APROVADO",
            "CLASSIFICAR",
            "2026-04",
            "08884180000123",
            "2026-04-05",
            20,
            "FORNECEDOR TESTE",
            "-",
            "PIX Transferencias",
            "4.01.012 - Agua e Esgoto",
            "4.01.012",
            "Agua e Esgoto",
            "FORNECEDOR TESTE",
            "",
            "OK",
            "COMPROVANTE",
            99,
            999,
            "old",
            "preservar",
        ]
    )
    workbook.save(path)


def _write_rules(tmp_path: Path) -> Path:
    path = tmp_path / "regras.yaml"
    path.write_text("regras: []\nexclusoes: []\n", encoding="utf-8")
    return path


def _write_plan(tmp_path: Path) -> Path:
    path = tmp_path / "plano.yaml"
    path.write_text(
        """contas:
- id: "4.01.012"
  nome: Agua e Esgoto
""",
        encoding="utf-8",
    )
    return path


def _write_mapping(tmp_path: Path) -> Path:
    path = tmp_path / "mapeamento.yaml"
    path.write_text(
        """mapeamento:
  "4.01.012":
    dre_linha_id: "06.02"
    entra_dre: true
""",
        encoding="utf-8",
    )
    return path
