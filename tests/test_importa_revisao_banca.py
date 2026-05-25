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
    _existing_key,
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


def test_importa_revisao_banca_adds_decision_dropdown_validations(
    monkeypatch,
    tmp_path,
):
    bank_raw_dir, cnpj_dir = _make_bank_raw_month(tmp_path)
    (cnpj_dir / "ConsultaPagamentos.xls").write_bytes(b"placeholder")
    (cnpj_dir / "Extrato_Lancamentos.xlsx").write_bytes(b"placeholder")
    output_dir = tmp_path / "bank_raw" / "_decisoes_regras"
    _patch_bank_sources(
        monkeypatch,
        [_receipt("Fornecedor Teste", raw_hash="receipt-1")],
        [_launch("PIX enviado", "Fornecedor Teste", raw_hash="launch-1")],
    )

    importar_revisao_banca(**_import_kwargs(tmp_path, bank_raw_dir, output_dir))

    comprovantes_workbook = load_workbook(output_dir / COMPROVANTES_XLSX, data_only=False)
    comprovantes = comprovantes_workbook[COMPROVANTES_SHEET]
    _assert_dropdown(
        comprovantes,
        "STATUS_DECISAO",
        '"PENDENTE,APROVADO,REVISAR,NAO_CRIAR_REGRA,BLOQUEADO"',
    )
    _assert_dropdown(
        comprovantes,
        "ACAO_REGRA",
        '"CLASSIFICAR,ESCLUDERE,XML,NAO_CRIAR_REGRA,REVISAR,BLOQUEADO"',
    )
    _assert_dropdown(comprovantes, "RISCO_DUPLICIDADE_XML", '"OK,DUVIDA"')
    _assert_dropdown(comprovantes, "ESCOPO_REGRA", '"COMPROVANTE,CONTA_CORRENTE"')
    _assert_dropdown_prefix(comprovantes, "CONTA_ESCOLHIDA", "=plano_contas!")

    conta_workbook = load_workbook(output_dir / CONTA_CORRENTE_XLSX, data_only=False)
    conta = conta_workbook[CONTA_CORRENTE_SHEET]
    conta_headers = [cell.value for cell in conta[1]]
    assert "POSSIVEL_COMPROVANTE_RELACIONADO" in conta_headers
    _assert_dropdown(
        conta,
        "STATUS_DECISAO",
        '"PENDENTE,APROVADO,REVISAR,NAO_CRIAR_REGRA,BLOQUEADO"',
    )
    _assert_dropdown(
        conta,
        "ACAO_REGRA",
        '"CLASSIFICAR,ESCLUDERE,XML,NAO_CRIAR_REGRA,REVISAR,BLOQUEADO"',
    )
    _assert_dropdown(conta, "USAR_SOMENTE_SEM_COMPROVANTE", '"SIM,NAO"')
    _assert_dropdown(conta, "RISCO_DUPLICIDADE_XML", '"OK,DUVIDA"')
    _assert_dropdown(conta, "ESCOPO_REGRA", '"COMPROVANTE,CONTA_CORRENTE"')
    _assert_dropdown_prefix(conta, "CONTA_ESCOLHIDA", "=plano_contas!")


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


def test_importa_revisao_banca_derives_account_fields_from_choice_not_old_formulas(
    monkeypatch,
    tmp_path,
):
    bank_raw_dir = tmp_path / "bank_raw" / "2026" / "2026-04"
    cnpj_dir = bank_raw_dir / "08884180000123"
    cnpj_dir.mkdir(parents=True)
    (cnpj_dir / "ConsultaPagamentos.xls").write_bytes(b"placeholder")
    output_dir = tmp_path / "bank_raw" / "_decisoes_regras"
    output_dir.mkdir(parents=True)
    _write_existing_comprovantes_with_stale_formula(output_dir / COMPROVANTES_XLSX)

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

    importar_revisao_banca(
        bank_raw_dir=bank_raw_dir,
        output_dir=output_dir,
        regras_path=_write_rules(tmp_path),
        plano_contas_path=_write_plan(tmp_path),
        mapeamento_dre_path=_write_mapping(tmp_path),
    )

    workbook = load_workbook(output_dir / COMPROVANTES_XLSX, data_only=False)
    sheet = workbook[COMPROVANTES_SHEET]
    headers = [cell.value for cell in sheet[1]]
    conta_id_col = headers.index("CONTA_ID_APROVADA") + 1
    conta_nome_col = headers.index("CONTA_NOME_APROVADA") + 1

    assert sheet.max_row == 2
    assert sheet.cell(row=2, column=conta_id_col).value == "4.01.012"
    assert sheet.cell(row=2, column=conta_nome_col).value == "Agua e Esgoto"


def test_existing_key_for_comprovantes_without_beneficiary_does_not_fall_through():
    assert (
        _existing_key(
            {
                "BENEFICIARIO_FAVORECIDO": "",
                "LANCAMENTO": "PIX ENVIADO",
                "RAZAO_SOCIAL": "Fornecedor Teste",
                "VALOR": "-20.00",
            },
            "comprovantes",
        )
        == ""
    )


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


def test_importar_revisao_banca_idempotente(monkeypatch, tmp_path):
    bank_raw_dir, cnpj_dir = _make_bank_raw_month(tmp_path)
    (cnpj_dir / "ConsultaPagamentos.xls").write_bytes(b"placeholder")
    (cnpj_dir / "Extrato_Lancamentos.xlsx").write_bytes(b"placeholder")
    output_dir = tmp_path / "bank_raw" / "_decisoes_regras"
    receipt = _receipt("Fornecedor Teste", raw_hash="receipt-1")
    launch = _launch("PIX Fornecedor Teste", "Fornecedor Teste", raw_hash="launch-1")
    _patch_bank_sources(monkeypatch, [receipt], [launch])

    kwargs = _import_kwargs(tmp_path, bank_raw_dir, output_dir)
    importar_revisao_banca(**kwargs)
    first_comprovantes = _sheet_rows(output_dir / COMPROVANTES_XLSX, COMPROVANTES_SHEET)
    first_conta = _sheet_rows(output_dir / CONTA_CORRENTE_XLSX, CONTA_CORRENTE_SHEET)

    importar_revisao_banca(**kwargs)
    second_comprovantes = _sheet_rows(output_dir / COMPROVANTES_XLSX, COMPROVANTES_SHEET)
    second_conta = _sheet_rows(output_dir / CONTA_CORRENTE_XLSX, CONTA_CORRENTE_SHEET)

    assert second_comprovantes == first_comprovantes
    assert second_conta == first_conta


def test_importar_revisao_banca_preserva_decisoes_apos_rerun(monkeypatch, tmp_path):
    bank_raw_dir, cnpj_dir = _make_bank_raw_month(tmp_path)
    (cnpj_dir / "ConsultaPagamentos.xls").write_bytes(b"placeholder")
    (cnpj_dir / "Extrato_Lancamentos.xlsx").write_bytes(b"placeholder")
    output_dir = tmp_path / "bank_raw" / "_decisoes_regras"
    receipt = _receipt("Fornecedor Teste", raw_hash="receipt-1")
    launch = _launch("PIX Fornecedor Teste", "Fornecedor Teste", raw_hash="launch-1")
    _patch_bank_sources(monkeypatch, [receipt], [launch])

    kwargs = _import_kwargs(tmp_path, bank_raw_dir, output_dir)
    importar_revisao_banca(**kwargs)
    _set_cell_by_header(
        output_dir / COMPROVANTES_XLSX,
        COMPROVANTES_SHEET,
        row=2,
        values={
            "STATUS_DECISAO": "APROVADO",
            "ACAO_REGRA": "CLASSIFICAR",
            "CONTA_ESCOLHIDA": "4.01.012 - Agua e Esgoto",
            "OBSERVACAO_OPERADOR": "decisao preservada",
            "RISCO_DUPLICIDADE_XML": "OK",
            "ESCOPO_REGRA": "COMPROVANTE",
        },
    )
    _set_cell_by_header(
        output_dir / CONTA_CORRENTE_XLSX,
        CONTA_CORRENTE_SHEET,
        row=2,
        values={
            "STATUS_DECISAO": "APROVADO",
            "ACAO_REGRA": "CLASSIFICAR",
            "CONTA_ESCOLHIDA": "4.01.012 - Agua e Esgoto",
            "USAR_SOMENTE_SEM_COMPROVANTE": "SIM",
            "OBSERVACAO_OPERADOR": "conta preservada",
        },
    )

    importar_revisao_banca(**kwargs)

    comprovante = _row_by_index(
        output_dir / COMPROVANTES_XLSX,
        COMPROVANTES_SHEET,
        row=2,
    )
    conta = _row_by_index(output_dir / CONTA_CORRENTE_XLSX, CONTA_CORRENTE_SHEET, row=2)
    assert comprovante["STATUS_DECISAO"] == "APROVADO"
    assert comprovante["ACAO_REGRA"] == "CLASSIFICAR"
    assert comprovante["CONTA_ID_APROVADA"] == "4.01.012"
    assert comprovante["CONTA_NOME_APROVADA"] == "Agua e Esgoto"
    assert comprovante["OBSERVACAO_OPERADOR"] == "decisao preservada"
    assert comprovante["RISCO_DUPLICIDADE_XML"] == "OK"
    assert comprovante["ESCOPO_REGRA"] == "COMPROVANTE"
    assert conta["USAR_SOMENTE_SEM_COMPROVANTE"] == "SIM"
    assert conta["OBSERVACAO_OPERADOR"] == "conta preservada"


def test_importar_revisao_banca_aggiunge_nova_chave_sem_perder_decisoes(
    monkeypatch,
    tmp_path,
):
    bank_raw_dir, cnpj_dir = _make_bank_raw_month(tmp_path)
    (cnpj_dir / "ConsultaPagamentos.xls").write_bytes(b"placeholder")
    output_dir = tmp_path / "bank_raw" / "_decisoes_regras"
    receipts = [_receipt("Fornecedor A", raw_hash="receipt-a")]
    _patch_bank_sources(monkeypatch, receipts, [])

    kwargs = _import_kwargs(tmp_path, bank_raw_dir, output_dir)
    importar_revisao_banca(**kwargs)
    _set_cell_by_header(
        output_dir / COMPROVANTES_XLSX,
        COMPROVANTES_SHEET,
        row=2,
        values={"STATUS_DECISAO": "APROVADO", "OBSERVACAO_OPERADOR": "manter"},
    )
    receipts.append(_receipt("Fornecedor B", raw_hash="receipt-b"))

    importar_revisao_banca(**kwargs)

    rows = _data_rows(output_dir / COMPROVANTES_XLSX, COMPROVANTES_SHEET)
    assert len(rows) == 2
    fornecedor_a = next(row for row in rows if row["BENEFICIARIO_FAVORECIDO"] == "FORNECEDOR A")
    fornecedor_b = next(row for row in rows if row["BENEFICIARIO_FAVORECIDO"] == "FORNECEDOR B")
    assert fornecedor_a["STATUS_DECISAO"] == "APROVADO"
    assert fornecedor_a["OBSERVACAO_OPERADOR"] == "manter"
    assert fornecedor_b["STATUS_DECISAO"] in (None, "")


def test_importar_revisao_banca_remove_chave_desaparecida_do_operacional(
    monkeypatch,
    tmp_path,
):
    bank_raw_dir, cnpj_dir = _make_bank_raw_month(tmp_path)
    (cnpj_dir / "ConsultaPagamentos.xls").write_bytes(b"placeholder")
    output_dir = tmp_path / "bank_raw" / "_decisoes_regras"
    receipts = [
        _receipt("Fornecedor A", raw_hash="receipt-a"),
        _receipt("Fornecedor B", raw_hash="receipt-b"),
    ]
    _patch_bank_sources(monkeypatch, receipts, [])

    kwargs = _import_kwargs(tmp_path, bank_raw_dir, output_dir)
    importar_revisao_banca(**kwargs)
    receipts.pop()

    importar_revisao_banca(**kwargs)

    rows = _data_rows(output_dir / COMPROVANTES_XLSX, COMPROVANTES_SHEET)
    assert [row["BENEFICIARIO_FAVORECIDO"] for row in rows] == ["FORNECEDOR A"]


def test_importar_revisao_banca_skip_classificado_por_yaml(monkeypatch, tmp_path):
    bank_raw_dir, cnpj_dir = _make_bank_raw_month(tmp_path)
    (cnpj_dir / "ConsultaPagamentos.xls").write_bytes(b"placeholder")
    output_dir = tmp_path / "bank_raw" / "_decisoes_regras"
    _patch_bank_sources(
        monkeypatch,
        [_receipt("Fornecedor Teste", raw_hash="receipt-1")],
        [],
    )

    result = importar_revisao_banca(
        bank_raw_dir=bank_raw_dir,
        output_dir=output_dir,
        regras_path=_write_rules_with_supplier(tmp_path, "FORNECEDOR TESTE"),
        plano_contas_path=_write_plan(tmp_path),
        mapeamento_dre_path=_write_mapping(tmp_path),
    )

    assert result.comprovantes_candidatos == 0
    data_rows = [
        row
        for row in _data_rows(output_dir / COMPROVANTES_XLSX, COMPROVANTES_SHEET)
        if any(value not in (None, "") for value in row.values())
    ]
    assert data_rows == []


def test_importar_revisao_banca_escreve_manifest_historico_por_mes(
    monkeypatch,
    tmp_path,
):
    bank_raw_dir, cnpj_dir = _make_bank_raw_month(tmp_path)
    (cnpj_dir / "ConsultaPagamentos.xls").write_bytes(b"placeholder")
    output_dir = tmp_path / "bank_raw" / "_decisoes_regras"
    _patch_bank_sources(
        monkeypatch,
        [_receipt("Fornecedor Teste", raw_hash="receipt-1")],
        [],
    )

    result = importar_revisao_banca(**_import_kwargs(tmp_path, bank_raw_dir, output_dir))

    current = output_dir / "manifest_importa_revisao_banca.json"
    historic = output_dir / "manifest_importa_revisao_banca_2026-04.json"
    assert result.manifest_path == current
    assert result.manifest_historico_path == historic
    assert current.exists()
    assert historic.exists()
    assert json.loads(current.read_text(encoding="utf-8")) == json.loads(
        historic.read_text(encoding="utf-8")
    )


def test_importar_revisao_banca_indica_possivel_comprovante_relacionado(
    monkeypatch,
    tmp_path,
):
    bank_raw_dir, cnpj_dir = _make_bank_raw_month(tmp_path)
    (cnpj_dir / "ConsultaPagamentos.xls").write_bytes(b"placeholder")
    (cnpj_dir / "Extrato_Lancamentos.xlsx").write_bytes(b"placeholder")
    output_dir = tmp_path / "bank_raw" / "_decisoes_regras"
    _patch_bank_sources(
        monkeypatch,
        [_receipt("Fornecedor Teste", raw_hash="receipt-1")],
        [_launch("PIX enviado", "Fornecedor Teste", raw_hash="launch-1")],
    )

    result = importar_revisao_banca(**_import_kwargs(tmp_path, bank_raw_dir, output_dir))

    row = _row_by_index(output_dir / CONTA_CORRENTE_XLSX, CONTA_CORRENTE_SHEET, row=2)
    assert row["POSSIVEL_COMPROVANTE_RELACIONADO"] == "SIM"
    assert row["STATUS_DECISAO"] in (None, "")
    assert row["CONTA_ESCOLHIDA"] in (None, "")
    assert result.possiveis_comprovantes_relacionados == 1
    assert result.alertas
    manifest = json.loads((output_dir / "manifest_importa_revisao_banca.json").read_text())
    assert manifest["possiveis_comprovantes_relacionados"] == 1
    assert manifest["alertas"]


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


def _write_existing_comprovantes_with_stale_formula(path: Path) -> None:
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
            "2026-04-04",
            10,
            "OUTRO FORNECEDOR",
            "-",
            "PIX Transferencias",
            "7.01.001 - Tarifas Bancarias",
            '=IFERROR(LEFT(J2,FIND(" - ",J2)-1),"")',
            '=IFERROR(MID(J2,FIND(" - ",J2)+3,999),"")',
            "OUTRO FORNECEDOR",
            "",
            "OK",
            "COMPROVANTE",
            1,
            10,
            "old",
            "",
        ]
    )
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
            '=IFERROR(LEFT(J2,FIND(" - ",J2)-1),"")',
            '=IFERROR(MID(J2,FIND(" - ",J2)+3,999),"")',
            "FORNECEDOR TESTE",
            "",
            "OK",
            "COMPROVANTE",
            1,
            20,
            "old",
            "",
        ]
    )
    workbook.save(path)


def _write_rules(tmp_path: Path) -> Path:
    path = tmp_path / "regras.yaml"
    path.write_text("regras: []\nexclusoes: []\n", encoding="utf-8")
    return path


def _write_rules_with_supplier(tmp_path: Path, supplier: str) -> Path:
    path = tmp_path / "regras.yaml"
    path.write_text(
        f"""regras:
- regra_id: REGRA_FORNECEDOR_TESTE
  conta_id: "4.01.012"
  conta_nome: Agua e Esgoto
  dre_linha_id: "06.02"
  entra_dre: true
  match:
    modo: TODAS_AS_CONDICOES_NAO_VAZIAS
    beneficiario_favorecido_contem_algum:
    - {supplier}
    referencia_empresa_contem_algum: []
exclusoes: []
""",
        encoding="utf-8",
    )
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


def _make_bank_raw_month(tmp_path: Path) -> tuple[Path, Path]:
    bank_raw_dir = tmp_path / "bank_raw" / "2026" / "2026-04"
    cnpj_dir = bank_raw_dir / "08884180000123"
    cnpj_dir.mkdir(parents=True)
    return bank_raw_dir, cnpj_dir


def _receipt(
    beneficiary: str,
    *,
    raw_hash: str,
    reference: str = "-",
    amount: str = "20.00",
) -> PaymentReceiptRaw:
    return PaymentReceiptRaw(
        payment_date=date(2026, 4, 5),
        amount=Decimal(amount),
        beneficiary_name=beneficiary,
        beneficiary_document="12.345.678/0001-90",
        payment_type="PIX Transferencias",
        reference_id=reference,
        status="Efetuado",
        source_file="ConsultaPagamentos.xls",
        raw_hash=raw_hash,
    )


def _launch(
    description: str,
    counterparty: str,
    *,
    raw_hash: str,
    amount: str = "-20.00",
) -> BankLaunchEnrichment:
    return BankLaunchEnrichment(
        transaction_date=date(2026, 4, 5),
        amount=Decimal(amount),
        description=description,
        counterparty_name=counterparty,
        counterparty_document="12.345.678/0001-90",
        row_balance=None,
        source_file="Extrato_Lancamentos.xlsx",
        raw_hash=raw_hash,
    )


def _patch_bank_sources(
    monkeypatch,
    receipts: list[PaymentReceiptRaw],
    launches: list[BankLaunchEnrichment],
) -> None:
    monkeypatch.setattr(
        "fantozzi.banking.revisao_banca.parse_payment_receipts",
        lambda path: list(receipts),
    )
    monkeypatch.setattr(
        "fantozzi.banking.revisao_banca.parse_bank_launches",
        lambda path: (list(launches), []),
    )


def _import_kwargs(tmp_path: Path, bank_raw_dir: Path, output_dir: Path) -> dict[str, Path]:
    return {
        "bank_raw_dir": bank_raw_dir,
        "output_dir": output_dir,
        "regras_path": _write_rules(tmp_path),
        "plano_contas_path": _write_plan(tmp_path),
        "mapeamento_dre_path": _write_mapping(tmp_path),
    }


def _sheet_rows(path: Path, sheet_name: str) -> list[tuple[object, ...]]:
    workbook = load_workbook(path, data_only=False)
    return list(workbook[sheet_name].iter_rows(values_only=True))


def _data_rows(path: Path, sheet_name: str) -> list[dict[str, object]]:
    rows = _sheet_rows(path, sheet_name)
    headers = [str(value or "") for value in rows[0]]
    return [
        {headers[index]: value for index, value in enumerate(row)}
        for row in rows[1:]
    ]


def _row_by_index(path: Path, sheet_name: str, *, row: int) -> dict[str, object]:
    workbook = load_workbook(path, data_only=False)
    sheet = workbook[sheet_name]
    headers = [cell.value for cell in sheet[1]]
    return {
        str(headers[index]): sheet.cell(row=row, column=index + 1).value
        for index in range(len(headers))
    }


def _set_cell_by_header(
    path: Path,
    sheet_name: str,
    *,
    row: int,
    values: dict[str, object],
) -> None:
    workbook = load_workbook(path, data_only=False)
    sheet = workbook[sheet_name]
    headers = [cell.value for cell in sheet[1]]
    for header, value in values.items():
        sheet.cell(row=row, column=headers.index(header) + 1).value = value
    workbook.save(path)


def _assert_dropdown(sheet, header: str, formula1: str, *, row: int = 2) -> None:
    assert _has_dropdown(sheet, header, formula1, row=row)


def _assert_dropdown_prefix(sheet, header: str, formula_prefix: str, *, row: int = 2) -> None:
    assert _has_dropdown_prefix(sheet, header, formula_prefix, row=row)


def _has_dropdown(sheet, header: str, formula1: str, *, row: int = 2) -> bool:
    cell = _cell_ref(sheet, header, row)
    return any(
        validation.type == "list"
        and validation.formula1 == formula1
        and cell in str(validation.sqref)
        for validation in sheet.data_validations.dataValidation
    )


def _has_dropdown_prefix(sheet, header: str, formula_prefix: str, *, row: int = 2) -> bool:
    cell = _cell_ref(sheet, header, row)
    return any(
        validation.type == "list"
        and str(validation.formula1).startswith(formula_prefix)
        and cell in str(validation.sqref)
        for validation in sheet.data_validations.dataValidation
    )


def _cell_ref(sheet, header: str, row: int) -> str:
    headers = [cell.value for cell in sheet[1]]
    column = chr(ord("A") + headers.index(header))
    return f"{column}{row}"


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
