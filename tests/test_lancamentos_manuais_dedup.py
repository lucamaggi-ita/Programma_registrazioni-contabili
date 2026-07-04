from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

from fantozzi.excel_writer import HEADERS, LANCAMENTOS_SHEET
from fantozzi.lancamentos_manuais_dedup import (
    auditar_possiveis_duplicatas_manuais_xml,
)


def test_manual_xml_mesma_data_mesmo_valor_sinaliza(tmp_path):
    manual_xlsx = tmp_path / "manual.xlsx"
    xml_xlsx = tmp_path / "xml.xlsx"
    audit_path = tmp_path / "audit.txt"
    _save_lancamentos_xlsx(
        manual_xlsx,
        [
            _row(
                origem="MANUAL_GERENCIAL",
                tipo_documento="MANUAL_GERENCIAL",
                chave_origem="manual-1:D",
                hash_registro="manual-d",
                tipo_lancamento="MANUAL_GERENCIAL_DEBITO_TOTALIZAVEL",
            ),
            _row(
                origem="MANUAL_GERENCIAL",
                tipo_documento="MANUAL_GERENCIAL",
                chave_origem="manual-1:C",
                hash_registro="manual-c",
                tipo_lancamento="MANUAL_GERENCIAL_CREDITO",
            ),
        ],
    )
    _save_lancamentos_xlsx(xml_xlsx, [_row()])

    result = auditar_possiveis_duplicatas_manuais_xml(
        manual_xlsx,
        [xml_xlsx],
        audit_path,
    )

    assert result.possiveis_duplicatas == 1
    assert result.manuais_lidos == 1
    text = audit_path.read_text(encoding="utf-8")
    assert "POSSIVEL_DUPLICADO_MANUAL_XML" in text
    assert "nao blocante" in text


def test_manual_xml_mesma_data_valor_diferente_nao_sinaliza(tmp_path):
    result = _run_single_case(tmp_path, xml_valor=Decimal("99.99"))

    assert result.possiveis_duplicatas == 0


def test_manual_xml_mesmo_valor_data_diferente_nao_sinaliza(tmp_path):
    result = _run_single_case(tmp_path, xml_data=date(2026, 5, 17))

    assert result.possiveis_duplicatas == 0


def test_multiplos_xml_candidatos_sao_match_ambiguo_sem_bloquear(tmp_path):
    manual_xlsx = tmp_path / "manual.xlsx"
    xml_xlsx = tmp_path / "xml.xlsx"
    audit_path = tmp_path / "audit.txt"
    _save_lancamentos_xlsx(
        manual_xlsx,
        [
            _row(
                origem="MANUAL_GERENCIAL",
                tipo_documento="MANUAL_GERENCIAL",
                chave_origem="manual-1:D",
                tipo_lancamento="MANUAL_GERENCIAL_DEBITO_TOTALIZAVEL",
            )
        ],
    )
    _save_lancamentos_xlsx(
        xml_xlsx,
        [
            _row(documento="100"),
            _row(documento="101", chave_origem="xml-2", hash_registro="xml-hash-2"),
        ],
    )

    result = auditar_possiveis_duplicatas_manuais_xml(
        manual_xlsx,
        [xml_xlsx],
        audit_path,
    )

    assert result.possiveis_duplicatas == 1
    assert len(result.matches[0].xml_candidatos) == 2
    assert "xml_candidatos: 2" in audit_path.read_text(encoding="utf-8")


def test_zero_match_produz_audit_com_conteggio_zero(tmp_path):
    result = _run_single_case(tmp_path, xml_valor=Decimal("99.99"))

    text = result.output_path.read_text(encoding="utf-8")
    assert "possiveis_duplicatas_manuais_xml: 0" in text
    assert "nenhuma possivel duplicata encontrada" in text


def _run_single_case(
    tmp_path: Path,
    *,
    xml_data: date = date(2026, 5, 16),
    xml_valor: Decimal = Decimal("10.50"),
):
    manual_xlsx = tmp_path / "manual.xlsx"
    xml_xlsx = tmp_path / "xml.xlsx"
    audit_path = tmp_path / "audit.txt"
    _save_lancamentos_xlsx(
        manual_xlsx,
        [
            _row(
                origem="MANUAL_GERENCIAL",
                tipo_documento="MANUAL_GERENCIAL",
                chave_origem="manual-1:D",
                tipo_lancamento="MANUAL_GERENCIAL_DEBITO_TOTALIZAVEL",
            )
        ],
    )
    _save_lancamentos_xlsx(xml_xlsx, [_row(data_documento=xml_data, valor=xml_valor)])
    return auditar_possiveis_duplicatas_manuais_xml(
        manual_xlsx,
        [xml_xlsx],
        audit_path,
    )


def _save_lancamentos_xlsx(path: Path, rows: list[dict[str, object]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = LANCAMENTOS_SHEET
    sheet.append(HEADERS)
    for row in rows:
        sheet.append([row.get(header, "") for header in HEADERS])
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def _row(**overrides: object) -> dict[str, object]:
    row: dict[str, object] = {
        "data_documento": date(2026, 5, 16),
        "competencia": "2026-05",
        "cnpj_empresa": "08884180000123",
        "nome_empresa": "AMAMI RESTAURANTE LTDA",
        "fornecedor_ou_favorecido": "FORNECEDOR TESTE",
        "cnpj_cpf_favorecido": "48901288000100",
        "documento": "100",
        "tipo_documento": "XML_NFE",
        "conta_id": "2.02.001",
        "conta_nome": "Carnes Bovinas",
        "grupo_conta": "Custos",
        "descricao_itens": "Despesa teste",
        "valor": Decimal("10.50"),
        "origem": "XML_NFE",
        "arquivo_origem": "nfe.xml",
        "chave_origem": "xml-1",
        "hash_registro": "xml-hash-1",
        "status_revisao": "OK",
        "tipo_lancamento": "custo",
        "duplicatas_json": "[]",
    }
    row.update(overrides)
    return row
