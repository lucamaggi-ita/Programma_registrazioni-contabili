import csv
from pathlib import Path

from openpyxl import Workbook
import pytest

from fantozzi.classification_review_export import (
    CSV_HEADERS,
    ClassificationReviewExportError,
    export_classification_review_csv,
)
from fantozzi.excel_writer import HEADERS, LANCAMENTOS_SHEET, REVISAO_SHEET


def test_export_csv_with_candidates(tmp_path):
    excel_path = tmp_path / "lancamentos.xlsx"
    csv_path = tmp_path / "revisao_classificacao.csv"
    _save_workbook(excel_path, revisao_rows=[_row()])

    result = export_classification_review_csv(excel_path, csv_path)
    rows = _read_csv(csv_path)

    assert result.exported_candidates == 1
    assert csv_path.exists()
    assert rows[0]["tipo_candidato"] == "CLASSIFICACAO_PRODUTO"
    assert rows[0]["descricao_normalizada"] == "produto desconhecido"
    assert rows[0]["descricao_exemplo"] == "PRODUTO DESCONHECIDO"
    assert rows[0]["ocorrencias"] == "1"
    assert rows[0]["fornecedores"] == "Fornecedor teste"
    assert rows[0]["documentos"] == "NF-1"
    assert rows[0]["valor_total"] == "R$ 10,00"
    assert rows[0]["conta_atual_id"] == "13.01.001"
    assert rows[0]["conta_atual_nome"] == "Revisar - Produto Nao Classificado"
    assert rows[0]["status_revisao"] == "REVISAR"
    assert rows[0]["alertas"] == "CLASSIFICACAO_NAO_ENCONTRADA"
    assert rows[0]["cfop"] == "5910"
    assert rows[0]["natureza_operacao"] == "BONIFICACAO"
    assert rows[0]["tipo_operacao_especial"] == "BONIFICACAO_RECEBIDA"


def test_export_csv_uses_semicolon_separator_and_utf8_sig(tmp_path):
    excel_path = tmp_path / "lancamentos.xlsx"
    csv_path = tmp_path / "revisao_classificacao.csv"
    _save_workbook(excel_path, revisao_rows=[_row()])

    export_classification_review_csv(excel_path, csv_path)

    raw = csv_path.read_bytes()
    first_line = csv_path.read_text(encoding="utf-8-sig").splitlines()[0]
    assert raw.startswith(b"\xef\xbb\xbf")
    assert ";" in first_line
    assert "," not in first_line
    assert first_line == ";".join(CSV_HEADERS)


def test_manual_fields_are_empty(tmp_path):
    excel_path = tmp_path / "lancamentos.xlsx"
    csv_path = tmp_path / "revisao_classificacao.csv"
    _save_workbook(excel_path, revisao_rows=[_row()])

    export_classification_review_csv(excel_path, csv_path)
    row = _read_csv(csv_path)[0]

    assert row["acao_manual"] == ""
    assert row["conta_id_sugerida"] == ""
    assert row["conta_nome_sugerida"] == ""
    assert row["palavras_chave_sugeridas"] == ""
    assert row["observacao_manual"] == ""


def test_export_csv_without_candidates_writes_header_only(tmp_path):
    excel_path = tmp_path / "sem_candidatos.xlsx"
    csv_path = tmp_path / "revisao_classificacao.csv"
    _save_workbook(excel_path, lancamento_rows=[_row(status_revisao="OK")])

    result = export_classification_review_csv(excel_path, csv_path)
    text = csv_path.read_text(encoding="utf-8-sig")

    assert result.exported_candidates == 0
    assert text.splitlines() == [";".join(CSV_HEADERS)]


def test_export_creates_missing_output_directory(tmp_path):
    excel_path = tmp_path / "lancamentos.xlsx"
    csv_path = tmp_path / "reports" / "revisao_classificacao.csv"
    _save_workbook(excel_path, revisao_rows=[_row()])

    result = export_classification_review_csv(excel_path, csv_path)

    assert result.exported_candidates == 1
    assert csv_path.exists()


def test_missing_excel_raises_controlled_error(tmp_path):
    missing_excel = tmp_path / "inexistente.xlsx"
    csv_path = tmp_path / "revisao_classificacao.csv"

    with pytest.raises(ClassificationReviewExportError, match="Arquivo Excel nao encontrado"):
        export_classification_review_csv(missing_excel, csv_path)


def _read_csv(csv_path: Path) -> list[dict[str, str]]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file, delimiter=";"))


def _save_workbook(
    excel_path: Path,
    *,
    lancamento_rows: list[list[object]] | None = None,
    revisao_rows: list[list[object]] | None = None,
) -> None:
    workbook = Workbook()
    lancamentos_sheet = workbook.active
    lancamentos_sheet.title = LANCAMENTOS_SHEET
    revisao_sheet = workbook.create_sheet(REVISAO_SHEET)

    lancamentos_sheet.append(HEADERS)
    for row in lancamento_rows or []:
        lancamentos_sheet.append(row)
    revisao_sheet.append(HEADERS)
    for row in revisao_rows or []:
        revisao_sheet.append(row)

    workbook.save(excel_path)


def _row(**overrides: object) -> list[object]:
    values = {
        "data_documento": "2026-05-05",
        "fornecedor_ou_favorecido": "Fornecedor teste",
        "documento": "NF-1",
        "conta_id": "13.01.001",
        "conta_nome": "Revisar - Produto Nao Classificado",
        "descricao_itens": "PRODUTO DESCONHECIDO",
        "valor": 10,
        "origem": "XML_NFE",
        "chave_origem": "chave-1",
        "hash_registro": "hash-1",
        "status_revisao": "REVISAR",
        "alertas": "CLASSIFICACAO_NAO_ENCONTRADA",
        "cfop": "5910",
        "natureza_operacao": "BONIFICACAO",
        "tipo_operacao_especial": "BONIFICACAO_RECEBIDA",
        "confianca_operacao_especial": "FORTE",
        "tratamento_operacao_especial": "FORA_CUSTO_COMPRA",
        "duplicatas_json": "[]",
    }
    values.update(overrides)
    return [values.get(header) for header in HEADERS]
