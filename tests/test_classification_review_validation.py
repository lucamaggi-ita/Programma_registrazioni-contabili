import csv
from pathlib import Path

import pytest

from fantozzi.classification_review_export import CSV_HEADERS
from fantozzi.classification_review_validation import (
    ClassificationReviewValidationError,
    report_text,
    save_report,
    validate_classification_review_csv,
)


def test_missing_csv_raises_clear_error(tmp_path):
    csv_path = tmp_path / "inexistente.csv"
    plano_path = _save_plano_contas(tmp_path)

    with pytest.raises(ClassificationReviewValidationError, match="Arquivo CSV nao encontrado"):
        validate_classification_review_csv(csv_path, plano_path)


def test_empty_csv_raises_clear_error(tmp_path):
    csv_path = tmp_path / "vazio.csv"
    csv_path.write_text("", encoding="utf-8")
    plano_path = _save_plano_contas(tmp_path)

    with pytest.raises(ClassificationReviewValidationError, match="CSV de revisao esta vazio"):
        validate_classification_review_csv(csv_path, plano_path)


def test_csv_without_header_raises_clear_error(tmp_path):
    csv_path = tmp_path / "sem_header.csv"
    csv_path.write_text("\n", encoding="utf-8")
    plano_path = _save_plano_contas(tmp_path)

    with pytest.raises(ClassificationReviewValidationError, match="cabecalho valido"):
        validate_classification_review_csv(csv_path, plano_path)


def test_csv_with_zero_rows_is_valid_report(tmp_path):
    csv_path = tmp_path / "sem_linhas.csv"
    _save_csv(csv_path, [])
    plano_path = _save_plano_contas(tmp_path)

    result = validate_classification_review_csv(csv_path, plano_path)
    text = report_text(result)

    assert result.total_rows == 0
    assert result.valid_rows == 0
    assert result.pending_rows == 0
    assert result.invalid_rows == 0
    assert "Nenhuma linha encontrada no CSV de revisao." in text
    assert "Status: PRONTO PARA PROXIMA FASE" in text


def test_row_without_suggested_account_is_pending(tmp_path):
    csv_path = tmp_path / "pendente.csv"
    _save_csv(csv_path, [_row(acao_manual="CRIAR_ALIAS")])
    plano_path = _save_plano_contas(tmp_path)

    result = validate_classification_review_csv(csv_path, plano_path)

    assert result.pending_rows == 1
    assert result.valid_rows == 0
    assert result.invalid_rows == 0
    assert result.problems[0].problem == "conta_id_sugerida pendente"


def test_row_with_existing_account_and_known_action_is_valid(tmp_path):
    csv_path = tmp_path / "valido.csv"
    _save_csv(
        csv_path,
        [_row(conta_id_sugerida="2.01.004", acao_manual="CRIAR_ALIAS")],
    )
    plano_path = _save_plano_contas(tmp_path)

    result = validate_classification_review_csv(csv_path, plano_path)

    assert result.valid_rows == 1
    assert result.pending_rows == 0
    assert result.invalid_rows == 0
    assert result.status == "PRONTO PARA PROXIMA FASE"


def test_row_with_unknown_account_is_invalid(tmp_path):
    csv_path = tmp_path / "conta_desconhecida.csv"
    _save_csv(
        csv_path,
        [_row(conta_id_sugerida="9.99.999", acao_manual="CRIAR_ALIAS")],
    )
    plano_path = _save_plano_contas(tmp_path)

    result = validate_classification_review_csv(csv_path, plano_path)

    assert result.invalid_rows == 1
    assert result.unknown_accounts == 1
    assert result.problems[0].problem == "conta_id_sugerida desconhecida"


def test_known_manual_action_is_valid(tmp_path):
    csv_path = tmp_path / "acao_valida.csv"
    _save_csv(
        csv_path,
        [_row(conta_id_sugerida="2.01.004", acao_manual="MANTER_REVISAO")],
    )
    plano_path = _save_plano_contas(tmp_path)

    result = validate_classification_review_csv(csv_path, plano_path)

    assert result.valid_rows == 1


def test_unknown_manual_action_is_invalid(tmp_path):
    csv_path = tmp_path / "acao_invalida.csv"
    _save_csv(
        csv_path,
        [_row(conta_id_sugerida="2.01.004", acao_manual="FAZER_SOZINHO")],
    )
    plano_path = _save_plano_contas(tmp_path)

    result = validate_classification_review_csv(csv_path, plano_path)

    assert result.invalid_rows == 1
    assert result.unknown_actions == 1
    assert result.problems[0].problem == "acao_manual desconhecida"


def test_report_contains_summary_and_details(tmp_path):
    csv_path = tmp_path / "problemas.csv"
    _save_csv(csv_path, [_row(conta_id_sugerida="9.99.999", acao_manual="INVALIDA")])
    plano_path = _save_plano_contas(tmp_path)

    result = validate_classification_review_csv(csv_path, plano_path)
    text = report_text(result)

    assert "Total de linhas: 1" in text
    assert "Linhas invalidas: 1" in text
    assert "LINHAS COM PROBLEMAS" in text
    assert "Linha CSV: 2" in text
    assert "Produto/descricao: PRODUTO DESCONHECIDO" in text


def test_missing_required_columns_raise_clear_error(tmp_path):
    csv_path = tmp_path / "colunas_faltando.csv"
    csv_path.write_text("descricao_normalizada;valor_total\nproduto;R$ 10,00\n", encoding="utf-8")
    plano_path = _save_plano_contas(tmp_path)

    with pytest.raises(ClassificationReviewValidationError, match="Colunas obrigatorias ausentes"):
        validate_classification_review_csv(csv_path, plano_path)


def test_report_txt_preserves_utf8_accents(tmp_path):
    report_path = tmp_path / "validacao.txt"

    save_report(report_path, "acao manual: criação de classificação")

    assert report_path.read_text(encoding="utf-8") == "acao manual: criação de classificação"


def _save_csv(csv_path: Path, rows: list[dict[str, str]]) -> None:
    with csv_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=CSV_HEADERS, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


def _row(**overrides: str) -> dict[str, str]:
    row = {
        "tipo_candidato": "CLASSIFICACAO_PRODUTO",
        "descricao_normalizada": "produto desconhecido",
        "descricao_exemplo": "PRODUTO DESCONHECIDO",
        "ocorrencias": "1",
        "fornecedores": "Fornecedor teste",
        "documentos": "NF-1",
        "valor_total": "R$ 10,00",
        "conta_atual_id": "13.01.001",
        "conta_atual_nome": "Revisar - Produto Nao Classificado",
        "status_revisao": "REVISAR",
        "alertas": "CLASSIFICACAO_NAO_ENCONTRADA",
        "acao_manual": "",
        "conta_id_sugerida": "",
        "conta_nome_sugerida": "",
        "palavras_chave_sugeridas": "",
        "observacao_manual": "",
    }
    row.update(overrides)
    return row


def _save_plano_contas(tmp_path: Path) -> Path:
    plano_path = tmp_path / "plano_contas_gerencial.yaml"
    plano_path.write_text(
        """\
contas:
  - id: "2.01.004"
    nome: "Refrigerantes"
    grupo: "Custos de Mercadorias e Insumos"
    tipo: "custo"
  - id: "13.01.001"
    nome: "Revisar - Produto Nao Classificado"
    grupo: "Contas de Revisao e Classificacao Pendente"
    tipo: "revisao"
  - id: "13.01.002"
    nome: "Revisar - Classificacao Ambigua"
    grupo: "Contas de Revisao e Classificacao Pendente"
    tipo: "revisao"
  - id: "13.01.006"
    nome: "Revisar - Valor Inconsistente"
    grupo: "Contas de Revisao e Classificacao Pendente"
    tipo: "revisao"
  - id: "13.01.009"
    nome: "Revisar - Possivel Investimento"
    grupo: "Contas de Revisao e Classificacao Pendente"
    tipo: "revisao"
  - id: "13.01.010"
    nome: "Revisar - Pagamento de Pessoal a Regularizar"
    grupo: "Contas de Revisao e Classificacao Pendente"
    tipo: "revisao"
  - id: "13.01.999"
    nome: "Revisar - Outros Casos"
    grupo: "Contas de Revisao e Classificacao Pendente"
    tipo: "revisao"
""",
        encoding="utf-8",
    )
    return plano_path
