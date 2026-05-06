import csv
from pathlib import Path

from openpyxl import Workbook
import pytest

from fantozzi.classification_review_export import CSV_HEADERS
from fantozzi.classification_review_workbook_importer import (
    ClassificationReviewWorkbookImportError,
    importar_planilha_revisao_classificacao,
)


def test_importa_xlsx_valido_e_gera_csv(tmp_path):
    xlsx_path = _save_review_workbook(tmp_path)
    plano_path = _save_plano_contas(tmp_path)
    csv_path = tmp_path / "revisao_classificacao_revisada.csv"

    result = importar_planilha_revisao_classificacao(
        xlsx_path, csv_path, plano_path
    )

    assert result.is_ok
    assert result.linhas_lidas == 2
    assert result.linhas_exportadas == 2
    assert csv_path.exists()


def test_csv_gerado_mantem_cabecalhos_tecnicos(tmp_path):
    csv_path = _import_valid_workbook(tmp_path)

    with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.reader(file, delimiter=";")
        headers = next(reader)

    assert headers == CSV_HEADERS


def test_acao_manual_e_normalizada_para_maiusculas(tmp_path):
    xlsx_path = _save_review_workbook(
        tmp_path,
        rows=[
            _row(acao_manual=" classificar ", conta_id_sugerida="1.01.001"),
        ],
    )
    plano_path = _save_plano_contas(tmp_path)
    csv_path = tmp_path / "revisao.csv"

    importar_planilha_revisao_classificacao(xlsx_path, csv_path, plano_path)

    with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
        row = next(csv.DictReader(file, delimiter=";"))
    assert row["acao_manual"] == "CLASSIFICAR"


def test_falha_se_aba_revisao_nao_existe(tmp_path):
    xlsx_path = _save_review_workbook(tmp_path, sheet_name="Outra")
    plano_path = _save_plano_contas(tmp_path)

    with pytest.raises(ClassificationReviewWorkbookImportError) as exc_info:
        importar_planilha_revisao_classificacao(
            xlsx_path, tmp_path / "revisao.csv", plano_path
        )

    assert "Aba obrigatoria nao encontrada: Revisao" in str(exc_info.value)


def test_falha_se_coluna_acao_manual_nao_existe(tmp_path):
    headers = [header for header in CSV_HEADERS if header != "acao_manual"]
    xlsx_path = _save_review_workbook(tmp_path, headers=headers)
    plano_path = _save_plano_contas(tmp_path)

    with pytest.raises(ClassificationReviewWorkbookImportError) as exc_info:
        importar_planilha_revisao_classificacao(
            xlsx_path, tmp_path / "revisao.csv", plano_path
        )

    assert "Coluna obrigatoria ausente na aba Revisao: acao_manual" in str(
        exc_info.value
    )


def test_falha_se_acao_manual_invalida(tmp_path):
    xlsx_path = _save_review_workbook(
        tmp_path, rows=[_row(acao_manual="APROVAR", conta_id_sugerida="1.01.001")]
    )
    plano_path = _save_plano_contas(tmp_path)
    csv_path = tmp_path / "revisao.csv"

    with pytest.raises(ClassificationReviewWorkbookImportError) as exc_info:
        importar_planilha_revisao_classificacao(xlsx_path, csv_path, plano_path)

    assert "Linha 2: acao_manual invalida: APROVAR" in str(exc_info.value)
    assert not csv_path.exists()


def test_falha_se_classificar_sem_conta_id_sugerida(tmp_path):
    xlsx_path = _save_review_workbook(
        tmp_path, rows=[_row(acao_manual="CLASSIFICAR", conta_id_sugerida="")]
    )
    plano_path = _save_plano_contas(tmp_path)

    with pytest.raises(ClassificationReviewWorkbookImportError) as exc_info:
        importar_planilha_revisao_classificacao(
            xlsx_path, tmp_path / "revisao.csv", plano_path
        )

    assert "conta_id_sugerida obrigatoria para acao_manual CLASSIFICAR" in str(
        exc_info.value
    )


def test_falha_se_criar_regra_sem_conta_id_sugerida(tmp_path):
    xlsx_path = _save_review_workbook(
        tmp_path, rows=[_row(acao_manual="CRIAR_REGRA", conta_id_sugerida="")]
    )
    plano_path = _save_plano_contas(tmp_path)

    with pytest.raises(ClassificationReviewWorkbookImportError) as exc_info:
        importar_planilha_revisao_classificacao(
            xlsx_path, tmp_path / "revisao.csv", plano_path
        )

    assert "conta_id_sugerida obrigatoria para acao_manual CRIAR_REGRA" in str(
        exc_info.value
    )


def test_falha_se_conta_id_sugerida_nao_existe_no_plano(tmp_path):
    xlsx_path = _save_review_workbook(
        tmp_path, rows=[_row(acao_manual="CLASSIFICAR", conta_id_sugerida="9.99.999")]
    )
    plano_path = _save_plano_contas(tmp_path)

    with pytest.raises(ClassificationReviewWorkbookImportError) as exc_info:
        importar_planilha_revisao_classificacao(
            xlsx_path, tmp_path / "revisao.csv", plano_path
        )

    assert (
        "Linha 2: conta_id_sugerida nao existe no plano de contas: 9.99.999"
        in str(exc_info.value)
    )


def test_ignora_linhas_totalmente_vazias(tmp_path):
    xlsx_path = _save_review_workbook(
        tmp_path,
        rows=[
            _row(acao_manual="CLASSIFICAR", conta_id_sugerida="1.01.001"),
            {header: "" for header in CSV_HEADERS},
        ],
    )
    plano_path = _save_plano_contas(tmp_path)
    csv_path = tmp_path / "revisao.csv"

    result = importar_planilha_revisao_classificacao(xlsx_path, csv_path, plano_path)

    assert result.linhas_lidas == 1
    with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
        rows = list(csv.DictReader(file, delimiter=";"))
    assert len(rows) == 1


def test_nao_modifica_xlsx_original(tmp_path):
    xlsx_path = _save_review_workbook(tmp_path)
    plano_path = _save_plano_contas(tmp_path)
    original_bytes = xlsx_path.read_bytes()

    importar_planilha_revisao_classificacao(
        xlsx_path, tmp_path / "revisao.csv", plano_path
    )

    assert xlsx_path.read_bytes() == original_bytes


def test_nao_modifica_yaml_original(tmp_path):
    xlsx_path = _save_review_workbook(tmp_path)
    plano_path = _save_plano_contas(tmp_path)
    original_yaml = plano_path.read_text(encoding="utf-8")

    importar_planilha_revisao_classificacao(
        xlsx_path, tmp_path / "revisao.csv", plano_path
    )

    assert plano_path.read_text(encoding="utf-8") == original_yaml


def test_gera_relatorio_txt_quando_fornecido(tmp_path):
    xlsx_path = _save_review_workbook(tmp_path)
    plano_path = _save_plano_contas(tmp_path)
    csv_path = tmp_path / "revisao.csv"
    report_path = tmp_path / "reports" / "validacao.txt"

    result = importar_planilha_revisao_classificacao(
        xlsx_path, csv_path, plano_path, report_path
    )

    assert result.relatorio_path == report_path
    assert report_path.exists()
    text = report_path.read_text(encoding="utf-8")
    assert "VALIDACAO DA PLANILHA ASSISTIDA" in text
    assert "Resultado: OK" in text


def _import_valid_workbook(tmp_path: Path) -> Path:
    xlsx_path = _save_review_workbook(tmp_path)
    plano_path = _save_plano_contas(tmp_path)
    csv_path = tmp_path / "revisao.csv"
    importar_planilha_revisao_classificacao(xlsx_path, csv_path, plano_path)
    return csv_path


def _save_review_workbook(
    tmp_path: Path,
    *,
    rows: list[dict[str, str]] | None = None,
    headers: list[str] | None = None,
    sheet_name: str = "Revisao",
) -> Path:
    xlsx_path = tmp_path / "revisao_classificacao_assistida.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    headers = headers or CSV_HEADERS
    sheet.append(headers)
    rows = rows or [
        _row(acao_manual="CLASSIFICAR", conta_id_sugerida="1.01.001"),
        _row(
            descricao_normalizada="freezer comodato",
            descricao_exemplo="FREEZER 300L",
            fornecedores="Fornecedor B",
            documentos="999",
            valor_total="0.00",
            alertas="POSSIVEL_COMODATO",
            acao_manual="ANALISAR_COM_CONTADOR",
            conta_id_sugerida="",
            observacao_manual="validar com contador",
        ),
    ]
    for row in rows:
        sheet.append([row.get(header, "") for header in headers])
    workbook.save(xlsx_path)
    return xlsx_path


def _row(**overrides: str) -> dict[str, str]:
    row = {
        "tipo_candidato": "CLASSIFICACAO_PRODUTO",
        "descricao_normalizada": "agua mineral",
        "descricao_exemplo": "AGUA MINERAL 500ML",
        "ocorrencias": "3",
        "fornecedores": "Fornecedor A",
        "documentos": "123",
        "valor_total": "45.90",
        "conta_atual_id": "",
        "conta_atual_nome": "",
        "status_revisao": "REVISAR",
        "alertas": "",
        "acao_manual": "CLASSIFICAR",
        "conta_id_sugerida": "1.01.001",
        "conta_nome_sugerida": "Mercadorias para Revenda",
        "palavras_chave_sugeridas": "agua mineral",
        "observacao_manual": "ok",
    }
    row.update(overrides)
    return row


def _save_plano_contas(tmp_path: Path) -> Path:
    plano_path = tmp_path / "plano_contas_gerencial.yaml"
    plano_path.write_text(
        """\
contas:
  - id: "1.01.001"
    nome: "Mercadorias para Revenda"
    grupo: "Compras"
  - id: "1.01.002"
    nome: "Bebidas"
    grupo: "Compras"
""",
        encoding="utf-8",
    )
    return plano_path
