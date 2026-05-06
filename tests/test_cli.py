from pathlib import Path
from shutil import copyfile

from openpyxl import Workbook, load_workbook
import pytest

from fantozzi.cli import build_parser, main
from fantozzi.classification_review_export import CSV_HEADERS
from fantozzi.excel_writer import HEADERS, LANCAMENTOS_SHEET, REVISAO_SHEET


ROOT = Path(__file__).resolve().parents[1]
SAMPLE_XML = ROOT / "samples" / "nfe_exemplo" / "nfe_oesa_12566_20260505.xml"


def test_parser_accepts_importar_nfe_command():
    parser = build_parser()

    args = parser.parse_args(
        ["importar-nfe", "--xml", "sample.xml", "--excel", "output.xlsx"]
    )

    assert args.command == "importar-nfe"
    assert args.xml == "sample.xml"
    assert args.excel == "output.xlsx"


def test_parser_accepts_importar_pasta_nfe_command():
    parser = build_parser()

    args = parser.parse_args(
        ["importar-pasta-nfe", "--pasta", "xmls", "--excel", "output.xlsx"]
    )

    assert args.command == "importar-pasta-nfe"
    assert args.pasta == "xmls"
    assert args.excel == "output.xlsx"


def test_parser_accepts_diagnosticar_excel_command():
    parser = build_parser()

    args = parser.parse_args(
        [
            "diagnosticar-excel",
            "--excel",
            "output.xlsx",
            "--relatorio",
            "reports/diagnostico.txt",
        ]
    )

    assert args.command == "diagnosticar-excel"
    assert args.excel == "output.xlsx"
    assert args.relatorio == "reports/diagnostico.txt"


def test_parser_accepts_relatorio_revisao_command():
    parser = build_parser()

    args = parser.parse_args(
        [
            "relatorio-revisao",
            "--excel",
            "output.xlsx",
            "--saida",
            "reports/revisao.txt",
        ]
    )

    assert args.command == "relatorio-revisao"
    assert args.excel == "output.xlsx"
    assert args.saida == "reports/revisao.txt"


def test_parser_accepts_relatorio_candidatos_classificacao_command():
    parser = build_parser()

    args = parser.parse_args(
        [
            "relatorio-candidatos-classificacao",
            "--excel",
            "output.xlsx",
            "--saida",
            "reports/candidatos.txt",
        ]
    )

    assert args.command == "relatorio-candidatos-classificacao"
    assert args.excel == "output.xlsx"
    assert args.saida == "reports/candidatos.txt"


def test_parser_accepts_exportar_revisao_classificacao_command():
    parser = build_parser()

    args = parser.parse_args(
        [
            "exportar-revisao-classificacao",
            "--excel",
            "output.xlsx",
            "--saida",
            "reports/revisao_classificacao.csv",
        ]
    )

    assert args.command == "exportar-revisao-classificacao"
    assert args.excel == "output.xlsx"
    assert args.saida == "reports/revisao_classificacao.csv"


def test_parser_accepts_validar_revisao_classificacao_command():
    parser = build_parser()

    args = parser.parse_args(
        [
            "validar-revisao-classificacao",
            "--csv",
            "reports/revisao_classificacao.csv",
            "--relatorio",
            "reports/validacao_revisao_classificacao.txt",
        ]
    )

    assert args.command == "validar-revisao-classificacao"
    assert args.csv == "reports/revisao_classificacao.csv"
    assert args.relatorio == "reports/validacao_revisao_classificacao.txt"


def test_general_help_returns_zero_and_lists_commands(capsys):
    with pytest.raises(SystemExit) as exc_info:
        main(["--help"])

    captured = capsys.readouterr()
    assert exc_info.value.code == 0
    assert "Fantozzi" in captured.out
    assert "importar-nfe" in captured.out
    assert "importar-pasta-nfe" in captured.out
    assert "diagnosticar-excel" in captured.out
    assert "relatorio-revisao" in captured.out
    assert "relatorio-candidatos-classificacao" in captured.out
    assert "exportar-revisao-classificacao" in captured.out
    assert "validar-revisao-classificacao" in captured.out


def test_importar_nfe_help_returns_zero_and_lists_options(capsys):
    with pytest.raises(SystemExit) as exc_info:
        main(["importar-nfe", "--help"])

    captured = capsys.readouterr()
    assert exc_info.value.code == 0
    assert "--xml" in captured.out
    assert "--excel" in captured.out
    assert "--plano-contas" in captured.out
    assert "--regras-produtos" in captured.out


def test_importar_pasta_nfe_help_returns_zero_and_lists_options(capsys):
    with pytest.raises(SystemExit) as exc_info:
        main(["importar-pasta-nfe", "--help"])

    captured = capsys.readouterr()
    assert exc_info.value.code == 0
    assert "--pasta" in captured.out
    assert "--excel" in captured.out
    assert "--plano-contas" in captured.out
    assert "--regras-produtos" in captured.out


def test_diagnosticar_excel_help_returns_zero_and_lists_options(capsys):
    with pytest.raises(SystemExit) as exc_info:
        main(["diagnosticar-excel", "--help"])

    captured = capsys.readouterr()
    assert exc_info.value.code == 0
    assert "--excel" in captured.out
    assert "--relatorio" in captured.out


def test_relatorio_revisao_help_returns_zero_and_lists_options(capsys):
    with pytest.raises(SystemExit) as exc_info:
        main(["relatorio-revisao", "--help"])

    captured = capsys.readouterr()
    assert exc_info.value.code == 0
    assert "--excel" in captured.out
    assert "--saida" in captured.out


def test_relatorio_candidatos_classificacao_help_returns_zero_and_lists_options(capsys):
    with pytest.raises(SystemExit) as exc_info:
        main(["relatorio-candidatos-classificacao", "--help"])

    captured = capsys.readouterr()
    assert exc_info.value.code == 0
    assert "--excel" in captured.out
    assert "--saida" in captured.out


def test_exportar_revisao_classificacao_help_returns_zero_and_lists_options(capsys):
    with pytest.raises(SystemExit) as exc_info:
        main(["exportar-revisao-classificacao", "--help"])

    captured = capsys.readouterr()
    assert exc_info.value.code == 0
    assert "--excel" in captured.out
    assert "--saida" in captured.out


def test_validar_revisao_classificacao_help_returns_zero_and_lists_options(capsys):
    with pytest.raises(SystemExit) as exc_info:
        main(["validar-revisao-classificacao", "--help"])

    captured = capsys.readouterr()
    assert exc_info.value.code == 0
    assert "--csv" in captured.out
    assert "--relatorio" in captured.out


def test_importar_nfe_command_writes_excel_with_sample_xml(tmp_path):
    excel_path = tmp_path / "lancamentos_cli.xlsx"

    result = main(
        ["importar-nfe", "--xml", str(SAMPLE_XML), "--excel", str(excel_path)]
    )

    assert result == 0
    assert excel_path.exists()
    workbook = load_workbook(excel_path)
    assert "Lancamentos" in workbook.sheetnames
    assert workbook["Lancamentos"].max_row >= 2


def test_importar_nfe_command_is_idempotent(tmp_path):
    excel_path = tmp_path / "lancamentos_cli.xlsx"
    argv = ["importar-nfe", "--xml", str(SAMPLE_XML), "--excel", str(excel_path)]

    first_result = main(argv)
    first_rows = load_workbook(excel_path)["Lancamentos"].max_row
    second_result = main(argv)
    second_rows = load_workbook(excel_path)["Lancamentos"].max_row

    assert first_result == 0
    assert second_result == 0
    assert second_rows == first_rows


def test_importar_nfe_command_returns_error_for_missing_xml(tmp_path):
    excel_path = tmp_path / "missing.xlsx"
    missing_xml = tmp_path / "arquivo_inexistente.xml"

    result = main(
        ["importar-nfe", "--xml", str(missing_xml), "--excel", str(excel_path)]
    )

    assert result != 0
    assert not excel_path.exists()


def test_importar_pasta_nfe_command_writes_single_excel(tmp_path):
    xml_dir = _xml_dir_with_sample(tmp_path)
    excel_path = tmp_path / "lancamentos_pasta.xlsx"

    result = main(
        ["importar-pasta-nfe", "--pasta", str(xml_dir), "--excel", str(excel_path)]
    )

    assert result == 0
    assert excel_path.exists()
    workbook = load_workbook(excel_path)
    assert "Lancamentos" in workbook.sheetnames
    assert workbook["Lancamentos"].max_row >= 2


def test_importar_pasta_nfe_command_is_idempotent(tmp_path):
    xml_dir = _xml_dir_with_sample(tmp_path)
    excel_path = tmp_path / "lancamentos_pasta.xlsx"
    argv = ["importar-pasta-nfe", "--pasta", str(xml_dir), "--excel", str(excel_path)]

    first_result = main(argv)
    first_rows = load_workbook(excel_path)["Lancamentos"].max_row
    second_result = main(argv)
    second_rows = load_workbook(excel_path)["Lancamentos"].max_row

    assert first_result == 0
    assert second_result == 0
    assert second_rows == first_rows


def test_importar_pasta_nfe_command_returns_error_for_missing_folder(tmp_path):
    excel_path = tmp_path / "missing_folder.xlsx"
    missing_folder = tmp_path / "pasta_inexistente"

    result = main(
        [
            "importar-pasta-nfe",
            "--pasta",
            str(missing_folder),
            "--excel",
            str(excel_path),
        ]
    )

    assert result != 0
    assert not excel_path.exists()


def test_importar_pasta_nfe_command_returns_error_when_folder_has_no_xml(tmp_path):
    xml_dir = tmp_path / "sem_xml"
    xml_dir.mkdir()
    (xml_dir / "nota.txt").write_text("nao e xml", encoding="utf-8")
    excel_path = tmp_path / "sem_xml.xlsx"

    result = main(
        ["importar-pasta-nfe", "--pasta", str(xml_dir), "--excel", str(excel_path)]
    )

    assert result != 0
    assert not excel_path.exists()


def test_importar_pasta_nfe_command_ignores_invalid_xml_and_imports_valid(
    tmp_path, capsys
):
    xml_dir = _xml_dir_with_sample(tmp_path, filename="valid_sample.XML")
    (xml_dir / "invalid.xml").write_text("<root></root>", encoding="utf-8")
    excel_path = tmp_path / "mixed_xml.xlsx"

    result = main(
        ["importar-pasta-nfe", "--pasta", str(xml_dir), "--excel", str(excel_path)]
    )

    assert result == 0
    assert excel_path.exists()
    assert load_workbook(excel_path)["Lancamentos"].max_row >= 2
    captured = capsys.readouterr()
    assert "XML ignorado" in captured.out


def test_relatorio_revisao_command_with_saida_returns_success(tmp_path):
    excel_path = tmp_path / "lancamentos.xlsx"
    report_path = tmp_path / "reports" / "revisao.txt"
    _save_revision_workbook(excel_path)

    result = main(
        [
            "relatorio-revisao",
            "--excel",
            str(excel_path),
            "--saida",
            str(report_path),
        ]
    )

    assert result == 0
    assert report_path.exists()


def test_relatorio_revisao_command_returns_error_for_missing_excel(tmp_path):
    missing_excel = tmp_path / "inexistente.xlsx"

    result = main(["relatorio-revisao", "--excel", str(missing_excel)])

    assert result == 1


def test_relatorio_candidatos_command_with_saida_returns_success(tmp_path):
    excel_path = tmp_path / "lancamentos.xlsx"
    report_path = tmp_path / "reports" / "candidatos.txt"
    _save_candidates_workbook(excel_path)

    result = main(
        [
            "relatorio-candidatos-classificacao",
            "--excel",
            str(excel_path),
            "--saida",
            str(report_path),
        ]
    )

    assert result == 0
    assert report_path.exists()


def test_relatorio_candidatos_command_without_saida_prints_report(tmp_path, capsys):
    excel_path = tmp_path / "lancamentos.xlsx"
    _save_candidates_workbook(excel_path)

    result = main(["relatorio-candidatos-classificacao", "--excel", str(excel_path)])

    captured = capsys.readouterr()
    assert result == 0
    assert "RELATORIO DE CANDIDATOS PARA CLASSIFICACAO" in captured.out
    assert "PRODUTO DESCONHECIDO" in captured.out


def test_relatorio_candidatos_command_returns_error_for_missing_excel(tmp_path):
    missing_excel = tmp_path / "inexistente.xlsx"

    result = main(
        ["relatorio-candidatos-classificacao", "--excel", str(missing_excel)]
    )

    assert result == 1


def test_exportar_revisao_classificacao_command_returns_success(tmp_path):
    excel_path = tmp_path / "lancamentos.xlsx"
    csv_path = tmp_path / "reports" / "revisao_classificacao.csv"
    _save_candidates_workbook(excel_path)

    result = main(
        [
            "exportar-revisao-classificacao",
            "--excel",
            str(excel_path),
            "--saida",
            str(csv_path),
        ]
    )

    assert result == 0
    assert csv_path.exists()


def test_exportar_revisao_classificacao_command_returns_error_for_missing_excel(
    tmp_path,
):
    missing_excel = tmp_path / "inexistente.xlsx"
    csv_path = tmp_path / "reports" / "revisao_classificacao.csv"

    result = main(
        [
            "exportar-revisao-classificacao",
            "--excel",
            str(missing_excel),
            "--saida",
            str(csv_path),
        ]
    )

    assert result == 1
    assert not csv_path.exists()


def test_validar_revisao_classificacao_command_returns_success(tmp_path, capsys):
    csv_path = tmp_path / "revisao_classificacao.csv"
    _save_review_csv(csv_path)

    result = main(["validar-revisao-classificacao", "--csv", str(csv_path)])

    captured = capsys.readouterr()
    assert result == 0
    assert "VALIDACAO DA REVISAO MANUAL DE CLASSIFICACAO" in captured.out
    assert "Total de linhas: 1" in captured.out


def test_validar_revisao_classificacao_command_saves_report(tmp_path):
    csv_path = tmp_path / "revisao_classificacao.csv"
    report_path = tmp_path / "reports" / "validacao.txt"
    _save_review_csv(csv_path)

    result = main(
        [
            "validar-revisao-classificacao",
            "--csv",
            str(csv_path),
            "--relatorio",
            str(report_path),
        ]
    )

    assert result == 0
    assert report_path.exists()


def test_validar_revisao_classificacao_command_returns_error_for_missing_csv(
    tmp_path,
):
    csv_path = tmp_path / "inexistente.csv"

    result = main(["validar-revisao-classificacao", "--csv", str(csv_path)])

    assert result == 1


def _xml_dir_with_sample(tmp_path: Path, *, filename: str = "sample.xml") -> Path:
    xml_dir = tmp_path / "nfe_xml"
    xml_dir.mkdir()
    copyfile(SAMPLE_XML, xml_dir / filename)
    return xml_dir


def _save_revision_workbook(excel_path: Path) -> None:
    workbook = Workbook()
    lancamentos_sheet = workbook.active
    lancamentos_sheet.title = LANCAMENTOS_SHEET
    revisao_sheet = workbook.create_sheet(REVISAO_SHEET)
    lancamentos_sheet.append(HEADERS)
    revisao_sheet.append(HEADERS)
    lancamentos_sheet.append(
        [
            {
                "data_documento": "2026-05-05",
                "fornecedor_ou_favorecido": "Fornecedor teste",
                "documento": "NF-1",
                "conta_id": "2.01.004",
                "conta_nome": "Refrigerantes",
                "descricao_itens": "COCA COLA",
                "valor": 10,
                "origem": "XML_NFE",
                "chave_origem": "chave-1",
                "status_revisao": "OK",
                "hash_registro": "hash-1",
                "duplicatas_json": "[]",
            }.get(header)
            for header in HEADERS
        ]
    )
    workbook.save(excel_path)


def _save_candidates_workbook(excel_path: Path) -> None:
    workbook = Workbook()
    lancamentos_sheet = workbook.active
    lancamentos_sheet.title = LANCAMENTOS_SHEET
    revisao_sheet = workbook.create_sheet(REVISAO_SHEET)
    lancamentos_sheet.append(HEADERS)
    revisao_sheet.append(HEADERS)
    revisao_sheet.append(
        [
            {
                "data_documento": "2026-05-05",
                "fornecedor_ou_favorecido": "Fornecedor teste",
                "documento": "NF-1",
                "conta_id": "13.01.001",
                "conta_nome": "Revisar - Produto Nao Classificado",
                "descricao_itens": "PRODUTO DESCONHECIDO",
                "valor": 10,
                "origem": "XML_NFE",
                "chave_origem": "chave-1",
                "status_revisao": "REVISAR",
                "hash_registro": "hash-1",
                "alertas": "CLASSIFICACAO_NAO_ENCONTRADA",
                "duplicatas_json": "[]",
            }.get(header)
            for header in HEADERS
        ]
    )
    workbook.save(excel_path)


def _save_review_csv(csv_path: Path) -> None:
    csv_path.write_text(
        ";".join(CSV_HEADERS)
        + "\n"
        + ";".join(
            [
                "CLASSIFICACAO_PRODUTO",
                "produto desconhecido",
                "PRODUTO DESCONHECIDO",
                "1",
                "Fornecedor teste",
                "NF-1",
                "R$ 10,00",
                "13.01.001",
                "Revisar - Produto Nao Classificado",
                "REVISAR",
                "CLASSIFICACAO_NAO_ENCONTRADA",
                "CRIAR_ALIAS",
                "2.01.004",
                "Refrigerantes",
                "produto desconhecido",
                "",
            ]
        )
        + "\n",
        encoding="utf-8-sig",
    )
