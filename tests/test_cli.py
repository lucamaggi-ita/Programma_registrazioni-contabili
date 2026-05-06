from pathlib import Path

from openpyxl import load_workbook

from fantozzi.cli import build_parser, main


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
