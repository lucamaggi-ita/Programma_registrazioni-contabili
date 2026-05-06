from pathlib import Path
from shutil import copyfile

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

    args = parser.parse_args(["diagnosticar-excel", "--excel", "output.xlsx"])

    assert args.command == "diagnosticar-excel"
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


def _xml_dir_with_sample(tmp_path: Path, *, filename: str = "sample.xml") -> Path:
    xml_dir = tmp_path / "nfe_xml"
    xml_dir.mkdir()
    copyfile(SAMPLE_XML, xml_dir / filename)
    return xml_dir
