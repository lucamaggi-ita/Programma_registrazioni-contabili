import csv
import json
import shutil
from pathlib import Path

from fantozzi.cli import build_parser, main
from fantozzi.nfstock_ingest import (
    STATUS_COPIADO,
    STATUS_COPIARIA,
    STATUS_DUPLICADO,
    STATUS_JA_EXISTE,
    STATUS_IGNORADO_CNPJ_TOMADOR_NAO_SUPORTADO,
    STATUS_IGNORADO_XML_NAO_SUPORTADO,
    ingest_nfstock_xml,
)


EMPRESA_CNPJ = "08884180000123"
OUTRA_EMPRESA_CNPJ = "55746722000181"
FORNECEDOR_CNPJ = "42870341000168"
NFSE_FIXTURES = Path(__file__).parent / "fixtures" / "nfse"


def _nfe_xml(
    *,
    chave_suffix: str,
    cnpj_emitente: str = FORNECEDOR_CNPJ,
    cnpj_destinatario: str = EMPRESA_CNPJ,
    numero: str = "123",
    data_emissao: str = "2026-04-10",
    valor: str = "1000.00",
) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<nfeProc versao="4.00" xmlns="http://www.portalfiscal.inf.br/nfe">
  <NFe xmlns="http://www.portalfiscal.inf.br/nfe">
    <infNFe versao="4.00" Id="NFe{chave_suffix}">
      <ide>
        <cUF>29</cUF>
        <natOp>VENDA</natOp>
        <serie>1</serie>
        <nNF>{numero}</nNF>
        <dhEmi>{data_emissao}T10:00:00-03:00</dhEmi>
      </ide>
      <emit>
        <CNPJ>{cnpj_emitente}</CNPJ>
        <xNome>FORNECEDOR TESTE LTDA</xNome>
      </emit>
      <dest>
        <CNPJ>{cnpj_destinatario}</CNPJ>
        <xNome>EMPRESA TESTE</xNome>
      </dest>
      <det nItem="1">
        <prod>
          <cProd>001</cProd>
          <xProd>ITEM TESTE</xProd>
          <NCM>00000000</NCM>
          <CFOP>5102</CFOP>
          <uCom>UN</uCom>
          <qCom>1.0000</qCom>
          <vUnCom>{valor}</vUnCom>
          <vProd>{valor}</vProd>
        </prod>
      </det>
      <total>
        <ICMSTot>
          <vProd>{valor}</vProd>
          <vNF>{valor}</vNF>
        </ICMSTot>
      </total>
    </infNFe>
  </NFe>
</nfeProc>
"""


def _write_nfe(path: Path, **kwargs) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(_nfe_xml(**kwargs), encoding="utf-8")
    return path


def _copy_nfse(path: Path, fixture_name: str = "17755939000169202610000000111CFEQ5H12-nfse.xml") -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(NFSE_FIXTURES / fixture_name, path)
    return path


def _statuses(result):
    return [row.status for row in result.rows]


def test_dry_run_non_copia_file(tmp_path):
    source = tmp_path / "source"
    xml_base = tmp_path / "xml_base"
    output = tmp_path / "out"
    xml_base.mkdir()
    _write_nfe(source / "nfe.xml", chave_suffix="1" * 44)

    result = ingest_nfstock_xml(source, xml_base, output_dir=output)

    dest = xml_base / EMPRESA_CNPJ / "xml" / "nfe.xml"
    assert _statuses(result) == [STATUS_COPIARIA]
    assert not dest.exists()
    assert result.csv_path.is_file()
    assert result.json_path.is_file()


def test_apply_copia_file_na_destinacao_correta(tmp_path):
    source = tmp_path / "source"
    xml_base = tmp_path / "xml_base"
    output = tmp_path / "out"
    xml_base.mkdir()
    _write_nfe(source / "nfe.xml", chave_suffix="2" * 44)

    result = ingest_nfstock_xml(source, xml_base, output_dir=output, apply=True)

    dest = xml_base / EMPRESA_CNPJ / "xml" / "nfe.xml"
    assert _statuses(result) == [STATUS_COPIADO]
    assert dest.is_file()
    assert result.rows[0].dest_path == str(dest)


def test_nfe_destinada_a_cnpj_xml(tmp_path):
    source = tmp_path / "source"
    xml_base = tmp_path / "xml_base"
    xml_base.mkdir()
    _write_nfe(
        source / "nfe_trattoria.xml",
        chave_suffix="3" * 44,
        cnpj_destinatario=OUTRA_EMPRESA_CNPJ,
    )

    result = ingest_nfstock_xml(source, xml_base, output_dir=tmp_path / "out")

    assert result.rows[0].tipo_documento == "NFE"
    assert result.rows[0].dest_path.endswith(
        f"{OUTRA_EMPRESA_CNPJ}\\xml\\nfe_trattoria.xml"
    ) or result.rows[0].dest_path.endswith(
        f"{OUTRA_EMPRESA_CNPJ}/xml/nfe_trattoria.xml"
    )


def test_nfse_destinada_a_cnpj_xml(tmp_path):
    source = tmp_path / "source"
    xml_base = tmp_path / "xml_base"
    xml_base.mkdir()
    _copy_nfse(source / "nfse.xml")

    result = ingest_nfstock_xml(source, xml_base, output_dir=tmp_path / "out")

    assert result.rows[0].tipo_documento == "NFSE"
    assert result.rows[0].cnpj_tomador == EMPRESA_CNPJ
    assert result.rows[0].status == STATUS_COPIARIA
    assert result.rows[0].dest_path.endswith(f"{EMPRESA_CNPJ}\\xml\\nfse.xml") or result.rows[
        0
    ].dest_path.endswith(f"{EMPRESA_CNPJ}/xml/nfse.xml")


def test_duplicado_non_viene_copiato(tmp_path):
    source = tmp_path / "source"
    xml_base = tmp_path / "xml_base"
    output = tmp_path / "out"
    xml_base.mkdir()
    xml_text = _nfe_xml(chave_suffix="4" * 44)
    (source / "a.xml").parent.mkdir(parents=True)
    (source / "a.xml").write_text(xml_text, encoding="utf-8")
    (source / "b.xml").write_text(xml_text, encoding="utf-8")

    result = ingest_nfstock_xml(source, xml_base, output_dir=output, apply=True)

    assert _statuses(result) == [STATUS_COPIADO, STATUS_DUPLICADO]
    assert (xml_base / EMPRESA_CNPJ / "xml" / "a.xml").is_file()
    assert not (xml_base / EMPRESA_CNPJ / "xml" / "b.xml").exists()


def test_duplicado_em_destino_com_nome_diferente_non_viene_copiato(tmp_path):
    source = tmp_path / "source"
    xml_base = tmp_path / "xml_base"
    existing = xml_base / EMPRESA_CNPJ / "xml" / "ja_importado.xml"
    xml_text = _nfe_xml(chave_suffix="a" * 44)
    source.mkdir(parents=True)
    source.joinpath("novo_nome.xml").write_text(xml_text, encoding="utf-8")
    existing.parent.mkdir(parents=True)
    existing.write_text(xml_text, encoding="utf-8")

    result = ingest_nfstock_xml(source, xml_base, output_dir=tmp_path / "out", apply=True)

    assert _statuses(result) == [STATUS_DUPLICADO]
    assert existing.is_file()
    assert not (xml_base / EMPRESA_CNPJ / "xml" / "novo_nome.xml").exists()


def test_file_destinazione_esistente_non_viene_sovrascritto(tmp_path):
    source = tmp_path / "source"
    xml_base = tmp_path / "xml_base"
    dest = xml_base / EMPRESA_CNPJ / "xml" / "nfe.xml"
    _write_nfe(source / "nfe.xml", chave_suffix="9" * 44)
    dest.parent.mkdir(parents=True)
    dest.write_text("gia presente", encoding="utf-8")

    result = ingest_nfstock_xml(source, xml_base, output_dir=tmp_path / "out", apply=True)

    assert _statuses(result) == [STATUS_JA_EXISTE]
    assert dest.read_text(encoding="utf-8") == "gia presente"


def test_cnpj_tomador_nao_suportado_viene_ignorado(tmp_path):
    source = tmp_path / "source"
    xml_base = tmp_path / "xml_base"
    xml_base.mkdir()
    _write_nfe(
        source / "fora.xml",
        chave_suffix="5" * 44,
        cnpj_destinatario="99999999000199",
    )

    result = ingest_nfstock_xml(source, xml_base, output_dir=tmp_path / "out")

    assert _statuses(result) == [STATUS_IGNORADO_CNPJ_TOMADOR_NAO_SUPORTADO]


def test_xml_nao_suportado_nao_blocca_tudo(tmp_path):
    source = tmp_path / "source"
    xml_base = tmp_path / "xml_base"
    xml_base.mkdir()
    _write_nfe(source / "ok.xml", chave_suffix="6" * 44)
    (source / "evento.xml").write_text("<evento><id>1</id></evento>", encoding="utf-8")

    result = ingest_nfstock_xml(source, xml_base, output_dir=tmp_path / "out")

    assert sorted(_statuses(result)) == sorted(
        [STATUS_COPIARIA, STATUS_IGNORADO_XML_NAO_SUPORTADO]
    )


def test_manifest_contem_status_corretos(tmp_path):
    source = tmp_path / "source"
    xml_base = tmp_path / "xml_base"
    output = tmp_path / "out"
    xml_base.mkdir()
    _write_nfe(source / "nfe.xml", chave_suffix="7" * 44)
    (source / "outro.xml").write_text("<outro />", encoding="utf-8")

    result = ingest_nfstock_xml(source, xml_base, output_dir=output)

    with result.csv_path.open(encoding="utf-8", newline="") as file:
        csv_rows = list(csv.DictReader(file, delimiter=";"))
    json_payload = json.loads(result.json_path.read_text(encoding="utf-8"))

    assert {row["status"] for row in csv_rows} == {
        STATUS_COPIARIA,
        STATUS_IGNORADO_XML_NAO_SUPORTADO,
    }
    assert json_payload["totais"][STATUS_COPIARIA] == 1
    assert json_payload["totais"][STATUS_IGNORADO_XML_NAO_SUPORTADO] == 1


def test_cli_has_ingestar_nfstock_xml_command():
    parser = build_parser()
    args = parser.parse_args(
        [
            "ingestar-nfstock-xml",
            "--source-dir",
            "source",
            "--xml-base",
            "xml_base",
        ]
    )
    assert args.command == "ingestar-nfstock-xml"
    assert args.apply is False


def test_cli_dry_run_gera_manifest(tmp_path, capsys):
    source = tmp_path / "source"
    xml_base = tmp_path / "xml_base"
    output = tmp_path / "out"
    xml_base.mkdir()
    _write_nfe(source / "nfe.xml", chave_suffix="8" * 44)

    exit_code = main(
        [
            "ingestar-nfstock-xml",
            "--source-dir",
            str(source),
            "--xml-base",
            str(xml_base),
            "--output-dir",
            str(output),
        ]
    )

    assert exit_code == 0
    assert (output / "manifest_nfstock_ingest.csv").is_file()
    assert (output / "manifest_nfstock_ingest.json").is_file()
    assert not (xml_base / EMPRESA_CNPJ / "xml" / "nfe.xml").exists()
    captured = capsys.readouterr()
    assert "[DRY-RUN] Nenhum XML foi copiado." in captured.out
