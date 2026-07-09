from datetime import date
from decimal import Decimal
from pathlib import Path
import shutil

import pytest

from fantozzi.audit_xml_comprovantes import (
    AuditXmlComprovantesError,
    STATUS_DADOS_INSUFICIENTES,
    STATUS_MATCH_ALTA,
    STATUS_MATCH_MEDIA,
    STATUS_MULTIPLOS_CANDIDATOS,
    STATUS_SEM_XML_NO_DATASET,
    build_resumo_por_fornecedor,
    load_comprovantes,
    run_audit,
    scan_xml_root,
    write_outputs,
)
from fantozzi.cli import build_parser, main

EMITENTE_CNPJ = "42870341000168"
DESTINATARIO_CNPJ = "55746722000181"
OUTRA_EMPRESA_CNPJ = "08884180000123"
NFSE_PRESTADOR_CNPJ = "17755939000169"
NFSE_FIXTURES = Path(__file__).parent / "fixtures" / "nfse"


def _nfe_xml(
    *,
    chave_suffix: str,
    cnpj_emitente: str,
    cnpj_destinatario: str,
    data_emissao: str,
    valor: str,
) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<nfeProc versao="4.00" xmlns="http://www.portalfiscal.inf.br/nfe">
  <NFe xmlns="http://www.portalfiscal.inf.br/nfe">
    <infNFe versao="4.00" Id="NFe{chave_suffix}">
      <ide>
        <cUF>29</cUF>
        <natOp>VENDA DE MERCADORIAS</natOp>
        <serie>001</serie>
        <nNF>1</nNF>
        <dhEmi>{data_emissao}T10:00:00-03:00</dhEmi>
      </ide>
      <emit>
        <CNPJ>{cnpj_emitente}</CNPJ>
        <xNome>HIMEROS COMERCIO DE PESCADOS LTDA</xNome>
      </emit>
      <dest>
        <CNPJ>{cnpj_destinatario}</CNPJ>
        <xNome>EMPRESA DESTINATARIA</xNome>
      </dest>
      <det nItem="1">
        <prod>
          <cProd>001</cProd>
          <xProd>PESCADOS</xProd>
          <NCM>03000000</NCM>
          <CFOP>5102</CFOP>
          <uCom>KG</uCom>
          <qCom>1.0000</qCom>
          <vUnCom>{valor}</vUnCom>
          <vProd>{valor}</vProd>
        </prod>
      </det>
      <total>
        <ICMSTot>
          <vBC>0.00</vBC>
          <vICMS>0.00</vICMS>
          <vProd>{valor}</vProd>
          <vNF>{valor}</vNF>
        </ICMSTot>
      </total>
    </infNFe>
  </NFe>
</nfeProc>
"""


def _write_xml(path: Path, **kwargs) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(_nfe_xml(**kwargs), encoding="utf-8")


CSV_HEADER = (
    "cnpj;mes;data_pagamento;beneficiario_normalizado;referencia;valor;"
    "arquivo_origem;status_classificacao;motivo_revisao;regras_candidatas;"
    "contas_candidatas;DECISAO_OPERADOR;CONTA_ID_APROVADA;CRIAR_REGRA_FUTURA;"
    "ALIAS_O_KEYWORD_APROVADO;OBSERVACOES;payment_id;beneficiary_name;"
    "beneficiary_document;payment_type;status_pagamento;source_file;raw_hash"
)


def _comprovante_row(
    *,
    payment_id: str,
    cnpj_pagador: str,
    beneficiary_document: str,
    beneficiary_name: str = "HIMEROS COMERCIO DE PESCADOS L",
    data_pagamento: str,
    valor: str,
    status_classificacao: str = "REVISAR_SEM_REGRA",
) -> str:
    fields = [
        cnpj_pagador,
        "2026-01",
        data_pagamento,
        beneficiary_name,
        "-",
        valor,
        "arquivo_origem.csv",
        status_classificacao,
        "SEM_REGRA",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        payment_id,
        beneficiary_name,
        beneficiary_document,
        "PIX",
        "Efetuado",
        "ConsultaPagamentos.xls",
        f"hash_{payment_id}",
    ]
    return ";".join(fields)


def _write_comprovantes_csv(path: Path, rows: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join([CSV_HEADER, *rows]) + "\n", encoding="utf-8-sig")


def test_load_comprovantes_parses_semicolon_csv(tmp_path):
    csv_path = tmp_path / "comprovantes_a_revisar.csv"
    _write_comprovantes_csv(
        csv_path,
        [
            _comprovante_row(
                payment_id="p1",
                cnpj_pagador=DESTINATARIO_CNPJ,
                beneficiary_document="42.870.341/0001-68",
                data_pagamento="2026-01-07",
                valor="6857",
            )
        ],
    )
    records = load_comprovantes(csv_path)
    assert len(records) == 1
    rec = records[0]
    assert rec.cnpj_pagador == DESTINATARIO_CNPJ
    assert rec.cnpj_beneficiario == EMITENTE_CNPJ
    assert rec.valor == Decimal("6857")
    assert rec.data_pagamento == date(2026, 1, 7)


def test_load_comprovantes_missing_file_raises(tmp_path):
    with pytest.raises(AuditXmlComprovantesError):
        load_comprovantes(tmp_path / "nao_existe.csv")


def test_scan_xml_root_extracts_cnpj_valor_data(tmp_path):
    xml_root = tmp_path / "xml"
    _write_xml(
        xml_root / DESTINATARIO_CNPJ / "chave1-nfe.xml",
        chave_suffix="1" * 44,
        cnpj_emitente=EMITENTE_CNPJ,
        cnpj_destinatario=DESTINATARIO_CNPJ,
        data_emissao="2026-01-05",
        valor="6857.00",
    )
    records, xml_por_pasta, manifest, ignoradas = scan_xml_root(xml_root)
    assert len(records) == 1
    assert records[0].cnpj_emitente == EMITENTE_CNPJ
    assert records[0].cnpj_destinatario == DESTINATARIO_CNPJ
    assert records[0].valor_total == Decimal("6857.00")
    assert records[0].data_emissao == date(2026, 1, 5)
    assert records[0].tipo_documento == "NFE"
    assert xml_por_pasta[DESTINATARIO_CNPJ] == 1
    assert len(manifest) == 1
    assert ignoradas == 0


def test_scan_xml_root_extracts_nfse_candidate(tmp_path):
    xml_root = tmp_path / "xml"
    destino = xml_root / OUTRA_EMPRESA_CNPJ / "nfse.xml"
    destino.parent.mkdir(parents=True)
    shutil.copyfile(
        NFSE_FIXTURES / "17755939000169202610000000111CFEQ5H12-nfse.xml",
        destino,
    )

    records, xml_por_pasta, manifest, ignoradas = scan_xml_root(xml_root)

    assert len(records) == 1
    assert records[0].tipo_documento == "NFSE"
    assert records[0].cnpj_emitente == NFSE_PRESTADOR_CNPJ
    assert records[0].cnpj_destinatario == OUTRA_EMPRESA_CNPJ
    assert records[0].valor_total == Decimal("1830")
    assert records[0].data_emissao == date(2026, 2, 25)
    assert records[0].chave == "29253031217755939000169261000000011126020070216971"
    assert xml_por_pasta[OUTRA_EMPRESA_CNPJ] == 1
    assert len(manifest) == 1
    assert ignoradas == 0


def test_scan_xml_root_missing_dir_raises(tmp_path):
    with pytest.raises(AuditXmlComprovantesError):
        scan_xml_root(tmp_path / "nao_existe")


def _build_dataset(tmp_path):
    xml_root = tmp_path / "xml"
    _write_xml(
        xml_root / DESTINATARIO_CNPJ / "chave1-nfe.xml",
        chave_suffix="1" * 44,
        cnpj_emitente=EMITENTE_CNPJ,
        cnpj_destinatario=DESTINATARIO_CNPJ,
        data_emissao="2026-01-05",
        valor="6857.00",
    )
    csv_path = tmp_path / "comprovantes_a_revisar.csv"
    rows = [
        _comprovante_row(
            payment_id="match_alta",
            cnpj_pagador=DESTINATARIO_CNPJ,
            beneficiary_document="42.870.341/0001-68",
            data_pagamento="2026-01-07",
            valor="6857",
        ),
        _comprovante_row(
            payment_id="sem_xml",
            cnpj_pagador=OUTRA_EMPRESA_CNPJ,
            beneficiary_document="42.870.341/0001-68",
            data_pagamento="2026-01-15",
            valor="4200",
        ),
    ]
    _write_comprovantes_csv(csv_path, rows)
    return csv_path, xml_root


def test_run_audit_match_alta_and_sem_xml(tmp_path):
    csv_path, xml_root = _build_dataset(tmp_path)
    result = run_audit(csv_path, xml_root, dias_tolerancia=7)

    by_id = {row.comprovante_id: row for row in result.rows}
    assert by_id["match_alta"].status_audit == STATUS_MATCH_ALTA
    assert by_id["match_alta"].xml_chave.endswith("1" * 44)
    assert by_id["match_alta"].tipo_documento == "NFE"
    assert by_id["sem_xml"].status_audit == STATUS_SEM_XML_NO_DATASET
    assert "Nenhuma XML encontrada" in by_id["sem_xml"].observacao


def test_run_audit_match_alta_with_nfse_candidate(tmp_path):
    xml_root = tmp_path / "xml"
    destino = xml_root / OUTRA_EMPRESA_CNPJ / "nfse.xml"
    destino.parent.mkdir(parents=True)
    shutil.copyfile(
        NFSE_FIXTURES / "17755939000169202610000000111CFEQ5H12-nfse.xml",
        destino,
    )
    csv_path = tmp_path / "comprovantes_a_revisar.csv"
    _write_comprovantes_csv(
        csv_path,
        [
            _comprovante_row(
                payment_id="nfse_match",
                cnpj_pagador=OUTRA_EMPRESA_CNPJ,
                beneficiary_document="17.755.939/0001-69",
                beneficiary_name="FC CONTABILIDADE LTDA ME",
                data_pagamento="2026-02-26",
                valor="1830",
            )
        ],
    )

    result = run_audit(csv_path, xml_root, dias_tolerancia=7)

    assert result.rows[0].status_audit == STATUS_MATCH_ALTA
    assert result.rows[0].tipo_documento == "NFSE"
    assert result.rows[0].xml_chave == "29253031217755939000169261000000011126020070216971"


def test_run_audit_does_not_pick_a_decision_for_sem_xml(tmp_path):
    csv_path, xml_root = _build_dataset(tmp_path)
    result = run_audit(csv_path, xml_root, dias_tolerancia=7)
    sem_xml_row = next(
        row for row in result.rows if row.comprovante_id == "sem_xml"
    )
    assert sem_xml_row.status_classificacao == "REVISAR_SEM_REGRA"
    assert sem_xml_row.xml_chave == ""


def test_run_audit_cnpj_destinatario_diverge_no_match(tmp_path):
    xml_root = tmp_path / "xml"
    _write_xml(
        xml_root / DESTINATARIO_CNPJ / "chave1-nfe.xml",
        chave_suffix="2" * 44,
        cnpj_emitente=EMITENTE_CNPJ,
        cnpj_destinatario=DESTINATARIO_CNPJ,
        data_emissao="2026-01-05",
        valor="6857.00",
    )
    csv_path = tmp_path / "comprovantes_a_revisar.csv"
    _write_comprovantes_csv(
        csv_path,
        [
            _comprovante_row(
                payment_id="pagador_diferente",
                cnpj_pagador=OUTRA_EMPRESA_CNPJ,
                beneficiary_document="42.870.341/0001-68",
                data_pagamento="2026-01-07",
                valor="6857",
            )
        ],
    )
    result = run_audit(csv_path, xml_root, dias_tolerancia=7)
    assert result.rows[0].status_audit == STATUS_SEM_XML_NO_DATASET


def test_run_audit_detects_multiplos_candidatos(tmp_path):
    xml_root = tmp_path / "xml"
    _write_xml(
        xml_root / DESTINATARIO_CNPJ / "chave1-nfe.xml",
        chave_suffix="3" * 44,
        cnpj_emitente=EMITENTE_CNPJ,
        cnpj_destinatario=DESTINATARIO_CNPJ,
        data_emissao="2026-01-05",
        valor="1000.00",
    )
    _write_xml(
        xml_root / DESTINATARIO_CNPJ / "chave2-nfe.xml",
        chave_suffix="4" * 44,
        cnpj_emitente=EMITENTE_CNPJ,
        cnpj_destinatario=DESTINATARIO_CNPJ,
        data_emissao="2026-01-06",
        valor="1000.00",
    )
    csv_path = tmp_path / "comprovantes_a_revisar.csv"
    _write_comprovantes_csv(
        csv_path,
        [
            _comprovante_row(
                payment_id="duplo",
                cnpj_pagador=DESTINATARIO_CNPJ,
                beneficiary_document="42.870.341/0001-68",
                data_pagamento="2026-01-07",
                valor="1000",
            )
        ],
    )
    result = run_audit(csv_path, xml_root, dias_tolerancia=7)
    duplo_rows = [row for row in result.rows if row.comprovante_id == "duplo"]
    assert len(duplo_rows) == 2
    assert all(row.status_audit == STATUS_MULTIPLOS_CANDIDATOS for row in duplo_rows)


def test_run_audit_dados_insuficientes_without_cnpj(tmp_path):
    xml_root = tmp_path / "xml"
    xml_root.mkdir()
    csv_path = tmp_path / "comprovantes_a_revisar.csv"
    _write_comprovantes_csv(
        csv_path,
        [
            _comprovante_row(
                payment_id="sem_cnpj",
                cnpj_pagador=DESTINATARIO_CNPJ,
                beneficiary_document="",
                data_pagamento="2026-01-07",
                valor="1000",
            )
        ],
    )
    result = run_audit(csv_path, xml_root, dias_tolerancia=7)
    assert result.rows[0].status_audit == STATUS_DADOS_INSUFICIENTES


def test_run_audit_fornecedor_cnpj_filter(tmp_path):
    csv_path, xml_root = _build_dataset(tmp_path)
    result = run_audit(
        csv_path, xml_root, dias_tolerancia=7, fornecedor_cnpj="42.870.341/0001-68"
    )
    assert result.total_comprovantes_avaliados == 2
    assert all(
        row.cnpj_beneficiario == EMITENTE_CNPJ for row in result.rows
    )


def test_run_audit_somente_revisar_filter(tmp_path):
    xml_root = tmp_path / "xml"
    xml_root.mkdir()
    csv_path = tmp_path / "comprovantes.csv"
    rows = [
        _comprovante_row(
            payment_id="revisar1",
            cnpj_pagador=DESTINATARIO_CNPJ,
            beneficiary_document="42.870.341/0001-68",
            data_pagamento="2026-01-07",
            valor="1000",
            status_classificacao="REVISAR_SEM_REGRA",
        ),
        _comprovante_row(
            payment_id="classificado1",
            cnpj_pagador=DESTINATARIO_CNPJ,
            beneficiary_document="42.870.341/0001-68",
            data_pagamento="2026-01-08",
            valor="2000",
            status_classificacao="CLASSIFICADO_POR_REGRA",
        ),
    ]
    _write_comprovantes_csv(csv_path, rows)
    result = run_audit(csv_path, xml_root, dias_tolerancia=7, somente_revisar=True)
    assert result.total_comprovantes_avaliados == 1
    assert result.rows[0].comprovante_id == "revisar1"


def test_build_resumo_por_fornecedor_aggregates_by_cnpj():
    from fantozzi.audit_xml_comprovantes import ComprovanteRecord

    comp1 = ComprovanteRecord(
        comprovante_id="a",
        data_pagamento=date(2026, 1, 1),
        valor=Decimal("100"),
        cnpj_pagador=DESTINATARIO_CNPJ,
        cnpj_beneficiario=EMITENTE_CNPJ,
        beneficiario="HIMEROS",
        status_classificacao="REVISAR_SEM_REGRA",
    )
    comp2 = ComprovanteRecord(
        comprovante_id="b",
        data_pagamento=date(2026, 1, 2),
        valor=Decimal("50"),
        cnpj_pagador=OUTRA_EMPRESA_CNPJ,
        cnpj_beneficiario=EMITENTE_CNPJ,
        beneficiario="HIMEROS",
        status_classificacao="REVISAR_SEM_REGRA",
    )
    resumo = build_resumo_por_fornecedor(
        [(comp1, STATUS_MATCH_ALTA), (comp2, STATUS_SEM_XML_NO_DATASET)]
    )
    assert len(resumo) == 1
    assert resumo[0].qtd_comprovantes == 2
    assert resumo[0].valor_total_comprovantes == Decimal("150")
    assert resumo[0].qtd_match_alta == 1
    assert resumo[0].qtd_sem_xml_dataset == 1


def test_write_outputs_creates_expected_files(tmp_path):
    csv_path, xml_root = _build_dataset(tmp_path)
    result = run_audit(csv_path, xml_root, dias_tolerancia=7)
    output_dir = tmp_path / "output_audit"
    outputs = write_outputs(result, output_dir)
    for key in ("markdown", "csv", "resumo", "manifest"):
        assert outputs[key].is_file()
    assert "SEM_XML_NO_DATASET" in outputs["markdown"].read_text(encoding="utf-8")


def test_cli_has_auditar_comprovantes_xml_cnpj_command():
    parser = build_parser()
    args = parser.parse_args(
        [
            "auditar-comprovantes-xml-cnpj",
            "--comprovantes-csv",
            "dummy.csv",
            "--xml-root",
            "dummy_xml",
        ]
    )
    assert args.command == "auditar-comprovantes-xml-cnpj"
    assert args.dias_tolerancia == 7


def test_cli_end_to_end_generates_outputs(tmp_path, capsys):
    csv_path, xml_root = _build_dataset(tmp_path)
    output_dir = tmp_path / "saida_cli"
    exit_code = main(
        [
            "auditar-comprovantes-xml-cnpj",
            "--comprovantes-csv",
            str(csv_path),
            "--xml-root",
            str(xml_root),
            "--output-dir",
            str(output_dir),
        ]
    )
    assert exit_code == 0
    assert (output_dir / "audit_xml_comprovantes.md").is_file()
    assert (output_dir / "audit_xml_comprovantes.csv").is_file()
    assert (output_dir / "audit_xml_comprovantes_resumo_por_fornecedor.csv").is_file()
    assert (output_dir / "audit_xml_comprovantes_manifest.json").is_file()
    captured = capsys.readouterr()
    assert "MATCH_ALTA" in captured.out


def test_cli_missing_comprovantes_csv_returns_error(tmp_path, capsys):
    exit_code = main(
        [
            "auditar-comprovantes-xml-cnpj",
            "--comprovantes-csv",
            str(tmp_path / "nao_existe.csv"),
            "--xml-root",
            str(tmp_path),
            "--output-dir",
            str(tmp_path / "saida"),
        ]
    )
    assert exit_code == 1
    captured = capsys.readouterr()
    assert "[ERRO]" in captured.out
