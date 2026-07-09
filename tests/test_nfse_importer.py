from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from fantozzi.nfe_importer import parse_nfe
from fantozzi.nfse_importer import parse_nfse


FIXTURES = Path(__file__).parent / "fixtures" / "nfse"
NACIONAL_FIXTURES = Path(__file__).parent / "fixtures" / "nfse_nacional"
NFSE_TOMADOR = FIXTURES / "17755939000169202610000000111CFEQ5H12-nfse.xml"
NFSE_TOMADOR_SERVICO = FIXTURES / "422162670001604742ZSLDD6CV-nfse.xml"
NFSE_NACIONAL_AMAMI = NACIONAL_FIXTURES / "amami-nfse-nacional.xml"
NFSE_NACIONAL_TRATTORIA = NACIONAL_FIXTURES / "trattoria-nfse-nacional.xml"


def test_parse_nfse_abrasf_layout_com_tomador():
    nfse = parse_nfse(NFSE_TOMADOR)

    assert nfse.tipo_documento == "NFSE"
    assert nfse.numero == "202610000000111"
    assert nfse.codigo_verificacao == "CFEQ-5H12"
    assert nfse.chave_acesso == "29253031217755939000169261000000011126020070216971"
    assert nfse.data_emissao == date(2026, 2, 25)
    assert nfse.competencia == date(2026, 2, 1)
    assert nfse.cnpj_prestador == "17755939000169"
    assert nfse.razao_social_prestador == "FC CONTABILIDADE LTDA ME"
    assert nfse.cnpj_tomador == "08884180000123"
    assert nfse.razao_social_tomador == "AMAMI RESTAURANTE LTDA"


def test_parse_nfse_abrasf_layout_com_tomador_servico():
    nfse = parse_nfse(NFSE_TOMADOR_SERVICO)

    assert nfse.numero == "4742"
    assert nfse.codigo_verificacao == "ZSLDD6CV"
    assert nfse.data_emissao == date(2026, 3, 3)
    assert nfse.competencia == date(2026, 3, 1)
    assert nfse.cnpj_prestador == "42216267000160"
    assert nfse.razao_social_prestador == "ALEX STRAUCH BATISTA"
    assert nfse.cnpj_tomador == "08884180000123"
    assert nfse.razao_social_tomador == "AMAMI RESTAURANTE LTDA"
    assert "Acompanhamento de marca" in nfse.discriminacao


def test_parse_nfse_normaliza_cnpj():
    nfse = parse_nfse(NFSE_TOMADOR_SERVICO)

    assert nfse.cnpj_prestador == "42216267000160"
    assert nfse.cnpj_tomador == "08884180000123"


def test_parse_nfse_valor_com_ponto():
    nfse = parse_nfse(NFSE_TOMADOR)

    assert nfse.valor_servicos == Decimal("1830")
    assert nfse.valor_liquido == Decimal("1830")


def test_parse_nfse_valor_com_virgula():
    nfse = parse_nfse(NFSE_TOMADOR_SERVICO)

    assert nfse.valor_servicos == Decimal("1440.20")
    assert nfse.valor_liquido == Decimal("1440.20")


def test_parse_nfse_nacional_amami():
    nfse = parse_nfse(NFSE_NACIONAL_AMAMI)

    assert nfse.tipo_documento == "NFSE"
    assert nfse.numero == "2026100000321"
    assert nfse.chave_acesso == "29253031202938301000196202610000032126050073035066"
    assert nfse.cnpj_prestador == "02938301000196"
    assert nfse.razao_social_prestador == "INFORMATICA PORTO SEGURO 2000 LTDA"
    assert nfse.cnpj_tomador == "08884180000123"
    assert nfse.razao_social_tomador == "AMAMI RESTAURANTE LTDA"
    assert "TROCA PLACA" in nfse.discriminacao


def test_parse_nfse_nacional_trattoria():
    nfse = parse_nfse(NFSE_NACIONAL_TRATTORIA)

    assert nfse.tipo_documento == "NFSE"
    assert nfse.numero == "137145"
    assert nfse.chave_acesso == "43149022201579286000174000000013714526077317298417"
    assert nfse.cnpj_prestador == "01579286000174"
    assert nfse.razao_social_prestador == "SAFEWEB SEGURANCA DA INFORMACAO LTDA."
    assert nfse.cnpj_tomador == "55746722000181"
    assert nfse.razao_social_tomador == "Trattoria do Sossego Restaurante Ltda"
    assert "Certificado Digital" in nfse.discriminacao


def test_parse_nfse_nacional_normaliza_cnpj():
    nfse = parse_nfse(NFSE_NACIONAL_TRATTORIA)

    assert nfse.cnpj_prestador == "01579286000174"
    assert nfse.cnpj_tomador == "55746722000181"


def test_parse_nfse_nacional_extrai_valor():
    nfse = parse_nfse(NFSE_NACIONAL_AMAMI)

    assert nfse.valor_servicos == Decimal("100.00")
    assert nfse.valor_liquido == Decimal("100.00")


def test_parse_nfse_nacional_extrai_data():
    nfse = parse_nfse(NFSE_NACIONAL_TRATTORIA)

    assert nfse.data_emissao == date(2026, 7, 6)
    assert nfse.competencia == date(2026, 7, 6)


def test_parse_nfse_nacional_nao_quebra_abrasf():
    nfse = parse_nfse(NFSE_TOMADOR)

    assert nfse.numero == "202610000000111"
    assert nfse.chave_acesso == "29253031217755939000169261000000011126020070216971"


def test_parse_nfse_nao_confunde_nfe_com_nfse(tmp_path):
    nfe_path = tmp_path / "nfe.xml"
    nfe_path.write_text(
        """<?xml version="1.0" encoding="UTF-8"?>
<nfeProc versao="4.00" xmlns="http://www.portalfiscal.inf.br/nfe">
  <NFe><infNFe Id="NFe123"><ide><nNF>1</nNF><serie>1</serie><dhEmi>2026-01-01</dhEmi></ide><emit /><dest /><total /></infNFe></NFe>
</nfeProc>
""",
        encoding="utf-8",
    )

    with pytest.raises(ValueError):
        parse_nfse(nfe_path)
    assert parse_nfe(nfe_path).numero == "1"
