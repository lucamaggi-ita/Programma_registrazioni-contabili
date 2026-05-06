from decimal import Decimal
from pathlib import Path

import pytest

from fantozzi.models import NFE
from fantozzi.nfe_importer import parse_nfe

SAMPLE_XML = (
    Path(__file__).parent.parent
    / "samples"
    / "nfe_exemplo"
    / "nfe_oesa_12566_20260505.xml"
)


@pytest.fixture(scope="module")
def nfe() -> NFE:
    return parse_nfe(SAMPLE_XML)


def test_returns_nfe_object(nfe):
    assert isinstance(nfe, NFE)


def test_numero_nfe(nfe):
    assert nfe.numero == "12566"


def test_valor_total(nfe):
    assert nfe.valor_total == Decimal("35.00")


def test_total_breakdown_fields(nfe):
    assert nfe.valor_produtos == Decimal("35.00")
    assert nfe.valor_frete == Decimal("0.00")
    assert nfe.valor_seguro == Decimal("0.00")
    assert nfe.valor_desconto == Decimal("0.00")
    assert nfe.valor_outras_despesas == Decimal("0.00")
    assert nfe.valor_tributos_inclusos == Decimal("0.00")


def test_at_least_three_produtos(nfe):
    assert len(nfe.produtos) >= 3


def test_cnpj_destinatario_populated(nfe):
    assert nfe.cnpj_destinatario != ""


def test_nome_destinatario_populated(nfe):
    assert nfe.nome_destinatario != ""


def test_two_duplicatas(nfe):
    assert len(nfe.duplicatas) == 2


def test_duplicatas_valor(nfe):
    for dup in nfe.duplicatas:
        assert dup.valor == Decimal("17.50")


def test_tem_duplicatas_true(nfe):
    assert nfe.tem_duplicatas is True


def test_arquivo_origem_populated(nfe):
    assert nfe.arquivo_origem != ""
