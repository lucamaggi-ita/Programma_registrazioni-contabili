from pathlib import Path

import pytest

from fantozzi.chart_of_accounts import ChartOfAccounts


CONFIG_DIR = Path(__file__).parent.parent / "config"


@pytest.fixture
def chart():
    return ChartOfAccounts.load(CONFIG_DIR / "plano_contas_gerencial.yaml")


def test_load_no_validation_errors(chart):
    assert chart.validate() == []


def test_get_by_id_refrigerantes(chart):
    conta = chart.get_by_id("2.01.004")
    assert conta.nome == "Refrigerantes"


def test_all_six_review_accounts_exist(chart):
    review_ids = [
        "13.01.001",
        "13.01.002",
        "13.01.006",
        "13.01.009",
        "13.01.010",
        "13.01.999",
    ]
    for rid in review_ids:
        conta = chart.get_by_id(rid)
        assert conta.id == rid


def test_get_review_account_nao_classificado(chart):
    conta = chart.get_review_account("NAO_CLASSIFICADO")
    assert conta.id == "13.01.001"
