from pathlib import Path

import pytest

from fantozzi.chart_of_accounts import ChartOfAccounts
from fantozzi.expense_classifier import ExpenseClassifier
from fantozzi.product_classifier import ProductClassifier


CONFIG_DIR = Path(__file__).parent.parent / "config"


@pytest.fixture
def chart():
    return ChartOfAccounts.load(CONFIG_DIR / "plano_contas_gerencial.yaml")


@pytest.fixture
def product_clf(chart):
    return ProductClassifier(chart).load(CONFIG_DIR / "regras_classificacao_produtos.yaml")


@pytest.fixture
def expense_clf(chart):
    return ExpenseClassifier(chart).load(CONFIG_DIR / "regras_classificacao_despesas.yaml")


def test_product_coca_cola(product_clf):
    result = product_clf.classify("COCA-COLA KS 290ML")
    assert result.conta.id == "2.01.004"
    assert result.confianca == "OK"


def test_product_vinho_rio_sol(product_clf):
    result = product_clf.classify("VINHO RIO SOL TINTO SECO 750ML")
    assert result.conta.id == "2.01.002"
    assert result.confianca == "OK"


def test_product_esponja(product_clf):
    result = product_clf.classify("ESPONJA DUPLA FACE LIMPEZA")
    assert result.conta.id == "2.04.001"
    assert result.confianca == "OK"


def test_product_unknown(product_clf):
    result = product_clf.classify("PRODUTO XYZ DESCONHECIDO")
    assert result.conta.id == "13.01.001"
    assert result.confianca == "NAO_CLASSIFICADO"


def test_expense_energia_eletrica(expense_clf):
    result = expense_clf.classify("energia eletrica")
    assert result.conta.id == "4.01.011"


def test_expense_fgts(expense_clf):
    result = expense_clf.classify("fgts guia")
    assert result.conta.id == "3.03.001"
