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


def test_product_cigarros_tabacaria(product_clf):
    result = product_clf.classify("DUNHILL CARLTON BLEND 2.0 BOX 10X1")
    assert result.conta.id == "2.02.016"
    assert result.confianca == "OK"


def test_product_copos_descartaveis(product_clf):
    result = product_clf.classify("COPO COPOSCHIO 500ML")
    assert result.conta.id == "2.03.001"
    assert result.confianca == "OK"


def test_product_pa_lixo_limpeza(product_clf):
    result = product_clf.classify("PA LIXO PLASTICA")
    assert result.conta.id == "2.04.001"
    assert result.confianca == "OK"


def test_product_sol_premium_cerveja(product_clf):
    result = product_clf.classify("SOL PREMIUM L/N 330ML")
    assert result.conta.id == "2.01.001"
    assert result.confianca == "OK"


def test_product_descontos_obtidos_nao_vira_produto(product_clf):
    result = product_clf.classify("Descontos obtidos identificados no total da NF-e")
    assert result.conta.id == "13.01.001"
    assert result.confianca == "NAO_CLASSIFICADO"


def test_product_unknown(product_clf):
    result = product_clf.classify("PRODUTO XYZ DESCONHECIDO")
    assert result.conta.id == "13.01.001"
    assert result.confianca == "NAO_CLASSIFICADO"


def test_product_canela_tempero(product_clf):
    result = product_clf.classify("CANELA EM PA 1KG")
    assert result.conta.id == "2.02.012"
    assert result.confianca == "OK"


def test_product_cravo_tempero(product_clf):
    result = product_clf.classify("CRAVO DA INDIA 100G")
    assert result.conta.id == "2.02.012"
    assert result.confianca == "OK"


def test_product_essencia_coala_limpeza(product_clf):
    result = product_clf.classify("ESSENCIA COALA 120ML CRAVO CANELA")
    assert result.conta.id == "2.04.001"
    assert result.confianca == "OK"


@pytest.mark.parametrize(
    ("descricao", "conta_id"),
    [
        ("SHINE & DRY 20 LITROS", "2.04.001"),
        ("MOLHO SALADA ROSE 12X236ML", "2.02.012"),
        ("MOLHO PIM. VERDE C/ ESPC. CREM. 12X450ML", "2.02.012"),
        ("LEITE COND SEMID ITALAC 395G", "2.02.006"),
        ("CREME LEITE ITALAC", "2.02.006"),
        ("VINAG BALS VILLA GRIMELLI 500ML", "2.02.012"),
        ("OVOS BRANCO CRIVO 30 DZ", "2.02.007"),
        ("MASSA GAB REFRATARIA 10KG", "8.01.001"),
        ("SACO DE LIXO MULTIPOP 60L", "2.04.001"),
    ],
)
def test_product_aliases_especificos_pos_12b(product_clf, descricao, conta_id):
    result = product_clf.classify(descricao)
    assert result.conta.id == conta_id
    assert result.confianca == "OK"


@pytest.mark.parametrize(
    ("descricao", "conta_id"),
    [
        ("CC SEM ACUCAR LT 350ML COLORCX12", "2.01.004"),
        ("COCA COLA ZERO ACUCAR PET 1L", "2.01.004"),
        ("POLPA SEMPRE VIVA MARACUJA 100GR", "2.01.007"),
        ("H2OH LIMAO C/GAS PET 500ML", "2.01.004"),
        ("M/B BLUECASE BMBH470-D3HGU-M2 REV 2.0 BOX", "10.01.005"),
        ("EQUIPAMENTO ELETRONICO GABINETE CPU", "10.01.005"),
    ],
)
def test_product_casos_sensiveis_pos_12b(product_clf, descricao, conta_id):
    result = product_clf.classify(descricao)
    assert result.conta.id == conta_id
    assert result.confianca == "OK"


@pytest.mark.parametrize(
    "descricao",
    [
        "REFRIGERADOR VERTICAL PEPSI CAPACIDADE 324L",
        "CERVEJEIRA BLK PRIN PM CVIS 8 CX 220V",
        "REFR VT VN PC SCHIN 8CX 220V FRICON",
    ],
)
def test_product_comodato_nao_classifica_como_custo_ou_investimento(product_clf, descricao):
    result = product_clf.classify(descricao)
    assert result.conta.id == "13.01.001"
    assert result.confianca == "NAO_CLASSIFICADO"


def test_product_pataxo_fornecedor_nao_vira_manutencao(product_clf):
    result = product_clf.classify("PATAXO MATERIAIS DE CONSTRUCAO LTDA EPP")
    assert result.conta.id == "13.01.001"
    assert result.confianca == "NAO_CLASSIFICADO"


@pytest.mark.parametrize(
    "fornecedor",
    [
        "PATAXO MATERIAIS DE CONSTRUCAO",
        "pataxo materiais de construcao",
        "Pataxo Materiais de Construcao",
        "Pataxó Materiais de Construção",
    ],
)
def test_product_fornecedor_pataxo_classifica_fallback_manutencao(
    product_clf, fornecedor
):
    result = product_clf.classify("ITEM GENERICO DE OBRA", fornecedor=fornecedor)
    assert result.conta.id == "8.01.001"
    assert result.confianca == "OK"


def test_product_pataxo_na_descricao_nao_dispara_regra_fornecedor(product_clf):
    result = product_clf.classify(
        "ALGUM PRODUTO PATAXO", fornecedor="OUTRO FORNECEDOR"
    )
    assert result.conta.id == "13.01.001"
    assert result.confianca == "NAO_CLASSIFICADO"


def test_product_especifico_prevalece_sobre_fornecedor_pataxo(product_clf):
    result = product_clf.classify(
        "SACO DE LIXO MULTIPOP", fornecedor="PATAXO MATERIAIS DE CONSTRUCAO"
    )
    assert result.conta.id == "2.04.001"
    assert result.confianca == "OK"


def test_product_shine_dry_prevalece_sobre_fornecedor_pataxo(product_clf):
    result = product_clf.classify(
        "SHINE & DRY 20 LITROS", fornecedor="PATAXO MATERIAIS DE CONSTRUCAO"
    )
    assert result.conta.id == "2.04.001"
    assert result.confianca == "OK"


@pytest.mark.parametrize(
    ("fornecedor", "conta_id"),
    [
        ("CENTRAL DE ALIMENTOS DA BAHIA - CAB LTDA", "2.02.008"),
        ("DATERRA DISTRIBUIDORA DE ALIMENTOS LTDA", "2.02.008"),
        ("REI DAS BOMBAS MANUT E REPA DE MAQ LTDA", "8.01.001"),
    ],
)
def test_product_fornecedores_recorrentes_pos_12g(product_clf, fornecedor, conta_id):
    result = product_clf.classify("ITEM GENERICO SEM ALIAS", fornecedor=fornecedor)
    assert result.conta.id == conta_id
    assert result.confianca == "OK"


@pytest.mark.parametrize(
    ("descricao", "conta_id"),
    [
        ("PAPEL TOALHA BOBINA FS 20CM X 200M 30G 6 RLS", "2.03.001"),
        (
            "PALETA CX COM 24UN 120G MORANGO COM LEITE COND PALETITAS",
            "2.02.015",
        ),
        (
            "PICOLE CX COM 24UN 65G PALETITO ZERO AD DE ACUCAR PALETITAS",
            "2.02.015",
        ),
        ("PICOLE CX COM 27UN 68G SPACE MONSTERS PALETITAS", "2.02.015"),
        ("ATUM RAL C/OLEO E CALDO VEG POUCH 500 GR 88", "2.02.013"),
        ("MONIN XAROPE PESSEGO 700ML", "2.01.007"),
        ("AGUA TONICA FYS ZERO LT 350ML", "2.01.004"),
        ("REFR FYS LIMAO ZERO 0,350LT DESC 12UN", "2.01.004"),
        ("FYS GUARANA 350ML", "2.01.004"),
        ("ADOCANTE ZERO-CAL 100ML", "2.02.010"),
        ("SAL LEBRE SACHE 2000X0,8G", "2.02.012"),
        ("CANETA PRETA 1.0 RETRO PROJETOR PILOT", "6.01.005"),
        ("CANETA PRETA PILOT", "6.01.005"),
        ("DESINF LIMPO 5LTS", "2.04.001"),
    ],
)
def test_product_aliases_especificos_pos_12g(product_clf, descricao, conta_id):
    result = product_clf.classify(descricao)
    assert result.conta.id == conta_id
    assert result.confianca == "OK"


def test_product_checon_nao_vira_fornecedor_broad(product_clf):
    result = product_clf.classify(
        "ITEM GENERICO SEM ALIAS",
        fornecedor="Checon Distribuidora e Transportadora Lt",
    )
    assert result.conta.id == "13.01.001"
    assert result.confianca == "NAO_CLASSIFICADO"


def test_product_refrigerante_checon_classifica_por_produto(product_clf):
    result = product_clf.classify(
        "FYS GUARANA 350ML",
        fornecedor="Checon Distribuidora e Transportadora Lt",
    )
    assert result.conta.id == "2.01.004"
    assert result.confianca == "OK"


@pytest.mark.parametrize(
    "descricao",
    [
        "PIRUL PEGADAS DO BANZE BRILHANTE 30UN DROPS HALLS CEREJA",
        "CHIC TORTUG HUEVITOS 12UN DROPS HALLS CEREJA 1X21UN",
        "CHOC BATON AO LEITE 30X16G",
    ],
)
def test_product_delmax_classifica_fallback_sobremesas(product_clf, descricao):
    result = product_clf.classify(
        descricao,
        fornecedor="DELMAX SUPERMERCADO COMERCIO DE ALIMENTOS LTDA",
    )
    assert result.conta.id == "2.02.015"
    assert result.confianca == "OK"


def test_product_especifico_prevalece_sobre_fornecedor_delmax(product_clf):
    result = product_clf.classify(
        "COPO COPOSCHIO 500ML",
        fornecedor="DELMAX SUPERMERCADO COMERCIO DE ALIMENTOS LTDA",
    )
    assert result.conta.id == "2.03.001"
    assert result.confianca == "OK"


def test_product_item_generico_sem_fornecedor_delmax_nao_classifica(product_clf):
    result = product_clf.classify(
        "ITEM GENERICO SEM ALIAS",
        fornecedor="OUTRO FORNECEDOR",
    )
    assert result.conta.id == "13.01.001"
    assert result.confianca == "NAO_CLASSIFICADO"


def test_product_valor_inconsistente_nao_e_resolvido_por_alias(product_clf):
    result = product_clf.classify("Descontos obtidos identificados no total da NF-e")
    assert result.conta.id == "13.01.001"
    assert result.confianca == "NAO_CLASSIFICADO"


def test_expense_energia_eletrica(expense_clf):
    result = expense_clf.classify("energia eletrica")
    assert result.conta.id == "4.01.011"


def test_expense_fgts(expense_clf):
    result = expense_clf.classify("fgts guia")
    assert result.conta.id == "3.03.001"
