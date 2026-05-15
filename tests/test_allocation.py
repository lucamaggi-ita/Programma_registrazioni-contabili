from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from fantozzi.allocation import allocate_nfe_to_lancamentos
from fantozzi.chart_of_accounts import ChartOfAccounts
from fantozzi.models import Duplicata, Lancamento, NFE, ProdutoNFE
from fantozzi.nfe_importer import parse_nfe
from fantozzi.product_classifier import ProductClassifier


PROJECT_ROOT = Path(__file__).parent.parent
CONFIG_DIR = PROJECT_ROOT / "config"
SAMPLE_XML = (
    PROJECT_ROOT / "samples" / "nfe_exemplo" / "nfe_oesa_12566_20260505.xml"
)


@pytest.fixture
def product_classifier():
    chart = ChartOfAccounts.load(CONFIG_DIR / "plano_contas_gerencial.yaml")
    return ProductClassifier(chart).load(CONFIG_DIR / "regras_classificacao_produtos.yaml")


@pytest.fixture
def sample_nfe():
    return parse_nfe(SAMPLE_XML)


@pytest.fixture
def sample_lancamentos(sample_nfe, product_classifier):
    return allocate_nfe_to_lancamentos(sample_nfe, product_classifier)


def test_allocate_returns_lancamentos(sample_lancamentos):
    assert isinstance(sample_lancamentos, list)
    assert all(isinstance(lancamento, Lancamento) for lancamento in sample_lancamentos)


def test_groups_products_by_expected_accounts(sample_lancamentos):
    by_conta = {lancamento.conta_id: lancamento for lancamento in sample_lancamentos}

    assert by_conta["2.01.004"].conta_nome == "Refrigerantes"
    assert by_conta["2.01.002"].conta_nome == "Vinhos e Espumantes"
    assert by_conta["2.04.001"].conta_nome == "Produtos de Limpeza e Higiene"
    assert len(sample_lancamentos) == 3


def test_lancamento_total_matches_nfe_total(sample_lancamentos):
    total = sum((lancamento.valor for lancamento in sample_lancamentos), Decimal("0"))
    assert total == Decimal("35.00")


def test_required_fields_are_populated(sample_lancamentos):
    for lancamento in sample_lancamentos:
        assert lancamento.cnpj_empresa
        assert lancamento.nome_empresa
        assert lancamento.fornecedor_ou_favorecido
        assert lancamento.chave_origem
        assert lancamento.arquivo_origem
        assert lancamento.hash_registro
        assert lancamento.data_documento == date(2026, 5, 5)
        assert lancamento.competencia == "2026-05"
        assert lancamento.tipo_documento == "XML_NFE"
        assert lancamento.origem == "XML_NFE"


def test_duplicatas_are_allocated_proportionally(sample_lancamentos):
    by_conta = {lancamento.conta_id: lancamento for lancamento in sample_lancamentos}

    assert [d.valor for d in by_conta["2.01.004"].duplicatas] == [
        Decimal("5.00"),
        Decimal("5.00"),
    ]
    assert [d.valor for d in by_conta["2.01.002"].duplicatas] == [
        Decimal("10.00"),
        Decimal("10.00"),
    ]
    assert [d.valor for d in by_conta["2.04.001"].duplicatas] == [
        Decimal("2.50"),
        Decimal("2.50"),
    ]

    for lancamento in sample_lancamentos:
        assert sum((dup.valor for dup in lancamento.duplicatas), Decimal("0")) == (
            lancamento.valor
        )


def test_duplicata_totals_by_vencimento_match_original(sample_lancamentos):
    totals_by_vencimento = {}
    for lancamento in sample_lancamentos:
        for duplicata in lancamento.duplicatas:
            totals_by_vencimento.setdefault(duplicata.vencimento, Decimal("0"))
            totals_by_vencimento[duplicata.vencimento] += duplicata.valor

    assert set(totals_by_vencimento.values()) == {Decimal("17.50")}
    assert len(totals_by_vencimento) == 2


def test_status_revisao_ok_for_classified_sample(sample_lancamentos):
    assert {lancamento.status_revisao for lancamento in sample_lancamentos} == {"OK"}
    assert all(lancamento.alertas == "" for lancamento in sample_lancamentos)


def test_unclassified_product_generates_revisar_status(product_classifier):
    nfe = NFE(
        chave="NFE-UNKNOWN-001",
        numero="999",
        serie="1",
        data_emissao=date(2026, 5, 6),
        cnpj_emitente="12345678000195",
        nome_emitente="FORNECEDOR TESTE",
        cnpj_destinatario="55746722000181",
        nome_destinatario="TRATTORIA DO SOSSEGO RESTAURANTE LTDA",
        valor_total=Decimal("12.00"),
        produtos=[
            ProdutoNFE(
                n_item=1,
                codigo="UNK",
                descricao="PRODUTO XYZ DESCONHECIDO",
                ncm="00000000",
                cfop="5102",
                unidade="UN",
                quantidade=Decimal("1"),
                valor_unitario=Decimal("12.00"),
                valor_total=Decimal("12.00"),
            )
        ],
        duplicatas=[
            Duplicata(
                numero="001",
                vencimento=date(2026, 5, 20),
                valor=Decimal("12.00"),
            )
        ],
        tem_duplicatas=True,
        arquivo_origem="sample_unknown.xml",
    )

    lancamentos = allocate_nfe_to_lancamentos(nfe, product_classifier)

    assert len(lancamentos) == 1
    assert lancamentos[0].conta_id == "13.01.001"
    assert lancamentos[0].status_revisao == "REVISAR"
    assert "NAO_CLASSIFICADO" in lancamentos[0].alertas


def test_supplier_rule_classifies_unmatched_product(product_classifier):
    nfe = _synthetic_nfe(
        chave="NFE-SUPPLIER-RULE-001",
        valor_total=Decimal("100.00"),
        duplicata_valor=Decimal("100.00"),
        fornecedor="PATAXO MATERIAIS DE CONSTRUCAO LTDA EPP",
        produto_descricao="ITEM GENERICO DE OBRA",
    )

    lancamentos = allocate_nfe_to_lancamentos(nfe, product_classifier)

    assert len(lancamentos) == 1
    assert lancamentos[0].conta_id == "8.01.001"
    assert lancamentos[0].conta_nome == "Manutencao Predial"
    assert lancamentos[0].status_revisao == "OK"
    assert lancamentos[0].alertas == ""


def test_comodato_cfop_goes_to_special_review_before_product_alias(product_classifier):
    nfe = _synthetic_nfe(
        chave="NFE-COMODATO-001",
        valor_total=Decimal("100.00"),
        duplicata_valor=Decimal("100.00"),
        natureza_operacao="COMODATO",
        produto_descricao="FREEZER METALFRIO",
        produto_cfop="6908",
    )

    lancamentos = allocate_nfe_to_lancamentos(nfe, product_classifier)

    assert len(lancamentos) == 1
    assert lancamentos[0].conta_id == "13.01.001"
    assert lancamentos[0].status_revisao == "REVISAR"
    assert "OPERACAO_ESPECIAL_COMODATO" in lancamentos[0].alertas
    assert lancamentos[0].tipo_operacao_especial == "COMODATO"
    assert lancamentos[0].cfop == "6908"


def test_bonificacao_cfop_goes_to_special_review_before_product_alias(product_classifier):
    nfe = _synthetic_nfe(
        chave="NFE-BONIFICACAO-001",
        valor_total=Decimal("100.00"),
        duplicata_valor=Decimal("100.00"),
        natureza_operacao="BONIFICACAO",
        produto_descricao="REFR GUARANA DA AMAZONIA FYS",
        produto_cfop="5910",
    )

    lancamentos = allocate_nfe_to_lancamentos(nfe, product_classifier)

    assert len(lancamentos) == 1
    assert lancamentos[0].conta_id == "13.01.001"
    assert lancamentos[0].status_revisao == "REVISAR"
    assert "OPERACAO_ESPECIAL_BONIFICACAO" in lancamentos[0].alertas
    assert lancamentos[0].tipo_operacao_especial == "BONIFICACAO_RECEBIDA"
    assert lancamentos[0].tratamento_operacao_especial == "FORA_CUSTO_COMPRA"


def test_normal_cfop_still_uses_product_classifier(product_classifier):
    nfe = _synthetic_nfe(
        chave="NFE-CFOP-NORMAL-001",
        valor_total=Decimal("100.00"),
        duplicata_valor=Decimal("100.00"),
        produto_descricao="COCA-COLA KS 290ML",
        produto_cfop="5405",
    )

    lancamentos = allocate_nfe_to_lancamentos(nfe, product_classifier)

    assert len(lancamentos) == 1
    assert lancamentos[0].conta_id == "2.01.004"
    assert lancamentos[0].status_revisao == "OK"
    assert lancamentos[0].tipo_operacao_especial == "COMPRA_NORMAL"


def test_mixed_nfe_keeps_normal_item_and_routes_special_item(product_classifier):
    nfe = _synthetic_nfe_with_products(
        chave="NFE-MISTA-001",
        natureza_operacao="VENDA DE MERCADORIAS",
        produtos=[
            _produto("COCA-COLA KS 290ML", "5405", Decimal("100.00")),
            _produto("REFR GUARANA DA AMAZONIA FYS", "5910", Decimal("100.00")),
        ],
    )

    lancamentos = allocate_nfe_to_lancamentos(nfe, product_classifier)
    by_operacao = {lancamento.tipo_operacao_especial: lancamento for lancamento in lancamentos}

    assert by_operacao["COMPRA_NORMAL"].conta_id == "2.01.004"
    assert by_operacao["COMPRA_NORMAL"].status_revisao == "OK"
    assert by_operacao["COMPRA_NORMAL"].cfop == "5405"
    assert by_operacao["BONIFICACAO_RECEBIDA"].conta_id == "13.01.001"
    assert by_operacao["BONIFICACAO_RECEBIDA"].status_revisao == "REVISAR"
    assert by_operacao["BONIFICACAO_RECEBIDA"].cfop == "5910"


def test_positive_accessory_cost_creates_separate_lancamento(product_classifier):
    nfe = _synthetic_nfe(
        chave="NFE-ACCESSORY-001",
        valor_total=Decimal("110.00"),
        valor_frete=Decimal("10.00"),
        duplicata_valor=Decimal("110.00"),
    )

    lancamentos = allocate_nfe_to_lancamentos(nfe, product_classifier)
    by_conta = {lancamento.conta_id: lancamento for lancamento in lancamentos}

    assert by_conta["2.05.001"].conta_nome == (
        "Fretes, Seguros e Outras Despesas sobre Compras"
    )
    assert by_conta["2.05.001"].valor == Decimal("10.00")
    assert by_conta["2.05.001"].status_revisao == "OK"
    assert sum((lancamento.valor for lancamento in lancamentos), Decimal("0")) == (
        Decimal("110.00")
    )


def test_discount_creates_negative_lancamento_without_positive_accessory_cost(
    product_classifier,
):
    nfe = _synthetic_nfe(
        chave="NFE-DISCOUNT-001",
        valor_total=Decimal("90.00"),
        valor_desconto=Decimal("10.00"),
        duplicata_valor=Decimal("90.00"),
    )

    lancamentos = allocate_nfe_to_lancamentos(nfe, product_classifier)
    by_conta = {lancamento.conta_id: lancamento for lancamento in lancamentos}

    assert "2.05.001" not in by_conta
    assert by_conta["2.05.002"].conta_nome == "Descontos Obtidos sobre Compras"
    assert by_conta["2.05.002"].valor == Decimal("-10.00")
    assert by_conta["2.05.002"].status_revisao == "OK"
    assert sum((lancamento.valor for lancamento in lancamentos), Decimal("0")) == (
        Decimal("90.00")
    )


def _synthetic_nfe(
    *,
    chave: str,
    valor_total: Decimal,
    duplicata_valor: Decimal,
    valor_frete: Decimal = Decimal("0.00"),
    valor_desconto: Decimal = Decimal("0.00"),
    fornecedor: str = "FORNECEDOR TESTE",
    produto_descricao: str = "COCA-COLA KS 290ML",
    produto_cfop: str = "5102",
    natureza_operacao: str = "VENDA DE MERCADORIAS",
) -> NFE:
    return NFE(
        chave=chave,
        numero="777",
        serie="1",
        data_emissao=date(2026, 5, 6),
        cnpj_emitente="12345678000195",
        nome_emitente=fornecedor,
        cnpj_destinatario="55746722000181",
        nome_destinatario="TRATTORIA DO SOSSEGO RESTAURANTE LTDA",
        valor_total=valor_total,
        natureza_operacao=natureza_operacao,
        valor_produtos=Decimal("100.00"),
        valor_frete=valor_frete,
        valor_desconto=valor_desconto,
        produtos=[
            ProdutoNFE(
                n_item=1,
                codigo="REFR",
                descricao=produto_descricao,
                ncm="22021000",
                cfop=produto_cfop,
                unidade="CX",
                quantidade=Decimal("1"),
                valor_unitario=Decimal("100.00"),
                valor_total=Decimal("100.00"),
            )
        ],
        duplicatas=[
            Duplicata(
                numero="001",
                vencimento=date(2026, 5, 20),
                valor=duplicata_valor,
            )
        ],
        tem_duplicatas=True,
        arquivo_origem=f"{chave}.xml",
    )


def _synthetic_nfe_with_products(
    *,
    chave: str,
    natureza_operacao: str,
    produtos: list[ProdutoNFE],
) -> NFE:
    total = sum((produto.valor_total for produto in produtos), Decimal("0"))
    return NFE(
        chave=chave,
        numero="777",
        serie="1",
        data_emissao=date(2026, 5, 6),
        cnpj_emitente="12345678000195",
        nome_emitente="FORNECEDOR TESTE",
        cnpj_destinatario="55746722000181",
        nome_destinatario="TRATTORIA DO SOSSEGO RESTAURANTE LTDA",
        valor_total=total,
        natureza_operacao=natureza_operacao,
        valor_produtos=total,
        produtos=produtos,
        duplicatas=[
            Duplicata(
                numero="001",
                vencimento=date(2026, 5, 20),
                valor=total,
            )
        ],
        tem_duplicatas=True,
        arquivo_origem=f"{chave}.xml",
    )


def _produto(descricao: str, cfop: str, valor: Decimal) -> ProdutoNFE:
    return ProdutoNFE(
        n_item=1,
        codigo="P",
        descricao=descricao,
        ncm="22021000",
        cfop=cfop,
        unidade="UN",
        quantidade=Decimal("1"),
        valor_unitario=valor,
        valor_total=valor,
    )
