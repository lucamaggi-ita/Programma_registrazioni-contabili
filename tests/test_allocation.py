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
    assert lancamentos[0].conta_id == "2.05.003"
    assert lancamentos[0].conta_nome == "Reducao de Custos - Bonificacoes Recebidas"
    assert lancamentos[0].status_revisao == "REVISAR"
    assert "OPERACAO_ESPECIAL_BONIFICACAO" in lancamentos[0].alertas
    assert lancamentos[0].tipo_operacao_especial == "BONIFICACAO_RECEBIDA"
    assert lancamentos[0].tratamento_operacao_especial == "REDUCAO_CUSTO"


@pytest.mark.parametrize(
    ("descricao", "cfop", "conta_custo_normal"),
    [
        ("REFR GUARANA DA AMAZONIA FYS", "5910", "2.01.004"),
        ("MASSA ESPAGUETE", "5910", "2.02.010"),
        ("BATATA 9MM", "5910", "2.02.014"),
        ("REFR GUARANA DA AMAZONIA FYS", "6910", "2.01.004"),
    ],
)
def test_bonificacao_cfop_never_goes_to_normal_cost_accounts(
    product_classifier, descricao, cfop, conta_custo_normal
):
    nfe = _synthetic_nfe(
        chave=f"NFE-BONIFICACAO-{conta_custo_normal}",
        valor_total=Decimal("100.00"),
        duplicata_valor=Decimal("100.00"),
        natureza_operacao="BONIFICACAO",
        produto_descricao=descricao,
        produto_cfop=cfop,
    )

    lancamentos = allocate_nfe_to_lancamentos(nfe, product_classifier)

    assert {lancamento.conta_id for lancamento in lancamentos} == {"2.05.003"}
    assert conta_custo_normal not in {lancamento.conta_id for lancamento in lancamentos}


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
    assert by_operacao["BONIFICACAO_RECEBIDA"].conta_id == "2.05.003"
    assert by_operacao["BONIFICACAO_RECEBIDA"].status_revisao == "REVISAR"
    assert by_operacao["BONIFICACAO_RECEBIDA"].cfop == "5910"
    assert by_operacao["BONIFICACAO_RECEBIDA"].tratamento_operacao_especial == "REDUCAO_CUSTO"


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


def test_mixed_cfop_nfe_with_item_discounts_uses_payable_duplicata_total(
    product_classifier,
):
    nfe = NFE(
        chave="NFE-CHECON-2052542",
        numero="2052542",
        serie="2",
        data_emissao=date(2026, 1, 19),
        cnpj_emitente="40463846000182",
        nome_emitente="Checon Distribuidora e Transportadora Lt",
        cnpj_destinatario="08884180000123",
        nome_destinatario="AMAMI RESTAURANTE LTDA",
        valor_total=Decimal("4116.41"),
        natureza_operacao="Venda / Bonificacao",
        valor_produtos=Decimal("4308.61"),
        valor_desconto=Decimal("192.20"),
        produtos=[
            _produto("CERV DEVASSA LAGER 06LGFA RT 24UN", "5405", Decimal("127.75")),
            _produto(
                "CERV HEINEKEN PIL 0,60GFA RT 24UN",
                "5405",
                Decimal("2491.30"),
                valor_desconto=Decimal("115.10"),
            ),
            _produto(
                "CERV AMSTEL LAGER 0,60L GFA RT 24UN",
                "5405",
                Decimal("1512.50"),
                valor_desconto=Decimal("77.10"),
            ),
            _produto("FYS GUARANA 350ML", "5910", Decimal("28.14")),
            _produto(
                "CERV HEINEKEN PIL 0,330GFA DESC 4X6UNPBR",
                "5405",
                Decimal("148.92"),
            ),
        ],
        duplicatas=[
            Duplicata(numero="001", vencimento=date(2026, 1, 26), valor=Decimal("2044.13")),
            Duplicata(numero="002", vencimento=date(2026, 2, 2), valor=Decimal("2044.14")),
        ],
        tem_duplicatas=True,
        tem_cobranca=True,
        valor_cobranca=Decimal("4088.27"),
        arquivo_origem="checon_2052542.xml",
    )

    lancamentos = allocate_nfe_to_lancamentos(nfe, product_classifier)
    by_conta = {lancamento.conta_id: lancamento for lancamento in lancamentos}

    assert by_conta["2.01.001"].valor == Decimal("4280.47")
    assert by_conta["2.05.003"].valor == Decimal("28.14")
    assert by_conta["2.05.002"].valor == Decimal("-192.20")
    assert all("VALOR_INCONSISTENTE" not in l.alertas for l in lancamentos)
    assert sum((d.valor for d in by_conta["2.01.001"].duplicatas), Decimal("0")) == (
        Decimal("4280.47")
    )
    assert sum((d.valor for d in by_conta["2.05.002"].duplicatas), Decimal("0")) == (
        Decimal("-192.20")
    )
    assert by_conta["2.05.003"].duplicatas == []


def test_wide_remessa_nature_does_not_create_valor_inconsistente_when_cfops_close(
    product_classifier,
):
    nfe = NFE(
        chave="NFE-PETROPOLIS-331181",
        numero="331181",
        serie="2",
        data_emissao=date(2026, 1, 24),
        cnpj_emitente="73410326018884",
        nome_emitente="CERVEJARIA PETROPOLIS S/A EM REC. JUDICIAL",
        cnpj_destinatario="08884180000123",
        nome_destinatario="AMAMI RESTAURANTE LTDA",
        valor_total=Decimal("676.59"),
        natureza_operacao="VENDA MERC RECEBTERC C/ STRIBUT CONT SUBSTITUIDO / REMESSA",
        valor_produtos=Decimal("659.94"),
        valor_tributos_inclusos=Decimal("16.65"),
        produtos=[
            _produto("CERV ITA PILS GRF RET 600ML", "5405", Decimal("290.29")),
            _produto("CERV ITA PILS GRF RET 600ML", "5910", Decimal("58.06")),
            _produto("CERV ITA PREM GRF RET 600ML", "5405", Decimal("141.54")),
            _produto("CERV PETRA PM GRF RET 600ML", "5405", Decimal("114.25")),
            _produto("CERV PETRA PM GRF RET 600ML", "5910", Decimal("22.85")),
            _produto("ENERG TNT ORG LT 269ML PACK 6", "5405", Decimal("32.95")),
        ],
        duplicatas=[
            Duplicata(numero="001", vencimento=date(2026, 2, 1), valor=Decimal("595.68")),
        ],
        tem_duplicatas=True,
        tem_cobranca=True,
        valor_cobranca=Decimal("595.68"),
        arquivo_origem="petropolis_331181.xml",
    )

    lancamentos = allocate_nfe_to_lancamentos(nfe, product_classifier)
    by_conta = {lancamento.conta_id: lancamento for lancamento in lancamentos}

    assert by_conta["2.01.001"].valor == Decimal("546.08")
    assert by_conta["2.05.003"].valor == Decimal("80.91")
    assert by_conta["2.05.001"].valor == Decimal("16.65")
    assert all("VALOR_INCONSISTENTE" not in l.alertas for l in lancamentos)
    assert by_conta["2.05.003"].duplicatas == []


def test_cfop_1949_with_safe_nfref_reduces_origin_account(product_classifier):
    origem = _synthetic_nfe_with_products(
        chave="NFE-ORIGEM-CERVEJA",
        natureza_operacao="VENDA DE MERCADORIAS",
        produtos=[
            _produto(
                "CERV HEINEKEN PIL 0,60GFA RT 24UN",
                "5405",
                Decimal("249.13"),
                codigo="CERV1",
            )
        ],
    )
    estorno = _synthetic_nfe_with_products(
        chave="NFE-ESTORNO-CERVEJA",
        natureza_operacao="Merc nao entregue ou recusada",
        produtos=[
            _produto(
                "CERV HEINEKEN PIL 0,60GFA RT 24UN",
                "1949",
                Decimal("249.13"),
                codigo="CERV1",
            )
        ],
    )
    estorno.numero = "359680"
    estorno.referenced_nfe_chaves = (origem.chave,)

    lancamentos = allocate_nfe_to_lancamentos(
        estorno, product_classifier, nfe_index={origem.chave: origem}
    )

    assert len(lancamentos) == 1
    assert lancamentos[0].conta_id == "2.01.001"
    assert lancamentos[0].valor == Decimal("-249.13")
    assert lancamentos[0].status_revisao == "OK"
    assert lancamentos[0].cfop == "1949"
    assert lancamentos[0].tipo_operacao_especial == "ESTORNO_NF_REFERENCIADA"


def test_cfop_1949_with_multiple_origin_accounts_reduces_each_account(
    product_classifier,
):
    origem = _synthetic_nfe_with_products(
        chave="NFE-ORIGEM-MULTI",
        natureza_operacao="VENDA DE MERCADORIAS",
        produtos=[
            _produto("CERV LOK PILS LT 350ML", "5405", Decimal("54.58"), codigo="CERV"),
            _produto(
                "CACHACA CABARE PRATA GRF 700ML",
                "5102",
                Decimal("27.46"),
                codigo="DEST",
            ),
        ],
    )
    estorno = _synthetic_nfe_with_products(
        chave="NFE-ESTORNO-MULTI",
        natureza_operacao="Merc nao entregue ou recusada",
        produtos=[
            _produto("CERV LOK PILS LT 350ML", "1949", Decimal("54.58"), codigo="CERV"),
            _produto(
                "CACHACA CABARE PRATA GRF 700ML",
                "1949",
                Decimal("27.46"),
                codigo="DEST",
            ),
        ],
    )
    estorno.referenced_nfe_chaves = (origem.chave,)

    lancamentos = allocate_nfe_to_lancamentos(
        estorno, product_classifier, nfe_index={origem.chave: origem}
    )
    by_conta = {lancamento.conta_id: lancamento for lancamento in lancamentos}

    assert by_conta["2.01.001"].valor == Decimal("-54.58")
    assert by_conta["2.01.003"].valor == Decimal("-27.46")
    assert {lancamento.status_revisao for lancamento in lancamentos} == {"OK"}


def test_cfop_1949_without_safe_nfref_stays_revisar(product_classifier):
    estorno = _synthetic_nfe_with_products(
        chave="NFE-ESTORNO-SEM-REF",
        natureza_operacao="Merc nao entregue ou recusada",
        produtos=[
            _produto("CERV HEINEKEN PIL 0,60GFA RT 24UN", "1949", Decimal("249.13"))
        ],
    )

    lancamentos = allocate_nfe_to_lancamentos(estorno, product_classifier)

    assert len(lancamentos) == 1
    assert lancamentos[0].conta_id == "13.01.001"
    assert lancamentos[0].status_revisao == "REVISAR"
    assert "CFOP_1949_SEM_VINCULO_SEGURO" in lancamentos[0].alertas


def test_cfop_1949_with_missing_origin_stays_revisar(product_classifier):
    estorno = _synthetic_nfe_with_products(
        chave="NFE-ESTORNO-ORIGEM-AUSENTE",
        natureza_operacao="Merc nao entregue ou recusada",
        produtos=[
            _produto("CERV HEINEKEN PIL 0,60GFA RT 24UN", "1949", Decimal("249.13"))
        ],
    )
    estorno.referenced_nfe_chaves = ("NFE-ORIGEM-INEXISTENTE",)

    lancamentos = allocate_nfe_to_lancamentos(
        estorno, product_classifier, nfe_index={}
    )

    assert len(lancamentos) == 1
    assert lancamentos[0].status_revisao == "REVISAR"
    assert "CFOP_1949_SEM_VINCULO_SEGURO" in lancamentos[0].alertas


def test_cfop_1949_with_unmatched_origin_product_stays_revisar(product_classifier):
    origem = _synthetic_nfe_with_products(
        chave="NFE-ORIGEM-DIVERGENTE",
        natureza_operacao="VENDA DE MERCADORIAS",
        produtos=[
            _produto(
                "CERV HEINEKEN PIL 0,60GFA RT 24UN",
                "5405",
                Decimal("249.13"),
                codigo="CERV1",
            )
        ],
    )
    estorno = _synthetic_nfe_with_products(
        chave="NFE-ESTORNO-DIVERGENTE",
        natureza_operacao="Merc nao entregue ou recusada",
        produtos=[
            _produto(
                "CERV AMSTEL LAGER 0,60L GFA RT 24UN",
                "1949",
                Decimal("249.13"),
                codigo="CERV1",
            )
        ],
    )
    estorno.referenced_nfe_chaves = (origem.chave,)

    lancamentos = allocate_nfe_to_lancamentos(
        estorno, product_classifier, nfe_index={origem.chave: origem}
    )

    assert len(lancamentos) == 1
    assert lancamentos[0].status_revisao == "REVISAR"
    assert "CFOP_1949_SEM_VINCULO_SEGURO" in lancamentos[0].alertas


def test_cfop_1949_quantity_difference_hidden_by_money_rounding_stays_revisar(
    product_classifier,
):
    origem = _synthetic_nfe_with_products(
        chave="NFE-ORIGEM-QTD-PRECISA",
        natureza_operacao="VENDA DE MERCADORIAS",
        produtos=[
            _produto(
                "CERV HEINEKEN PIL 0,60GFA RT 24UN",
                "5405",
                Decimal("249.13"),
                codigo="CERV1",
                quantidade=Decimal("1.001"),
            )
        ],
    )
    estorno = _synthetic_nfe_with_products(
        chave="NFE-ESTORNO-QTD-PRECISA",
        natureza_operacao="Merc nao entregue ou recusada",
        produtos=[
            _produto(
                "CERV HEINEKEN PIL 0,60GFA RT 24UN",
                "1949",
                Decimal("249.13"),
                codigo="CERV1",
                quantidade=Decimal("1.002"),
            )
        ],
    )
    estorno.referenced_nfe_chaves = (origem.chave,)

    lancamentos = allocate_nfe_to_lancamentos(
        estorno, product_classifier, nfe_index={origem.chave: origem}
    )

    assert len(lancamentos) == 1
    assert lancamentos[0].status_revisao == "REVISAR"
    assert lancamentos[0].valor == Decimal("249.13")
    assert "CFOP_1949_SEM_VINCULO_SEGURO" in lancamentos[0].alertas


def test_cfop_1949_with_unsafe_quadrature_stays_revisar(product_classifier):
    origem = _synthetic_nfe_with_products(
        chave="NFE-ORIGEM-QUADRATURA",
        natureza_operacao="VENDA DE MERCADORIAS",
        produtos=[
            _produto(
                "CERV HEINEKEN PIL 0,60GFA RT 24UN",
                "5405",
                Decimal("100.00"),
                codigo="CERV1",
            )
        ],
    )
    origem.valor_total = Decimal("110.00")
    estorno = _synthetic_nfe_with_products(
        chave="NFE-ESTORNO-QUADRATURA",
        natureza_operacao="Merc nao entregue ou recusada",
        produtos=[
            _produto(
                "CERV HEINEKEN PIL 0,60GFA RT 24UN",
                "1949",
                Decimal("100.00"),
                codigo="CERV1",
            )
        ],
    )
    estorno.valor_total = Decimal("110.00")
    estorno.valor_frete = Decimal("5.00")
    estorno.referenced_nfe_chaves = (origem.chave,)

    lancamentos = allocate_nfe_to_lancamentos(
        estorno, product_classifier, nfe_index={origem.chave: origem}
    )

    assert len(lancamentos) == 1
    assert lancamentos[0].status_revisao == "REVISAR"
    assert lancamentos[0].valor == Decimal("110.00")
    assert "CFOP_1949_QUADRATURA_INSEGURA" in lancamentos[0].alertas


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


def _produto(
    descricao: str,
    cfop: str,
    valor: Decimal,
    *,
    codigo: str = "P",
    quantidade: Decimal = Decimal("1"),
    valor_desconto: Decimal = Decimal("0.00"),
) -> ProdutoNFE:
    return ProdutoNFE(
        n_item=1,
        codigo=codigo,
        descricao=descricao,
        ncm="22021000",
        cfop=cfop,
        unidade="UN",
        quantidade=quantidade,
        valor_unitario=valor,
        valor_total=valor,
        valor_desconto=valor_desconto,
    )
