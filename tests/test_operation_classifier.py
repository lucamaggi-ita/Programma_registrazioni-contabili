from datetime import date
from decimal import Decimal

from fantozzi.models import Duplicata, NFE, ProdutoNFE
from fantozzi.operation_classifier import (
    TIPO_BONIFICACAO,
    TIPO_COMODATO,
    TIPO_COMPRA_NORMAL,
    TIPO_OUTRA_REVISAO,
    classificar_operacao_especial,
    normalizar_texto_operacao,
)


def test_cfop_5908_with_comodato_nature_is_strong_comodato():
    result = classificar_operacao_especial(_nfe("Comodato", ["5908"]))

    assert result.tipo == TIPO_COMODATO
    assert result.confianca == "FORTE"
    assert result.destino_gerencial == "FORA_DRE"
    assert result.status_revisao == "ANALISAR_COM_CONTADOR"


def test_cfop_6908_with_comodato_nature_is_strong_comodato():
    result = classificar_operacao_especial(_nfe("COMODATO - SAIDA", ["6908"]))

    assert result.tipo == TIPO_COMODATO
    assert result.confianca == "FORTE"


def test_cfop_5909_with_retorno_comodato_is_comodato():
    result = classificar_operacao_especial(
        _nfe("RETORNO DE BEM RECEBIDO POR COMODATO", ["5909"])
    )

    assert result.tipo == TIPO_COMODATO
    assert result.confianca == "FORTE"


def test_comodato_nature_with_unexpected_cfop_requires_review():
    result = classificar_operacao_especial(_nfe("Comodato", ["5102"]))

    assert result.tipo == TIPO_COMODATO
    assert result.confianca == "REVISAR"
    assert "CFOP_NAO_COMPATIVEL_COM_COMODATO" in result.alertas


def test_cfop_5910_with_bonificacao_nature_is_strong_bonificacao():
    result = classificar_operacao_especial(_nfe("Bonificacao", ["5910"]))

    assert result.tipo == TIPO_BONIFICACAO
    assert result.confianca == "FORTE"
    assert result.destino_gerencial == "FORA_CUSTO_COMPRA"


def test_cfop_6910_with_brinde_nature_is_strong_bonificacao():
    result = classificar_operacao_especial(_nfe("BRINDE", ["6910"]))

    assert result.tipo == TIPO_BONIFICACAO
    assert result.confianca == "FORTE"


def test_doacao_with_accent_is_detected():
    assert normalizar_texto_operacao("Doação") == "DOACAO"

    result = classificar_operacao_especial(_nfe("Doação", ["5102"]))

    assert result.tipo == TIPO_BONIFICACAO
    assert result.confianca == "REVISAR"


def test_refrigerador_description_is_secondary_alert_only():
    result = classificar_operacao_especial(
        _nfe("Venda", ["5102"], descriptions=["REFRIGERADOR SKOL NORMAL 220W"])
    )

    assert result.tipo == TIPO_COMPRA_NORMAL
    assert result.confianca == "REVISAR"
    assert "POSSIVEL_EQUIPAMENTO_OU_BEM_DE_TERCEIRO" in result.alertas


def test_absence_of_duplicata_without_special_cfop_or_nature_does_not_classify():
    result = classificar_operacao_especial(
        _nfe("Venda", ["5102"], tem_duplicatas=False)
    )

    assert result.tipo == TIPO_COMPRA_NORMAL
    assert result.status_revisao == "OK"
    assert not result.alertas


def test_normal_purchase_is_compra_normal():
    result = classificar_operacao_especial(_nfe("Venda de mercadorias", ["5102"]))

    assert result.tipo == TIPO_COMPRA_NORMAL
    assert result.confianca == "NAO_APLICAVEL"
    assert result.destino_gerencial == "DRE"


def test_conflict_between_comodato_and_bonificacao_requires_review():
    result = classificar_operacao_especial(_nfe("Comodato e Bonificacao", ["5908", "5910"]))

    assert result.tipo == TIPO_OUTRA_REVISAO
    assert result.confianca == "REVISAR"
    assert result.status_revisao == "ANALISAR_COM_CONTADOR"


def _nfe(
    natureza: str,
    cfops: list[str],
    *,
    descriptions: list[str] | None = None,
    tem_duplicatas: bool = True,
) -> NFE:
    descriptions = descriptions or [f"Produto {index + 1}" for index in range(len(cfops))]
    produtos = [
        ProdutoNFE(
            n_item=index + 1,
            codigo=f"P{index + 1}",
            descricao=description,
            ncm="00000000",
            cfop=cfop,
            unidade="UN",
            quantidade=Decimal("1"),
            valor_unitario=Decimal("10.00"),
            valor_total=Decimal("10.00"),
        )
        for index, (cfop, description) in enumerate(zip(cfops, descriptions))
    ]
    duplicatas = (
        [Duplicata(numero="001", vencimento=None, valor=Decimal("10.00"))]
        if tem_duplicatas
        else []
    )
    return NFE(
        chave="chave",
        numero="1",
        serie="1",
        data_emissao=date(2026, 1, 1),
        cnpj_emitente="11111111000111",
        nome_emitente="Fornecedor",
        cnpj_destinatario="22222222000122",
        nome_destinatario="Empresa",
        valor_total=sum((produto.valor_total for produto in produtos), Decimal("0")),
        natureza_operacao=natureza,
        produtos=produtos,
        duplicatas=duplicatas,
        tem_duplicatas=tem_duplicatas,
    )
