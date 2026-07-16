from __future__ import annotations

from decimal import Decimal
from datetime import date
from pathlib import Path

import pytest

from fantozzi.nfse_importer import NFSE
from fantozzi.nfse_importer import parse_nfse
from fantozzi.nfse_service_classifier import (
    ORIGEM_XML_NFSE,
    STATUS_CLASSIFICADO_XML_NFSE,
    STATUS_COMPROVANTE_SOLO_RECONCILIACAO,
    STATUS_DUPLICADO_XML_NFSE,
    STATUS_REVISAR_CHAVE_NFSE_INSUFICIENTE,
    STATUS_REVISAR_FORNECEDOR_XML_ONLY_BLOCCATO,
    STATUS_REVISAR_SEM_REGRA_NFSE,
    STATUS_REVISAR_SEM_XML_AUTORIZZATO,
    classify_nfse_service,
    classify_nfse_service_files,
    load_nfse_service_rules,
    nfse_dedup_key,
    reconcile_xml_only_receipt,
    NfseServiceRulesError,
)


ROOT = Path(__file__).resolve().parents[1]
CONFIG = ROOT / "config"
RULES = CONFIG / "regras_classificacao_nfse_servicos.yaml"
PLANO = CONFIG / "plano_contas_gerencial.yaml"
MAPEAMENTO = CONFIG / "mapeamento_dre.yaml"


def _rules():
    return load_nfse_service_rules(RULES, PLANO, MAPEAMENTO)


def _nfse_xml(
    path: Path,
    *,
    numero: str,
    codigo: str,
    prestador_cnpj: str,
    prestador_nome: str,
    tomador_cnpj: str = "08.884.180/0001-23",
    tomador_nome: str = "AMAMI RESTAURANTE LTDA",
    valor: str = "350,00",
    discriminacao: str,
) -> Path:
    path.write_text(
        f"""<CompNfse xmlns="http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd">
  <Nfse>
    <InfNfse>
      <Numero>{numero}</Numero>
      <CodigoVerificacao>{codigo}</CodigoVerificacao>
      <DataEmissao>2026-01-13T10:00:00-03:00</DataEmissao>
      <Competencia>2026-01-01T00:00:00-03:00</Competencia>
      <Servico>
        <Valores>
          <ValorServicos>{valor}</ValorServicos>
          <ValorLiquidoNfse>{valor}</ValorLiquidoNfse>
        </Valores>
        <Discriminacao>{discriminacao}</Discriminacao>
      </Servico>
      <PrestadorServico>
        <IdentificacaoPrestador><Cnpj>{prestador_cnpj}</Cnpj></IdentificacaoPrestador>
        <RazaoSocial>{prestador_nome}</RazaoSocial>
      </PrestadorServico>
      <TomadorServico>
        <IdentificacaoTomador><CpfCnpj><Cnpj>{tomador_cnpj}</Cnpj></CpfCnpj></IdentificacaoTomador>
        <RazaoSocial>{tomador_nome}</RazaoSocial>
      </TomadorServico>
    </InfNfse>
  </Nfse>
</CompNfse>
""",
        encoding="utf-8",
    )
    return path


def test_parse_nfse_exposes_fields_needed_by_service_rules(tmp_path):
    xml = _nfse_xml(
        tmp_path / "akira.xml",
        numero="9",
        codigo="AKIRA9",
        prestador_cnpj="07.690.972/0001-02",
        prestador_nome="AKIRA DESINSETIZADORA E DESENTUPIDORA LTDA",
        discriminacao="DEDETIZACAO PREVENTIVA",
    )

    nfse = parse_nfse(xml)

    assert nfse.cnpj_prestador == "07690972000102"
    assert nfse.cnpj_tomador == "08884180000123"
    assert nfse.valor_liquido == Decimal("350.00")
    assert nfse.codigo_verificacao == "AKIRA9"
    assert nfse.discriminacao == "DEDETIZACAO PREVENTIVA"


def test_akira_nfse_classifies_to_canonical_account_dre_and_cost_center(tmp_path):
    xml = _nfse_xml(
        tmp_path / "akira.xml",
        numero="9",
        codigo="AKIRA9",
        prestador_cnpj="07.690.972/0001-02",
        prestador_nome="AKIRA DESINSETIZADORA E DESENTUPIDORA LTDA",
        discriminacao="DEDETIZACAO PREVENTIVA NO RESTAURANTE",
    )

    result = classify_nfse_service(parse_nfse(xml), _rules())

    assert result.status == STATUS_CLASSIFICADO_XML_NFSE
    assert result.regra_id == "NFSE_SERVICO_AKIRA_DEDETIZACAO_4_01_015"
    assert result.lancamento is not None
    assert result.lancamento.origem == ORIGEM_XML_NFSE
    assert result.lancamento.conta_id == "4.01.015"
    assert result.lancamento.dre_linha_id == "06.02"
    assert result.lancamento.centro_custo == "CC-006 - Limpeza e Higiene"
    assert result.lancamento.valor == Decimal("350.00")


def test_alex_inpi_future_nfse_classifies_without_sample_verification_code(tmp_path):
    xml = _nfse_xml(
        tmp_path / "alex.xml",
        numero="9999",
        codigo="FUTURO123",
        prestador_cnpj="42.216.267/0001-60",
        prestador_nome="ALEX STRAUCH BATISTA",
        valor="1440,20",
        discriminacao="Acompanhamento de marca junto ao INPI",
    )

    result = classify_nfse_service(parse_nfse(xml), _rules())

    assert result.status == STATUS_CLASSIFICADO_XML_NFSE
    assert result.regra_id == "NFSE_SERVICO_ALEX_INPI_6_01_002"
    assert result.lancamento is not None
    assert result.lancamento.conta_id == "6.01.002"
    assert result.lancamento.dre_linha_id == "08.01"
    assert result.lancamento.centro_custo == "CC-005 - Administracao"
    assert result.lancamento.valor == Decimal("1440.20")


def test_alex_generic_service_does_not_match_inpi_rule(tmp_path):
    xml = _nfse_xml(
        tmp_path / "alex_generic.xml",
        numero="4743",
        codigo="OUTRO",
        prestador_cnpj="42.216.267/0001-60",
        prestador_nome="ALEX STRAUCH BATISTA",
        valor="1000,00",
        discriminacao="Consultoria administrativa mensal",
    )

    result = classify_nfse_service(parse_nfse(xml), _rules())

    assert result.status == STATUS_REVISAR_SEM_REGRA_NFSE
    assert result.lancamento is None


def test_same_codigo_verificacao_from_different_providers_does_not_collide(tmp_path):
    first = parse_nfse(
        _nfse_xml(
            tmp_path / "first.xml",
            numero="1",
            codigo="CODIGO-IGUAL",
            prestador_cnpj="42.216.267/0001-60",
            prestador_nome="ALEX STRAUCH BATISTA",
            discriminacao="Acompanhamento de marca junto ao INPI",
        )
    )
    second = parse_nfse(
        _nfse_xml(
            tmp_path / "second.xml",
            numero="1",
            codigo="CODIGO-IGUAL",
            prestador_cnpj="07.690.972/0001-02",
            prestador_nome="AKIRA DESINSETIZADORA E DESENTUPIDORA LTDA",
            discriminacao="DEDETIZACAO PREVENTIVA",
        )
    )

    assert nfse_dedup_key(first) != nfse_dedup_key(second)
    assert "CODIGO-IGUAL" not in nfse_dedup_key(first)


def test_nfse_with_insufficient_key_stays_review():
    nfse = NFSE(
        tipo_documento="NFSE",
        numero="",
        codigo_verificacao="",
        chave_acesso="",
        data_emissao=date(2026, 1, 13),
        competencia=date(2026, 1, 1),
        cnpj_prestador="07690972000102",
        razao_social_prestador="AKIRA DESINSETIZADORA E DESENTUPIDORA LTDA",
        cnpj_tomador="08884180000123",
        razao_social_tomador="AMAMI RESTAURANTE LTDA",
        valor_servicos=Decimal("350.00"),
        valor_liquido=Decimal("350.00"),
        discriminacao="DEDETIZACAO PREVENTIVA",
        arquivo_origem="memoria.xml",
    )

    result = classify_nfse_service(nfse, _rules())

    assert result.status == STATUS_REVISAR_CHAVE_NFSE_INSUFICIENTE
    assert result.lancamento is None


def test_xml_only_receipt_without_xml_stays_review():
    diagnostic = reconcile_xml_only_receipt(
        {
            "beneficiary_document": "07.690.972/0001-02",
            "beneficiary_name": "AKIRA DESINSETIZADORA E DESENT",
            "amount": "350.00",
        },
        [],
        _rules(),
    )

    assert diagnostic.status == STATUS_REVISAR_SEM_XML_AUTORIZZATO
    assert diagnostic.motivo == "XML_CLASSIFICABILE_ASSENTE"


def test_xml_and_receipt_create_one_cost_and_receipt_only_reconciliation(tmp_path):
    xml = _nfse_xml(
        tmp_path / "akira.xml",
        numero="9",
        codigo="AKIRA9",
        prestador_cnpj="07.690.972/0001-02",
        prestador_nome="AKIRA DESINSETIZADORA E DESENTUPIDORA LTDA",
        discriminacao="DEDETIZACAO PREVENTIVA",
    )
    batch = classify_nfse_service_files([xml], _rules())

    diagnostic = reconcile_xml_only_receipt(
        {
            "beneficiary_document": "07.690.972/0001-02",
            "beneficiary_name": "AKIRA DESINSETIZADORA E DESENT",
            "amount": "350.00",
        },
        batch.diagnostics,
        _rules(),
    )

    assert len(batch.lancamentos) == 1
    assert diagnostic.status == STATUS_COMPROVANTE_SOLO_RECONCILIACAO
    assert diagnostic.xml_valor == Decimal("350.00")
    assert diagnostic.comprovante_valor == Decimal("350.00")
    assert diagnostic.diferenca == Decimal("0.00")
    assert diagnostic.diferenca_efeito_contabil == "NENHUM"


def test_reexecution_batch_does_not_duplicate_same_nfse(tmp_path):
    xml = _nfse_xml(
        tmp_path / "akira.xml",
        numero="9",
        codigo="AKIRA9",
        prestador_cnpj="07.690.972/0001-02",
        prestador_nome="AKIRA DESINSETIZADORA E DESENTUPIDORA LTDA",
        discriminacao="DEDETIZACAO PREVENTIVA",
    )

    batch = classify_nfse_service_files([xml, xml], _rules())

    assert len(batch.lancamentos) == 1
    assert [item.status for item in batch.diagnostics] == [
        STATUS_CLASSIFICADO_XML_NFSE,
        STATUS_DUPLICADO_XML_NFSE,
    ]


def test_alex_small_difference_is_diagnostic_only(tmp_path):
    xml = _nfse_xml(
        tmp_path / "alex.xml",
        numero="4742",
        codigo="ZSLDD6CV",
        prestador_cnpj="42.216.267/0001-60",
        prestador_nome="ALEX STRAUCH BATISTA",
        valor="1440,20",
        discriminacao="Acompanhamento de marca junto ao INPI",
    )
    batch = classify_nfse_service_files([xml], _rules())

    diagnostic = reconcile_xml_only_receipt(
        {
            "beneficiary_document": "42.216.267/0001-60",
            "beneficiary_name": "ALEX STRAUCH BATISTA",
            "amount": "1440.05",
        },
        batch.diagnostics,
        _rules(),
    )

    assert len(batch.lancamentos) == 1
    assert batch.lancamentos[0].valor == Decimal("1440.20")
    assert diagnostic.status == STATUS_COMPROVANTE_SOLO_RECONCILIACAO
    assert diagnostic.diferenca == Decimal("0.15")
    assert diagnostic.diferenca_efeito_contabil == "NENHUM"


def test_rei_policy_is_registered_but_blocked_for_005e_b():
    diagnostic = reconcile_xml_only_receipt(
        {
            "beneficiary_document": "56.198.626/0001-09",
            "beneficiary_name": "REI DAS BOMBAS MANUTENCAO E RE",
            "amount": "2980.93",
        },
        [],
        _rules(),
    )

    assert diagnostic.status == STATUS_REVISAR_FORNECEDOR_XML_ONLY_BLOCCATO
    assert diagnostic.motivo == "POLICY_XML_ONLY_BLOCCATA_005E_B"


def test_nfse_rules_reject_duplicate_rule_id(tmp_path):
    rules_text = RULES.read_text(encoding="utf-8").replace(
        "NFSE_SERVICO_ALEX_INPI_6_01_002",
        "NFSE_SERVICO_AKIRA_DEDETIZACAO_4_01_015",
    )
    rules_path = tmp_path / "regras.yaml"
    rules_path.write_text(rules_text, encoding="utf-8")

    with pytest.raises(NfseServiceRulesError, match="regra_id duplicado"):
        load_nfse_service_rules(rules_path, PLANO, MAPEAMENTO)


def test_nfse_rules_reject_generic_discrimination_term(tmp_path):
    rules_text = RULES.read_text(encoding="utf-8").replace(
        "- INPI",
        "- SERVICO",
    )
    rules_path = tmp_path / "regras.yaml"
    rules_path.write_text(rules_text, encoding="utf-8")

    with pytest.raises(NfseServiceRulesError, match="generica demais"):
        load_nfse_service_rules(rules_path, PLANO, MAPEAMENTO)
