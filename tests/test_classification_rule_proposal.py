import csv
from pathlib import Path

import pytest
import yaml

from fantozzi.classification_review_export import CSV_HEADERS
from fantozzi.classification_rule_proposal import (
    ClassificationRuleProposalError,
    escrever_proposta_yaml,
    gerar_proposta_classificacao,
    gerar_relatorio_proposta_classificacao,
)


def test_csv_valido_gera_proposta_yaml(tmp_path):
    csv_path = tmp_path / "revisao_classificacao.csv"
    yaml_path = tmp_path / "reports" / "proposta.yaml"
    plano_path = _save_plano_contas(tmp_path)
    regras_path = _save_regras_produtos(tmp_path)
    _save_csv(
        csv_path,
        [
            _row(
                descricao_normalizada="refrigerante teste",
                acao_manual="CRIAR_ALIAS",
                conta_id_sugerida="2.01.004",
                palavras_chave_sugeridas="refri teste",
            )
        ],
    )

    result = gerar_proposta_classificacao(csv_path, plano_path, regras_path)
    escrever_proposta_yaml(result, yaml_path)

    assert len(result.propostas) == 1
    assert yaml_path.exists()
    data = yaml.safe_load(yaml_path.read_text(encoding="utf-8"))
    assert data["propostas"][0]["tipo"] == "alias"
    assert data["propostas"][0]["termo"] == "refri teste"
    assert data["propostas"][0]["conta_id"] == "2.01.004"
    assert "PROPOSTA GERADA AUTOMATICAMENTE" in yaml_path.read_text(encoding="utf-8")


def test_linha_pendente_e_ignorada(tmp_path):
    csv_path = tmp_path / "pendente.csv"
    plano_path = _save_plano_contas(tmp_path)
    regras_path = _save_regras_produtos(tmp_path)
    _save_csv(csv_path, [_row(acao_manual="CRIAR_ALIAS", conta_id_sugerida="")])

    result = gerar_proposta_classificacao(csv_path, plano_path, regras_path)

    assert result.linhas_pendentes == 1
    assert result.propostas == ()


def test_conta_id_sugerida_desconhecida_nao_entra_como_proposta_valida(tmp_path):
    csv_path = tmp_path / "conta_desconhecida.csv"
    plano_path = _save_plano_contas(tmp_path)
    regras_path = _save_regras_produtos(tmp_path)
    _save_csv(
        csv_path,
        [_row(acao_manual="CRIAR_ALIAS", conta_id_sugerida="9.99.999")],
    )

    result = gerar_proposta_classificacao(csv_path, plano_path, regras_path)

    assert result.linhas_invalidas == 1
    assert result.contas_desconhecidas == 1
    assert result.propostas == ()
    assert result.status == "NAO FOI POSSIVEL GERAR PROPOSTA COMPLETA"


def test_acao_desconhecida_e_bloqueada_pela_validacao_5d(tmp_path):
    csv_path = tmp_path / "acao_desconhecida.csv"
    plano_path = _save_plano_contas(tmp_path)
    regras_path = _save_regras_produtos(tmp_path)
    _save_csv(
        csv_path,
        [_row(acao_manual="INVENTAR", conta_id_sugerida="2.01.004")],
    )

    result = gerar_proposta_classificacao(csv_path, plano_path, regras_path)

    assert result.linhas_invalidas == 1
    assert result.acoes_desconhecidas == 1
    assert result.propostas == ()


def test_acao_valida_sem_proposta_e_reportada(tmp_path):
    csv_path = tmp_path / "ignorar.csv"
    plano_path = _save_plano_contas(tmp_path)
    regras_path = _save_regras_produtos(tmp_path)
    _save_csv(
        csv_path,
        [_row(acao_manual="IGNORAR", conta_id_sugerida="2.01.004")],
    )

    result = gerar_proposta_classificacao(csv_path, plano_path, regras_path)
    text = gerar_relatorio_proposta_classificacao(result)

    assert result.linhas_validas_sem_proposta == 1
    assert result.propostas == ()
    assert "acao_manual valida sem proposta YAML" in text


def test_duplicidade_no_csv_e_sinalizada(tmp_path):
    csv_path = tmp_path / "duplicidade.csv"
    plano_path = _save_plano_contas(tmp_path)
    regras_path = _save_regras_produtos(tmp_path)
    _save_csv(
        csv_path,
        [
            _row(
                acao_manual="CRIAR_REGRA",
                conta_id_sugerida="2.01.004",
                palavras_chave_sugeridas="produto repetido",
            ),
            _row(
                acao_manual="CRIAR_REGRA",
                conta_id_sugerida="2.02.010",
                palavras_chave_sugeridas="produto repetido",
            ),
        ],
    )

    result = gerar_proposta_classificacao(csv_path, plano_path, regras_path)

    assert len(result.propostas) == 2
    assert result.duplicidades_potenciais >= 1
    assert "TERMO_DUPLICADO_NA_PROPOSTA" in result.propostas[1].alertas
    assert "CONTA_CONFLITANTE" in result.propostas[1].alertas


def test_mesma_descricao_normalizada_com_contas_diferentes_e_sinalizada(tmp_path):
    csv_path = tmp_path / "descricao_conflitante.csv"
    plano_path = _save_plano_contas(tmp_path)
    regras_path = _save_regras_produtos(tmp_path)
    _save_csv(
        csv_path,
        [
            _row(
                descricao_normalizada="produto sem regra",
                acao_manual="CRIAR_ALIAS",
                conta_id_sugerida="2.01.004",
                palavras_chave_sugeridas="termo bebida",
            ),
            _row(
                descricao_normalizada="produto sem regra",
                acao_manual="CRIAR_ALIAS",
                conta_id_sugerida="2.02.010",
                palavras_chave_sugeridas="termo mercearia",
            ),
        ],
    )

    result = gerar_proposta_classificacao(csv_path, plano_path, regras_path)

    assert len(result.propostas) == 2
    assert result.duplicidades_potenciais == 1
    assert result.propostas[1].alertas == ("CONTA_CONFLITANTE",)


def test_termo_existente_em_regras_e_sinalizado(tmp_path):
    csv_path = tmp_path / "termo_existente.csv"
    plano_path = _save_plano_contas(tmp_path)
    regras_path = _save_regras_produtos(tmp_path)
    _save_csv(
        csv_path,
        [
            _row(
                acao_manual="CRIAR_ALIAS",
                conta_id_sugerida="2.01.004",
                palavras_chave_sugeridas="coca cola",
            )
        ],
    )

    result = gerar_proposta_classificacao(csv_path, plano_path, regras_path)

    assert result.propostas[0].alertas == ("TERMO_JA_EXISTE_EM_REGRAS",)


def test_arquivo_yaml_original_nao_e_modificado(tmp_path):
    csv_path = tmp_path / "revisao.csv"
    plano_path = _save_plano_contas(tmp_path)
    regras_path = _save_regras_produtos(tmp_path)
    original_rules = regras_path.read_text(encoding="utf-8")
    original_plan = plano_path.read_text(encoding="utf-8")
    _save_csv(
        csv_path,
        [_row(acao_manual="CRIAR_ALIAS", conta_id_sugerida="2.01.004")],
    )

    gerar_proposta_classificacao(csv_path, plano_path, regras_path)

    assert regras_path.read_text(encoding="utf-8") == original_rules
    assert plano_path.read_text(encoding="utf-8") == original_plan


def test_report_txt_contem_totais_e_status_correto(tmp_path):
    csv_path = tmp_path / "revisao.csv"
    plano_path = _save_plano_contas(tmp_path)
    regras_path = _save_regras_produtos(tmp_path)
    _save_csv(
        csv_path,
        [_row(acao_manual="CRIAR_REGRA", conta_id_sugerida="2.01.004")],
    )

    result = gerar_proposta_classificacao(csv_path, plano_path, regras_path)
    text = gerar_relatorio_proposta_classificacao(
        result,
        csv_path,
        tmp_path / "reports" / "proposta.yaml",
    )

    assert "Total de linhas no CSV: 1" in text
    assert "Propostas geradas: 1" in text
    assert "Propostas de regra: 1" in text
    assert "Status: PROPOSTA GERADA PARA REVISAO MANUAL" in text


def test_missing_csv_raises_clear_error(tmp_path):
    plano_path = _save_plano_contas(tmp_path)
    regras_path = _save_regras_produtos(tmp_path)

    with pytest.raises(ClassificationRuleProposalError, match="Arquivo CSV nao encontrado"):
        gerar_proposta_classificacao(tmp_path / "inexistente.csv", plano_path, regras_path)


def _save_csv(csv_path: Path, rows: list[dict[str, str]]) -> None:
    with csv_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=CSV_HEADERS, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


def _row(**overrides: str) -> dict[str, str]:
    row = {
        "tipo_candidato": "CLASSIFICACAO_PRODUTO",
        "descricao_normalizada": "produto desconhecido",
        "descricao_exemplo": "PRODUTO DESCONHECIDO",
        "ocorrencias": "1",
        "fornecedores": "Fornecedor teste",
        "documentos": "NF-1",
        "valor_total": "R$ 10,00",
        "conta_atual_id": "13.01.001",
        "conta_atual_nome": "Revisar - Produto Nao Classificado",
        "status_revisao": "REVISAR",
        "alertas": "CLASSIFICACAO_NAO_ENCONTRADA",
        "acao_manual": "",
        "conta_id_sugerida": "",
        "conta_nome_sugerida": "",
        "palavras_chave_sugeridas": "",
        "observacao_manual": "",
    }
    row.update(overrides)
    return row


def _save_plano_contas(tmp_path: Path) -> Path:
    plano_path = tmp_path / "plano_contas_gerencial.yaml"
    plano_path.write_text(
        """\
contas:
  - id: "2.01.004"
    nome: "Refrigerantes"
    grupo: "Custos de Mercadorias e Insumos"
    tipo: "custo"
  - id: "2.02.010"
    nome: "Mercearia Seca"
    grupo: "Custos de Mercadorias e Insumos"
    tipo: "custo"
  - id: "13.01.001"
    nome: "Revisar - Produto Nao Classificado"
    grupo: "Contas de Revisao e Classificacao Pendente"
    tipo: "revisao"
  - id: "13.01.002"
    nome: "Revisar - Classificacao Ambigua"
    grupo: "Contas de Revisao e Classificacao Pendente"
    tipo: "revisao"
  - id: "13.01.006"
    nome: "Revisar - Valor Inconsistente"
    grupo: "Contas de Revisao e Classificacao Pendente"
    tipo: "revisao"
  - id: "13.01.009"
    nome: "Revisar - Possivel Investimento"
    grupo: "Contas de Revisao e Classificacao Pendente"
    tipo: "revisao"
  - id: "13.01.010"
    nome: "Revisar - Pagamento de Pessoal a Regularizar"
    grupo: "Contas de Revisao e Classificacao Pendente"
    tipo: "revisao"
  - id: "13.01.999"
    nome: "Revisar - Outros Casos"
    grupo: "Contas de Revisao e Classificacao Pendente"
    tipo: "revisao"
""",
        encoding="utf-8",
    )
    return plano_path


def _save_regras_produtos(tmp_path: Path) -> Path:
    regras_path = tmp_path / "regras_classificacao_produtos.yaml"
    regras_path.write_text(
        """\
versao: "1.0"
regras:
  - conta_id: "2.01.004"
    conta_nome: "Refrigerantes"
    aliases:
      - coca cola
regras_por_fornecedor: []
""",
        encoding="utf-8",
    )
    return regras_path
