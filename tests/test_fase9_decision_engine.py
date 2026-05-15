import csv
from pathlib import Path

from fantozzi.fase9_decision_engine import (
    Fase9Decision,
    Fase9DecisionReason,
    evaluate_fase9_decisions,
    load_fase9_decisions_from_csv,
)


VALID_ACCOUNTS = frozenset({"2.01.004", "2.02.010", "2.05.003"})


def test_classificar_valido_produz_promotion():
    result = evaluate_fase9_decisions(
        [_decision(action="CLASSIFICAR", account_id="2.01.004")],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert len(result.promotions) == 1
    assert result.promotions[0].reason is Fase9DecisionReason.PROMOTION_CANDIDATE
    assert result.promotions[0].account_id == "2.01.004"
    assert result.promotions[0].description == "REFRIGERANTE COLA 350ML"
    assert not result.blocked
    assert not result.ignored


def test_criar_alias_valido_produz_promotion():
    result = evaluate_fase9_decisions(
        [
            _decision(
                action="CRIAR_ALIAS",
                description="COCA COLA ORIGINAL LATA 350ML",
                descricao_normalizada="coca cola original lata 350ml",
                account_id="2.01.004",
            )
        ],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert len(result.promotions) == 1
    assert result.promotions[0].action == "CRIAR_ALIAS"
    assert result.promotions[0].rule_terms == ("coca cola original lata 350ml",)


def test_criar_regra_valido_produz_promotion():
    result = evaluate_fase9_decisions(
        [
            _decision(
                action="CRIAR_REGRA",
                account_id="2.02.010",
                keywords=("arroz arborio", "feijao carioca"),
            )
        ],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert len(result.promotions) == 1
    assert result.promotions[0].rule_terms == ("arroz arborio", "feijao carioca")


def test_manter_revisao_nao_produz_promotion():
    result = evaluate_fase9_decisions(
        [_decision(action="MANTER_REVISAO", account_id="2.01.004")],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert not result.promotions
    assert result.blocked[0].reason is Fase9DecisionReason.MANTER_REVISAO


def test_analisar_com_contador_nao_produz_promotion():
    result = evaluate_fase9_decisions(
        [_decision(action="ANALISAR_COM_CONTADOR", account_id="2.01.004")],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert not result.promotions
    assert result.blocked[0].reason is Fase9DecisionReason.ANALISAR_COM_CONTADOR


def test_ignorar_nao_produz_promotion_e_fica_explicito():
    result = evaluate_fase9_decisions(
        [_decision(action="IGNORAR", account_id="2.01.004")],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert not result.promotions
    assert not result.blocked
    assert result.ignored[0].reason is Fase9DecisionReason.IGNORAR


def test_comodato_resta_bloqueado_mesmo_com_classificar():
    result = evaluate_fase9_decisions(
        [
            _decision(
                action="CLASSIFICAR",
                account_id="2.01.004",
                tipo_operacao_especial="COMODATO",
                tratamento_operacao_especial="FORA_DRE",
            )
        ],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert not result.promotions
    assert result.blocked[0].reason is Fase9DecisionReason.COMODATO_FORA_DRE


def test_nf_359680_resta_bloqueada():
    result = evaluate_fase9_decisions(
        [
            _decision(
                action="CLASSIFICAR",
                account_id="2.05.003",
                numero_nota="359680",
            )
        ],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert not result.promotions
    assert result.blocked[0].reason is Fase9DecisionReason.NF_359680_CONTROLE_MANUAL


def test_cfop_1949_resta_bloqueado():
    result = evaluate_fase9_decisions(
        [_decision(action="CLASSIFICAR", account_id="2.05.003", cfop="1949")],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert not result.promotions
    assert result.blocked[0].reason is Fase9DecisionReason.CFOP_1949_CONTROLE_MANUAL


def test_conta_id_sugerida_inexistente_bloqueia():
    result = evaluate_fase9_decisions(
        [_decision(action="CLASSIFICAR", account_id="9.99.999")],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert not result.promotions
    assert result.blocked[0].reason is Fase9DecisionReason.CONTA_INVALIDA


def test_conta_nome_incoerente_bloqueia_quando_mapa_fornecido():
    result = evaluate_fase9_decisions(
        [
            _decision(
                action="CLASSIFICAR",
                account_id="2.01.004",
                account_name="Mercearia Seca",
            )
        ],
        valid_account_ids=VALID_ACCOUNTS,
        account_names_by_id={"2.01.004": "Refrigerantes"},
    )

    assert not result.promotions
    assert result.blocked[0].reason is Fase9DecisionReason.CONTA_NOME_INCOERENTE


def test_status_ok_nao_e_repromovido():
    result = evaluate_fase9_decisions(
        [_decision(action="CLASSIFICAR", account_id="2.01.004", status_revisao="OK")],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert not result.promotions
    assert result.blocked[0].reason is Fase9DecisionReason.JA_OK


def test_status_diferente_de_revisar_bloqueia():
    result = evaluate_fase9_decisions(
        [
            _decision(
                action="CLASSIFICAR",
                account_id="2.01.004",
                status_revisao="PENDENTE",
            )
        ],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert not result.promotions
    assert result.blocked[0].reason is Fase9DecisionReason.STATUS_NAO_REVISAR


def test_criar_regra_com_somente_palavras_genericas_bloqueia():
    result = evaluate_fase9_decisions(
        [
            _decision(
                action="CRIAR_REGRA",
                account_id="2.02.010",
                keywords=("papel", "saco", "kg"),
            )
        ],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert not result.promotions
    assert result.blocked[0].reason is Fase9DecisionReason.PALAVRAS_GENERICAS


def test_criar_regra_com_keyword_especifica_valida_passa():
    result = evaluate_fase9_decisions(
        [
            _decision(
                action="CRIAR_REGRA",
                account_id="2.02.010",
                keywords=("papel", "arroz arborio"),
            )
        ],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert len(result.promotions) == 1


def test_criar_regra_sem_palavras_bloqueia():
    result = evaluate_fase9_decisions(
        [_decision(action="CRIAR_REGRA", account_id="2.02.010", keywords=())],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert not result.promotions
    assert result.blocked[0].reason is Fase9DecisionReason.PALAVRAS_AUSENTES


def test_criar_alias_com_descricao_generica_bloqueia():
    result = evaluate_fase9_decisions(
        [
            _decision(
                action="CRIAR_ALIAS",
                description="PAPEL",
                descricao_normalizada="papel",
                account_id="2.02.010",
            )
        ],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert not result.promotions
    assert result.blocked[0].reason is Fase9DecisionReason.ALIAS_GENERICO


def test_candidato_com_alerta_nao_resolvido_bloqueia():
    result = evaluate_fase9_decisions(
        [
            _decision(
                action="CLASSIFICAR",
                account_id="2.01.004",
                alertas="CLASSIFICACAO_AMBIGUA",
            )
        ],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert not result.promotions
    assert result.blocked[0].reason is Fase9DecisionReason.ALERTA_NAO_RESOLVIDO


def test_evaluate_nao_modifica_input_e_e_idempotente():
    decisions = (
        _decision(action="CLASSIFICAR", account_id="2.01.004"),
        _decision(action="IGNORAR", account_id="2.01.004"),
    )
    before = decisions

    first = evaluate_fase9_decisions(decisions, valid_account_ids=VALID_ACCOUNTS)
    second = evaluate_fase9_decisions(decisions, valid_account_ids=VALID_ACCOUNTS)

    assert decisions == before
    assert first == second


def test_load_fase9_decisions_from_csv_nao_escreve_arquivo(tmp_path):
    csv_path = tmp_path / "revisao_classificacao_revisada.csv"
    _write_csv(
        csv_path,
        [
            {
                "acao_manual": " classificar ",
                "descricao_normalizada": "agua mineral",
                "fornecedores": "Fornecedor A",
                "documentos": "123",
                "conta_id_sugerida": " 2.01.004 ",
                "conta_nome_sugerida": "Refrigerantes",
                "palavras_chave_sugeridas": "agua mineral; garrafa",
                "status_revisao": " revisar ",
                "alertas": "",
                "tipo_operacao_especial": "compra_normal",
                "tratamento_operacao_especial": "dre",
                "cfop": "5102",
                "valor_total": "R$ 10,00",
            }
        ],
    )
    original_bytes = csv_path.read_bytes()

    decisions = load_fase9_decisions_from_csv(csv_path)

    assert csv_path.read_bytes() == original_bytes
    assert len(decisions) == 1
    assert decisions[0].line_number == 2
    assert decisions[0].action == "CLASSIFICAR"
    assert decisions[0].account_id == "2.01.004"
    assert decisions[0].status_revisao == "REVISAR"
    assert decisions[0].keywords == ("agua mineral", "garrafa")


def test_load_fase9_decisions_from_csv_suporta_colunas_variaveis(tmp_path):
    csv_path = tmp_path / "revisao.csv"
    _write_csv(
        csv_path,
        [
            {
                "acao": "CRIAR_ALIAS",
                "descricao": "COCA COLA LATA 350ML",
                "fornecedor": "Fornecedor B",
                "nota_fiscal": "456",
                "conta_id_sugerida": "2.01.004",
                "status_revisao": "REVISAR",
                "cfop": "5405",
                "cnpj": "00000000000100",
                "valor": "12.34",
            }
        ],
    )

    decision = load_fase9_decisions_from_csv(csv_path)[0]

    assert decision.action == "CRIAR_ALIAS"
    assert decision.description == "COCA COLA LATA 350ML"
    assert decision.supplier == "Fornecedor B"
    assert decision.numero_nota == "456"
    assert decision.valor == "12.34"


def test_resultado_tem_estrutura_de_promocoes_bloqueios_e_ignorados():
    result = evaluate_fase9_decisions(
        [
            _decision(action="CLASSIFICAR", account_id="2.01.004"),
            _decision(action="MANTER_REVISAO", account_id="2.01.004"),
            _decision(action="IGNORAR", account_id="2.01.004"),
        ],
        valid_account_ids=VALID_ACCOUNTS,
    )

    assert result.promotions[0].reason is Fase9DecisionReason.PROMOTION_CANDIDATE
    assert result.promotions[0].account_id == "2.01.004"
    assert result.promotions[0].description
    assert result.blocked[0].reason is Fase9DecisionReason.MANTER_REVISAO
    assert result.ignored[0].reason is Fase9DecisionReason.IGNORAR


def _decision(**overrides: object) -> Fase9Decision:
    data = {
        "line_number": 2,
        "action": "CLASSIFICAR",
        "description": "REFRIGERANTE COLA 350ML",
        "descricao_normalizada": "refrigerante cola 350ml",
        "supplier": "Fornecedor A",
        "document": "123",
        "account_id": "2.01.004",
        "account_name": "Refrigerantes",
        "keywords": (),
        "status_revisao": "REVISAR",
        "alertas": "",
        "motivo_revisao": "",
        "tipo_operacao_especial": "COMPRA_NORMAL",
        "tratamento_operacao_especial": "DRE",
        "cfop": "5102",
        "numero_nota": "123",
        "chave_nfe": "chave",
        "cnpj": "00000000000100",
        "valor": "10.00",
        "source_row": (),
    }
    data.update(overrides)
    return Fase9Decision(**data)


def _write_csv(csv_path: Path, rows: list[dict[str, str]]) -> None:
    headers = list(rows[0].keys())
    with csv_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=headers, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)
