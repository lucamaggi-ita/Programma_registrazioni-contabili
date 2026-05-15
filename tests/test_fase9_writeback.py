import csv
from pathlib import Path

import pytest

from fantozzi.fase9_decision_engine import (
    evaluate_fase9_decisions,
    load_fase9_decisions_from_csv,
)
from fantozzi.fase9_writeback import (
    STATUS_BLOQUEADO_FASE9,
    STATUS_JA_OK_OU_FORA_ESCOPO,
    STATUS_PROMOVIDO_OK,
    apply_fase9_writeback,
    render_fase9_writeback_csv,
    write_fase9_applied_csv,
)


VALID_ACCOUNTS = frozenset({"2.01.004", "2.02.010", "2.05.003"})


def test_writeback_cria_novo_csv_sem_alterar_input(tmp_path):
    input_csv = tmp_path / "revisao_classificacao_revisada.csv"
    output_csv = tmp_path / "lancamentos_aplicados_fase9.csv"
    _write_rows(input_csv, [_row()])
    original_bytes = input_csv.read_bytes()
    result = _decision_result(input_csv)

    write_result = write_fase9_applied_csv(input_csv, output_csv, result)

    assert input_csv.read_bytes() == original_bytes
    assert output_csv.exists()
    assert write_result.output_path == output_csv
    assert write_result.promoted_rows == 1
    output_row = _read_rows(output_csv)[0]
    assert output_row["fase9_status"] == STATUS_PROMOVIDO_OK


def test_promocao_valida_muda_status_para_ok_so_no_output(tmp_path):
    rows = [_row(status_revisao="REVISAR")]
    input_csv = tmp_path / "revisao.csv"
    output_csv = tmp_path / "aplicado.csv"
    _write_rows(input_csv, rows)

    write_fase9_applied_csv(input_csv, output_csv, _decision_result(input_csv))

    assert _read_rows(input_csv)[0]["status_revisao"] == "REVISAR"
    output_row = _read_rows(output_csv)[0]
    assert output_row["status_revisao"] == "OK"
    assert output_row["fase9_aplicada"] == "true"
    assert output_row["acao_fase9"] == "CLASSIFICAR"
    assert output_row["conta_id_aplicada"] == "2.01.004"
    assert output_row["conta_nome_aplicada"] == "Refrigerantes"
    assert output_row["motivo_promocao"] == "PROMOTION_CANDIDATE"
    assert output_row["fonte_decisao"] == "csv_line:2"


def test_bloqueados_ficam_revisar_com_motivo(tmp_path):
    input_csv = tmp_path / "revisao.csv"
    _write_rows(
        input_csv,
        [_row(acao_manual="MANTER_REVISAO", conta_id_sugerida="2.01.004")],
    )

    result = apply_fase9_writeback(_read_rows(input_csv), _decision_result(input_csv))

    row = result.rows[0]
    assert row["status_revisao"] == "REVISAR"
    assert row["fase9_status"] == STATUS_BLOQUEADO_FASE9
    assert row["fase9_aplicada"] == "false"
    assert row["motivo_bloqueio_fase9"] == "MANTER_REVISAO"


def test_ja_ok_nao_vem_repromovido(tmp_path):
    input_csv = tmp_path / "revisao.csv"
    _write_rows(input_csv, [_row(status_revisao="OK")])

    result = apply_fase9_writeback(_read_rows(input_csv), _decision_result(input_csv))

    row = result.rows[0]
    assert row["status_revisao"] == "OK"
    assert row["fase9_status"] == STATUS_JA_OK_OU_FORA_ESCOPO
    assert row["fase9_aplicada"] == "false"
    assert row["motivo_bloqueio_fase9"] == "JA_OK"


def test_comodato_nf359680_cfop1949_restam_bloqueados(tmp_path):
    input_csv = tmp_path / "revisao.csv"
    _write_rows(
        input_csv,
        [
            _row(
                documentos="100",
                tipo_operacao_especial="COMODATO",
                tratamento_operacao_especial="FORA_DRE",
            ),
            _row(documentos="359680", numero_nota="359680"),
            _row(documentos="101", cfop="1949"),
        ],
    )

    result = apply_fase9_writeback(_read_rows(input_csv), _decision_result(input_csv))

    assert [row["status_revisao"] for row in result.rows] == ["REVISAR"] * 3
    assert [row["fase9_status"] for row in result.rows] == [STATUS_BLOQUEADO_FASE9] * 3
    assert [row["motivo_bloqueio_fase9"] for row in result.rows] == [
        "COMODATO_FORA_DRE",
        "NF_359680_CONTROLE_MANUAL",
        "CFOP_1949_CONTROLE_MANUAL",
    ]


def test_output_deterministico(tmp_path):
    input_csv = tmp_path / "revisao.csv"
    rows = [
        _row(),
        _row(documentos="200", acao_manual="IGNORAR"),
    ]
    _write_rows(input_csv, rows)
    result = _decision_result(input_csv)

    first = render_fase9_writeback_csv(_read_rows(input_csv), result)
    second = render_fase9_writeback_csv(_read_rows(input_csv), result)

    assert first == second


def test_render_e_apply_nao_escrevem_arquivos_sem_path_explicito(tmp_path):
    rows = [_row()]
    input_csv = tmp_path / "revisao.csv"
    _write_rows(input_csv, rows)
    before = sorted(path.name for path in tmp_path.iterdir())
    result = _decision_result(input_csv)

    apply_fase9_writeback(_read_rows(input_csv), result)
    render_fase9_writeback_csv(_read_rows(input_csv), result)

    assert sorted(path.name for path in tmp_path.iterdir()) == before


def test_writeback_recusa_saida_igual_ao_csv_original(tmp_path):
    input_csv = tmp_path / "revisao.csv"
    _write_rows(input_csv, [_row()])

    with pytest.raises(ValueError):
        write_fase9_applied_csv(input_csv, input_csv, _decision_result(input_csv))


def test_sem_decisao_fica_fora_escopo():
    rows = [_row()]
    empty_result = evaluate_fase9_decisions((), valid_account_ids=VALID_ACCOUNTS)

    result = apply_fase9_writeback(rows, empty_result)

    assert result.rows[0]["fase9_status"] == STATUS_JA_OK_OU_FORA_ESCOPO
    assert result.rows[0]["motivo_bloqueio_fase9"] == "SEM_DECISAO_FASE9"


def _decision_result(csv_path: Path):
    return evaluate_fase9_decisions(
        load_fase9_decisions_from_csv(csv_path),
        valid_account_ids=VALID_ACCOUNTS,
    )


def _row(**overrides: str) -> dict[str, str]:
    row = {
        "descricao_normalizada": "refrigerante cola 350ml",
        "descricao_exemplo": "REFRIGERANTE COLA 350ML",
        "fornecedores": "Fornecedor A",
        "documentos": "123",
        "numero_nota": "123",
        "status_revisao": "REVISAR",
        "alertas": "",
        "acao_manual": "CLASSIFICAR",
        "conta_atual_id": "13.01.001",
        "conta_atual_nome": "Revisar - Produto Nao Classificado",
        "conta_id_sugerida": "2.01.004",
        "conta_nome_sugerida": "Refrigerantes",
        "palavras_chave_sugeridas": "",
        "tipo_operacao_especial": "COMPRA_NORMAL",
        "tratamento_operacao_especial": "DRE",
        "cfop": "5102",
        "valor_total": "R$ 10,00",
    }
    row.update(overrides)
    return row


def _write_rows(path: Path, rows: list[dict[str, str]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=tuple(rows[0].keys()), delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


def _read_rows(path: Path) -> tuple[dict[str, str], ...]:
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        return tuple(csv.DictReader(file, delimiter=";"))
