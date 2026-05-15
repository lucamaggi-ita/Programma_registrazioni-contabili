import csv
from pathlib import Path

from fantozzi.cli import build_parser, main


def test_parser_aceita_aplicar_fase9_revisao():
    parser = build_parser()

    args = parser.parse_args(
        [
            "aplicar-fase9-revisao",
            "--csv",
            "revisao_classificacao_revisada.csv",
            "--saida",
            "lancamentos_aplicados_fase9.csv",
            "--dry-run",
        ]
    )

    assert args.command == "aplicar-fase9-revisao"
    assert args.csv == "revisao_classificacao_revisada.csv"
    assert args.saida == "lancamentos_aplicados_fase9.csv"
    assert args.dry_run is True


def test_aplicar_fase9_revisao_gera_csv_separado_sem_alterar_original(tmp_path, capsys):
    input_csv = tmp_path / "revisao_classificacao_revisada.csv"
    output_csv = tmp_path / "lancamentos_aplicados_fase9.csv"
    plano = _save_plano_contas(tmp_path)
    _write_rows(
        input_csv,
        [
            _row(),
            _row(documentos="200", acao_manual="MANTER_REVISAO"),
        ],
    )
    original_bytes = input_csv.read_bytes()

    result = main(
        [
            "aplicar-fase9-revisao",
            "--csv",
            str(input_csv),
            "--saida",
            str(output_csv),
            "--plano-contas",
            str(plano),
        ]
    )

    captured = capsys.readouterr()
    assert result == 0
    assert input_csv.read_bytes() == original_bytes
    assert output_csv.exists()
    assert "Promovidos: 1" in captured.out
    assert "Bloqueados: 1" in captured.out
    rows = _read_rows(output_csv)
    assert rows[0]["status_revisao"] == "OK"
    assert rows[0]["fase9_status"] == "PROMOVIDO_OK"
    assert rows[1]["status_revisao"] == "REVISAR"
    assert rows[1]["motivo_bloqueio_fase9"] == "MANTER_REVISAO"


def test_aplicar_fase9_revisao_dry_run_nao_escreve_saida(tmp_path, capsys):
    input_csv = tmp_path / "revisao.csv"
    output_csv = tmp_path / "aplicado.csv"
    plano = _save_plano_contas(tmp_path)
    _write_rows(input_csv, [_row()])

    result = main(
        [
            "aplicar-fase9-revisao",
            "--csv",
            str(input_csv),
            "--saida",
            str(output_csv),
            "--plano-contas",
            str(plano),
            "--dry-run",
        ]
    )

    captured = capsys.readouterr()
    assert result == 0
    assert not output_csv.exists()
    assert "[DRY-RUN] Nenhum arquivo foi escrito." in captured.out
    assert "Promovidos: 1" in captured.out


def test_aplicar_fase9_revisao_falha_se_input_igual_output(tmp_path, capsys):
    input_csv = tmp_path / "revisao.csv"
    plano = _save_plano_contas(tmp_path)
    _write_rows(input_csv, [_row()])

    result = main(
        [
            "aplicar-fase9-revisao",
            "--csv",
            str(input_csv),
            "--saida",
            str(input_csv),
            "--plano-contas",
            str(plano),
        ]
    )

    captured = capsys.readouterr()
    assert result == 1
    assert "--saida deve ser diferente" in captured.out


def test_aplicar_fase9_revisao_mantem_gates_fiscais_bloqueados(tmp_path):
    input_csv = tmp_path / "revisao.csv"
    output_csv = tmp_path / "aplicado.csv"
    plano = _save_plano_contas(tmp_path)
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

    result = main(
        [
            "aplicar-fase9-revisao",
            "--csv",
            str(input_csv),
            "--saida",
            str(output_csv),
            "--plano-contas",
            str(plano),
        ]
    )

    rows = _read_rows(output_csv)
    assert result == 0
    assert [row["status_revisao"] for row in rows] == ["REVISAR"] * 3
    assert [row["motivo_bloqueio_fase9"] for row in rows] == [
        "COMODATO_FORA_DRE",
        "NF_359680_CONTROLE_MANUAL",
        "CFOP_1949_CONTROLE_MANUAL",
    ]


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


def _read_rows(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file, delimiter=";"))


def _save_plano_contas(tmp_path: Path) -> Path:
    plano = tmp_path / "plano_contas_gerencial.yaml"
    plano.write_text(
        """\
contas:
  - id: "2.01.004"
    nome: "Refrigerantes"
    grupo: "Custos de Mercadorias e Insumos"
    tipo: "custo"
  - id: "13.01.001"
    nome: "Revisar - Produto Nao Classificado"
    grupo: "Contas de Revisao"
    tipo: "revisao"
  - id: "13.01.002"
    nome: "Revisar - Classificacao Ambigua"
    grupo: "Contas de Revisao"
    tipo: "revisao"
  - id: "13.01.006"
    nome: "Revisar - Valor Inconsistente"
    grupo: "Contas de Revisao"
    tipo: "revisao"
  - id: "13.01.009"
    nome: "Revisar - Possivel Investimento"
    grupo: "Contas de Revisao"
    tipo: "revisao"
  - id: "13.01.010"
    nome: "Revisar - Pagamento de Pessoal a Regularizar"
    grupo: "Contas de Revisao"
    tipo: "revisao"
  - id: "13.01.999"
    nome: "Revisar - Outros Casos"
    grupo: "Contas de Revisao"
    tipo: "revisao"
""",
        encoding="utf-8",
    )
    return plano
