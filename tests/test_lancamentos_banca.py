from __future__ import annotations

import json
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook

from fantozzi.banking.lancamentos_banca import (
    MANIFEST_NAME,
    REJEITADOS_CSV_NAME,
    gerar_lancamentos_banca_aprovados,
)
from fantozzi.banking.revisao_banca import (
    COMPROVANTES_HEADERS,
    COMPROVANTES_SHEET,
    CONTA_CORRENTE_HEADERS,
    CONTA_CORRENTE_SHEET,
    TECNICO_CONTA_CORRENTE_HEADERS,
    TECNICO_CONTA_CORRENTE_SHEET,
)
from fantozzi.cli import build_parser, main
from fantozzi.excel_writer import HEADERS, LANCAMENTOS_SHEET


def test_gerar_lancamentos_banca_filters_decisions_and_writes_standard_output(
    tmp_path,
):
    comprovantes = _write_comprovantes(
        tmp_path / "Analisi_regole_registrazioni_banca_comprovanti.xlsx",
        [
            _comprovante_row("APROVADO", "CLASSIFICAR", "2026-02-05", 2000, "IZALTINO JOSE ZANI JUNIOR"),
            _comprovante_row("PENDENTE", "CLASSIFICAR", "2026-02-05", 10, "PENDENTE LTDA"),
            _comprovante_row("APROVADO", "XML", "2026-02-06", 11, "XML LTDA"),
            _comprovante_row("APROVADO", "ESCLUDERE", "2026-02-07", 12, "ESCLUDERE LTDA"),
            _comprovante_row("APROVADO", "SISPAG", "2026-02-08", 13, "SISPAG LTDA"),
            _comprovante_row("BLOQUEADO", "CLASSIFICAR", "2026-02-09", 14, "BLOQUEADO LTDA"),
            _comprovante_row("REVISAR", "CLASSIFICAR", "2026-02-10", 15, "REVISAR LTDA"),
            _comprovante_row("NAO_CRIAR_REGRA", "CLASSIFICAR", "2026-02-11", 16, "NAO CRIAR LTDA"),
            _comprovante_row("APROVADO", "CLASSIFICAR", "2026-03-01", 17, "FORA PERIODO LTDA"),
            _comprovante_row(
                "APROVADO",
                "CLASSIFICAR",
                "2026-02-12",
                18,
                "RISCO LTDA",
                risco="DUVIDA",
            ),
        ],
    )
    conta_corrente = _write_conta_corrente(
        tmp_path / "Analisi_regole_registrazioni_banca_conto_corrente.xlsx",
        [
            _conta_row(
                "APROVADO",
                "CLASSIFICAR",
                "2026-02-18",
                -3290,
                "PIX ENVIADO",
                "TOP LEVEL DIGITAL",
                "9.01.001",
                "Publicidade e Propaganda",
            ),
            _conta_row(
                "APROVADO",
                "CLASSIFICAR",
                "2026-02-19",
                "-1.234,56",
                "DA COELBA",
                "COELBA",
                "4.01.011",
                "Energia Eletrica",
                risco="OK",
            ),
        ],
    )
    original_comprovantes = comprovantes.read_bytes()
    original_conta = conta_corrente.read_bytes()

    output = gerar_lancamentos_banca_aprovados(
        comprovantes,
        conta_corrente,
        tmp_path / "lancamentos_banca_aprovados.xlsx",
        date(2026, 2, 1),
        date(2026, 2, 28),
    )

    rows = _lancamentos(output)
    assert [row["fornecedor_ou_favorecido"] for row in rows] == [
        "IZALTINO JOSE ZANI JUNIOR",
        "TOP LEVEL DIGITAL",
        "COELBA",
    ]
    assert rows[0]["origem"] == "COMPROVANTES_BANCA"
    assert rows[0]["conta_id"] == "6.01.002"
    assert rows[0]["valor"] == 2000
    assert rows[1]["origem"] == "CONTA_CORRENTE"
    assert rows[1]["conta_id"] == "9.01.001"
    assert rows[1]["valor"] == 3290
    assert rows[2]["valor"] == 1234.56

    workbook = load_workbook(output, data_only=True)
    headers = [cell.value for cell in workbook[LANCAMENTOS_SHEET][1]]
    assert headers[: len(HEADERS)] == HEADERS

    manifest = json.loads((tmp_path / MANIFEST_NAME).read_text(encoding="utf-8"))
    assert manifest["linhas_lidas_comprovantes"] == 10
    assert manifest["linhas_lidas_conta_corrente"] == 2
    assert manifest["linhas_aprovadas_incluidas"] == 3
    assert manifest["linhas_excluidas_xml"] == 1
    assert manifest["linhas_excluidas_escludere"] == 1
    assert manifest["linhas_excluidas_sispag"] == 1
    assert manifest["linhas_rejeitadas_status"] == 4
    assert manifest["linhas_fora_periodo"] == 1
    assert manifest["linhas_rejeitadas_risco_duplicidade"] == 1
    assert (tmp_path / REJEITADOS_CSV_NAME).exists()
    assert comprovantes.read_bytes() == original_comprovantes
    assert conta_corrente.read_bytes() == original_conta


def test_cli_gerar_lancamentos_banca_aprovados(tmp_path):
    comprovantes = _write_comprovantes(
        tmp_path / "comprovantes.xlsx",
        [_comprovante_row("APROVADO", "CLASSIFICAR", "2026-02-05", 100, "M A BRITO COMERCIO LTDA")],
    )
    conta_corrente = _write_conta_corrente(
        tmp_path / "conta.xlsx",
        [_conta_row("APROVADO", "CLASSIFICAR", "2026-02-18", -3290, "PIX", "TOP LEVEL DIGITAL")],
    )
    output = tmp_path / "out.xlsx"

    result = main(
        [
            "gerar-lancamentos-banca-aprovados",
            "--comprovantes-xlsx",
            str(comprovantes),
            "--conta-corrente-xlsx",
            str(conta_corrente),
            "--saida",
            str(output),
            "--periodo-ini",
            "2026-02-01",
            "--periodo-fim",
            "2026-02-28",
        ]
    )

    assert result == 0
    assert output.exists()
    assert len(_lancamentos(output)) == 2


def test_parser_accepts_gerar_lancamentos_banca_aprovados_command():
    args = build_parser().parse_args(
        [
            "gerar-lancamentos-banca-aprovados",
            "--saida",
            "output/lancamentos_banca_aprovados.xlsx",
            "--periodo-ini",
            "2026-02-01",
            "--periodo-fim",
            "2026-02-28",
        ]
    )

    assert args.command == "gerar-lancamentos-banca-aprovados"
    assert args.saida == "output/lancamentos_banca_aprovados.xlsx"


def _write_comprovantes(path: Path, rows: list[dict[str, object]]) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = COMPROVANTES_SHEET
    sheet.append(COMPROVANTES_HEADERS)
    for row in rows:
        sheet.append([row.get(header, "") for header in COMPROVANTES_HEADERS])
    workbook.save(path)
    return path


def _write_conta_corrente(path: Path, rows: list[dict[str, object]]) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = CONTA_CORRENTE_SHEET
    sheet.append(CONTA_CORRENTE_HEADERS)
    for row in rows:
        sheet.append([row.get(header, "") for header in CONTA_CORRENTE_HEADERS])
    tecnico = workbook.create_sheet(TECNICO_CONTA_CORRENTE_SHEET)
    tecnico.append(TECNICO_CONTA_CORRENTE_HEADERS)
    for row in rows:
        tecnico.append(
            [
                _tecnico_value(row, header)
                for header in TECNICO_CONTA_CORRENTE_HEADERS
            ]
        )
    workbook.save(path)
    return path


def _comprovante_row(
    status: str,
    acao: str,
    data_pagamento: str,
    valor: object,
    beneficiario: str,
    *,
    risco: str = "OK",
) -> dict[str, object]:
    return {
        "STATUS_DECISAO": status,
        "ACAO_REGRA": acao,
        "PERIODO": "2026-02",
        "CNPJ": "08884180000123",
        "DATA_PAGAMENTO": data_pagamento,
        "VALOR": valor,
        "BENEFICIARIO_FAVORECIDO": beneficiario,
        "REFERENCIA_EMPRESA": "REFERENCIA",
        "TIPO_PAGAMENTO": "PIX",
        "CONTA_ID_APROVADA": "6.01.002",
        "CONTA_NOME_APROVADA": "Honorarios Contabeis",
        "CHAVE_REGRA_BENEFICIARIO": beneficiario,
        "CHAVE_REGRA_REFERENCIA": "REFERENCIA",
        "RISCO_DUPLICIDADE_XML": risco,
        "ESCOPO_REGRA": "COMPROVANTE",
    }


def _conta_row(
    status: str,
    acao: str,
    data_movimento: str,
    valor: object,
    lancamento: str,
    razao_social: str,
    conta_id: str = "9.01.001",
    conta_nome: str = "Publicidade e Propaganda",
    *,
    risco: str = "",
) -> dict[str, object]:
    return {
        "STATUS_DECISAO": status,
        "ACAO_REGRA": acao,
        "DATA": data_movimento,
        "VALOR": valor,
        "LANCAMENTO": lancamento,
        "RAZAO_SOCIAL": razao_social,
        "CONTA_ID_APROVADA": conta_id,
        "CONTA_NOME_APROVADA": conta_nome,
        "CHAVE_REGRA_LANCAMENTO": lancamento,
        "CHAVE_REGRA_RAZAO_SOCIAL": razao_social,
        "RISCO_DUPLICIDADE_XML": risco,
        "ESCOPO_REGRA": "CONTA_CORRENTE",
    }


def _tecnico_value(row: dict[str, object], header: str) -> object:
    if header == "CHAVE_ESTAVEL":
        return (
            f"CONTA_CORRENTE|HISTORICO={row['LANCAMENTO']}|"
            f"RAZAO={row['RAZAO_SOCIAL']}|TIPO=PIX_TRANSFERENCIA_SAIDA|"
            "DIRECAO=SAIDA"
        )
    if header == "CNPJ":
        return "08884180000123"
    return ""


def _lancamentos(path: Path) -> list[dict[str, object]]:
    workbook = load_workbook(path, data_only=True)
    sheet = workbook[LANCAMENTOS_SHEET]
    headers = [cell.value for cell in sheet[1]]
    return [
        {
            headers[index]: value
            for index, value in enumerate(row)
            if index < len(headers)
        }
        for row in sheet.iter_rows(min_row=2, values_only=True)
    ]
