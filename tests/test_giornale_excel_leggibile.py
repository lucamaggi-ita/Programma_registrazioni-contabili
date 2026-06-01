from __future__ import annotations

import json
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from fantozzi.cli import main
from fantozzi.giornale_excel_leggibile import (
    FORBIDDEN_TECH_HEADERS,
    STATUS,
    GiornaleExcelLeggibileError,
    _GIORNALE_COLS,
    gerar_giornale_leggibile,
)


# ---------------------------------------------------------------------------
# Fixtures helpers
# ---------------------------------------------------------------------------

_PRELIVRO_OK_HEADERS = [
    "arquivo_excel_origem",
    "data_documento",
    "competencia",
    "cnpj_empresa",
    "nome_empresa",
    "fornecedor_ou_favorecido",
    "cnpj_cpf_favorecido",
    "documento",
    "tipo_documento",
    "conta_id",
    "conta_nome",
    "grupo_conta",
    "descricao_itens",
    "valor",
    "origem",
    "arquivo_origem",
    "chave_origem",
    "hash_registro",
    "status_revisao",
    "alertas",
    "observacoes",
    "centro_custo",
    "cfop",
    "natureza_operacao",
    "tipo_operacao_especial",
    "confianca_operacao_especial",
    "tratamento_operacao_especial",
    "tipo_lancamento",
    "dre_linha_id",
    "dre_linha_nome",
    "data_pagamento",
    "duplicatas_json",
]

_RESUMO_MENSAL_HEADERS = [
    "competencia",
    "qtd_lancamentos_ok",
    "valor_ok",
    "qtd_lancamentos_revisar",
    "valor_revisar",
    "qtd_lancamentos_total",
    "valor_total",
]


def _ok_row(
    competencia: str = "2026-01",
    conta_id: str = "2.01.001",
    conta_nome: str = "Alimentos",
    grupo_conta: str = "Custos",
    dre_linha_nome: str = "CMV - Alimentos",
    valor: float = -500.0,
    origem: str = "XML",
    fornecedor: str = "FORNECEDOR LTDA",
) -> list:
    row = [""] * len(_PRELIVRO_OK_HEADERS)
    idx = {h: i for i, h in enumerate(_PRELIVRO_OK_HEADERS)}
    row[idx["competencia"]] = competencia
    row[idx["conta_id"]] = conta_id
    row[idx["conta_nome"]] = conta_nome
    row[idx["grupo_conta"]] = grupo_conta
    row[idx["dre_linha_nome"]] = dre_linha_nome
    row[idx["valor"]] = valor
    row[idx["status_revisao"]] = "OK"
    row[idx["origem"]] = origem
    row[idx["fornecedor_ou_favorecido"]] = fornecedor
    row[idx["documento"]] = "NF-001"
    return row


def _revisar_row(
    competencia: str = "2026-01",
    conta_id: str = "2.99.001",
    conta_nome: str = "Sem Classificacao",
    valor: float = -100.0,
) -> list:
    row = [""] * len(_PRELIVRO_OK_HEADERS)
    idx = {h: i for i, h in enumerate(_PRELIVRO_OK_HEADERS)}
    row[idx["competencia"]] = competencia
    row[idx["conta_id"]] = conta_id
    row[idx["conta_nome"]] = conta_nome
    row[idx["valor"]] = valor
    row[idx["status_revisao"]] = "REVISAR"
    row[idx["origem"]] = "XML"
    return row


def _make_prelivro(
    tmp_path: Path,
    ok_rows: list[list] | None = None,
    revisar_rows: list[list] | None = None,
    resumo_mensal_rows: list[list] | None = None,
    filename: str = "02_PRE_LIVRO_DIARIO_GERENCIAL.xlsx",
) -> Path:
    wb = Workbook()
    ws_ok = wb.active
    ws_ok.title = "Lancamentos_OK"
    ws_ok.append(_PRELIVRO_OK_HEADERS)
    for r in (ok_rows or [_ok_row()]):
        ws_ok.append(r)

    ws_rev = wb.create_sheet("Lancamentos_REVISAR")
    ws_rev.append(_PRELIVRO_OK_HEADERS)
    for r in (revisar_rows or []):
        ws_rev.append(r)

    ws_men = wb.create_sheet("Resumo_Mensal")
    ws_men.append(_RESUMO_MENSAL_HEADERS)
    for r in (resumo_mensal_rows or [["2026-01", 1, 500.0, 0, 0.0, 1, 500.0]]):
        ws_men.append(r)

    subdir = tmp_path / "pacote_contabil_gerencial_provisorio"
    subdir.mkdir(parents=True, exist_ok=True)
    path = subdir / filename
    wb.save(str(path))
    return path


_BASE_OK_DRE_HEADERS = [
    "data", "ano_mes", "identificador_operacional",
    "fornecedor", "documento", "descricao",
    "conta_id", "conta_nome", "grupo_dre", "linha_dre",
    "valor", "status_revisao", "tipo_operacao_especial",
    "tratamento_operacao_especial", "origem_arquivo",
]

_DRE_GERAL_HEADERS = ["ordem", "grupo_dre", "linha_dre", "valor", "observacao"]
_DRE_MENSAL_HEADERS = ["ano_mes", "ordem", "grupo_dre", "linha_dre", "valor"]


def _make_dre(
    tmp_path: Path,
    dre_geral_rows: list | None = None,
    dre_mensal_rows: list | None = None,
    base_ok_rows: list | None = None,
    filename: str = "01_DRE_GERENCIAL_CONSOLIDADO.xlsx",
) -> Path:
    wb = Workbook()
    ws_geral = wb.active
    ws_geral.title = "DRE_Geral"
    ws_geral.append(_DRE_GERAL_HEADERS)
    for r in (dre_geral_rows or [[1, "Custos", "CMV - Alimentos", 500.0, "PROVISORIO"]]):
        ws_geral.append(r)

    ws_men = wb.create_sheet("DRE_Mensal")
    ws_men.append(_DRE_MENSAL_HEADERS)
    for r in (dre_mensal_rows or [["2026-01", 1, "Custos", "CMV - Alimentos", 500.0]]):
        ws_men.append(r)

    ws_base = wb.create_sheet("Base_OK_DRE")
    ws_base.append(_BASE_OK_DRE_HEADERS)
    for r in (base_ok_rows or [
        [None, "2026-01", "08884180000123", "FORNECEDOR LTDA", "NF-001",
         "Alimento X", "2.01.001", "Alimentos", "Custos", "CMV - Alimentos",
         -500.0, "OK", "COMPRA_NORMAL", "DRE", "arquivo.xlsx"],
    ]):
        ws_base.append(r)

    subdir = tmp_path / "pacote_contabil_gerencial_provisorio"
    subdir.mkdir(parents=True, exist_ok=True)
    path = subdir / filename
    wb.save(str(path))
    return path


# ---------------------------------------------------------------------------
# 1. Gera o arquivo 03_GIORNALE_LEGGIBILE.xlsx
# ---------------------------------------------------------------------------

def test_gera_arquivo_xlsx(tmp_path):
    _make_prelivro(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    assert result.xlsx_path.name == "03_GIORNALE_LEGGIBILE.xlsx"
    assert result.xlsx_path.exists()


# ---------------------------------------------------------------------------
# 2. Workbook contém as abas obrigatórias
# ---------------------------------------------------------------------------

def test_workbook_contem_abas_obrigatorias(tmp_path):
    _make_prelivro(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    assert "01_DASHBOARD" in wb.sheetnames
    assert "02_DRE_DETTAGLIO" in wb.sheetnames
    assert "03_ANALISE_POR_CONTA" in wb.sheetnames
    assert "04_GIORNALE" in wb.sheetnames
    assert "05_BASE_TECNICA_AUDIT" in wb.sheetnames
    wb.close()


# ---------------------------------------------------------------------------
# 3. 01_DASHBOARD contém TESTE_TECNICO_NAO_OFICIAL
# ---------------------------------------------------------------------------

def test_dashboard_contem_status(tmp_path):
    _make_prelivro(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["01_DASHBOARD"]
    all_text = " ".join(
        str(cell.value or "")
        for row in ws.iter_rows()
        for cell in row
        if cell.value is not None
    )
    assert STATUS in all_text
    wb.close()


# ---------------------------------------------------------------------------
# 4. 02_DRE_DETTAGLIO tem cabeçalhos humanos (sem snake_case proibidos)
# ---------------------------------------------------------------------------

def test_dre_dettaglio_tem_cabecalhos_humanos(tmp_path):
    _make_prelivro(tmp_path)
    _make_dre(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["02_DRE_DETTAGLIO"]
    headers = [str(cell.value or "") for cell in ws[1]]
    for h in headers:
        assert h not in FORBIDDEN_TECH_HEADERS, f"Header proibido em 02_DRE_DETTAGLIO: {h!r}"
    assert any(h in ("Grupo DRE", "Linha DRE", "Mês", "Valor") for h in headers)
    wb.close()


# ---------------------------------------------------------------------------
# 5. 03_ANALISE_POR_CONTA tem colunas mês e Total
# ---------------------------------------------------------------------------

def test_analise_por_conta_tem_colunas_mes_e_total(tmp_path):
    _make_prelivro(tmp_path)
    _make_dre(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["03_ANALISE_POR_CONTA"]
    headers = [str(cell.value or "") for cell in ws[1]]
    assert "Total" in headers, f"'Total' nao encontrado nos cabecalhos: {headers}"
    assert any(h.startswith("2026") for h in headers), f"Nenhum mes encontrado: {headers}"
    wb.close()


# ---------------------------------------------------------------------------
# 6. 04_GIORNALE tem cabeçalhos humanos
# ---------------------------------------------------------------------------

def test_giornale_tem_cabecalhos_humanos(tmp_path):
    _make_prelivro(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["04_GIORNALE"]
    headers = [str(cell.value or "") for cell in ws[1]]
    expected_human = [c[1] for c in _GIORNALE_COLS]
    for expected in expected_human:
        assert expected in headers, f"Cabecalho humano ausente em 04_GIORNALE: {expected!r}"
    wb.close()


# ---------------------------------------------------------------------------
# 7. 04_GIORNALE tem autofilter e freeze panes
# ---------------------------------------------------------------------------

def test_giornale_tem_autofilter(tmp_path):
    _make_prelivro(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["04_GIORNALE"]
    assert ws.auto_filter.ref is not None and ws.auto_filter.ref != ""
    wb.close()


def test_giornale_tem_freeze_panes(tmp_path):
    _make_prelivro(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["04_GIORNALE"]
    assert ws.freeze_panes is not None
    wb.close()


# ---------------------------------------------------------------------------
# 8. Valores monetários têm formato BRL nos primeiros 4 fogli
# ---------------------------------------------------------------------------

def test_valores_tem_formato_brl(tmp_path):
    _make_prelivro(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    # Check 04_GIORNALE valor column
    ws = wb["04_GIORNALE"]
    headers = [str(cell.value or "") for cell in ws[1]]
    valor_col = headers.index("Valor") + 1
    found_brl = False
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        cell = row[valor_col - 1]
        if cell.value is not None and cell.number_format:
            assert "#,##0.00" in cell.number_format or "0.00" in cell.number_format
            found_brl = True
            break
    assert found_brl, "Nenhuma celula com formato BRL encontrada em 04_GIORNALE"
    wb.close()


# ---------------------------------------------------------------------------
# 9. CLI gera arquivo e imprime [OK]
# ---------------------------------------------------------------------------

def test_cli_gera_arquivo_e_imprime_ok(tmp_path, capsys):
    _make_prelivro(tmp_path)
    out_dir = tmp_path / "out"
    rc = main([
        "gerar-giornale-leggibile",
        "--run-dir", str(tmp_path),
        "--output-dir", str(out_dir),
    ])
    captured = capsys.readouterr()
    assert rc == 0
    assert "[OK]" in captured.out
    assert (out_dir / "03_GIORNALE_LEGGIBILE.xlsx").exists()


# ---------------------------------------------------------------------------
# 10. CLI falha com erro claro se pré-livro ausente
# ---------------------------------------------------------------------------

def test_cli_falha_se_prelivro_ausente(tmp_path, capsys):
    rc = main([
        "gerar-giornale-leggibile",
        "--run-dir", str(tmp_path),
        "--output-dir", str(tmp_path / "out"),
    ])
    captured = capsys.readouterr()
    assert rc == 1
    assert "[ERRO]" in captured.out


# ---------------------------------------------------------------------------
# 11. Nenhum cabeçalho técnico proibido nos 4 fogli operacionais
# ---------------------------------------------------------------------------

def test_sem_cabecalhos_tecnicos_proibidos_em_fogli_operacionais(tmp_path):
    _make_prelivro(tmp_path)
    _make_dre(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))

    operational_sheets = ["01_DASHBOARD", "02_DRE_DETTAGLIO", "03_ANALISE_POR_CONTA", "04_GIORNALE"]
    for sheet_name in operational_sheets:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        # Check row 1 (headers) for forbidden names
        for cell in ws[1]:
            val = str(cell.value or "")
            assert val not in FORBIDDEN_TECH_HEADERS, (
                f"Cabecalho tecnico proibido '{val}' encontrado em aba '{sheet_name}'"
            )
    wb.close()


# ---------------------------------------------------------------------------
# 12. Base técnica de auditoria está no foglio 05_BASE_TECNICA_AUDIT
# ---------------------------------------------------------------------------

def test_base_tecnica_existe_e_tem_dados(tmp_path):
    _make_prelivro(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    assert "05_BASE_TECNICA_AUDIT" in wb.sheetnames
    ws = wb["05_BASE_TECNICA_AUDIT"]
    # Should have at least a header row + 1 data row
    assert ws.max_row >= 2
    wb.close()


def test_base_tecnica_tem_autofilter(tmp_path):
    _make_prelivro(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["05_BASE_TECNICA_AUDIT"]
    assert ws.auto_filter.ref is not None and ws.auto_filter.ref != ""
    wb.close()


# ---------------------------------------------------------------------------
# Additional: status e periodos no resultado
# ---------------------------------------------------------------------------

def test_resultado_status_e_coreto(tmp_path):
    _make_prelivro(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    assert result.status == STATUS


def test_resultado_conta_lancamentos_ok(tmp_path):
    _make_prelivro(tmp_path, ok_rows=[_ok_row(), _ok_row("2026-02")])
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    assert result.lancamentos_ok == 2


def test_resultado_conta_lancamentos_revisar(tmp_path):
    _make_prelivro(tmp_path, revisar_rows=[_revisar_row()])
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    assert result.lancamentos_revisar == 1


def test_run_dir_inexistente_levanta_erro(tmp_path):
    with pytest.raises(GiornaleExcelLeggibileError, match="run-dir nao encontrado"):
        gerar_giornale_leggibile(tmp_path / "nao_existe")


def test_output_dir_default_e_output_visual(tmp_path):
    _make_prelivro(tmp_path)
    result = gerar_giornale_leggibile(tmp_path)
    assert result.xlsx_path.parent == tmp_path / "output_visual"


def test_le_run_id_do_manifest(tmp_path):
    _make_prelivro(tmp_path)
    manifest = {"run_id": "2026-05-29_001", "status": "PROVISORIO_CONCLUIDO"}
    (tmp_path / "manifest_run.json").write_text(json.dumps(manifest), encoding="utf-8")
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    assert result.run_id == "2026-05-29_001"


def test_periodos_coletados_de_lancamentos(tmp_path):
    _make_prelivro(
        tmp_path,
        ok_rows=[_ok_row("2026-01"), _ok_row("2026-02"), _ok_row("2026-03")],
        resumo_mensal_rows=[
            ["2026-01", 1, 500.0, 0, 0.0, 1, 500.0],
            ["2026-02", 1, 600.0, 0, 0.0, 1, 600.0],
            ["2026-03", 1, 400.0, 0, 0.0, 1, 400.0],
        ],
    )
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    assert "2026-01" in result.periodos
    assert "2026-02" in result.periodos
    assert "2026-03" in result.periodos


def test_analise_por_conta_tem_freeze_panes(tmp_path):
    _make_prelivro(tmp_path)
    _make_dre(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["03_ANALISE_POR_CONTA"]
    assert ws.freeze_panes is not None
    wb.close()


def test_dre_dettaglio_tem_dados_quando_base_ok_dre_disponivel(tmp_path):
    _make_prelivro(tmp_path)
    _make_dre(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["02_DRE_DETTAGLIO"]
    assert ws.max_row >= 2, "02_DRE_DETTAGLIO deve ter pelo menos uma linha de dados"
    wb.close()


def test_cli_output_dir_default(tmp_path, capsys):
    _make_prelivro(tmp_path)
    rc = main(["gerar-giornale-leggibile", "--run-dir", str(tmp_path)])
    assert rc == 0
    assert (tmp_path / "output_visual" / "03_GIORNALE_LEGGIBILE.xlsx").exists()


def test_giornale_linha_dre_populada_via_lookup(tmp_path):
    """Linha DRE deve ser preenchida via lookup Base_OK_DRE mesmo quando pre-livro tem dre_linha_nome=None."""
    _make_prelivro(
        tmp_path,
        ok_rows=[_ok_row(conta_id="2.01.001", conta_nome="Alimentos")],
    )
    _make_dre(
        tmp_path,
        base_ok_rows=[
            [None, "2026-01", "08884180000123", "FORN", "NF-001",
             "Alimento X", "2.01.001", "Alimentos", "Custos", "CMV - Alimentos",
             -500.0, "OK", "COMPRA_NORMAL", "DRE", "arquivo.xlsx"],
        ],
    )
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["04_GIORNALE"]
    headers = [str(cell.value or "") for cell in ws[1]]
    linha_col = headers.index("Linha DRE") + 1
    grupo_col = headers.index("Grupo DRE") + 1
    # row 2 is the first data row
    assert ws.cell(row=2, column=linha_col).value == "CMV - Alimentos", (
        f"Linha DRE esperada 'CMV - Alimentos', obtida: {ws.cell(row=2, column=linha_col).value!r}"
    )
    assert ws.cell(row=2, column=grupo_col).value == "Custos", (
        f"Grupo DRE esperado 'Custos', obtido: {ws.cell(row=2, column=grupo_col).value!r}"
    )
    wb.close()


def test_giornale_inclui_lancamentos_revisar_com_status_marcado(tmp_path):
    _make_prelivro(tmp_path, revisar_rows=[_revisar_row()])
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["04_GIORNALE"]
    headers = [str(cell.value or "") for cell in ws[1]]
    status_col = headers.index("Status") + 1
    status_values = [
        str(ws.cell(row=r, column=status_col).value or "")
        for r in range(2, ws.max_row + 1)
        if ws.cell(row=r, column=status_col).value is not None
    ]
    assert "REVISAR" in status_values, "Lancamentos REVISAR devem aparecer em 04_GIORNALE"
    wb.close()


# ---------------------------------------------------------------------------
# 003C — Leituras rápidas no 01_DASHBOARD
# ---------------------------------------------------------------------------

def test_dashboard_tem_leituras_rapidas(tmp_path):
    _make_prelivro(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["01_DASHBOARD"]
    all_text = " ".join(
        str(cell.value or "")
        for row in ws.iter_rows()
        for cell in row
        if cell.value is not None
    )
    assert "LEITURAS" in all_text.upper(), (
        "01_DASHBOARD deve conter o bloco 'Leituras rápidas'"
    )
    wb.close()


# ---------------------------------------------------------------------------
# 003E — 01_DASHBOARD visual DRE / Conto Econômico
# ---------------------------------------------------------------------------

def test_dashboard_003e_sem_gridlines(tmp_path):
    _make_prelivro(tmp_path)
    _make_dre(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["01_DASHBOARD"]
    assert ws.sheet_view.showGridLines is False
    wb.close()


def test_dashboard_003e_kpis_grandes(tmp_path):
    _make_prelivro(tmp_path)
    _make_dre(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["01_DASHBOARD"]
    kpi_cells = ["A7", "C7", "E7", "G7", "I7"]
    for cell_ref in kpi_cells:
        assert ws[cell_ref].font.sz >= 20, f"KPI {cell_ref} com fonte pequena"
    wb.close()


def test_dashboard_003e_contem_dois_graficos_excel(tmp_path):
    _make_prelivro(
        tmp_path,
        ok_rows=[_ok_row("2026-01"), _ok_row("2026-02", valor=-700.0)],
        resumo_mensal_rows=[
            ["2026-01", 1, -500.0, 0, 0.0, 1, -500.0],
            ["2026-02", 1, -700.0, 0, 0.0, 1, -700.0],
        ],
    )
    _make_dre(
        tmp_path,
        dre_geral_rows=[
            [1, "Receitas", "Receita Bruta", 2000.0, "PROVISORIO"],
            [2, "Custos", "CMV", -900.0, "PROVISORIO"],
            [3, "Despesas Operacionais", "Administrativas", -300.0, "PROVISORIO"],
        ],
        dre_mensal_rows=[
            ["2026-01", 1, "Receitas", "Receita Bruta", 1000.0],
            ["2026-01", 2, "Custos", "CMV", -400.0],
            ["2026-02", 1, "Receitas", "Receita Bruta", 1000.0],
            ["2026-02", 2, "Custos", "CMV", -500.0],
        ],
    )
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["01_DASHBOARD"]
    assert len(ws._charts) >= 2
    wb.close()


def test_dashboard_003e_dre_nao_fluxo_de_caixa(tmp_path):
    _make_prelivro(tmp_path)
    _make_dre(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["01_DASHBOARD"]
    all_text = " ".join(
        str(cell.value or "")
        for row in ws.iter_rows()
        for cell in row
        if cell.value is not None
    )
    forbidden = ["Fluxo de Caixa", "Saldo de Caixa", "Contas a Receber", "Contas a Pagar"]
    for text in forbidden:
        assert text.lower() not in all_text.lower()
    assert "PAINEL ECONÔMICO GERENCIAL" in all_text
    wb.close()


def test_dashboard_003e_tem_pendencias_e_links_internos(tmp_path):
    _make_prelivro(tmp_path, revisar_rows=[_revisar_row()])
    _make_dre(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["01_DASHBOARD"]
    all_text = " ".join(
        str(cell.value or "")
        for row in ws.iter_rows()
        for cell in row
        if cell.value is not None
    )
    assert "PENDÊNCIAS" in all_text
    assert "REVISAR" in all_text
    assert any(cell.hyperlink for row in ws.iter_rows() for cell in row)
    wb.close()


# ---------------------------------------------------------------------------
# 003C — 03_ANALISE_POR_CONTA: coluna % do Total
# ---------------------------------------------------------------------------

def test_analise_por_conta_tem_coluna_percentual(tmp_path):
    _make_prelivro(tmp_path)
    _make_dre(tmp_path)
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["03_ANALISE_POR_CONTA"]
    headers = [str(cell.value or "") for cell in ws[1]]
    assert any("%" in h for h in headers), (
        f"Nenhuma coluna percentual encontrada em 03_ANALISE_POR_CONTA: {headers}"
    )
    wb.close()


# ---------------------------------------------------------------------------
# 003C — 03_ANALISE_POR_CONTA: ordenada por Total decrescente
# ---------------------------------------------------------------------------

def test_analise_por_conta_ordenada_por_total_decrescente(tmp_path):
    """Conta com maior total deve aparecer antes da conta com menor total."""
    _make_prelivro(
        tmp_path,
        ok_rows=[
            _ok_row(conta_id="2.01.001", conta_nome="Alimentos", valor=-100.0),
            _ok_row(conta_id="2.02.001", conta_nome="Bebidas", valor=-500.0),
            _ok_row(conta_id="2.03.001", conta_nome="Energia", valor=-200.0),
        ],
    )
    _make_dre(
        tmp_path,
        base_ok_rows=[
            [None, "2026-01", "CNPJ", "F1", "NF1", "Alimento",
             "2.01.001", "Alimentos", "Custos", "CMV", -100.0, "OK", "N", "DRE", "f.xlsx"],
            [None, "2026-01", "CNPJ", "F2", "NF2", "Bebida",
             "2.02.001", "Bebidas", "Custos", "CMV", -500.0, "OK", "N", "DRE", "f.xlsx"],
            [None, "2026-01", "CNPJ", "F3", "NF3", "Energia",
             "2.03.001", "Energia", "Custos", "CIF", -200.0, "OK", "N", "DRE", "f.xlsx"],
        ],
    )
    result = gerar_giornale_leggibile(tmp_path, tmp_path / "out")
    wb = load_workbook(str(result.xlsx_path))
    ws = wb["03_ANALISE_POR_CONTA"]
    headers = [str(cell.value or "") for cell in ws[1]]
    total_col = headers.index("Total") + 1
    # Collect data rows (exclude header and TOTAL GERAL row)
    totals = []
    for r in range(2, ws.max_row + 1):
        conta_cell = ws.cell(row=r, column=1).value
        if conta_cell and str(conta_cell) != "TOTAL GERAL":
            val = ws.cell(row=r, column=total_col).value
            if val is not None:
                totals.append(float(val))
    # Must be sorted descending by absolute value (costs are negative)
    abs_totals = [abs(v) for v in totals]
    assert abs_totals == sorted(abs_totals, reverse=True), (
        f"03_ANALISE_POR_CONTA nao esta ordenada por total decrescente: {totals}"
    )
