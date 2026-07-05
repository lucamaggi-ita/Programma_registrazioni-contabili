from __future__ import annotations

import csv
import hashlib
import json
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import Workbook

from fantozzi.cli import build_parser
from fantozzi.excel_writer import HEADERS, LANCAMENTOS_SHEET
from fantozzi.models import Lancamento
from fantozzi.reprocessamento_pipeline import (
    MODO_DEFINITIVO_BLOQUEADO,
    ETAPAS_PREVISTAS,
    PipelineContext,
    PipelineSteps,
    ReprocessamentoPipelineError,
    StepResult,
    _aplicar_excecoes_contabeis_pontuais,
    _step_gerar_pacote_revisao,
    descobrir_fontes_xml,
    reprocessar_importacao_real_provisoria,
)


def _make_xml_base(tmp_path: Path) -> Path:
    xml_base = tmp_path / "XML_fatture_acquisto"
    for cnpj in ("55746722000181", "8884180000123"):
        xml_dir = xml_base / cnpj / "xml"
        xml_dir.mkdir(parents=True)
        (xml_dir / "nota.xml").write_text("<xml />", encoding="utf-8")
    return xml_base


def _fake_steps(calls: list[tuple[str, Path]], fail_at: str | None = None) -> PipelineSteps:
    def _record(name: str, path: Path) -> None:
        calls.append((name, path))
        if fail_at == name:
            raise RuntimeError(f"falha controlada em {name}")

    def importar_xml(context):
        _record("importar_xml", context.output_run_dir)
        arquivos = []
        for fonte in context.fontes_xml:
            path = (
                context.output_run_dir
                / fonte.cnpj_pasta
                / "lancamentos_fantozzi_real.xlsx"
            )
            _save_xml_lancamentos_xlsx(path)
            arquivos.append(path)
        return StepResult(
            arquivos_gerados=tuple(arquivos),
            totais={
                "xml_processados": 2,
                "lancamentos_gerados": 4,
                "valor_total_importado": "100.00",
            },
        )

    def gerar_diagnosticos(context):
        _record("gerar_diagnosticos", context.output_run_dir)
        arquivos = []
        for xlsx in context.lancamentos_xlsx:
            report_dir = context.reports_run_dir / xlsx.parent.name
            report_dir.mkdir(parents=True, exist_ok=True)
            for name in (
                "diagnostico_excel_real.txt",
                "revisao_lancamentos_real.txt",
                "candidatos_classificacao_real.txt",
                "revisao_classificacao_real.csv",
            ):
                path = report_dir / name
                path.write_text(name, encoding="utf-8")
                arquivos.append(path)
        return StepResult(
            arquivos_gerados=tuple(arquivos),
            classificacao={"candidatos_manter_revisao": 2},
        )

    def auditar_operacoes(context):
        _record("auditar_operacoes", context.output_run_dir)
        arquivos = []
        for fonte in context.fontes_xml:
            path = context.reports_run_dir / fonte.cnpj_pasta / "operacoes_especiais.txt"
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text("operacoes", encoding="utf-8")
            arquivos.append(path)
        return StepResult(
            arquivos_gerados=tuple(arquivos),
            totais={"operacoes_especiais": 1},
        )

    def gerar_pre_livro(context):
        _record("gerar_pre_livro", context.output_run_dir)
        xlsx = context.reports_run_dir / "pre_livro_diario_gerencial_teste.xlsx"
        txt = context.reports_run_dir / "auditoria_pre_livro_diario_gerencial_teste.txt"
        xlsx.write_text(
            "\n".join(str(path) for path in context.lancamentos_xlsx),
            encoding="utf-8",
        )
        txt.write_text("audit", encoding="utf-8")
        return StepResult(
            arquivos_gerados=(xlsx, txt),
            totais={"lancamentos_ok": 3, "lancamentos_revisar": 1},
        )

    def gerar_pacote_revisao(context):
        _record("gerar_pacote_revisao", context.output_run_dir)
        xlsx = context.reports_run_dir / "pacote_revisao_contabil_fase8_1.xlsx"
        txt = context.reports_run_dir / "pacote_revisao_contabil_fase8_1.txt"
        xlsx.write_text("pacote", encoding="utf-8")
        txt.write_text("pacote", encoding="utf-8")
        return StepResult(
            arquivos_gerados=(xlsx, txt),
            classificacao={
                "fase9_executada": False,
                "regras_aplicadas": 0,
                "candidatos_manter_revisao": 2,
            },
        )

    def gerar_relatorios_ok(context):
        _record("gerar_relatorios_ok", context.output_run_dir)
        xlsx = context.reports_run_dir / "relatorios_gerenciais_ok_provisorios.xlsx"
        txt = context.reports_run_dir / "relatorios_gerenciais_ok_provisorios.txt"
        xlsx.write_text("ok", encoding="utf-8")
        txt.write_text("ok", encoding="utf-8")
        return StepResult(
            arquivos_gerados=(xlsx, txt),
            totais={
                "lancamentos_ok": 3,
                "lancamentos_revisar": 1,
                "valor_total_incluido_ok": "75.00",
            },
        )

    def auditar_relatorios_ok(context):
        _record("auditar_relatorios_ok", context.output_run_dir)
        xlsx = (
            context.reports_run_dir
            / "auditoria_relatorios_gerenciais_ok_provisorios.xlsx"
        )
        txt = (
            context.reports_run_dir
            / "auditoria_relatorios_gerenciais_ok_provisorios.txt"
        )
        xlsx.write_text("audit ok", encoding="utf-8")
        txt.write_text("audit ok", encoding="utf-8")
        return StepResult(arquivos_gerados=(xlsx, txt))

    def gerar_dre_ok(context):
        _record("gerar_dre_ok", context.output_run_dir)
        xlsx = context.reports_run_dir / "dre_gerencial_ok_provisoria.xlsx"
        txt = context.reports_run_dir / "dre_gerencial_ok_provisoria.txt"
        xlsx.write_text("dre", encoding="utf-8")
        txt.write_text("dre", encoding="utf-8")
        return StepResult(arquivos_gerados=(xlsx, txt))

    return PipelineSteps(
        importar_xml=importar_xml,
        gerar_diagnosticos=gerar_diagnosticos,
        auditar_operacoes=auditar_operacoes,
        gerar_pre_livro=gerar_pre_livro,
        gerar_pacote_revisao=gerar_pacote_revisao,
        gerar_relatorios_ok=gerar_relatorios_ok,
        auditar_relatorios_ok=auditar_relatorios_ok,
        gerar_dre_ok=gerar_dre_ok,
    )


def _run_pipeline(tmp_path: Path, *, fail_at: str | None = None):
    calls: list[tuple[str, Path]] = []
    xml_base = _make_xml_base(tmp_path)
    saida_output = tmp_path / "output" / "importacao_real"
    saida_reports = tmp_path / "reports" / "importacao_real"
    result = reprocessar_importacao_real_provisoria(
        xml_base,
        saida_output,
        saida_reports,
        project_root=tmp_path,
        steps=_fake_steps(calls, fail_at=fail_at),
    )
    return result, calls, saida_output, saida_reports


def _run_pipeline_com_fase9(tmp_path: Path, fase9_csv_path):
    calls: list[tuple[str, Path]] = []
    xml_base = _make_xml_base(tmp_path)
    saida_output = tmp_path / "output" / "importacao_real"
    saida_reports = tmp_path / "reports" / "importacao_real"
    result = reprocessar_importacao_real_provisoria(
        xml_base,
        saida_output,
        saida_reports,
        project_root=tmp_path,
        steps=_fake_steps(calls),
        today=date(2026, 5, 15),
        fase9_csv_path=fase9_csv_path,
    )
    return result, calls, saida_output, saida_reports


def _lancamento_xml(
    *,
    fornecedor: str,
    documento: str,
    valor: Decimal,
    descricao: str,
    origem: str = "XML_NFE",
    tipo_documento: str = "XML_NFE",
) -> Lancamento:
    return Lancamento(
        data_documento=date(2026, 2, 17),
        competencia="2026-02",
        cnpj_empresa="08884180000123",
        nome_empresa="AMAMI RESTAURANTE LTDA",
        fornecedor_ou_favorecido=fornecedor,
        cnpj_cpf_favorecido="48901288000100",
        documento=documento,
        tipo_documento=tipo_documento,
        conta_id="4.01.013",
        conta_nome="Gas de Cozinha",
        grupo_conta="Ocupacao, Utilidades e Funcionamento",
        descricao_itens=descricao,
        valor=valor,
        origem=origem,
        arquivo_origem="Database_Dev/XML_fatture_acquisto/08884180000123/xml/nfe.xml",
        chave_origem="29260248901288000100550010000026841126527091",
        hash_registro="hash",
        status_revisao="OK",
    )


def test_excecao_pontual_m_a_brito_remove_apenas_xml_2684_glp_p45():
    excecao = _lancamento_xml(
        fornecedor="M. A BRITO COMERCIO LTDA",
        documento="2684",
        valor=Decimal("12000.00"),
        descricao="GLP EM CILINDRO P45 (12000.00)",
    )
    pagamento_banca = _lancamento_xml(
        fornecedor="M A BRITO COMERCIO LTDA",
        documento="PIX 1000",
        valor=Decimal("1000.00"),
        descricao="PIX pagamento gas",
        origem="COMPROVANTES_CSV",
        tipo_documento="BANCA_COMPROVANTE_CSV",
    )
    outro_xml = _lancamento_xml(
        fornecedor="M A BRITO COMERCIO LTDA",
        documento="2685",
        valor=Decimal("12000.00"),
        descricao="GLP EM CILINDRO P45 (12000.00)",
    )

    mantidos, excluidos = _aplicar_excecoes_contabeis_pontuais(
        [excecao, pagamento_banca, outro_xml]
    )

    assert mantidos == [pagamento_banca, outro_xml]
    assert len(excluidos) == 1
    assert (
        excluidos[0]["motivo"]
        == "EXCECAO_CONTABIL_PONTUAL_M_A_BRITO_USAR_PAGAMENTOS"
    )
    assert excluidos[0]["documento"] == "2684"


def test_pipeline_provisoria_cria_run(tmp_path):
    result, _, saida_output, saida_reports = _run_pipeline(tmp_path)

    assert result.status_run == "PROVISORIO_CONCLUIDO"
    assert result.output_run_dir == saida_output / "runs" / result.run_id
    assert result.reports_run_dir == saida_reports / "runs" / result.run_id
    assert result.output_run_dir.is_dir()
    assert result.reports_run_dir.is_dir()
    assert (saida_output / "current_run.txt").read_text(encoding="utf-8").strip() == result.run_id
    assert (saida_reports / "current_run.txt").read_text(encoding="utf-8").strip() == result.run_id
    assert result.manifest_path.exists()


def test_pipeline_provisoria_usa_paths_da_run(tmp_path):
    result, calls, saida_output, _ = _run_pipeline(tmp_path)

    for _, path in calls:
        assert path == result.output_run_dir
        assert path != saida_output
        assert path != saida_output / "runs"


def test_pipeline_provisoria_registra_checkpoints_na_ordem(tmp_path):
    result, _, _, _ = _run_pipeline(tmp_path)

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    checkpoints = [item["descricao"] for item in manifest["checkpoints"]]
    assert checkpoints == [
        "CRIACAO_DA_RUN",
        "IMPORTACAO_XML_CONCLUIDA",
        "DIAGNOSTICOS_CONCLUIDOS",
        "OPERACOES_ESPECIAIS_AUDITADAS",
        "PRE_LIVRO_GERADO",
        "AUDITORIA_PRE_LIVRO_CONCLUIDA",
        "PACOTE_REVISAO_CONTABIL_GERADO",
        "RELATORIOS_OK_GERADOS",
        "AUDITORIA_RELATORIOS_OK_CONCLUIDA",
        "DRE_OK_PROVISORIA_GERADA",
        "FINALIZACAO_PROVISORIA",
    ]


def test_pipeline_provisoria_atualiza_manifesto_final(tmp_path):
    result, _, _, _ = _run_pipeline(tmp_path)

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert manifest["status_run"] == "PROVISORIO_CONCLUIDO"
    assert manifest["totais"]["xml_processados"] == 2
    assert manifest["totais"]["lancamentos_gerados"] == 4
    assert manifest["totais"]["lancamentos_ok"] == 3
    assert manifest["totais"]["lancamentos_revisar"] == 1
    assert manifest["totais"]["operacoes_especiais"] == 1
    assert manifest["totais"]["valor_total_incluido_ok"] == "75.00"
    assert manifest["avisos"]


def test_pipeline_manifest_final_registra_hash_regras_comprovantes(tmp_path):
    config_dir = tmp_path / "config"
    config_dir.mkdir()
    for file_name in (
        "plano_contas_gerencial.yaml",
        "regras_classificacao_produtos.yaml",
        "mapeamento_dre.yaml",
    ):
        (config_dir / file_name).write_text("regras: []\n", encoding="utf-8")
    regras_path = config_dir / "regras_classificacao_comprovantes.yaml"
    regras_path.write_text(
        "versao: teste\nregras:\n  - id: REGRA_TESTE\n",
        encoding="utf-8",
    )
    expected_hash = hashlib.sha256(regras_path.read_bytes()).hexdigest()

    result, _, _, _ = _run_pipeline(tmp_path)

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    entry = manifest["config_usado"]["regras_classificacao_comprovantes"]
    assert manifest["status_run"] == "PROVISORIO_CONCLUIDO"
    assert entry["path_relativo"] == "config/regras_classificacao_comprovantes.yaml"
    assert entry["exists"] is True
    assert entry["sha256"] == expected_hash
    assert entry["hash_sha256"] == expected_hash
    assert "totais" in manifest
    assert "classificacao" in manifest


def test_pipeline_provisoria_nao_executa_fase9(tmp_path):
    result, _, _, _ = _run_pipeline(tmp_path)

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert manifest["classificacao"]["fase9_executada"] is False
    assert not (result.reports_run_dir / "lancamentos_aplicados_fase9.csv").exists()


def test_pipeline_provisoria_nao_aplica_regras(tmp_path):
    result, _, _, _ = _run_pipeline(tmp_path)

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert manifest["classificacao"]["regras_aplicadas"] == 0


def test_pipeline_sem_csv_manual_registra_lancamentos_manuais_nao_executados(tmp_path):
    result, _, _, _ = _run_pipeline(tmp_path)

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert manifest["lancamentos_manuais_gerenciais"] == {"executado": False}


def test_pipeline_default_nao_inclui_lancamentos_banca_aprovados(tmp_path):
    result, _, _, _ = _run_pipeline(tmp_path)

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    pre_livro_inputs = (
        result.reports_run_dir / "pre_livro_diario_gerencial_teste.xlsx"
    ).read_text(encoding="utf-8")

    assert manifest["lancamentos_banca_aprovados"] == {"incluido": False}
    assert "BANCA_APROVADA" not in pre_livro_inputs


def test_pipeline_com_lancamentos_banca_aprovados_inclui_xlsx_e_manifest(tmp_path):
    calls: list[tuple[str, Path]] = []
    xml_base = _make_xml_base(tmp_path)
    saida_output = tmp_path / "output" / "importacao_real"
    saida_reports = tmp_path / "reports" / "importacao_real"
    banca_xlsx = tmp_path / "lancamentos_banca_aprovados.xlsx"
    _save_lancamentos_banca_xlsx(banca_xlsx)

    result = reprocessar_importacao_real_provisoria(
        xml_base,
        saida_output,
        saida_reports,
        project_root=tmp_path,
        steps=_fake_steps(calls),
        lancamentos_banca_aprovados_path=banca_xlsx,
    )

    output_xlsx = (
        result.output_run_dir / "BANCA_APROVADA" / "lancamentos_banca_aprovados.xlsx"
    )
    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    banca_manifest = manifest["lancamentos_banca_aprovados"]
    pre_livro_inputs = (
        result.reports_run_dir / "pre_livro_diario_gerencial_teste.xlsx"
    ).read_text(encoding="utf-8")
    checkpoints = [item["descricao"] for item in manifest["checkpoints"]]

    assert output_xlsx.exists()
    assert banca_manifest["incluido"] is True
    assert banca_manifest["input_xlsx"] == str(banca_xlsx)
    assert banca_manifest["output_xlsx_run"] == str(output_xlsx)
    assert banca_manifest["linhas_lidas"] == 2
    assert str(output_xlsx) in manifest["relatorios_gerados"]
    assert str(output_xlsx) in pre_livro_inputs
    assert checkpoints.index("LANCAMENTOS_BANCA_APROVADOS_INCLUIDOS") < (
        checkpoints.index("PRE_LIVRO_GERADO")
    )


def test_pipeline_lancamentos_banca_aprovados_path_inexistente_bloqueia(tmp_path):
    xml_base = _make_xml_base(tmp_path)

    with pytest.raises(
        ReprocessamentoPipelineError,
        match="Excel de lancamentos banca aprovados nao encontrado",
    ):
        reprocessar_importacao_real_provisoria(
            xml_base,
            tmp_path / "output",
            tmp_path / "reports",
            project_root=tmp_path,
            steps=_fake_steps([]),
            lancamentos_banca_aprovados_path=tmp_path / "ausente.xlsx",
        )


def test_parser_accepts_lancamentos_banca_aprovados_flag():
    args = build_parser().parse_args(
        [
            "reprocessar-importacao-real",
            "--xml-base",
            "Database_Dev/XML_fatture_acquisto",
            "--saida-output",
            "output/importacao_real",
            "--saida-reports",
            "reports/importacao_real",
            "--lancamentos-banca-aprovados",
            "output/lancamentos_banca_aprovados_bridge_001.xlsx",
        ]
    )

    assert args.command == "reprocessar-importacao-real"
    assert (
        args.lancamentos_banca_aprovados
        == "output/lancamentos_banca_aprovados_bridge_001.xlsx"
    )


def test_pipeline_com_csv_manual_gera_outputs_e_manifest(tmp_path):
    _save_config_9n(tmp_path)
    csv_manual = tmp_path / "manual.csv"
    _save_manual_csv_9n(csv_manual)
    calls: list[tuple[str, Path]] = []
    xml_base = _make_xml_base(tmp_path)
    saida_output = tmp_path / "output" / "importacao_real"
    saida_reports = tmp_path / "reports" / "importacao_real"

    result = reprocessar_importacao_real_provisoria(
        xml_base,
        saida_output,
        saida_reports,
        project_root=tmp_path,
        steps=_fake_steps(calls),
        lancamentos_manuais_gerenciais_csv_path=csv_manual,
    )

    output_xlsx = (
        result.output_run_dir
        / "MANUAL_GERENCIAL"
        / "lancamentos_manuais_gerenciais.xlsx"
    )
    output_csv = result.reports_run_dir / "lancamentos_manuais_gerenciais_aplicados.csv"
    auditoria_txt = (
        result.reports_run_dir / "auditoria_lancamentos_manuais_gerenciais.txt"
    )
    auditoria_duplicatas = (
        result.reports_run_dir
        / "07_AUDITORIA_POSSIVEIS_DUPLICATAS_MANUAIS_XML.txt"
    )
    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    manual = manifest["lancamentos_manuais_gerenciais"]
    checkpoints = [item["descricao"] for item in manifest["checkpoints"]]

    assert output_xlsx.exists()
    assert output_csv.exists()
    assert auditoria_txt.exists()
    assert auditoria_duplicatas.exists()
    assert manual["executado"] is True
    assert manual["input_csv"] == str(csv_manual)
    assert manual["output_xlsx_tecnico"] == str(output_xlsx)
    assert manual["output_csv"] == str(output_csv)
    assert manual["auditoria_txt"] == str(auditoria_txt)
    assert manual["lancamentos_aceitos"] == 1
    assert manual["linhas_tecnicas_geradas"] == 2
    assert manual["total_manual"] == "10.50"
    assert manual["fonte"] == "MANUAL_GERENCIAL"
    assert manual["possiveis_duplicatas_manuais_xml"] == 1
    assert manual["auditoria_possiveis_duplicatas_manuais_xml_txt"] == str(
        auditoria_duplicatas
    )
    assert str(output_xlsx) in manifest["relatorios_gerados"]
    assert str(output_csv) in manifest["relatorios_gerados"]
    assert str(auditoria_duplicatas) in manifest["relatorios_gerados"]
    assert checkpoints.index("OPERACOES_ESPECIAIS_AUDITADAS") < checkpoints.index(
        "LANCAMENTOS_MANUAIS_GERENCIAIS_APLICADOS"
    )
    assert checkpoints.index("LANCAMENTOS_MANUAIS_GERENCIAIS_APLICADOS") < (
        checkpoints.index("PRE_LIVRO_GERADO")
    )
    assert manifest["classificacao"]["fase9_executada"] is False


def test_pipeline_com_flag_pacote_contabil_gera_pacote_e_manifest(tmp_path):
    calls: list[tuple[str, Path]] = []
    xml_base = _make_xml_base(tmp_path)
    saida_output = tmp_path / "output" / "importacao_real"
    saida_reports = tmp_path / "reports" / "importacao_real"

    result = reprocessar_importacao_real_provisoria(
        xml_base,
        saida_output,
        saida_reports,
        project_root=tmp_path,
        steps=_fake_steps(calls),
        gerar_pacote_contabil_gerencial=True,
    )

    pacote_dir = result.reports_run_dir / "pacote_contabil_gerencial_provisorio"
    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    checkpoints = [item["descricao"] for item in manifest["checkpoints"]]
    pacote_manifest = pacote_dir / "manifest_pacote_contabil_gerencial.json"
    indice = pacote_dir / "00_INDICE_PACOTE_CONTABIL_GERENCIAL.txt"

    assert pacote_dir.is_dir()
    assert pacote_manifest.exists()
    assert indice.exists()
    assert manifest["pacote_contabil_gerencial"]["executado"] is True
    assert manifest["pacote_contabil_gerencial"]["diretorio"] == str(pacote_dir)
    assert manifest["classificacao"]["fase9_executada"] is False
    assert checkpoints.index("DRE_OK_PROVISORIA_GERADA") < checkpoints.index(
        "PACOTE_CONTABIL_GERENCIAL_GERADO"
    )
    assert str(pacote_manifest) in manifest["relatorios_gerados"]
    assert str(indice) in manifest["relatorios_gerados"]


def test_pipeline_dry_run_nao_valida_csv_manual(tmp_path):
    calls: list[tuple[str, Path]] = []
    xml_base = _make_xml_base(tmp_path)
    saida_output = tmp_path / "output"
    saida_reports = tmp_path / "reports"

    result = reprocessar_importacao_real_provisoria(
        xml_base,
        saida_output,
        saida_reports,
        dry_run=True,
        project_root=tmp_path,
        steps=_fake_steps(calls),
        lancamentos_manuais_gerenciais_csv_path=tmp_path / "inexistente.csv",
    )

    assert result.dry_run is True
    assert calls == []
    assert not result.manifest_path.exists()


def test_pipeline_com_fase9_explicita_gera_csv_separado_e_manifest(tmp_path):
    _save_plano_contas_fase9(tmp_path)
    fase9_csv = tmp_path / "revisao_classificacao_revisada.csv"
    _save_fase9_csv(
        fase9_csv,
        [
            _fase9_row(),
            _fase9_row(documentos="200", acao_manual="MANTER_REVISAO"),
            _fase9_row(
                documentos="300",
                tipo_operacao_especial="COMODATO",
                tratamento_operacao_especial="FORA_DRE",
            ),
            _fase9_row(documentos="359680", numero_nota="359680"),
            _fase9_row(documentos="400", cfop="1949"),
        ],
    )
    original_csv = fase9_csv.read_bytes()

    result, _, _, _ = _run_pipeline_com_fase9(tmp_path, fase9_csv)

    output_csv = result.reports_run_dir / "lancamentos_aplicados_fase9.csv"
    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    rows = _read_csv_rows(output_csv)
    assert fase9_csv.read_bytes() == original_csv
    assert output_csv.exists()
    assert rows[0]["status_revisao"] == "OK"
    assert rows[0]["fase9_status"] == "PROMOVIDO_OK"
    assert [row["status_revisao"] for row in rows[1:]] == ["REVISAR"] * 4
    assert [row["motivo_bloqueio_fase9"] for row in rows[2:]] == [
        "COMODATO_FORA_DRE",
        "NF_359680_CONTROLE_MANUAL",
        "CFOP_1949_CONTROLE_MANUAL",
    ]
    classificacao = manifest["classificacao"]
    assert classificacao["fase9_executada"] is True
    assert classificacao["input_csv"] == str(fase9_csv)
    assert classificacao["output_csv"] == str(output_csv)
    assert classificacao["promovidos"] == 1
    assert classificacao["bloqueados"] == 4
    assert classificacao["fora_escopo"] == 0
    assert classificacao["csvs_processados"] == [
        {
            "input_csv": str(fase9_csv),
            "output_csv": str(output_csv),
            "promovidos": 1,
            "bloqueados": 4,
            "fora_escopo": 0,
        }
    ]
    assert str(output_csv) in manifest["relatorios_gerados"]


def test_pipeline_com_varios_csv_fase9_gera_outputs_distintos_e_manifest_agregado(
    tmp_path,
):
    _save_plano_contas_fase9(tmp_path)
    primeiro_csv = (
        tmp_path / "55746722000181" / "revisao_classificacao_revisada.csv"
    )
    segundo_csv = (
        tmp_path / "08884180000123" / "revisao_classificacao_revisada.csv"
    )
    _save_fase9_csv(primeiro_csv, [_fase9_row(documentos="100")])
    _save_fase9_csv(
        segundo_csv,
        [
            _fase9_row(documentos="200"),
            _fase9_row(documentos="201", cfop="1949"),
        ],
    )

    result, _, _, _ = _run_pipeline_com_fase9(
        tmp_path,
        [primeiro_csv, segundo_csv],
    )

    primeiro_output = (
        result.reports_run_dir
        / "lancamentos_aplicados_fase9_55746722000181.csv"
    )
    segundo_output = (
        result.reports_run_dir
        / "lancamentos_aplicados_fase9_08884180000123.csv"
    )
    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    classificacao = manifest["classificacao"]
    assert primeiro_output.exists()
    assert segundo_output.exists()
    assert not (result.reports_run_dir / "lancamentos_aplicados_fase9.csv").exists()
    assert _read_csv_rows(segundo_output)[1]["motivo_bloqueio_fase9"] == (
        "CFOP_1949_CONTROLE_MANUAL"
    )
    assert classificacao["fase9_executada"] is True
    assert classificacao["promovidos"] == 2
    assert classificacao["bloqueados"] == 1
    assert classificacao["fora_escopo"] == 0
    assert classificacao["regras_aplicadas"] == 2
    assert classificacao["csvs_processados"] == [
        {
            "input_csv": str(primeiro_csv),
            "output_csv": str(primeiro_output),
            "promovidos": 1,
            "bloqueados": 0,
            "fora_escopo": 0,
        },
        {
            "input_csv": str(segundo_csv),
            "output_csv": str(segundo_output),
            "promovidos": 1,
            "bloqueados": 1,
            "fora_escopo": 0,
        },
    ]
    assert str(primeiro_output) in manifest["relatorios_gerados"]
    assert str(segundo_output) in manifest["relatorios_gerados"]


def test_pipeline_com_fase9_registra_checkpoint_antes_dos_relatorios_ok(tmp_path):
    _save_plano_contas_fase9(tmp_path)
    fase9_csv = tmp_path / "revisao_classificacao_revisada.csv"
    _save_fase9_csv(fase9_csv, [_fase9_row()])

    result, calls, _, _ = _run_pipeline_com_fase9(tmp_path, fase9_csv)

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    checkpoints = [item["descricao"] for item in manifest["checkpoints"]]
    assert checkpoints.index("PACOTE_REVISAO_CONTABIL_GERADO") < checkpoints.index(
        "FASE9_APLICADA"
    )
    assert checkpoints.index("FASE9_APLICADA") < checkpoints.index(
        "RELATORIOS_OK_GERADOS"
    )
    assert [name for name, _ in calls][-3:] == [
        "gerar_relatorios_ok",
        "auditar_relatorios_ok",
        "gerar_dre_ok",
    ]


def test_pipeline_com_fase9_rejeita_input_igual_output_aplicado(tmp_path, monkeypatch):
    _save_plano_contas_fase9(tmp_path)
    xml_base = _make_xml_base(tmp_path)
    saida_output = tmp_path / "output" / "importacao_real"
    saida_reports = tmp_path / "reports" / "importacao_real"
    from fantozzi import run_manager

    monkeypatch.setattr(
        run_manager,
        "generate_run_id",
        lambda saida_output, saida_reports, today=None: "2026-05-15_001",
    )
    fase9_csv = (
        saida_reports
        / "runs"
        / "2026-05-15_001"
        / "lancamentos_aplicados_fase9.csv"
    )
    fase9_csv.parent.mkdir(parents=True)
    _save_fase9_csv(fase9_csv, [_fase9_row()])

    with pytest.raises(ReprocessamentoPipelineError, match="deve ser diferente"):
        reprocessar_importacao_real_provisoria(
            xml_base,
            saida_output,
            saida_reports,
            project_root=tmp_path,
            steps=_fake_steps([]),
            today=date(2026, 5, 15),
            fase9_csv_path=fase9_csv,
        )


def test_pipeline_provisoria_bloqueia_modo_definitivo(tmp_path):
    xml_base = _make_xml_base(tmp_path)

    with pytest.raises(ReprocessamentoPipelineError, match=MODO_DEFINITIVO_BLOQUEADO):
        reprocessar_importacao_real_provisoria(
            xml_base,
            tmp_path / "output",
            tmp_path / "reports",
            modo="definitivo",
            project_root=tmp_path,
        )


def test_pipeline_provisoria_nao_mistura_runs(tmp_path):
    xml_base = _make_xml_base(tmp_path)
    saida_output = tmp_path / "output"
    saida_reports = tmp_path / "reports"
    (saida_output / "runs" / "2026-05-07_001" / "antigo").mkdir(parents=True)
    (saida_reports / "runs" / "2026-05-07_001").mkdir(parents=True)
    calls: list[tuple[str, Path]] = []

    result = reprocessar_importacao_real_provisoria(
        xml_base,
        saida_output,
        saida_reports,
        project_root=tmp_path,
        steps=_fake_steps(calls),
        today=date(2026, 5, 7),
    )

    assert result.run_id != "2026-05-07_001"
    assert result.output_run_dir.name.endswith("_002")
    assert (saida_output / "runs" / "2026-05-07_001" / "antigo").exists()
    assert (saida_output / "current_run.txt").read_text(encoding="utf-8").strip() == result.run_id
    assert (saida_reports / "current_run.txt").read_text(encoding="utf-8").strip() == result.run_id
    assert result.manifest_path == result.reports_run_dir / "manifest_run.json"
    assert result.manifest_path.exists()
    assert not (saida_reports / "runs" / "2026-05-07_001" / "manifest_run.json").exists()
    assert {path for _, path in calls} == {result.output_run_dir}


def test_pipeline_provisoria_nao_usa_raiz_output_como_entrada_consolidada(tmp_path):
    _, calls, saida_output, _ = _run_pipeline(tmp_path)

    paths_recebidos = {path for _, path in calls}
    assert saida_output not in paths_recebidos
    assert saida_output / "runs" not in paths_recebidos


def test_pipeline_provisoria_registra_relatorios_gerados(tmp_path):
    result, _, _, _ = _run_pipeline(tmp_path)

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    relatorios = manifest["relatorios_gerados"]
    assert str(result.output_run_dir / "55746722000181" / "lancamentos_fantozzi_real.xlsx") in relatorios
    assert str(result.reports_run_dir / "dre_gerencial_ok_provisoria.txt") in relatorios


def test_pipeline_dry_run_nao_escreve_nada(tmp_path):
    calls: list[tuple[str, Path]] = []
    xml_base = _make_xml_base(tmp_path)
    saida_output = tmp_path / "output"
    saida_reports = tmp_path / "reports"

    result = reprocessar_importacao_real_provisoria(
        xml_base,
        saida_output,
        saida_reports,
        dry_run=True,
        project_root=tmp_path,
        steps=_fake_steps(calls),
    )

    assert result.dry_run is True
    assert result.etapas_executadas == ()
    assert result.etapas_previstas == ETAPAS_PREVISTAS
    assert calls == []
    assert not (saida_output / "runs").exists()
    assert not (saida_reports / "runs").exists()
    assert not (saida_output / "current_run.txt").exists()
    assert not (saida_reports / "current_run.txt").exists()
    assert not result.manifest_path.exists()


def test_pipeline_quando_etapa_falha_status_manifesto_erro(tmp_path):
    calls: list[tuple[str, Path]] = []
    xml_base = _make_xml_base(tmp_path)
    saida_output = tmp_path / "output"
    saida_reports = tmp_path / "reports"

    with pytest.raises(ReprocessamentoPipelineError):
        reprocessar_importacao_real_provisoria(
            xml_base,
            saida_output,
            saida_reports,
            project_root=tmp_path,
            steps=_fake_steps(calls, fail_at="gerar_diagnosticos"),
        )

    run_id = (saida_reports / "current_run.txt").read_text(encoding="utf-8").strip()
    manifest_path = saida_reports / "runs" / run_id / "manifest_run.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert manifest["status_run"] == "ERRO"
    assert manifest["erros"]
    assert manifest["checkpoints"][-1]["descricao"] == "DIAGNOSTICOS_CONCLUIDOS_ERRO"


def test_pipeline_quando_etapa_falha_nao_continua_etapas_dependentes(tmp_path):
    calls: list[tuple[str, Path]] = []
    xml_base = _make_xml_base(tmp_path)

    with pytest.raises(ReprocessamentoPipelineError):
        reprocessar_importacao_real_provisoria(
            xml_base,
            tmp_path / "output",
            tmp_path / "reports",
            project_root=tmp_path,
            steps=_fake_steps(calls, fail_at="gerar_diagnosticos"),
        )

    assert [name for name, _ in calls] == ["importar_xml", "gerar_diagnosticos"]


def test_pipeline_quando_etapa_falha_preserva_arquivos_ja_gerados(tmp_path):
    calls: list[tuple[str, Path]] = []
    xml_base = _make_xml_base(tmp_path)
    saida_output = tmp_path / "output"
    saida_reports = tmp_path / "reports"

    with pytest.raises(ReprocessamentoPipelineError):
        reprocessar_importacao_real_provisoria(
            xml_base,
            saida_output,
            saida_reports,
            project_root=tmp_path,
            steps=_fake_steps(calls, fail_at="gerar_diagnosticos"),
        )

    run_id = (saida_reports / "current_run.txt").read_text(encoding="utf-8").strip()
    assert (
        saida_output
        / "runs"
        / run_id
        / "55746722000181"
        / "lancamentos_fantozzi_real.xlsx"
    ).exists()


def test_descobrir_fontes_xml_aceita_xml_direto_na_pasta_cnpj(tmp_path):
    xml_base = tmp_path / "XML_fatture_acquisto"
    cnpj_dir = xml_base / "55746722000181"
    cnpj_dir.mkdir(parents=True)
    (cnpj_dir / "nota.xml").write_text("<xml />", encoding="utf-8")

    fontes = descobrir_fontes_xml(xml_base)

    assert len(fontes) == 1
    assert fontes[0].cnpj_pasta == "55746722000181"
    assert fontes[0].xml_dir == cnpj_dir


def test_descobrir_fontes_xml_aceita_xml_direto_com_subpasta_xml_vazia(tmp_path):
    xml_base = tmp_path / "XML_fatture_acquisto"
    cnpj_dir = xml_base / "55746722000181"
    (cnpj_dir / "xml").mkdir(parents=True)
    (cnpj_dir / "nota.xml").write_text("<xml />", encoding="utf-8")

    fontes = descobrir_fontes_xml(xml_base)

    assert len(fontes) == 1
    assert fontes[0].cnpj_pasta == "55746722000181"
    assert fontes[0].xml_dir == cnpj_dir


def test_gerar_pacote_revisao_falha_sem_csvs(tmp_path):
    context = PipelineContext(
        xml_base=tmp_path / "XML_fatture_acquisto",
        saida_output=tmp_path / "output",
        saida_reports=tmp_path / "reports",
        project_root=tmp_path,
        run_id="2026-05-07_001",
        output_run_dir=tmp_path / "output" / "runs" / "2026-05-07_001",
        reports_run_dir=tmp_path / "reports" / "runs" / "2026-05-07_001",
        manifest_path=tmp_path / "reports" / "runs" / "2026-05-07_001" / "manifest_run.json",
        fontes_xml=(),
    )
    context.reports_run_dir.mkdir(parents=True)

    with pytest.raises(ReprocessamentoPipelineError, match="Nenhum CSV"):
        _step_gerar_pacote_revisao(context)


def _save_plano_contas_fase9(tmp_path: Path) -> Path:
    config_dir = tmp_path / "config"
    config_dir.mkdir(parents=True, exist_ok=True)
    plano = config_dir / "plano_contas_gerencial.yaml"
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


def _save_config_9n(tmp_path: Path) -> None:
    config_dir = tmp_path / "config"
    config_dir.mkdir(parents=True, exist_ok=True)
    (config_dir / "plano_contas_gerencial.yaml").write_text(
        """\
contas:
  - {id: "2.02.001", nome: "Carnes Bovinas", grupo: "Custos", tipo: custo}
  - {id: "10.01.001", nome: "Equipamentos de Cozinha", grupo: "Investimentos", tipo: investimento}
  - {id: "13.01.001", nome: "A Revisar", grupo: "Revisao", tipo: revisao}
  - {id: "13.01.002", nome: "Ambiguo", grupo: "Revisao", tipo: revisao}
  - {id: "13.01.006", nome: "Valor Inconsistente", grupo: "Revisao", tipo: revisao}
  - {id: "13.01.009", nome: "Possivel Investimento", grupo: "Revisao", tipo: revisao}
  - {id: "13.01.010", nome: "Pessoal a Regularizar", grupo: "Revisao", tipo: revisao}
  - {id: "13.01.999", nome: "Outros", grupo: "Revisao", tipo: revisao}
""",
        encoding="utf-8",
    )
    (config_dir / "mapeamento_dre.yaml").write_text(
        """\
linhas_dre:
  - {id: "03.01", nome: "CMV", grupo: "Custos", sinal: "-", ordem: 1}
  - {id: "12.01", nome: "Investimentos", grupo: "Investimentos", sinal: "-", ordem: 2}
mapeamento:
  "2.02.001": {dre_linha_id: "03.01", entra_dre: true}
  "10.01.001": {dre_linha_id: "12.01", entra_dre: false}
""",
        encoding="utf-8",
    )


def _save_manual_csv_9n(path: Path) -> None:
    with path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.DictWriter(
            file,
            fieldnames=[
                "data",
                "descricao_da_despesa",
                "conta_debito_id",
                "conta_debito_nome",
                "conta_credito_id",
                "conta_credito_nome",
                "valor",
                "observacao",
            ],
            delimiter=";",
        )
        writer.writeheader()
        writer.writerow(
            {
                "data": "2026-05-16",
                "descricao_da_despesa": "Ajuste manual gerencial",
                "conta_debito_id": "2.02.001",
                "conta_debito_nome": "Carnes Bovinas",
                "conta_credito_id": "10.01.001",
                "conta_credito_nome": "Equipamentos de Cozinha",
                "valor": "10,50",
                "observacao": "fixture sintetica",
            }
        )


def _save_lancamentos_banca_xlsx(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = LANCAMENTOS_SHEET
    sheet.append(HEADERS)
    base_row = {header: "" for header in HEADERS}
    base_row.update(
        {
            "data_documento": date(2026, 2, 5),
            "competencia": "2026-02",
            "cnpj_empresa": "08884180000123",
            "fornecedor_ou_favorecido": "IZALTINO JOSE ZANI JUNIOR",
            "tipo_documento": "BANCA_COMPROVANTE",
            "conta_id": "6.01.002",
            "conta_nome": "Honorarios Juridicos",
            "valor": 2000,
            "origem": "COMPROVANTES_BANCA",
            "status_revisao": "OK",
            "tipo_lancamento": "custo",
            "data_pagamento": date(2026, 2, 5),
            "duplicatas_json": "[]",
        }
    )
    second_row = {
        **base_row,
        "fornecedor_ou_favorecido": "TOP LEVEL DIGITAL",
        "tipo_documento": "BANCA_CONTA_CORRENTE",
        "conta_id": "9.01.001",
        "conta_nome": "Publicidade e Propaganda",
        "valor": 3290,
        "origem": "CONTA_CORRENTE",
    }
    sheet.append([base_row[header] for header in HEADERS])
    sheet.append([second_row[header] for header in HEADERS])
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def _save_xml_lancamentos_xlsx(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = LANCAMENTOS_SHEET
    sheet.append(HEADERS)
    row = {header: "" for header in HEADERS}
    row.update(
        {
            "data_documento": date(2026, 5, 16),
            "competencia": "2026-05",
            "cnpj_empresa": "08884180000123",
            "nome_empresa": "AMAMI RESTAURANTE LTDA",
            "fornecedor_ou_favorecido": "FORNECEDOR XML TESTE",
            "documento": "NF-100",
            "tipo_documento": "XML_NFE",
            "conta_id": "2.02.001",
            "conta_nome": "Carnes Bovinas",
            "valor": Decimal("10.50"),
            "origem": "XML_NFE",
            "arquivo_origem": "nfe.xml",
            "chave_origem": f"xml-{path.parent.name}",
            "hash_registro": f"hash-{path.parent.name}",
            "status_revisao": "OK",
            "tipo_lancamento": "custo",
            "duplicatas_json": "[]",
        }
    )
    sheet.append([row[header] for header in HEADERS])
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def _fase9_row(**overrides: str) -> dict[str, str]:
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


def _save_fase9_csv(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=tuple(rows[0].keys()), delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


def _read_csv_rows(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file, delimiter=";"))
