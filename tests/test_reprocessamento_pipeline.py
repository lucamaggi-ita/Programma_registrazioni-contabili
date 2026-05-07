from __future__ import annotations

import json
from pathlib import Path

import pytest

from fantozzi.reprocessamento_pipeline import (
    MODO_DEFINITIVO_BLOQUEADO,
    PipelineSteps,
    ReprocessamentoPipelineError,
    StepResult,
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
            path.parent.mkdir(parents=True)
            path.write_text("fake xlsx", encoding="utf-8")
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
        xlsx.write_text("pre", encoding="utf-8")
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


def test_pipeline_provisoria_nao_executa_fase9(tmp_path):
    result, _, _, _ = _run_pipeline(tmp_path)

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert manifest["classificacao"]["fase9_executada"] is False


def test_pipeline_provisoria_nao_aplica_regras(tmp_path):
    result, _, _, _ = _run_pipeline(tmp_path)

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert manifest["classificacao"]["regras_aplicadas"] == 0


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
    )

    assert result.run_id != "2026-05-07_001"
    assert result.output_run_dir.name.endswith("_002")
    assert (saida_output / "runs" / "2026-05-07_001" / "antigo").exists()


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
    assert calls == []
    assert not (saida_output / "runs").exists()
    assert not (saida_reports / "runs").exists()
    assert not (saida_output / "current_run.txt").exists()
    assert not (saida_reports / "current_run.txt").exists()


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
