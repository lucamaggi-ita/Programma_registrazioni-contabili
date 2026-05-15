from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Callable, Mapping, Sequence

from fantozzi import run_manager
from fantozzi.allocation import allocate_nfe_to_lancamentos
from fantozzi.auditoria_relatorios_gerenciais_ok import (
    auditar_relatorios_gerenciais_ok_provisorios,
)
from fantozzi.chart_of_accounts import ChartOfAccounts
from fantozzi.classification_candidates import (
    generate_classification_candidates_report,
    save_report as save_candidates_report,
)
from fantozzi.classification_review_export import export_classification_review_csv
from fantozzi.classification_review_package import gerar_pacote_revisao_contabil
from fantozzi.dre_gerencial_provisoria import gerar_dre_gerencial_ok_provisoria
from fantozzi.excel_diagnostics import diagnose_excel
from fantozzi.excel_diagnostics import report_text as diagnostic_report_text
from fantozzi.excel_diagnostics import save_report as save_diagnostic_report
from fantozzi.excel_writer import write_lancamentos_to_excel
from fantozzi.fase9_decision_engine import (
    evaluate_fase9_decisions,
    load_fase9_decisions_from_csv,
)
from fantozzi.fase9_writeback import write_fase9_applied_csv
from fantozzi.nfe_importer import parse_nfe
from fantozzi.operation_classifier import (
    auditar_operacoes_especiais,
    relatorio_operacoes_especiais,
    salvar_relatorio as salvar_relatorio_operacoes,
)
from fantozzi.pipeline_paths import resolver_arquivos_lancamentos_xlsx
from fantozzi.pre_livro_diario import gerar_pre_livro_diario_teste
from fantozzi.product_classifier import ProductClassifier
from fantozzi.relatorios_gerenciais_ok import (
    gerar_relatorios_gerenciais_ok_provisorios,
)
from fantozzi.revision_report import generate_revision_report
from fantozzi.revision_report import save_report as save_revision_report


STATUS_PROVISORIO_CONCLUIDO = "PROVISORIO_CONCLUIDO"
MODO_DEFINITIVO_BLOQUEADO = (
    "Modo definitivo bloqueado: Fase 9 ainda nao foi aprovada."
)
AVISO_FASE9_BLOQUEADA = (
    "Fase 9 bloqueada. Revisao contabil real pendente. "
    "Relatorios gerados sao provisorios."
)

ETAPAS_PREVISTAS = (
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
)


class ReprocessamentoPipelineError(Exception):
    pass


@dataclass(frozen=True)
class XmlSource:
    cnpj_pasta: str
    xml_dir: Path


@dataclass(frozen=True)
class StepResult:
    arquivos_gerados: tuple[Path, ...] = ()
    totais: Mapping[str, object] = field(default_factory=dict)
    classificacao: Mapping[str, object] = field(default_factory=dict)
    avisos: tuple[str, ...] = ()


@dataclass
class PipelineContext:
    xml_base: Path
    saida_output: Path
    saida_reports: Path
    project_root: Path
    run_id: str
    output_run_dir: Path
    reports_run_dir: Path
    manifest_path: Path
    fontes_xml: tuple[XmlSource, ...]
    fase9_csv_paths: tuple[Path, ...] = ()
    lancamentos_xlsx: tuple[Path, ...] = ()


@dataclass(frozen=True)
class PipelineSteps:
    importar_xml: Callable[[PipelineContext], StepResult]
    gerar_diagnosticos: Callable[[PipelineContext], StepResult]
    auditar_operacoes: Callable[[PipelineContext], StepResult]
    gerar_pre_livro: Callable[[PipelineContext], StepResult]
    gerar_pacote_revisao: Callable[[PipelineContext], StepResult]
    gerar_relatorios_ok: Callable[[PipelineContext], StepResult]
    auditar_relatorios_ok: Callable[[PipelineContext], StepResult]
    gerar_dre_ok: Callable[[PipelineContext], StepResult]


@dataclass(frozen=True)
class ReprocessamentoResultado:
    run_id: str
    output_run_dir: Path
    reports_run_dir: Path
    manifest_path: Path
    dry_run: bool
    status_run: str
    etapas_executadas: tuple[str, ...]
    arquivos_gerados: tuple[Path, ...]
    avisos: tuple[str, ...]
    etapas_previstas: tuple[str, ...] = ()


def reprocessar_importacao_real_provisoria(
    xml_base: Path,
    saida_output: Path,
    saida_reports: Path,
    *,
    modo: str = "provisorio",
    dry_run: bool = False,
    project_root: Path | None = None,
    steps: PipelineSteps | None = None,
    today: date | None = None,
    fase9_csv_path: Path | str | Sequence[Path | str] | None = None,
) -> ReprocessamentoResultado:
    xml_base = Path(xml_base)
    saida_output = Path(saida_output)
    saida_reports = Path(saida_reports)
    project_root = Path(project_root or run_manager.PROJECT_ROOT)
    fase9_csv_paths = _normalize_fase9_csv_paths(fase9_csv_path)

    if modo == "definitivo":
        raise ReprocessamentoPipelineError(MODO_DEFINITIVO_BLOQUEADO)
    if modo != "provisorio":
        raise ReprocessamentoPipelineError(f"Modo invalido: {modo}")

    if dry_run:
        planned_run = run_manager.preparar_run(
            xml_base=xml_base,
            saida_output=saida_output,
            saida_reports=saida_reports,
            dry_run=True,
            project_root=project_root,
            today=today,
        )
        return ReprocessamentoResultado(
            run_id=planned_run.run_id,
            output_run_dir=planned_run.output_run_dir,
            reports_run_dir=planned_run.reports_run_dir,
            manifest_path=planned_run.manifest_path,
            dry_run=True,
            status_run=run_manager.STATUS_DRY_RUN,
            etapas_executadas=(),
            arquivos_gerados=(),
            avisos=(
                "Nenhum arquivo foi criado.",
                "Nenhum XML foi importado.",
                "Nenhum relatorio foi gerado.",
            ),
            etapas_previstas=ETAPAS_PREVISTAS,
        )

    fontes_xml = descobrir_fontes_xml(xml_base)
    if not fontes_xml:
        raise ReprocessamentoPipelineError(
            f"Nenhuma fonte XML encontrada em: {xml_base}"
        )

    run_result = run_manager.preparar_run(
        xml_base=xml_base,
        saida_output=saida_output,
        saida_reports=saida_reports,
        dry_run=False,
        project_root=project_root,
        today=today,
    )

    context = PipelineContext(
        xml_base=xml_base,
        saida_output=saida_output,
        saida_reports=saida_reports,
        project_root=project_root,
        run_id=run_result.run_id,
        output_run_dir=run_result.output_run_dir,
        reports_run_dir=run_result.reports_run_dir,
        manifest_path=run_result.manifest_path,
        fontes_xml=fontes_xml,
        fase9_csv_paths=fase9_csv_paths,
    )
    steps = steps or default_pipeline_steps()
    arquivos_gerados: list[Path] = []
    avisos: list[str] = []
    etapas_executadas: list[str] = []

    run_manager.add_checkpoint(context.manifest_path, "CRIACAO_DA_RUN")
    etapas_executadas.append("CRIACAO_DA_RUN")

    try:
        _executar_step(
            context,
            "IMPORTACAO_XML_CONCLUIDA",
            steps.importar_xml,
            arquivos_gerados,
            avisos,
            etapas_executadas,
        )
        context.lancamentos_xlsx = resolver_arquivos_lancamentos_xlsx(
            context.output_run_dir
        )
        _executar_step(
            context,
            "DIAGNOSTICOS_CONCLUIDOS",
            steps.gerar_diagnosticos,
            arquivos_gerados,
            avisos,
            etapas_executadas,
        )
        _executar_step(
            context,
            "OPERACOES_ESPECIAIS_AUDITADAS",
            steps.auditar_operacoes,
            arquivos_gerados,
            avisos,
            etapas_executadas,
        )
        _executar_step(
            context,
            ("PRE_LIVRO_GERADO", "AUDITORIA_PRE_LIVRO_CONCLUIDA"),
            steps.gerar_pre_livro,
            arquivos_gerados,
            avisos,
            etapas_executadas,
        )
        _executar_step(
            context,
            "PACOTE_REVISAO_CONTABIL_GERADO",
            steps.gerar_pacote_revisao,
            arquivos_gerados,
            avisos,
            etapas_executadas,
        )
        if context.fase9_csv_paths:
            _executar_step(
                context,
                "FASE9_APLICADA",
                _step_aplicar_fase9,
                arquivos_gerados,
                avisos,
                etapas_executadas,
            )
        _executar_step(
            context,
            "RELATORIOS_OK_GERADOS",
            steps.gerar_relatorios_ok,
            arquivos_gerados,
            avisos,
            etapas_executadas,
        )
        _executar_step(
            context,
            "AUDITORIA_RELATORIOS_OK_CONCLUIDA",
            steps.auditar_relatorios_ok,
            arquivos_gerados,
            avisos,
            etapas_executadas,
        )
        _executar_step(
            context,
            "DRE_OK_PROVISORIA_GERADA",
            steps.gerar_dre_ok,
            arquivos_gerados,
            avisos,
            etapas_executadas,
        )
    except Exception as exc:
        if isinstance(exc, ReprocessamentoPipelineError):
            raise
        raise ReprocessamentoPipelineError(str(exc)) from exc

    manifest = run_manager.load_manifest(context.manifest_path)
    manifest.setdefault("avisos", [])
    if not context.fase9_csv_paths and AVISO_FASE9_BLOQUEADA not in manifest["avisos"]:
        manifest["avisos"].append(AVISO_FASE9_BLOQUEADA)
    manifest.setdefault("classificacao", {})
    if not context.fase9_csv_paths:
        manifest["classificacao"].update(
            {
                "fase9_executada": False,
                "regras_aplicadas": 0,
                "nota": "Revisao contabil pendente. Nenhuma regra aprovada aplicada.",
            }
        )
    run_manager.save_manifest(context.manifest_path, _json_ready(manifest))
    run_manager.add_checkpoint(context.manifest_path, "FINALIZACAO_PROVISORIA")
    run_manager.update_manifest_status(
        context.manifest_path,
        STATUS_PROVISORIO_CONCLUIDO,
    )
    etapas_executadas.append("FINALIZACAO_PROVISORIA")

    final_manifest = run_manager.load_manifest(context.manifest_path)
    return ReprocessamentoResultado(
        run_id=context.run_id,
        output_run_dir=context.output_run_dir,
        reports_run_dir=context.reports_run_dir,
        manifest_path=context.manifest_path,
        dry_run=False,
        status_run=final_manifest["status_run"],
        etapas_executadas=tuple(etapas_executadas),
        arquivos_gerados=tuple(arquivos_gerados),
        avisos=tuple(final_manifest.get("avisos", ())),
        etapas_previstas=ETAPAS_PREVISTAS,
    )


def descobrir_fontes_xml(xml_base: Path) -> tuple[XmlSource, ...]:
    """Descobre fontes XML por CNPJ.

    Estrutura preferida: <xml_base>/<cnpj>/xml/*.xml.
    Tambem aceita XML diretamente em <xml_base>/<cnpj> para evitar silencio
    quando a subpasta xml ainda nao existir.
    """
    xml_base = Path(xml_base)
    if not xml_base.exists() or not xml_base.is_dir():
        return ()
    fontes: list[XmlSource] = []
    for subdir in sorted(xml_base.iterdir(), key=lambda path: path.name.lower()):
        if not subdir.is_dir():
            continue
        xml_dir = subdir / "xml"
        xml_dir_files = _find_xml_files(xml_dir) if xml_dir.is_dir() else []
        direct_xml_files = _find_xml_files(subdir)
        if xml_dir_files:
            fontes.append(XmlSource(cnpj_pasta=subdir.name, xml_dir=xml_dir))
        elif direct_xml_files:
            fontes.append(XmlSource(cnpj_pasta=subdir.name, xml_dir=subdir))
    return tuple(fontes)


def default_pipeline_steps() -> PipelineSteps:
    return PipelineSteps(
        importar_xml=_step_importar_xml,
        gerar_diagnosticos=_step_gerar_diagnosticos,
        auditar_operacoes=_step_auditar_operacoes,
        gerar_pre_livro=_step_gerar_pre_livro,
        gerar_pacote_revisao=_step_gerar_pacote_revisao,
        gerar_relatorios_ok=_step_gerar_relatorios_ok,
        auditar_relatorios_ok=_step_auditar_relatorios_ok,
        gerar_dre_ok=_step_gerar_dre_ok,
    )


def _executar_step(
    context: PipelineContext,
    checkpoints: str | tuple[str, ...],
    step: Callable[[PipelineContext], StepResult],
    arquivos_gerados: list[Path],
    avisos: list[str],
    etapas_executadas: list[str],
) -> None:
    checkpoint_tuple = (checkpoints,) if isinstance(checkpoints, str) else checkpoints
    try:
        result = step(context)
    except Exception as exc:
        _registrar_erro(context.manifest_path, checkpoint_tuple[0], exc)
        raise ReprocessamentoPipelineError(str(exc)) from exc

    arquivos_gerados.extend(result.arquivos_gerados)
    avisos.extend(result.avisos)
    _atualizar_manifesto_step(context.manifest_path, result)
    for checkpoint in checkpoint_tuple:
        run_manager.add_checkpoint(context.manifest_path, checkpoint)
        etapas_executadas.append(checkpoint)


def _registrar_erro(manifest_path: Path, checkpoint: str, exc: Exception) -> None:
    erro = f"{checkpoint}: {exc}"
    try:
        run_manager.add_checkpoint(manifest_path, f"{checkpoint}_ERRO")
        manifest = run_manager.load_manifest(manifest_path)
        manifest.setdefault("erros", []).append(erro)
        manifest["status_run"] = run_manager.STATUS_ERRO
        run_manager.save_manifest(manifest_path, _json_ready(manifest))
        run_manager.update_manifest_status(manifest_path, run_manager.STATUS_ERRO)
    except Exception:
        pass


def _atualizar_manifesto_step(manifest_path: Path, result: StepResult) -> None:
    manifest = run_manager.load_manifest(manifest_path)
    manifest.setdefault("totais", {}).update(result.totais)
    manifest.setdefault("classificacao", {}).update(result.classificacao)
    manifest.setdefault("avisos", [])
    for aviso in result.avisos:
        if aviso not in manifest["avisos"]:
            manifest["avisos"].append(aviso)
    manifest.setdefault("relatorios_gerados", [])
    for path in result.arquivos_gerados:
        path_text = str(path)
        if path_text not in manifest["relatorios_gerados"]:
            manifest["relatorios_gerados"].append(path_text)
    run_manager.save_manifest(manifest_path, _json_ready(manifest))


def _step_importar_xml(context: PipelineContext) -> StepResult:
    classifier = _load_product_classifier(
        context.project_root / "config" / "plano_contas_gerencial.yaml",
        context.project_root / "config" / "regras_classificacao_produtos.yaml",
    )
    arquivos: list[Path] = []
    avisos: list[str] = []
    xml_processados = 0
    lancamentos_gerados = 0
    valor_total = Decimal("0.00")

    for fonte in context.fontes_xml:
        xml_files = _find_xml_files(fonte.xml_dir)
        lancamentos = []
        for xml_path in xml_files:
            try:
                nfe = parse_nfe(xml_path)
                file_lancamentos = allocate_nfe_to_lancamentos(nfe, classifier)
            except Exception as exc:
                avisos.append(f"XML ignorado: {xml_path} - {exc}")
                continue
            xml_processados += 1
            valor_total += nfe.valor_total
            lancamentos_gerados += len(file_lancamentos)
            lancamentos.extend(file_lancamentos)

        excel_path = (
            context.output_run_dir
            / fonte.cnpj_pasta
            / "lancamentos_fantozzi_real.xlsx"
        )
        write_lancamentos_to_excel(excel_path, lancamentos)
        arquivos.append(excel_path)

    if xml_processados == 0:
        raise ReprocessamentoPipelineError("Nenhum XML valido importado.")

    return StepResult(
        arquivos_gerados=tuple(arquivos),
        totais={
            "xml_processados": xml_processados,
            "lancamentos_gerados": lancamentos_gerados,
            "valor_total_importado": valor_total,
        },
        avisos=tuple(avisos),
    )


def _step_gerar_diagnosticos(context: PipelineContext) -> StepResult:
    arquivos: list[Path] = []
    candidatos = 0
    for excel_path in _lancamentos_xlsx(context):
        cnpj_pasta = excel_path.parent.name
        report_dir = context.reports_run_dir / cnpj_pasta
        diagnostic_path = report_dir / "diagnostico_excel_real.txt"
        revision_path = report_dir / "revisao_lancamentos_real.txt"
        candidates_path = report_dir / "candidatos_classificacao_real.txt"
        review_csv_path = report_dir / "revisao_classificacao_real.csv"

        diagnostic = diagnose_excel(excel_path)
        save_diagnostic_report(diagnostic_path, diagnostic_report_text(diagnostic))

        _, revision_text = generate_revision_report(excel_path)
        save_revision_report(revision_path, revision_text)

        candidates_result, candidates_text = generate_classification_candidates_report(
            excel_path
        )
        save_candidates_report(candidates_path, candidates_text)
        export_result = export_classification_review_csv(excel_path, review_csv_path)
        candidatos += export_result.exported_candidates
        arquivos.extend(
            [diagnostic_path, revision_path, candidates_path, review_csv_path]
        )

    return StepResult(
        arquivos_gerados=tuple(arquivos),
        classificacao={"candidatos_manter_revisao": candidatos},
    )


def _step_auditar_operacoes(context: PipelineContext) -> StepResult:
    arquivos: list[Path] = []
    total_operacoes = 0
    for fonte in context.fontes_xml:
        result = auditar_operacoes_especiais(fonte.xml_dir)
        text = relatorio_operacoes_especiais(result)
        report_path = (
            context.reports_run_dir / fonte.cnpj_pasta / "operacoes_especiais.txt"
        )
        salvar_relatorio_operacoes(report_path, text)
        arquivos.append(report_path)
        total_operacoes += len(result.itens)
    return StepResult(
        arquivos_gerados=tuple(arquivos),
        totais={"operacoes_especiais": total_operacoes},
    )


def _step_gerar_pre_livro(context: PipelineContext) -> StepResult:
    output_path = context.reports_run_dir / "pre_livro_diario_gerencial_teste.xlsx"
    audit_path = (
        context.reports_run_dir / "auditoria_pre_livro_diario_gerencial_teste.txt"
    )
    result = gerar_pre_livro_diario_teste(
        _lancamentos_xlsx(context),
        output_path,
        _xml_dirs(context),
        audit_path,
    )
    return StepResult(
        arquivos_gerados=(output_path, audit_path),
        totais={
            "lancamentos_ok": result.lancamentos_ok,
            "lancamentos_revisar": result.lancamentos_revisar,
        },
    )


def _step_gerar_pacote_revisao(context: PipelineContext) -> StepResult:
    csv_paths = tuple(
        sorted(
            context.reports_run_dir.glob("*/revisao_classificacao_real.csv"),
            key=lambda path: str(path).lower(),
        )
    )
    if not csv_paths:
        raise ReprocessamentoPipelineError(
            "Nenhum CSV de revisao_classificacao_real.csv encontrado "
            "para gerar pacote de revisao contabil."
        )
    output_path = context.reports_run_dir / "pacote_revisao_contabil_fase8_1.xlsx"
    report_path = context.reports_run_dir / "pacote_revisao_contabil_fase8_1.txt"
    result = gerar_pacote_revisao_contabil(
        csv_paths,
        output_path,
        report_path,
    )
    return StepResult(
        arquivos_gerados=(output_path, report_path),
        classificacao={
            "fase9_executada": False,
            "regras_aplicadas": 0,
            "candidatos_manter_revisao": result.total_candidatos,
        },
    )


def _normalize_fase9_csv_paths(
    fase9_csv_path: Path | str | Sequence[Path | str] | None,
) -> tuple[Path, ...]:
    if fase9_csv_path is None:
        return ()
    if isinstance(fase9_csv_path, Path | str):
        return (Path(fase9_csv_path),)
    return tuple(Path(path) for path in fase9_csv_path)


def _fase9_output_paths(
    reports_run_dir: Path,
    fase9_csv_paths: tuple[Path, ...],
) -> tuple[Path, ...]:
    if len(fase9_csv_paths) == 1:
        return (reports_run_dir / "lancamentos_aplicados_fase9.csv",)

    outputs: list[Path] = []
    used_names: set[str] = set()
    for csv_path in fase9_csv_paths:
        identifier = _fase9_csv_identifier(csv_path)
        candidate = reports_run_dir / f"lancamentos_aplicados_fase9_{identifier}.csv"
        suffix = 2
        while candidate.name.lower() in used_names:
            candidate = (
                reports_run_dir
                / f"lancamentos_aplicados_fase9_{identifier}_{suffix}.csv"
            )
            suffix += 1
        used_names.add(candidate.name.lower())
        outputs.append(candidate)
    return tuple(outputs)


def _fase9_csv_identifier(csv_path: Path) -> str:
    parent_name = csv_path.parent.name
    if parent_name.isdigit():
        return _sanitize_fase9_identifier(parent_name)
    return _sanitize_fase9_identifier(csv_path.stem)


def _sanitize_fase9_identifier(value: str) -> str:
    sanitized = "".join(
        char if char.isalnum() or char in ("-", "_") else "_"
        for char in value.strip()
    ).strip("_")
    return sanitized or "csv"


def _step_aplicar_fase9(context: PipelineContext) -> StepResult:
    if not context.fase9_csv_paths:
        return StepResult()

    chart = ChartOfAccounts.load(
        context.project_root / "config" / "plano_contas_gerencial.yaml"
    )
    validation_errors = chart.validate()
    if validation_errors:
        raise ReprocessamentoPipelineError(
            "Plano de contas invalido: " + "; ".join(validation_errors)
        )
    valid_account_ids, account_names_by_id = _fase9_account_maps(chart)
    csv_outputs = _fase9_output_paths(
        context.reports_run_dir,
        context.fase9_csv_paths,
    )

    arquivos: list[Path] = []
    csvs_processados: list[dict[str, object]] = []
    promovidos = 0
    bloqueados = 0
    fora_escopo = 0
    for input_csv, output_csv in zip(
        context.fase9_csv_paths,
        csv_outputs,
        strict=True,
    ):
        if input_csv.resolve() == output_csv.resolve():
            raise ReprocessamentoPipelineError(
                "CSV de entrada da Fase 9 deve ser diferente do output aplicado."
            )
        if not input_csv.exists() or not input_csv.is_file():
            raise ReprocessamentoPipelineError(
                f"CSV revisionado Fase 9 nao encontrado: {input_csv}"
            )
        decisions = load_fase9_decisions_from_csv(input_csv)
        decision_result = evaluate_fase9_decisions(
            decisions,
            valid_account_ids=valid_account_ids,
            account_names_by_id=account_names_by_id,
        )
        writeback_result = write_fase9_applied_csv(
            input_csv,
            output_csv,
            decision_result,
        )
        arquivos.append(output_csv)
        promovidos += writeback_result.promoted_rows
        bloqueados += writeback_result.blocked_rows
        fora_escopo += writeback_result.out_of_scope_rows
        csvs_processados.append(
            {
                "input_csv": str(input_csv),
                "output_csv": str(output_csv),
                "promovidos": writeback_result.promoted_rows,
                "bloqueados": writeback_result.blocked_rows,
                "fora_escopo": writeback_result.out_of_scope_rows,
            }
        )

    classificacao: dict[str, object] = {
        "fase9_executada": True,
        "promovidos": promovidos,
        "bloqueados": bloqueados,
        "fora_escopo": fora_escopo,
        "regras_aplicadas": promovidos,
        "csvs_processados": csvs_processados,
        "nota": (
            "Fase 9 aplicada em CSV separado. "
            "Relatorios gerados permanecem provisorios."
        ),
    }
    if len(csvs_processados) == 1:
        classificacao["input_csv"] = csvs_processados[0]["input_csv"]
        classificacao["output_csv"] = csvs_processados[0]["output_csv"]

    return StepResult(
        arquivos_gerados=tuple(arquivos),
        classificacao=classificacao,
    )


def _step_gerar_relatorios_ok(context: PipelineContext) -> StepResult:
    output_path = (
        context.reports_run_dir / "relatorios_gerenciais_ok_provisorios.xlsx"
    )
    report_path = (
        context.reports_run_dir / "relatorios_gerenciais_ok_provisorios.txt"
    )
    result = gerar_relatorios_gerenciais_ok_provisorios(
        [context.output_run_dir],
        output_path,
        report_path,
        _xml_dirs(context),
    )
    return StepResult(
        arquivos_gerados=(output_path, report_path),
        totais={
            "lancamentos_ok": result.lancamentos_ok_incluidos,
            "lancamentos_revisar": result.lancamentos_revisar_excluidos,
            "valor_total_incluido_ok": result.valor_total_incluido,
        },
    )


def _step_auditar_relatorios_ok(context: PipelineContext) -> StepResult:
    output_path = (
        context.reports_run_dir
        / "auditoria_relatorios_gerenciais_ok_provisorios.xlsx"
    )
    report_path = (
        context.reports_run_dir
        / "auditoria_relatorios_gerenciais_ok_provisorios.txt"
    )
    auditar_relatorios_gerenciais_ok_provisorios(
        [context.output_run_dir],
        output_path,
        report_path,
        _xml_dirs(context),
    )
    return StepResult(arquivos_gerados=(output_path, report_path))


def _step_gerar_dre_ok(context: PipelineContext) -> StepResult:
    output_path = context.reports_run_dir / "dre_gerencial_ok_provisoria.xlsx"
    report_path = context.reports_run_dir / "dre_gerencial_ok_provisoria.txt"
    gerar_dre_gerencial_ok_provisoria(
        [context.output_run_dir],
        context.project_root / "config" / "mapeamento_dre.yaml",
        output_path,
        report_path,
        _xml_dirs(context),
    )
    return StepResult(arquivos_gerados=(output_path, report_path))


def _load_product_classifier(
    plano_contas_path: Path,
    regras_produtos_path: Path,
) -> ProductClassifier:
    chart = ChartOfAccounts.load(plano_contas_path)
    validation_errors = chart.validate()
    if validation_errors:
        raise ValueError("; ".join(validation_errors))
    return ProductClassifier(chart).load(regras_produtos_path)


def _fase9_account_maps(chart: ChartOfAccounts) -> tuple[frozenset[str], dict[str, str]]:
    accounts = tuple(chart._ordered)
    return (
        frozenset(account.id for account in accounts),
        {account.id: account.nome for account in accounts},
    )


def _find_xml_files(pasta_path: Path) -> list[Path]:
    return sorted(
        (
            path
            for path in pasta_path.iterdir()
            if path.is_file() and path.suffix.lower() == ".xml"
        ),
        key=lambda path: path.name.lower(),
    )


def _lancamentos_xlsx(context: PipelineContext) -> tuple[Path, ...]:
    if context.lancamentos_xlsx:
        return context.lancamentos_xlsx
    context.lancamentos_xlsx = resolver_arquivos_lancamentos_xlsx(
        context.output_run_dir
    )
    return context.lancamentos_xlsx


def _xml_dirs(context: PipelineContext) -> tuple[Path, ...]:
    return tuple(fonte.xml_dir for fonte in context.fontes_xml)


def _json_ready(value: object) -> object:
    if isinstance(value, Decimal):
        return str(value.quantize(Decimal("0.01")))
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, Mapping):
        return {str(k): _json_ready(v) for k, v in value.items()}
    if isinstance(value, tuple | list):
        return [_json_ready(item) for item in value]
    return value
