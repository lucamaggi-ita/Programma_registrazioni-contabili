from __future__ import annotations

import argparse
from pathlib import Path

from fantozzi.allocation import allocate_nfe_to_lancamentos
from fantozzi.chart_of_accounts import ChartOfAccounts
from fantozzi.classification_candidates import (
    ClassificationCandidatesError,
    console_lines as candidates_console_lines,
    generate_classification_candidates_report,
    save_report as save_candidates_report,
)
from fantozzi.classification_review_export import (
    ClassificationReviewExportError,
    console_lines as review_export_console_lines,
    export_classification_review_csv,
)
from fantozzi.classification_review_workbook import (
    ClassificationReviewWorkbookError,
    gerar_planilha_revisao_classificacao,
)
from fantozzi.classification_review_workbook_importer import (
    ClassificationReviewWorkbookImportError,
    console_lines as workbook_import_console_lines,
    importar_planilha_revisao_classificacao,
)
from fantozzi.classification_review_validation import (
    ClassificationReviewValidationError,
    report_text as validation_report_text,
    save_report as save_validation_report,
    validate_classification_review_csv,
)
from fantozzi.classification_rule_proposal import (
    ClassificationRuleProposalError,
    console_lines as proposal_console_lines,
    console_lines_real as proposal_real_console_lines,
    escrever_proposta_yaml,
    gerar_proposta_classificacao,
    gerar_proposta_regras_classificacao_real,
    gerar_relatorio_proposta_classificacao,
    salvar_relatorio_proposta,
)
from fantozzi.excel_diagnostics import (
    ExcelDiagnosticError,
    ExcelDiagnosticFailure,
    console_lines,
    diagnose_excel,
    report_text,
    save_report,
)
from fantozzi.excel_writer import write_lancamentos_to_excel
from fantozzi.logging_utils import error, info, warning
from fantozzi.nfe_importer import parse_nfe
from fantozzi.operation_classifier import (
    OperationAuditError,
    auditar_operacoes_especiais,
    console_lines as operation_audit_console_lines,
    relatorio_operacoes_especiais,
    salvar_relatorio as salvar_relatorio_operacoes,
)
from fantozzi.product_classifier import ProductClassifier
from fantozzi.revision_report import (
    RevisionReportError,
    console_lines as revision_console_lines,
    generate_revision_report,
    save_report as save_revision_report,
)


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PLANO_CONTAS = PROJECT_ROOT / "config" / "plano_contas_gerencial.yaml"
DEFAULT_REGRAS_PRODUTOS = (
    PROJECT_ROOT / "config" / "regras_classificacao_produtos.yaml"
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="fantozzi",
        description=(
            "Fantozzi importa XML NF-e, atualiza planilhas Excel de lancamentos "
            "e executa diagnosticos operacionais."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Exemplos PowerShell:\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli importar-nfe "
            "--xml samples\\nfe_exemplo\\nota.xml --excel output\\lancamentos.xlsx\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli importar-pasta-nfe "
            "--pasta samples\\nfe_exemplo --excel output\\lancamentos.xlsx\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli diagnosticar-excel "
            "--excel output\\lancamentos.xlsx --relatorio reports\\diagnostico.txt\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli relatorio-revisao "
            "--excel output\\lancamentos.xlsx --saida reports\\revisao_lancamentos.txt\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli "
            "relatorio-candidatos-classificacao --excel output\\lancamentos.xlsx "
            "--saida reports\\candidatos_classificacao.txt\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli "
            "exportar-revisao-classificacao --excel output\\lancamentos.xlsx "
            "--saida reports\\revisao_classificacao.csv\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli "
            "validar-revisao-classificacao --csv reports\\revisao_classificacao.csv "
            "--relatorio reports\\validacao_revisao_classificacao.txt\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli "
            "gerar-proposta-classificacao --csv reports\\revisao_classificacao.csv "
            "--saida-yaml reports\\proposta_regras_classificacao.yaml "
            "--relatorio reports\\proposta_regras_classificacao.txt\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli "
            "auditar-operacoes-especiais --pasta XML_fatture_acquisto\\xml "
            "--saida reports\\operacoes_especiais.txt\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli "
            "gerar-planilha-revisao-classificacao "
            "--csv reports\\revisao_classificacao.csv "
            "--saida reports\\revisao_classificacao_assistida.xlsx\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli "
            "importar-planilha-revisao-classificacao "
            "--xlsx reports\\revisao_classificacao_assistida.xlsx "
            "--saida reports\\revisao_classificacao_revisada.csv\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli "
            "gerar-proposta-regras-classificacao-real "
            "--csv reports\\revisao_classificacao_revisada.csv "
            "--saida reports\\proposta_regras_classificacao_real.yaml "
            "--relatorio reports\\resumo_proposta_regras_classificacao_real.txt"
        ),
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    importar_nfe = subparsers.add_parser(
        "importar-nfe",
        help="Importa um arquivo XML NF-e e atualiza um Excel de lancamentos.",
        description="Importa um arquivo XML NF-e e atualiza um Excel de lancamentos.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    importar_nfe.add_argument("--xml", required=True, help="Arquivo XML NF-e de entrada.")
    importar_nfe.add_argument(
        "--excel", required=True, help="Arquivo Excel .xlsx de saida/atualizacao."
    )
    _add_config_arguments(importar_nfe)
    importar_nfe.set_defaults(func=importar_nfe_command)

    importar_pasta_nfe = subparsers.add_parser(
        "importar-pasta-nfe",
        help="Importa uma pasta de arquivos XML NF-e e atualiza um Excel.",
        description="Importa todos os XML NF-e de uma pasta e atualiza um Excel.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    importar_pasta_nfe.add_argument(
        "--pasta", required=True, help="Pasta com arquivos XML NF-e de entrada."
    )
    importar_pasta_nfe.add_argument(
        "--excel", required=True, help="Arquivo Excel .xlsx de saida/atualizacao."
    )
    _add_config_arguments(importar_pasta_nfe)
    importar_pasta_nfe.set_defaults(func=importar_pasta_nfe_command)

    diagnosticar_excel = subparsers.add_parser(
        "diagnosticar-excel",
        help="Analisa um Excel gerado e imprime um diagnostico estrutural.",
        description=(
            "Analisa um Excel gerado pelo Fantozzi sem alterar o arquivo original."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    diagnosticar_excel.add_argument(
        "--excel", required=True, help="Arquivo Excel .xlsx para diagnosticar."
    )
    diagnosticar_excel.add_argument(
        "--relatorio",
        help="Arquivo TXT para salvar o relatorio da diagnostica.",
    )
    diagnosticar_excel.set_defaults(func=diagnosticar_excel_command)

    relatorio_revisao = subparsers.add_parser(
        "relatorio-revisao",
        help="Gera um relatorio business dos lancamentos em revisao.",
        description=(
            "Le o Excel gerado pelo Fantozzi e gera um relatorio read-only "
            "dos lancamentos que exigem revisao."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    relatorio_revisao.add_argument(
        "--excel", required=True, help="Arquivo Excel .xlsx para analisar."
    )
    relatorio_revisao.add_argument(
        "--saida",
        help="Arquivo TXT para salvar o relatorio de revisao.",
    )
    relatorio_revisao.set_defaults(func=relatorio_revisao_command)

    relatorio_candidatos = subparsers.add_parser(
        "relatorio-candidatos-classificacao",
        help="Gera relatorio de candidatos para novas regras de classificacao.",
        description=(
            "Gera relatorio read-only de candidatos para novas regras de "
            "classificacao a partir dos lancamentos em revisao."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    relatorio_candidatos.add_argument(
        "--excel", required=True, help="Arquivo Excel .xlsx para analisar."
    )
    relatorio_candidatos.add_argument(
        "--saida",
        help="Arquivo TXT para salvar o relatorio de candidatos.",
    )
    relatorio_candidatos.set_defaults(func=relatorio_candidatos_command)

    exportar_revisao = subparsers.add_parser(
        "exportar-revisao-classificacao",
        help="Exporta CSV para revisao manual de candidatos de classificacao.",
        description=(
            "Exporta um CSV de trabalho para revisao manual de candidatos de "
            "classificacao. Le o Excel em modo somente leitura e nao altera "
            "regras, YAML ou lancamentos."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    exportar_revisao.add_argument(
        "--excel",
        required=True,
        help="Caminho do arquivo Excel de lancamentos do Fantozzi.",
    )
    exportar_revisao.add_argument(
        "--saida",
        required=True,
        help="Caminho do CSV de saida para revisao manual.",
    )
    exportar_revisao.set_defaults(func=exportar_revisao_classificacao_command)

    validar_revisao = subparsers.add_parser(
        "validar-revisao-classificacao",
        help="Valida CSV de revisao manual de classificacao.",
        description=(
            "Valida o CSV de revisao manual de classificacao sem alterar CSV, "
            "Excel, YAML, regras ou aliases."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    validar_revisao.add_argument(
        "--csv",
        required=True,
        help="Caminho do CSV de revisao manual gerado pela Fase 5C.",
    )
    validar_revisao.add_argument(
        "--relatorio",
        help="Arquivo TXT opcional para salvar o relatorio de validacao.",
    )
    validar_revisao.set_defaults(func=validar_revisao_classificacao_command)

    gerar_proposta = subparsers.add_parser(
        "gerar-proposta-classificacao",
        help="Gera proposta YAML de regras a partir do CSV validado.",
        description=(
            "Le o CSV de revisao manual de classificacao e gera uma proposta "
            "separada de regras/aliases para revisao. Nao altera CSV, Excel ou "
            "arquivos YAML de config."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    gerar_proposta.add_argument(
        "--csv",
        required=True,
        help="Caminho do CSV de revisao manual gerado pela Fase 5C.",
    )
    gerar_proposta.add_argument(
        "--saida-yaml",
        help="Arquivo YAML separado para salvar a proposta.",
    )
    gerar_proposta.add_argument(
        "--relatorio",
        help="Arquivo TXT opcional para salvar o relatorio da proposta.",
    )
    gerar_proposta.add_argument(
        "--plano-contas",
        default=str(DEFAULT_PLANO_CONTAS),
        help="Arquivo YAML do plano de contas gerencial para validacao.",
    )
    gerar_proposta.add_argument(
        "--regras-produtos",
        default=str(DEFAULT_REGRAS_PRODUTOS),
        help="Arquivo YAML de regras de produtos consultado em modo leitura.",
    )
    gerar_proposta.set_defaults(func=gerar_proposta_classificacao_command)

    gerar_proposta_real = subparsers.add_parser(
        "gerar-proposta-regras-classificacao-real",
        help="Gera proposta YAML a partir da revisao real validada.",
        description=(
            "Le um ou mais CSV revisados de classificacao e gera uma proposta "
            "YAML separada para revisao humana. Nao altera CSV, Excel ou "
            "arquivos YAML de config."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    gerar_proposta_real.add_argument(
        "--csv",
        action="append",
        required=True,
        help="Caminho do CSV revisado gerado a partir da planilha assistida.",
    )
    gerar_proposta_real.add_argument(
        "--saida",
        required=True,
        help="Arquivo YAML separado para salvar a proposta real.",
    )
    gerar_proposta_real.add_argument(
        "--relatorio",
        help="Arquivo TXT opcional para salvar o resumo da proposta real.",
    )
    gerar_proposta_real.add_argument(
        "--plano-contas",
        default=str(DEFAULT_PLANO_CONTAS),
        help="Arquivo YAML do plano de contas gerencial para validacao.",
    )
    gerar_proposta_real.set_defaults(
        func=gerar_proposta_regras_classificacao_real_command
    )

    auditar_operacoes = subparsers.add_parser(
        "auditar-operacoes-especiais",
        help="Audita NF-e de comodato, bonificacao e outras operacoes especiais.",
        description=(
            "Le XML NF-e em modo somente leitura e gera relatorio TXT de "
            "operacoes especiais antes da classificacao de produtos. Nao altera "
            "XML, Excel, CSV, YAML ou DRE."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    auditar_operacoes.add_argument(
        "--pasta", required=True, help="Pasta com arquivos XML NF-e de entrada."
    )
    auditar_operacoes.add_argument(
        "--saida", required=True, help="Arquivo TXT de saida da auditoria."
    )
    auditar_operacoes.set_defaults(func=auditar_operacoes_especiais_command)

    gerar_planilha_revisao = subparsers.add_parser(
        "gerar-planilha-revisao-classificacao",
        help="Gera XLSX assistido para revisao manual de classificacao.",
        description=(
            "Le o CSV de revisao de classificacao e gera uma planilha XLSX "
            "assistida para preenchimento manual. Nao altera CSV, Excel de "
            "lancamentos, YAML ou regras."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    gerar_planilha_revisao.add_argument(
        "--csv",
        required=True,
        help="Caminho do CSV de revisao de classificacao.",
    )
    gerar_planilha_revisao.add_argument(
        "--saida",
        required=True,
        help="Caminho do arquivo XLSX assistido de saida.",
    )
    gerar_planilha_revisao.add_argument(
        "--plano-contas",
        default=str(DEFAULT_PLANO_CONTAS),
        help="Arquivo YAML do plano de contas gerencial consultado em modo leitura.",
    )
    gerar_planilha_revisao.set_defaults(
        func=gerar_planilha_revisao_classificacao_command
    )

    importar_planilha_revisao = subparsers.add_parser(
        "importar-planilha-revisao-classificacao",
        help="Importa XLSX assistido e gera CSV revisado de classificacao.",
        description=(
            "Le a planilha assistida de revisao de classificacao, valida os "
            "campos preenchidos manualmente e gera um CSV revisado. Nao altera "
            "XLSX, CSV original, Excel de lancamentos, YAML ou regras."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    importar_planilha_revisao.add_argument(
        "--xlsx",
        required=True,
        help="Caminho da planilha XLSX assistida de revisao.",
    )
    importar_planilha_revisao.add_argument(
        "--saida",
        required=True,
        help="Caminho do CSV revisado de saida.",
    )
    importar_planilha_revisao.add_argument(
        "--relatorio",
        help="Arquivo TXT opcional para salvar o relatorio de validacao.",
    )
    importar_planilha_revisao.add_argument(
        "--plano-contas",
        default=str(DEFAULT_PLANO_CONTAS),
        help="Arquivo YAML do plano de contas gerencial consultado em modo leitura.",
    )
    importar_planilha_revisao.set_defaults(
        func=importar_planilha_revisao_classificacao_command
    )

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


def importar_nfe_command(args: argparse.Namespace) -> int:
    xml_path = Path(args.xml)
    excel_path = Path(args.excel)
    plano_contas_path = Path(args.plano_contas)
    regras_produtos_path = Path(args.regras_produtos)

    try:
        _validate_file_exists(xml_path, "Arquivo XML")
        _validate_file_exists(plano_contas_path, "Plano de contas")
        _validate_file_exists(regras_produtos_path, "Regras de produtos")

        info(f"Importando NF-e: {xml_path}")
        nfe = parse_nfe(xml_path)

        classifier = _load_product_classifier(plano_contas_path, regras_produtos_path)
        lancamentos = allocate_nfe_to_lancamentos(nfe, classifier)

        info(f"Lancamentos gerados: {len(lancamentos)}")
        revisao_count = sum(
            1 for lancamento in lancamentos if lancamento.status_revisao != "OK"
        )
        if revisao_count:
            warning(f"Lancamentos para revisao: {revisao_count}")

        write_lancamentos_to_excel(excel_path, lancamentos)
        info(f"Excel atualizado: {excel_path}")
        info("Concluido com sucesso.")
        return 0
    except Exception as exc:
        error(str(exc))
        return 1


def importar_pasta_nfe_command(args: argparse.Namespace) -> int:
    pasta_path = Path(args.pasta)
    excel_path = Path(args.excel)
    plano_contas_path = Path(args.plano_contas)
    regras_produtos_path = Path(args.regras_produtos)

    try:
        _validate_directory_exists(pasta_path, "Pasta")
        _validate_file_exists(plano_contas_path, "Plano de contas")
        _validate_file_exists(regras_produtos_path, "Regras de produtos")

        xml_files = _find_xml_files(pasta_path)
        if not xml_files:
            error(f"Nenhum XML encontrado na pasta: {pasta_path}")
            return 1

        info(f"Importando pasta NF-e: {pasta_path}")
        info(f"XML encontrados: {len(xml_files)}")

        classifier = _load_product_classifier(plano_contas_path, regras_produtos_path)
        lancamentos = []
        imported_count = 0
        ignored_count = 0

        for xml_path in xml_files:
            info(f"Importando NF-e: {xml_path}")
            try:
                nfe = parse_nfe(xml_path)
                file_lancamentos = allocate_nfe_to_lancamentos(nfe, classifier)
                lancamentos.extend(file_lancamentos)
                imported_count += 1
            except Exception as exc:
                ignored_count += 1
                warning(f"XML ignorado: {xml_path} - {exc}")

        info(f"XML importados com sucesso: {imported_count}")
        if ignored_count:
            warning(f"XML ignorados/erro: {ignored_count}")

        if imported_count == 0:
            error("Nenhum XML valido importado.")
            return 1

        info(f"Lancamentos gerados: {len(lancamentos)}")
        revisao_count = sum(
            1 for lancamento in lancamentos if lancamento.status_revisao != "OK"
        )
        if revisao_count:
            warning(f"Lancamentos para revisao: {revisao_count}")

        write_lancamentos_to_excel(excel_path, lancamentos)
        info(f"Excel atualizado: {excel_path}")
        info("Concluido com sucesso.")
        return 0
    except Exception as exc:
        error(str(exc))
        return 1


def diagnosticar_excel_command(args: argparse.Namespace) -> int:
    report_path = Path(args.relatorio) if args.relatorio else None

    try:
        result = diagnose_excel(Path(args.excel))
    except ExcelDiagnosticError as exc:
        print(f"[ERRO] {exc}")
        if report_path and not _try_save_report(
            report_path,
            report_text(ExcelDiagnosticFailure(Path(args.excel), str(exc))),
        ):
            return 1
        if report_path:
            print(f"[OK] Relatorio salvo em: {report_path}")
        return 1

    for line in console_lines(result):
        print(line)

    if report_path:
        if not _try_save_report(report_path, report_text(result)):
            return 1
        print(f"[OK] Relatorio salvo em: {report_path}")

    return 0 if result.is_ok else 1


def relatorio_revisao_command(args: argparse.Namespace) -> int:
    report_path = Path(args.saida) if args.saida else None

    try:
        result, text = generate_revision_report(Path(args.excel))
    except RevisionReportError as exc:
        print(f"[ERRO] {exc}")
        return 1

    if report_path:
        for line in revision_console_lines(result):
            print(line)
        if not _try_save_revision_report(report_path, text):
            return 1
        print(f"[OK] Relatorio salvo em: {report_path}")
    else:
        print(text, end="")

    return 0


def relatorio_candidatos_command(args: argparse.Namespace) -> int:
    report_path = Path(args.saida) if args.saida else None

    try:
        result, text = generate_classification_candidates_report(Path(args.excel))
    except ClassificationCandidatesError as exc:
        print(f"[ERRO] {exc}")
        return 1

    if report_path:
        for line in candidates_console_lines(result):
            print(line)
        if not _try_save_candidates_report(report_path, text):
            return 1
        print(f"[OK] Relatorio salvo em: {report_path}")
    else:
        print(text, end="")

    return 0


def exportar_revisao_classificacao_command(args: argparse.Namespace) -> int:
    try:
        result = export_classification_review_csv(Path(args.excel), Path(args.saida))
    except ClassificationReviewExportError as exc:
        print(f"[ERRO] {exc}")
        return 1

    for line in review_export_console_lines(result):
        print(line)
    return 0


def validar_revisao_classificacao_command(args: argparse.Namespace) -> int:
    report_path = Path(args.relatorio) if args.relatorio else None

    try:
        result = validate_classification_review_csv(Path(args.csv))
    except ClassificationReviewValidationError as exc:
        print(f"[ERRO] {exc}")
        return 1

    text = validation_report_text(result)
    print(text, end="")

    if report_path:
        if not _try_save_validation_report(report_path, text):
            return 1
        print(f"[OK] Relatorio de validacao salvo em: {report_path}")

    return 0


def gerar_proposta_classificacao_command(args: argparse.Namespace) -> int:
    yaml_path = Path(args.saida_yaml) if args.saida_yaml else None
    report_path = Path(args.relatorio) if args.relatorio else None

    if not yaml_path and not report_path:
        print("[ERRO] Informe --saida-yaml e/ou --relatorio para gerar a proposta.")
        return 1

    try:
        result = gerar_proposta_classificacao(
            Path(args.csv),
            Path(args.plano_contas),
            Path(args.regras_produtos),
        )
    except ClassificationRuleProposalError as exc:
        print(f"[ERRO] {exc}")
        return 1

    if yaml_path:
        try:
            escrever_proposta_yaml(result, yaml_path)
        except Exception:
            print(f"[ERRO] Nao foi possivel salvar proposta YAML: {yaml_path}")
            return 1

    text = gerar_relatorio_proposta_classificacao(
        result,
        Path(args.csv),
        yaml_path,
    )
    if report_path:
        if not _try_save_proposal_report(report_path, text):
            return 1
    elif not yaml_path:
        print(text, end="")

    for line in proposal_console_lines(
        result, saida_yaml=yaml_path, relatorio=report_path
    ):
        print(line)

    return 0


def gerar_proposta_regras_classificacao_real_command(args: argparse.Namespace) -> int:
    csv_paths = [Path(csv_path) for csv_path in args.csv]
    yaml_path = Path(args.saida)
    report_path = Path(args.relatorio) if args.relatorio else None

    try:
        result = gerar_proposta_regras_classificacao_real(
            csv_paths,
            yaml_path,
            Path(args.plano_contas),
            report_path,
        )
    except ClassificationRuleProposalError as exc:
        print(f"[ERRO] {exc}")
        return 1
    except Exception:
        print(f"[ERRO] Nao foi possivel gerar proposta YAML real: {yaml_path}")
        return 1

    for line in proposal_real_console_lines(result):
        print(line)
    return 0


def auditar_operacoes_especiais_command(args: argparse.Namespace) -> int:
    try:
        result = auditar_operacoes_especiais(Path(args.pasta))
    except OperationAuditError as exc:
        print(f"[ERRO] {exc}")
        return 1

    text = relatorio_operacoes_especiais(result)
    try:
        salvar_relatorio_operacoes(Path(args.saida), text)
    except Exception:
        print(f"[ERRO] Nao foi possivel salvar relatorio: {args.saida}")
        return 1

    for line in operation_audit_console_lines(result):
        print(line)
    print(f"[OK] Relatorio salvo em: {args.saida}")
    return 0


def gerar_planilha_revisao_classificacao_command(args: argparse.Namespace) -> int:
    try:
        output_path = gerar_planilha_revisao_classificacao(
            Path(args.csv),
            Path(args.saida),
            Path(args.plano_contas),
        )
    except ClassificationReviewWorkbookError as exc:
        print(f"[ERRO] {exc}")
        return 1
    except Exception:
        print(f"[ERRO] Nao foi possivel gerar planilha assistida: {args.saida}")
        return 1

    print(f"[OK] Planilha assistida gerada: {output_path}")
    return 0


def importar_planilha_revisao_classificacao_command(args: argparse.Namespace) -> int:
    relatorio_path = Path(args.relatorio) if args.relatorio else None

    try:
        result = importar_planilha_revisao_classificacao(
            Path(args.xlsx),
            Path(args.saida),
            Path(args.plano_contas),
            relatorio_path,
        )
    except ClassificationReviewWorkbookImportError as exc:
        if exc.result:
            for line in workbook_import_console_lines(exc.result):
                print(line)
        else:
            print(f"[ERRO] {exc}")
        return 1
    except Exception:
        print(f"[ERRO] Nao foi possivel importar planilha assistida: {args.xlsx}")
        return 1

    for line in workbook_import_console_lines(result):
        print(line)
    if relatorio_path:
        print(f"[OK] Relatorio salvo em: {relatorio_path}")
    return 0


def _add_config_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--plano-contas",
        default=str(DEFAULT_PLANO_CONTAS),
        help="Arquivo YAML do plano de contas gerencial.",
    )
    parser.add_argument(
        "--regras-produtos",
        default=str(DEFAULT_REGRAS_PRODUTOS),
        help="Arquivo YAML com regras de classificacao de produtos.",
    )


def _load_product_classifier(
    plano_contas_path: Path, regras_produtos_path: Path
) -> ProductClassifier:
    chart = ChartOfAccounts.load(plano_contas_path)
    validation_errors = chart.validate()
    if validation_errors:
        raise ValueError("; ".join(validation_errors))
    return ProductClassifier(chart).load(regras_produtos_path)


def _find_xml_files(pasta_path: Path) -> list[Path]:
    return sorted(
        (
            path
            for path in pasta_path.iterdir()
            if path.is_file() and path.suffix.lower() == ".xml"
        ),
        key=lambda path: path.name.lower(),
    )


def _validate_file_exists(path: Path, label: str) -> None:
    if not path.exists():
        raise FileNotFoundError(f"{label} nao encontrado: {path}")
    if not path.is_file():
        raise ValueError(f"{label} nao e um arquivo: {path}")


def _validate_directory_exists(path: Path, label: str) -> None:
    if not path.exists():
        raise FileNotFoundError(f"{label} nao encontrada: {path}")
    if not path.is_dir():
        raise ValueError(f"{label} nao e uma pasta: {path}")


def _try_save_report(report_path: Path, text: str) -> bool:
    try:
        save_report(report_path, text)
    except Exception:
        print(f"[ERRO] Nao foi possivel salvar relatorio: {report_path}")
        return False
    return True


def _try_save_revision_report(report_path: Path, text: str) -> bool:
    try:
        save_revision_report(report_path, text)
    except Exception:
        print(f"[ERRO] Nao foi possivel salvar relatorio: {report_path}")
        return False
    return True


def _try_save_candidates_report(report_path: Path, text: str) -> bool:
    try:
        save_candidates_report(report_path, text)
    except Exception:
        print(f"[ERRO] Nao foi possivel salvar relatorio: {report_path}")
        return False
    return True


def _try_save_validation_report(report_path: Path, text: str) -> bool:
    try:
        save_validation_report(report_path, text)
    except Exception:
        print(f"[ERRO] Nao foi possivel salvar relatorio: {report_path}")
        return False
    return True


def _try_save_proposal_report(report_path: Path, text: str) -> bool:
    try:
        salvar_relatorio_proposta(report_path, text)
    except Exception:
        print(f"[ERRO] Nao foi possivel salvar relatorio: {report_path}")
        return False
    return True


if __name__ == "__main__":
    raise SystemExit(main())
