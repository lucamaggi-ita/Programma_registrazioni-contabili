from __future__ import annotations

import argparse
import csv
import json
from datetime import date
from pathlib import Path

from fantozzi.allocation import allocate_nfe_to_lancamentos
from fantozzi.auditoria_relatorios_gerenciais_ok import (
    AuditoriaRelatoriosGerenciaisOkError,
    auditar_relatorios_gerenciais_ok_provisorios,
    console_lines as auditoria_relatorios_ok_console_lines,
)
from fantozzi.banking.normalizer import import_bank_files
from fantozzi.banking.payment_classification_run import (
    PaymentClassificationRunError,
    execute_payment_classification_run,
)
from fantozzi.banking.payment_review_package import (
    PaymentReviewPackageError,
    console_lines as comprovantes_review_console_lines,
    gerar_pacote_revisao_comprovantes,
)
from fantozzi.banking.payment_rules import (
    PaymentRulesError,
    audit_payment_rules_file,
    classify_payment_receipts_file,
    payment_rules_audit_console_lines,
    payment_rules_audit_report_text,
)
from fantozzi.banking.lancamentos_banca import (
    LancamentosBancaError,
    MANIFEST_NAME,
    REJEITADOS_CSV_NAME,
    gerar_lancamentos_banca_aprovados,
)
from fantozzi.banking.revisao_banca import (
    ImportaRevisaoBancaError,
    console_lines as revisao_banca_console_lines,
    importar_revisao_banca,
    resolve_bank_raw_dir,
)
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
from fantozzi.classification_review_package import (
    ClassificationReviewPackageError,
    console_lines as review_package_console_lines,
    gerar_pacote_revisao_contabil,
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
from fantozzi.csv_utils import sniff_delimiter
from fantozzi.dre_gerencial_provisoria import (
    DEFAULT_MAPEAMENTO_DRE,
    DreGerencialProvisoriaError,
    console_lines as dre_provisoria_console_lines,
    gerar_dre_gerencial_ok_provisoria,
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
from fantozzi.fase9_decision_engine import (
    evaluate_fase9_decisions,
    load_fase9_decisions_from_csv,
)
from fantozzi.fase9_writeback import apply_fase9_writeback, write_fase9_applied_csv
from fantozzi.importa_revisao_xml import (
    ImportaRevisaoXmlError,
    console_lines as revisao_xml_console_lines,
    importar_revisao_xml,
    resolve_run_reports_dir,
)
from fantozzi.lancamentos_manuais_template import (
    LancamentosManuaisTemplateError,
    gerar_template_lancamentos_manuais_gerenciais,
    validar_lancamentos_manuais_gerenciais_operador,
)
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
from fantozzi.pre_livro_diario import (
    PreLivroDiarioError,
    console_lines as pre_livro_console_lines,
    gerar_pre_livro_diario_teste,
)
from fantozzi.revision_report import (
    RevisionReportError,
    console_lines as revision_console_lines,
    generate_revision_report,
    save_report as save_revision_report,
)
from fantozzi.relatorios_gerenciais_ok import (
    RelatoriosGerenciaisOkError,
    console_lines as relatorios_ok_console_lines,
    gerar_relatorios_gerenciais_ok_provisorios,
)
from fantozzi import reprocessamento_pipeline
from fantozzi.run_manager import RunManagerError, preparar_run


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PLANO_CONTAS = PROJECT_ROOT / "config" / "plano_contas_gerencial.yaml"
DEFAULT_REGRAS_PRODUTOS = (
    PROJECT_ROOT / "config" / "regras_classificacao_produtos.yaml"
)
DEFAULT_REGRAS_COMPROVANTES = (
    PROJECT_ROOT / "config" / "regras_classificacao_comprovantes.yaml"
)
DEFAULT_MAPEAMENTO_DRE_PATH = PROJECT_ROOT / "config" / "mapeamento_dre.yaml"
DEFAULT_TEMPLATE_MANUAIS_PATH = (
    PROJECT_ROOT / "Database_Dev" / "Registrazioni_manuali" / "registrazioni_manuali_base.xlsx"
)
DEFAULT_REGISTRAZIONI_MANUALI_BASE = (
    PROJECT_ROOT / "Database_Dev" / "Registrazioni_manuali" / "registrazioni_manuali_base.xlsx"
)
DEFAULT_TEMPLATE_MANUAIS_SAIDA_CSV = (
    PROJECT_ROOT / "Database_Dev" / "Registrazioni_manuali" / "registrazioni_manuali_tecnico.csv"
)
DEFAULT_XML_BASE = PROJECT_ROOT / "Database_Dev" / "XML_fatture_acquisto"
DEFAULT_OUTPUT_IMPORTACAO_REAL = PROJECT_ROOT / "output" / "importacao_real"
DEFAULT_REPORTS_IMPORTACAO_REAL = PROJECT_ROOT / "reports" / "importacao_real"
DEFAULT_BANK_DECISOES_DIR = PROJECT_ROOT / "bank_raw" / "_decisoes_regras"
DEFAULT_XML_DECISOES_DIR = PROJECT_ROOT / "Database_Dev" / "_decisoes_regras" / "xml"
DEFAULT_REPORTS_RUNS_ROOT = PROJECT_ROOT / "reports" / "importacao_real" / "runs"
DEFAULT_COMPROVANTES_REAL_INPUT = PROJECT_ROOT / "bank_raw" / "_audit" / "fase_13c"
DEFAULT_COMPROVANTES_REAL_REPORTS = PROJECT_ROOT / "reports" / "comprovantes_classificacao"


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
            "--relatorio reports\\resumo_proposta_regras_classificacao_real.txt\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli "
            "gerar-pre-livro-diario-teste "
            "--excel output\\lancamentos.xlsx "
            "--pasta-operacoes-especiais XML_fatture_acquisto\\xml "
            "--saida reports\\pre_livro_diario_teste.xlsx\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli "
            "gerar-pacote-revisao-contabil "
            "--csv reports\\revisao_classificacao_revisada.csv "
            "--saida-xlsx reports\\pacote_revisao_contabil.xlsx "
            "--relatorio reports\\pacote_revisao_contabil.txt\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli "
            "gerar-relatorios-gerenciais-ok-provisorios "
            "--entrada output\\importacao_real "
            "--saida-xlsx reports\\relatorios_gerenciais_ok_provisorios.xlsx "
            "--relatorio reports\\relatorios_gerenciais_ok_provisorios.txt\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli "
            "auditar-relatorios-gerenciais-ok-provisorios "
            "--entrada output\\importacao_real "
            "--saida-xlsx reports\\auditoria_relatorios_gerenciais_ok.xlsx "
            "--relatorio reports\\auditoria_relatorios_gerenciais_ok.txt\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli "
            "gerar-dre-gerencial-ok-provisoria "
            "--entrada output\\importacao_real "
            "--saida-xlsx reports\\dre_gerencial_ok_provisoria.xlsx "
            "--relatorio reports\\dre_gerencial_ok_provisoria.txt\n"
            "  .\\.venv\\Scripts\\python.exe -m fantozzi.cli "
            "aplicar-fase9-revisao "
            "--csv reports\\revisao_classificacao_revisada.csv "
            "--saida output\\lancamentos_aplicados_fase9.csv --dry-run"
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

    importar_extratos = subparsers.add_parser(
        "importar-extratos-bancarios",
        help="Importa extratos bancarios multi-fonte normalizados por CNPJ.",
        description=(
            "Normaliza OFX bancario, enriquecimento XLSX de lancamentos, "
            "comprovantes ConsultaPagamentos .xls e auditoria de saldos. "
            "Nao executa matching ou classificacao contabil."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    importar_extratos.add_argument("--cnpj", required=True, help="CNPJ da conta.")
    importar_extratos.add_argument(
        "--ofx",
        required=True,
        help="OFX primario com movimentos bancarios reais.",
    )
    importar_extratos.add_argument(
        "--xlsx-lancamentos",
        help="XLSX de lancamentos bancarios para enriquecimento/auditoria.",
    )
    importar_extratos.add_argument(
        "--consulta-pagamentos",
        help="Arquivo ConsultaPagamentos .xls com comprovantes.",
    )
    importar_extratos.add_argument(
        "--ofx-saldos",
        help="OFX mensal com linhas de saldo para auditoria.",
    )
    importar_extratos.add_argument(
        "--saida",
        default="output/banking_import",
        help="Pasta raiz de saida dos CSV/manifest normalizados.",
    )
    importar_extratos.set_defaults(func=importar_extratos_bancarios_command)

    importar_revisao = subparsers.add_parser(
        "importar-revisao-banca",
        help="Gera/atualiza Excel decisional para regras banca/comprovantes.",
        description=(
            "Le bank_raw/YYYY/YYYY-MM/CNPJ com ConsultaPagamentos .xls e "
            "Extrato_Lancamentos .xlsx, aplica regras atuais somente em memoria "
            "e gera planilhas decisorias. OFX e opcional e nao e usado."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    importar_revisao.add_argument(
        "--bank-raw-dir",
        help="Pasta do mes, por exemplo bank_raw/2026/2026-04.",
    )
    importar_revisao.add_argument(
        "--ano",
        help="Ano usado com --mes para construir bank_raw/<ano>/<YYYY-MM>.",
    )
    importar_revisao.add_argument(
        "--mes",
        help="Mes MM ou YYYY-MM usado com --ano.",
    )
    importar_revisao.add_argument(
        "--saida-dir",
        default=str(DEFAULT_BANK_DECISOES_DIR),
        help="Pasta dos Excel decisorios e manifest.",
    )
    importar_revisao.add_argument(
        "--regras",
        default=str(DEFAULT_REGRAS_COMPROVANTES),
        help="YAML com regras declaradas de comprovantes.",
    )
    importar_revisao.add_argument(
        "--plano-contas",
        default=str(DEFAULT_PLANO_CONTAS),
        help="Arquivo YAML do plano de contas gerencial.",
    )
    importar_revisao.add_argument(
        "--mapeamento-dre",
        default=str(DEFAULT_MAPEAMENTO_DRE_PATH),
        help="Arquivo YAML de mapeamento das contas para linhas DRE.",
    )
    importar_revisao.set_defaults(func=importar_revisao_banca_command)

    gerar_lancamentos_banca = subparsers.add_parser(
        "gerar-lancamentos-banca-aprovados",
        help="Converte decisoes bancarias APROVADO em Excel de lancamentos.",
        description=(
            "Le os Excel decisorios de comprovantes e conta corrente, inclui "
            "somente linhas APROVADO/CLASSIFICAR seguras no periodo informado "
            "e gera um .xlsx compativel com Lancamentos."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    gerar_lancamentos_banca.add_argument(
        "--comprovantes-xlsx",
        default=str(DEFAULT_BANK_DECISOES_DIR / "Analisi_regole_registrazioni_banca_comprovanti.xlsx"),
        help="Excel decisorio dos comprovantes.",
    )
    gerar_lancamentos_banca.add_argument(
        "--conta-corrente-xlsx",
        default=str(DEFAULT_BANK_DECISOES_DIR / "Analisi_regole_registrazioni_banca_conto_corrente.xlsx"),
        help="Excel decisorio da conta corrente.",
    )
    gerar_lancamentos_banca.add_argument(
        "--saida",
        required=True,
        help="Excel de lancamentos banca aprovados a gerar.",
    )
    gerar_lancamentos_banca.add_argument(
        "--periodo-ini",
        required=True,
        help="Data inicial inclusiva no formato YYYY-MM-DD.",
    )
    gerar_lancamentos_banca.add_argument(
        "--periodo-fim",
        required=True,
        help="Data final inclusiva no formato YYYY-MM-DD.",
    )
    gerar_lancamentos_banca.add_argument(
        "--comprovantes-csv",
        default=None,
        help=(
            "CSV de comprovantes ja classificados via YAML "
            "(reports/comprovantes_classificacao/runs/<RUN>/comprovantes_classificados.csv). "
            "Opcional."
        ),
    )
    gerar_lancamentos_banca.add_argument(
        "--conta-corrente-csv",
        default=None,
        help=(
            "CSV de conta corrente ja classificada via YAML "
            "(bank_raw/_decisoes_regras/conta_corrente_classificados.csv). "
            "Opcional."
        ),
    )
    gerar_lancamentos_banca.set_defaults(
        func=gerar_lancamentos_banca_aprovados_command
    )

    importar_revisao_xml_parser = subparsers.add_parser(
        "importar-revisao-xml",
        help="Gera/atualiza Excel decisional para regras XML/produtos.",
        description=(
            "Le reports/importacao_real/runs/<RUN_ID>/<CNPJ>/"
            "revisao_classificacao_real.csv, confere o hash do YAML de produtos "
            "usado na run e gera a planilha decisoria XML sem modificar YAML."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    importar_revisao_xml_parser.add_argument(
        "--run-reports-dir",
        help="Pasta da run em reports/importacao_real/runs/<RUN_ID>.",
    )
    importar_revisao_xml_parser.add_argument(
        "--run-id",
        help="Run id a resolver dentro de --reports-runs-root.",
    )
    importar_revisao_xml_parser.add_argument(
        "--reports-runs-root",
        default=str(DEFAULT_REPORTS_RUNS_ROOT),
        help="Raiz das runs de reports para resolver --run-id.",
    )
    importar_revisao_xml_parser.add_argument(
        "--saida-dir",
        default=str(DEFAULT_XML_DECISOES_DIR),
        help="Pasta do Excel decisorio XML e manifest.",
    )
    importar_revisao_xml_parser.add_argument(
        "--regras",
        default=str(DEFAULT_REGRAS_PRODUTOS),
        help="YAML atual de regras de produtos, usado somente para hash.",
    )
    importar_revisao_xml_parser.add_argument(
        "--plano-contas",
        default=str(DEFAULT_PLANO_CONTAS),
        help="Arquivo YAML do plano de contas gerencial.",
    )
    importar_revisao_xml_parser.add_argument(
        "--allow-stale-run",
        action="store_true",
        help="Permite gerar workbook de auditoria mesmo se o hash da run estiver obsoleto.",
    )
    importar_revisao_xml_parser.set_defaults(func=importar_revisao_xml_command)

    classificar_comprovantes = subparsers.add_parser(
        "classificar-comprovantes-pagamento",
        help="Classifica comprovantes por regras declaradas pelo operador.",
        description=(
            "Le payment_receipts_normalized.csv e aplica somente regras "
            "explicitas aprovadas pelo operador. Sem regra ou conflito vai "
            "para revisao; nao altera DRE, Fase 9 ou lancamentos definitivos."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    classificar_comprovantes.add_argument(
        "--comprovantes",
        required=True,
        help="CSV payment_receipts_normalized.csv gerado pela Fase 13A.",
    )
    classificar_comprovantes.add_argument(
        "--regras",
        default=str(DEFAULT_REGRAS_COMPROVANTES),
        help="YAML com regras declaradas de comprovantes.",
    )
    classificar_comprovantes.add_argument(
        "--plano-contas",
        default=str(DEFAULT_PLANO_CONTAS),
        help="Arquivo YAML do plano de contas gerencial.",
    )
    classificar_comprovantes.add_argument(
        "--mapeamento-dre",
        default=str(DEFAULT_MAPEAMENTO_DRE_PATH),
        help="Arquivo YAML de mapeamento das contas para linhas DRE.",
    )
    classificar_comprovantes.add_argument(
        "--saida",
        required=True,
        help="Pasta de saida dos CSV/manifest de classificacao.",
    )
    classificar_comprovantes.set_defaults(
        func=classificar_comprovantes_pagamento_command
    )

    auditar_regras_comprovantes = subparsers.add_parser(
        "auditar-regras-comprovantes",
        help="Audita YAML de regras de comprovantes sem alterar arquivos.",
        description=(
            "Le regras_classificacao_comprovantes.yaml em modo somente leitura "
            "e aponta regras invalidas, campos herdados de planilha, IDs baseados "
            "em linha, conflitos de criterio e condicoes potencialmente restritivas."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    auditar_regras_comprovantes.add_argument(
        "--regras",
        default=str(DEFAULT_REGRAS_COMPROVANTES),
        help="YAML com regras declaradas de comprovantes.",
    )
    auditar_regras_comprovantes.add_argument(
        "--relatorio",
        help="Arquivo TXT opcional para salvar o relatorio da auditoria.",
    )
    auditar_regras_comprovantes.set_defaults(
        func=auditar_regras_comprovantes_command
    )

    classificar_comprovantes_real = subparsers.add_parser(
        "executar-classificacao-comprovantes-real",
        help="Executa Fase 13S para comprovantes reais normalizados.",
        description=(
            "Descobre payment_receipts_normalized.csv por mes/CNPJ e classifica "
            "somente por regras declaradas. Nao executa matching bancario, nao "
            "infere contas e nao altera DRE, Fase 9 ou lancamentos definitivos."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    classificar_comprovantes_real.add_argument(
        "--entrada-raiz",
        default=str(DEFAULT_COMPROVANTES_REAL_INPUT),
        help="Pasta raiz onde procurar payment_receipts_normalized.csv.",
    )
    classificar_comprovantes_real.add_argument(
        "--saida-raiz",
        default=str(DEFAULT_COMPROVANTES_REAL_REPORTS),
        help="Pasta raiz dos reports da Fase 13S; a run sera criada em runs/<run_id>.",
    )
    classificar_comprovantes_real.add_argument(
        "--regras",
        default=str(DEFAULT_REGRAS_COMPROVANTES),
        help="YAML com regras declaradas de comprovantes.",
    )
    classificar_comprovantes_real.add_argument(
        "--plano-contas",
        default=str(DEFAULT_PLANO_CONTAS),
        help="Arquivo YAML do plano de contas gerencial.",
    )
    classificar_comprovantes_real.add_argument(
        "--mapeamento-dre",
        default=str(DEFAULT_MAPEAMENTO_DRE_PATH),
        help="Arquivo YAML de mapeamento das contas para linhas DRE.",
    )
    classificar_comprovantes_real.add_argument(
        "--run-id",
        help="Run id opcional. Se omitido, usa YYYY-MM-DD_NNN.",
    )
    classificar_comprovantes_real.set_defaults(
        func=executar_classificacao_comprovantes_real_command
    )

    pacote_revisao_comprovantes = subparsers.add_parser(
        "gerar-pacote-revisao-comprovantes",
        help="Gera XLSX/TXT 13T para revisao manual dos comprovantes.",
        description=(
            "Le comprovantes_a_revisar.csv de uma run 13S e gera pacote manual "
            "de revisao. SISPAG e marcado como FORA_ESCOPO_SISPAG e separado "
            "da aba operacional. Nao cria regras YAML e nao executa matching."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    pacote_revisao_comprovantes.add_argument(
        "--run-dir",
        required=True,
        help="Diretorio da run 13S com comprovantes_a_revisar.csv.",
    )
    pacote_revisao_comprovantes.add_argument(
        "--csv-revisao",
        help="CSV alternativo de comprovantes a revisar.",
    )
    pacote_revisao_comprovantes.add_argument(
        "--saida-xlsx",
        help="Caminho do XLSX de saida. Padrao: dentro do run-dir.",
    )
    pacote_revisao_comprovantes.add_argument(
        "--relatorio",
        help="Caminho do TXT de acompanhamento. Padrao: dentro do run-dir.",
    )
    pacote_revisao_comprovantes.set_defaults(
        func=gerar_pacote_revisao_comprovantes_command
    )

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

    aplicar_fase9 = subparsers.add_parser(
        "aplicar-fase9-revisao",
        help="Aplica decisoes revisadas da Fase 9 em CSV separado.",
        description=(
            "Le um CSV revisionado, avalia as decisoes aprovadas da Fase 9 "
            "e gera somente um CSV separado. Nao altera o CSV original, "
            "config YAML, pipeline, manifest ou modo definitivo. Com --dry-run, "
            "calcula o resumo sem escrever o CSV de saida."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    aplicar_fase9.add_argument(
        "--csv",
        required=True,
        help="CSV revisionado de entrada. O arquivo nao sera modificado.",
    )
    aplicar_fase9.add_argument(
        "--saida",
        required=True,
        help="CSV separado de saida para os lancamentos aplicados da Fase 9.",
    )
    aplicar_fase9.add_argument(
        "--plano-contas",
        default=str(DEFAULT_PLANO_CONTAS),
        help="Plano de contas consultado em modo somente leitura.",
    )
    aplicar_fase9.add_argument(
        "--dry-run",
        action="store_true",
        default=False,
        help="Calcula promocoes e bloqueios sem escrever o CSV de saida.",
    )
    aplicar_fase9.set_defaults(func=aplicar_fase9_revisao_command)

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

    gerar_pre_livro = subparsers.add_parser(
        "gerar-pre-livro-diario-teste",
        help="Gera XLSX de controle para pre-livro diario gerencial de teste.",
        description=(
            "Le um ou mais Excel de lancamentos em modo somente leitura e gera "
            "um XLSX de controle com lancamentos OK, lancamentos em revisao, "
            "operacoes especiais, resumo por conta, resumo mensal e instrucoes. "
            "Nao aplica regras novas e nao altera config, CSV ou Excel original."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    gerar_pre_livro.add_argument(
        "--excel",
        action="append",
        required=True,
        help="Arquivo Excel de lancamentos gerado pelo Fantozzi.",
    )
    gerar_pre_livro.add_argument(
        "--saida",
        required=True,
        help="Arquivo XLSX de saida do pre-livro diario gerencial de teste.",
    )
    gerar_pre_livro.add_argument(
        "--pasta-operacoes-especiais",
        action="append",
        default=[],
        help=(
            "Pasta XML opcional para auditar comodato, bonificacao e outras "
            "operacoes especiais dentro do pre-livro de teste."
        ),
    )
    gerar_pre_livro.add_argument(
        "--relatorio-auditoria",
        help="Arquivo TXT opcional para salvar os indicadores da aba Auditoria.",
    )
    gerar_pre_livro.set_defaults(func=gerar_pre_livro_diario_teste_command)

    template_manuais = subparsers.add_parser(
        "gerar-template-lancamentos-manuais-gerenciais",
        help="Gera template XLSX para lancamentos manuais gerenciais.",
        description=(
            "Gera uma planilha operacional para despesas manuais que nao vieram "
            "de XML ou comprovantes automaticos. O comando consulta plano de "
            "contas e mapeamento DRE em modo leitura e sinaliza blockers quando "
            "nao ha conta tecnica neutra fora DRE para contrapartida."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    template_manuais.add_argument(
        "--saida",
        default=str(DEFAULT_REGISTRAZIONI_MANUALI_BASE),
        help="Arquivo XLSX de saida do template operacional.",
    )
    template_manuais.add_argument(
        "--plano-contas",
        default=str(DEFAULT_PLANO_CONTAS),
        help="Arquivo YAML do plano de contas gerencial consultado em modo leitura.",
    )
    template_manuais.add_argument(
        "--mapeamento-dre",
        default=str(DEFAULT_MAPEAMENTO_DRE_PATH),
        help="Arquivo YAML de mapeamento DRE consultado em modo leitura.",
    )
    template_manuais.set_defaults(
        func=gerar_template_lancamentos_manuais_gerenciais_command
    )

    validar_manuais = subparsers.add_parser(
        "validar-lancamentos-manuais-gerenciais",
        help="Valida template/CSV operacional e gera CSV tecnico para pipeline.",
        description=(
            "Le XLSX ou CSV operacional de lancamentos manuais gerenciais, "
            "resolve categoria e forma de pagamento para contas debitadas e "
            "creditadas, e grava CSV tecnico UTF-8-SIG separado por ponto e "
            "virgula. Retorna erro quando ha blockers."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    validar_manuais.add_argument(
        "--entrada",
        required=True,
        help="Arquivo XLSX/CSV operacional preenchido pelo operador.",
    )
    validar_manuais.add_argument(
        "--saida-csv",
        default=str(DEFAULT_TEMPLATE_MANUAIS_SAIDA_CSV),
        help="Arquivo CSV tecnico de saida para reprocessamento.",
    )
    validar_manuais.add_argument(
        "--relatorio",
        help="Arquivo TXT opcional para salvar o relatorio de validacao.",
    )
    validar_manuais.add_argument(
        "--manifest",
        default=None,
        help=(
            "Arquivo JSON de manifest de validacao. "
            "Default: mesma pasta de --saida-csv com nome "
            "registrazioni_manuali_validazione_manifest.json."
        ),
    )
    validar_manuais.add_argument(
        "--plano-contas",
        default=str(DEFAULT_PLANO_CONTAS),
        help="Arquivo YAML do plano de contas gerencial consultado em modo leitura.",
    )
    validar_manuais.add_argument(
        "--mapeamento-dre",
        default=str(DEFAULT_MAPEAMENTO_DRE_PATH),
        help="Arquivo YAML de mapeamento DRE consultado em modo leitura.",
    )
    validar_manuais.set_defaults(
        func=validar_lancamentos_manuais_gerenciais_command
    )

    executar_manuais = subparsers.add_parser(
        "executar-lancamentos-manuais-gerenciais",
        help="Valida template XLSX, gera CSV tecnico e executa run completa com pacote.",
        description=(
            "Skill operativa 6G: valida o template de lancamentos manuais preenchido "
            "pelo operador, gera o CSV tecnico e executa a pipeline provisoria com "
            "XML reais + lancamentos manuais + pacote contabil gerencial. "
            "Para se houver erros de validacao. Nao requer CNPJ, fornecedor, "
            "documento, centro_custo, forma_pagamento nem categoria gerencial antiga."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    executar_manuais.add_argument(
        "--template",
        default=str(DEFAULT_TEMPLATE_MANUAIS_PATH),
        help="Template XLSX preenchido pelo operador.",
    )
    executar_manuais.add_argument(
        "--saida-csv",
        default=str(DEFAULT_TEMPLATE_MANUAIS_SAIDA_CSV),
        help="CSV tecnico intermediario gerado pela validacao.",
    )
    executar_manuais.add_argument(
        "--xml-base",
        default=str(DEFAULT_XML_BASE),
        help="Pasta raiz com subpastas de XML NF-e por CNPJ.",
    )
    executar_manuais.add_argument(
        "--saida-output",
        default=str(DEFAULT_OUTPUT_IMPORTACAO_REAL),
        help="Pasta raiz de output onde sera criada runs/{run_id}/.",
    )
    executar_manuais.add_argument(
        "--saida-reports",
        default=str(DEFAULT_REPORTS_IMPORTACAO_REAL),
        help="Pasta raiz de reports onde sera criada runs/{run_id}/.",
    )
    executar_manuais.add_argument(
        "--plano-contas",
        default=str(DEFAULT_PLANO_CONTAS),
        help="Arquivo YAML do plano de contas gerencial.",
    )
    executar_manuais.add_argument(
        "--mapeamento-dre",
        default=str(DEFAULT_MAPEAMENTO_DRE_PATH),
        help="Arquivo YAML de mapeamento DRE.",
    )
    executar_manuais.add_argument(
        "--relatorio-validacao",
        default=None,
        help="Arquivo TXT opcional para salvar o relatorio de validacao.",
    )
    executar_manuais.set_defaults(
        func=executar_lancamentos_manuais_gerenciais_command
    )

    gerar_pacote_revisao = subparsers.add_parser(
        "gerar-pacote-revisao-contabil",
        help="Gera XLSX e TXT de apoio para revisao contabil humana.",
        description=(
            "Le um ou mais CSV de revisao de classificacao e gera um pacote "
            "assistido com prioridades, top valores, top ocorrencias, visao por "
            "fornecedor e abas para contador. Nao altera CSV, XLSX existentes, "
            "XML, Excel de lancamentos ou arquivos YAML de config."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    gerar_pacote_revisao.add_argument(
        "--csv",
        action="append",
        required=True,
        help="Caminho do CSV de revisao/revisado. Pode ser informado varias vezes.",
    )
    gerar_pacote_revisao.add_argument(
        "--saida-xlsx",
        required=True,
        help="Arquivo XLSX de saida do pacote assistido.",
    )
    gerar_pacote_revisao.add_argument(
        "--relatorio",
        help="Arquivo TXT opcional para salvar o resumo do pacote.",
    )
    gerar_pacote_revisao.add_argument(
        "--limite-top",
        type=int,
        default=50,
        help="Quantidade maxima de linhas nas abas Top_Valor e Top_Ocorrencias.",
    )
    gerar_pacote_revisao.add_argument(
        "--separador",
        default=";",
        help="Separador dos CSV de entrada.",
    )
    gerar_pacote_revisao.set_defaults(func=gerar_pacote_revisao_contabil_command)

    gerar_relatorios_ok = subparsers.add_parser(
        "gerar-relatorios-gerenciais-ok-provisorios",
        help="Gera relatorios gerenciais provisorios somente com lancamentos OK.",
        description=(
            "Le Excel de lancamentos ou diretorios com Excel, inclui somente "
            "status_revisao OK e exclui REVISAR e operacoes especiais auditadas. "
            "Gera XLSX e TXT provisorios, sem alterar dados de origem ou config."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    gerar_relatorios_ok.add_argument(
        "--entrada",
        action="append",
        required=True,
        help="Arquivo Excel .xlsx ou diretorio com Excel de lancamentos.",
    )
    gerar_relatorios_ok.add_argument(
        "--saida-xlsx",
        required=True,
        help="Arquivo XLSX de saida dos relatorios gerenciais OK provisorios.",
    )
    gerar_relatorios_ok.add_argument(
        "--relatorio",
        help="Arquivo TXT opcional para salvar o resumo dos relatorios.",
    )
    gerar_relatorios_ok.add_argument(
        "--pasta-operacoes-especiais",
        action="append",
        default=[],
        help=(
            "Pasta XML opcional para auditar e excluir comodato, bonificacao "
            "e outras operacoes especiais."
        ),
    )
    gerar_relatorios_ok.set_defaults(
        func=gerar_relatorios_gerenciais_ok_provisorios_command
    )

    auditar_relatorios_ok = subparsers.add_parser(
        "auditar-relatorios-gerenciais-ok-provisorios",
        help="Audita criterios e reconciliacoes dos relatorios gerenciais OK.",
        description=(
            "Le a mesma base dos relatorios gerenciais OK provisorios e verifica "
            "reconciliacoes entre OK, REVISAR e operacoes especiais. Gera XLSX "
            "e TXT de auditoria sem alterar dados de origem."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    auditar_relatorios_ok.add_argument(
        "--entrada",
        action="append",
        required=True,
        help="Arquivo Excel .xlsx ou diretorio com Excel de lancamentos.",
    )
    auditar_relatorios_ok.add_argument(
        "--saida-xlsx",
        required=True,
        help="Arquivo XLSX de saida da auditoria.",
    )
    auditar_relatorios_ok.add_argument(
        "--relatorio",
        help="Arquivo TXT opcional para salvar o resumo da auditoria.",
    )
    auditar_relatorios_ok.add_argument(
        "--pasta-operacoes-especiais",
        action="append",
        default=[],
        help=(
            "Pasta XML opcional para auditar comodato, bonificacao e outras "
            "operacoes especiais."
        ),
    )
    auditar_relatorios_ok.set_defaults(
        func=auditar_relatorios_gerenciais_ok_provisorios_command
    )

    gerar_dre_provisoria = subparsers.add_parser(
        "gerar-dre-gerencial-ok-provisoria",
        help="Gera DRE gerencial provisoria somente com lancamentos OK.",
        description=(
            "Le Excel de lancamentos ou diretorios com Excel, usa somente "
            "status_revisao OK nao especial e aplica mapeamento DRE em modo "
            "somente leitura. Gera XLSX e TXT provisorios."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    gerar_dre_provisoria.add_argument(
        "--entrada",
        action="append",
        required=True,
        help="Arquivo Excel .xlsx ou diretorio com Excel de lancamentos.",
    )
    gerar_dre_provisoria.add_argument(
        "--mapeamento-dre",
        default=str(DEFAULT_MAPEAMENTO_DRE),
        help="Arquivo YAML de mapeamento DRE consultado em modo leitura.",
    )
    gerar_dre_provisoria.add_argument(
        "--saida-xlsx",
        required=True,
        help="Arquivo XLSX de saida da DRE gerencial OK provisoria.",
    )
    gerar_dre_provisoria.add_argument(
        "--relatorio",
        help="Arquivo TXT opcional para salvar o resumo da DRE.",
    )
    gerar_dre_provisoria.add_argument(
        "--pasta-operacoes-especiais",
        action="append",
        default=[],
        help=(
            "Pasta XML opcional para auditar e excluir comodato, bonificacao "
            "e outras operacoes especiais."
        ),
    )
    gerar_dre_provisoria.set_defaults(func=gerar_dre_gerencial_ok_provisoria_command)

    preparar_run_cmd = subparsers.add_parser(
        "preparar-run-reprocessamento",
        help="Prepara infraestrutura segura para uma nova run de reprocessamento.",
        description=(
            "Cria as pastas da run, current_run.txt e manifest_run.json inicial. "
            "Nao executa importacao, DRE, pre-livro nem qualquer reprocessamento. "
            "Com --dry-run, calcula run_id e caminhos sem escrever nada."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    preparar_run_cmd.add_argument(
        "--xml-base",
        required=True,
        help="Pasta raiz com subpastas de XML NF-e por CNPJ.",
    )
    preparar_run_cmd.add_argument(
        "--saida-output",
        required=True,
        help="Pasta raiz de output onde sera criada runs/{run_id}/.",
    )
    preparar_run_cmd.add_argument(
        "--saida-reports",
        required=True,
        help="Pasta raiz de reports onde sera criada runs/{run_id}/ e manifest_run.json.",
    )
    preparar_run_cmd.add_argument(
        "--dry-run",
        action="store_true",
        default=False,
        help="Calcula run_id e caminhos sem criar pastas nem escrever arquivos.",
    )
    preparar_run_cmd.set_defaults(func=preparar_run_reprocessamento_command)

    reprocessar_real = subparsers.add_parser(
        "reprocessar-importacao-real",
        help="Executa pipeline provisoria completa de reprocessamento real.",
        description=(
            "Cria uma nova run segura, importa XML reais dentro da run e gera "
            "relatorios provisorios. Modo definitivo permanece bloqueado."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    reprocessar_real.add_argument(
        "--xml-base",
        required=True,
        help="Pasta raiz com subpastas de XML NF-e por CNPJ.",
    )
    reprocessar_real.add_argument(
        "--saida-output",
        required=True,
        help="Pasta raiz de output onde sera criada runs/{run_id}/.",
    )
    reprocessar_real.add_argument(
        "--saida-reports",
        required=True,
        help="Pasta raiz de reports onde sera criada runs/{run_id}/.",
    )
    reprocessar_real.add_argument(
        "--modo",
        choices=("provisorio", "definitivo"),
        default="provisorio",
        help="Modo de execucao da pipeline.",
    )
    reprocessar_real.add_argument(
        "--dry-run",
        action="store_true",
        default=False,
        help="Calcula run e etapas sem escrever arquivos nem importar XML.",
    )
    reprocessar_real.add_argument(
        "--fase9-csv",
        action="append",
        metavar="PATH",
        help=(
            "CSV revisionado para aplicar Fase 9 em output separado dentro da run. "
            "Pode ser repetido para processar varios CNPJs. "
            "Se omitido, a Fase 9 nao e executada."
        ),
    )
    reprocessar_real.add_argument(
        "--lancamentos-manuais-gerenciais-csv",
        metavar="PATH",
        help=(
            "CSV de lancamentos manuais gerenciais controlados para aplicar na run. "
            "No dry-run v1, o CSV nao e lido nem validado."
        ),
    )
    reprocessar_real.add_argument(
        "--lancamentos-banca-aprovados",
        metavar="PATH",
        help=(
            "Excel de lancamentos banca aprovados ja gerado pelo bridge. "
            "Quando informado, e copiado para a run e incluido no pre-livro/DRE."
        ),
    )
    reprocessar_real.add_argument(
        "--gerar-pacote-contabil-gerencial",
        action="store_true",
        default=False,
        help=(
            "Gera pacote contabil gerencial provisorio ao final da run, com indice, "
            "auditorias e manifest dedicado."
        ),
    )
    reprocessar_real.set_defaults(func=reprocessar_importacao_real_command)

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
        nfes = []

        for xml_path in xml_files:
            info(f"Importando NF-e: {xml_path}")
            try:
                nfe = parse_nfe(xml_path)
                nfes.append(nfe)
            except Exception as exc:
                ignored_count += 1
                warning(f"XML ignorado: {xml_path} - {exc}")

        nfe_index = {nfe.chave: nfe for nfe in nfes if nfe.chave}
        for nfe in nfes:
            try:
                file_lancamentos = allocate_nfe_to_lancamentos(
                    nfe, classifier, nfe_index=nfe_index
                )
                lancamentos.extend(file_lancamentos)
                imported_count += 1
            except Exception as exc:
                ignored_count += 1
                warning(f"XML ignorado: {nfe.arquivo_origem} - {exc}")

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


def importar_extratos_bancarios_command(args: argparse.Namespace) -> int:
    try:
        manifest = import_bank_files(
            cnpj=args.cnpj,
            ofx=Path(args.ofx),
            xlsx_lancamentos=(
                Path(args.xlsx_lancamentos) if args.xlsx_lancamentos else None
            ),
            consulta_pagamentos=(
                Path(args.consulta_pagamentos) if args.consulta_pagamentos else None
            ),
            ofx_saldos=Path(args.ofx_saldos) if args.ofx_saldos else None,
            output_root=Path(args.saida),
        )
    except Exception as exc:
        print(f"[ERRO] Nao foi possivel importar extratos bancarios: {exc}")
        return 1

    print("[OK] Importacao bancaria normalizada concluida.")
    print(f"CNPJ: {manifest.cnpj}")
    print(f"Transacoes: {manifest.transaction_count}")
    print(f"Comprovantes: {manifest.payment_receipt_count}")
    print(f"Auditorias de saldo: {manifest.balance_audit_count}")
    print(f"Saida: {manifest.output_dir}")
    print("[INFO] Matching contabil nao executado nesta fase.")
    return 0


def importar_revisao_banca_command(args: argparse.Namespace) -> int:
    try:
        bank_raw_dir = resolve_bank_raw_dir(
            bank_raw_dir=Path(args.bank_raw_dir) if args.bank_raw_dir else None,
            ano=args.ano,
            mes=args.mes,
            project_root=PROJECT_ROOT,
        )
        result = importar_revisao_banca(
            bank_raw_dir=bank_raw_dir,
            output_dir=Path(args.saida_dir),
            regras_path=Path(args.regras),
            plano_contas_path=Path(args.plano_contas),
            mapeamento_dre_path=Path(args.mapeamento_dre),
        )
    except (
        OSError,
        PaymentRulesError,
        ImportaRevisaoBancaError,
        ValueError,
    ) as exc:
        print(f"[ERRO] Nao foi possivel importar revisao banca: {exc}")
        return 1

    for line in revisao_banca_console_lines(result):
        print(line)
    return 0


def gerar_lancamentos_banca_aprovados_command(args: argparse.Namespace) -> int:
    try:
        periodo_ini = date.fromisoformat(args.periodo_ini)
        periodo_fim = date.fromisoformat(args.periodo_fim)
        output_path = gerar_lancamentos_banca_aprovados(
            Path(args.comprovantes_xlsx),
            Path(args.conta_corrente_xlsx),
            Path(args.saida),
            periodo_ini,
            periodo_fim,
            comprovantes_csv=(
                Path(args.comprovantes_csv) if args.comprovantes_csv else None
            ),
            conta_corrente_csv=(
                Path(args.conta_corrente_csv) if args.conta_corrente_csv else None
            ),
        )
    except (OSError, LancamentosBancaError, ValueError) as exc:
        print(f"[ERRO] Nao foi possivel gerar lancamentos banca aprovados: {exc}")
        return 1

    print("[OK] Lancamentos banca aprovados gerados.")
    print(f"Saida: {output_path}")
    print(f"Rejeitados: {output_path.with_name(REJEITADOS_CSV_NAME)}")
    print(f"Manifest: {output_path.with_name(MANIFEST_NAME)}")
    return 0


def importar_revisao_xml_command(args: argparse.Namespace) -> int:
    try:
        run_reports_dir = resolve_run_reports_dir(
            run_reports_dir=(
                Path(args.run_reports_dir) if args.run_reports_dir else None
            ),
            run_id=args.run_id,
            reports_runs_root=Path(args.reports_runs_root),
        )
        result = importar_revisao_xml(
            run_reports_dir=run_reports_dir,
            output_dir=Path(args.saida_dir),
            regras_path=Path(args.regras),
            plano_contas_path=Path(args.plano_contas),
            allow_stale_run=args.allow_stale_run,
        )
    except (OSError, ImportaRevisaoXmlError, ValueError) as exc:
        print(f"[ERRO] Nao foi possivel importar revisao XML: {exc}")
        return 1

    for line in revisao_xml_console_lines(result):
        print(line)
    return 0


def classificar_comprovantes_pagamento_command(args: argparse.Namespace) -> int:
    try:
        manifest = classify_payment_receipts_file(
            comprovantes_path=Path(args.comprovantes),
            regras_path=Path(args.regras),
            plano_contas_path=Path(args.plano_contas),
            mapeamento_dre_path=Path(args.mapeamento_dre),
            output_dir=Path(args.saida),
        )
    except (OSError, PaymentRulesError, ValueError) as exc:
        print(f"[ERRO] Nao foi possivel classificar comprovantes: {exc}")
        return 1

    print("[OK] Classificacao de comprovantes por regras concluida.")
    print(f"Comprovantes: {manifest.receipt_count}")
    print(f"Classificados: {manifest.classified_count}")
    print(f"Sem regra: {manifest.review_without_rule_count}")
    print(f"Multiplas regras: {manifest.review_multiple_rules_count}")
    print(f"Saida: {manifest.output_dir}")
    print("[INFO] Nenhuma conta foi inferida sem regra explicita.")
    print("[INFO] Lancamentos definitivos, DRE e Fase 9 nao foram alterados.")
    return 0


def auditar_regras_comprovantes_command(args: argparse.Namespace) -> int:
    try:
        result = audit_payment_rules_file(Path(args.regras))
        if args.relatorio:
            Path(args.relatorio).parent.mkdir(parents=True, exist_ok=True)
            Path(args.relatorio).write_text(
                payment_rules_audit_report_text(result),
                encoding="utf-8",
            )
    except (OSError, PaymentRulesError, ValueError) as exc:
        print(f"[ERRO] Nao foi possivel auditar regras de comprovantes: {exc}")
        return 1

    for line in payment_rules_audit_console_lines(result):
        print(line)
    if args.relatorio:
        print(f"Relatorio: {args.relatorio}")
    return 0


def executar_classificacao_comprovantes_real_command(args: argparse.Namespace) -> int:
    try:
        manifest = execute_payment_classification_run(
            entrada_raiz=Path(args.entrada_raiz),
            saida_raiz=Path(args.saida_raiz),
            regras_path=Path(args.regras),
            plano_contas_path=Path(args.plano_contas),
            mapeamento_dre_path=Path(args.mapeamento_dre),
            run_id=args.run_id,
        )
    except (OSError, PaymentRulesError, PaymentClassificationRunError, ValueError) as exc:
        print(f"[ERRO] Nao foi possivel executar Fase 13S: {exc}")
        return 1

    print("[OK] Fase 13S concluida: comprovantes classificados por regras.")
    print(f"Run: {manifest.run_id}")
    print(f"CNPJs: {', '.join(manifest.cnpjs_processados)}")
    print(f"Meses: {', '.join(manifest.meses_processados)}")
    print(f"Comprovantes: {manifest.total_comprovantes}")
    print(f"Classificados: {manifest.total_classificados}")
    print(f"Excluidos XML: {manifest.total_excluidos_xml}")
    print(f"Excluidos ESCLUDERE: {manifest.total_excluidos_escludere}")
    print(f"A revisar sem regra: {manifest.total_a_revisar_sem_regra}")
    print(f"A revisar multiplas regras: {manifest.total_a_revisar_multiplas_regras}")
    print(f"Saida: {manifest.output_dir}")
    print("[INFO] matching_automatico_executado=false")
    print("[INFO] Nenhuma conta foi inferida sem regra explicita.")
    print("[INFO] Lancamentos definitivos, DRE e Fase 9 nao foram alterados.")
    return 0


def gerar_pacote_revisao_comprovantes_command(args: argparse.Namespace) -> int:
    try:
        result = gerar_pacote_revisao_comprovantes(
            Path(args.run_dir),
            csv_revisao=Path(args.csv_revisao) if args.csv_revisao else None,
            xlsx_path=Path(args.saida_xlsx) if args.saida_xlsx else None,
            report_path=Path(args.relatorio) if args.relatorio else None,
        )
    except (OSError, PaymentReviewPackageError, ValueError) as exc:
        print(f"[ERRO] Nao foi possivel gerar pacote 13T: {exc}")
        return 1

    for line in comprovantes_review_console_lines(result):
        print(line)
    return 0


def aplicar_fase9_revisao_command(args: argparse.Namespace) -> int:
    csv_path = Path(args.csv)
    output_path = Path(args.saida)
    plano_contas_path = Path(args.plano_contas)

    try:
        _validate_file_exists(csv_path, "CSV revisionado")
        _validate_file_exists(plano_contas_path, "Plano de contas")
        if csv_path.resolve() == output_path.resolve():
            print("[ERRO] --saida deve ser diferente do CSV original.")
            return 1

        chart = ChartOfAccounts.load(plano_contas_path)
        validation_errors = chart.validate()
        if validation_errors:
            print("[ERRO] Plano de contas invalido: " + "; ".join(validation_errors))
            return 1
        valid_account_ids, account_names_by_id = _fase9_account_maps(chart)

        decisions = load_fase9_decisions_from_csv(csv_path)
        decision_result = evaluate_fase9_decisions(
            decisions,
            valid_account_ids=valid_account_ids,
            account_names_by_id=account_names_by_id,
        )

        if args.dry_run:
            headers, rows = _read_csv_rows_for_fase9(csv_path)
            result = apply_fase9_writeback(
                rows,
                decision_result,
                headers=headers,
            )
            _print_fase9_summary(result, output_path=output_path, dry_run=True)
            print("[DRY-RUN] Nenhum arquivo foi escrito.")
            return 0

        result = write_fase9_applied_csv(csv_path, output_path, decision_result)
    except ValueError as exc:
        print(f"[ERRO] {exc}")
        return 1
    except Exception as exc:
        print(f"[ERRO] Nao foi possivel aplicar Fase 9: {exc}")
        return 1

    _print_fase9_summary(result, output_path=output_path, dry_run=False)
    print(f"[OK] CSV Fase 9 gerado: {output_path}")
    return 0


def _fase9_account_maps(chart: ChartOfAccounts) -> tuple[frozenset[str], dict[str, str]]:
    accounts = tuple(chart._ordered)
    return (
        frozenset(account.id for account in accounts),
        {account.id: account.nome for account in accounts},
    )


def _read_csv_rows_for_fase9(csv_path: Path) -> tuple[tuple[str, ...], tuple[dict[str, str], ...]]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
        sample = file.read(4096)
        file.seek(0)
        reader = csv.DictReader(
            file,
            delimiter=sniff_delimiter(sample),
            quotechar='"',
            doublequote=True,
        )
        headers = tuple(reader.fieldnames or ())
        rows = tuple(
            {header: _text(value) for header, value in row.items()}
            for row in reader
        )
    return headers, rows


def _print_fase9_summary(result, *, output_path: Path, dry_run: bool) -> None:
    prefix = "[DRY-RUN]" if dry_run else "[OK]"
    print(f"{prefix} Fase 9 revisao calculada.")
    print(f"Promovidos: {result.promoted_rows}")
    print(f"Bloqueados: {result.blocked_rows}")
    print(f"Fora escopo: {result.out_of_scope_rows}")
    print(f"Saida: {output_path}")


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


def gerar_pre_livro_diario_teste_command(args: argparse.Namespace) -> int:
    try:
        result = gerar_pre_livro_diario_teste(
            [Path(excel_path) for excel_path in args.excel],
            Path(args.saida),
            [
                Path(pasta_path)
                for pasta_path in (args.pasta_operacoes_especiais or [])
            ],
            Path(args.relatorio_auditoria) if args.relatorio_auditoria else None,
        )
    except PreLivroDiarioError as exc:
        print(f"[ERRO] {exc}")
        return 1
    except Exception:
        print(f"[ERRO] Nao foi possivel gerar pre-livro diario: {args.saida}")
        return 1

    for line in pre_livro_console_lines(result):
        print(line)
    if args.relatorio_auditoria:
        print(f"[OK] Relatorio de auditoria salvo em: {args.relatorio_auditoria}")
    return 0


def gerar_template_lancamentos_manuais_gerenciais_command(
    args: argparse.Namespace,
) -> int:
    try:
        result = gerar_template_lancamentos_manuais_gerenciais(
            Path(args.saida),
            Path(args.plano_contas),
            Path(args.mapeamento_dre),
        )
    except LancamentosManuaisTemplateError as exc:
        print(f"[ERRO] {exc}")
        return 1
    except Exception:
        print(f"[ERRO] Nao foi possivel gerar template: {args.saida}")
        return 1

    print(f"[OK] Template gerado: {result.output_xlsx}")
    print(f"[INFO] Categorias disponiveis: {len(result.categorias)}")
    if result.conta_credito_tecnica:
        print(
            "[INFO] Contrapartida tecnica: "
            f"{result.conta_credito_tecnica.conta_id} - "
            f"{result.conta_credito_tecnica.conta_nome}"
        )
    else:
        print("[BLOQUEIO] Contrapartida tecnica neutra fora DRE ausente.")
    for blocker in result.blockers:
        print(f"[BLOQUEIO] {blocker}")
    return 1 if result.blockers else 0


def validar_lancamentos_manuais_gerenciais_command(args: argparse.Namespace) -> int:
    manifest_path = (
        Path(args.manifest)
        if args.manifest
        else Path(args.saida_csv).parent / "registrazioni_manuali_validazione_manifest.json"
    )
    try:
        result = validar_lancamentos_manuais_gerenciais_operador(
            Path(args.entrada),
            Path(args.saida_csv),
            Path(args.relatorio) if args.relatorio else None,
            Path(args.plano_contas),
            Path(args.mapeamento_dre),
            manifest_path,
        )
    except LancamentosManuaisTemplateError as exc:
        print(f"[ERRO] {exc}")
        if args.relatorio:
            print(f"[INFO] Relatorio salvo em: {args.relatorio}")
        if manifest_path.exists():
            print(f"[INFO] Manifest salvo em: {manifest_path}")
        else:
            print(f"[BLOCCO] BLOCCO_MANIFEST_NAO_GERADO: {manifest_path}")
        return 1
    except Exception:
        print(f"[ERRO] Nao foi possivel validar lancamentos manuais: {args.entrada}")
        print(f"[BLOCCO] BLOCCO_MANIFEST_NAO_GERADO: {manifest_path}")
        return 1

    print(f"[OK] CSV tecnico gerado: {result.saida_csv}")
    print(f"[INFO] Linhas lidas: {result.linhas_lidas}")
    print(f"[INFO] Linhas validas: {result.linhas_validas}")
    print(f"[INFO] Avisos: {len(result.avisos)}")
    if result.relatorio:
        print(f"[OK] Relatorio salvo em: {result.relatorio}")
    print(f"[OK] Manifest salvo em: {manifest_path}")
    return 0


def executar_lancamentos_manuais_gerenciais_command(args: argparse.Namespace) -> int:
    template_path = Path(args.template)
    csv_path = Path(args.saida_csv)
    plano_path = Path(args.plano_contas)
    dre_path = Path(args.mapeamento_dre)
    relatorio_val = Path(args.relatorio_validacao) if args.relatorio_validacao else None

    print(f"[INFO] Template: {template_path}")
    print(f"[INFO] Validando lancamentos manuais gerenciais...")
    try:
        val_result = validar_lancamentos_manuais_gerenciais_operador(
            template_path,
            csv_path,
            relatorio_val,
            plano_path,
            dre_path,
        )
    except LancamentosManuaisTemplateError as exc:
        print(f"[ERRO] Validacao bloqueada: {exc}")
        if relatorio_val:
            print(f"[INFO] Relatorio de erros: {relatorio_val}")
        return 1
    except Exception as exc:
        print(f"[ERRO] Nao foi possivel validar lancamentos manuais: {exc}")
        return 1

    print(f"[OK] Validacao: {val_result.linhas_validas} linha(s) valida(s)")
    for aviso in val_result.avisos:
        print(f"[AVISO] {aviso}")

    print("[INFO] Executando run provisoria com lancamentos manuais gerenciais...")
    try:
        run_result = reprocessamento_pipeline.reprocessar_importacao_real_provisoria(
            xml_base=Path(args.xml_base),
            saida_output=Path(args.saida_output),
            saida_reports=Path(args.saida_reports),
            lancamentos_manuais_gerenciais_csv_path=csv_path,
            gerar_pacote_contabil_gerencial=True,
        )
    except reprocessamento_pipeline.ReprocessamentoPipelineError as exc:
        print(f"[ERRO] {exc}")
        return 1
    except RunManagerError as exc:
        print(f"[ERRO] {exc}")
        return 1
    except Exception as exc:
        print(f"[ERRO] Pipeline falhou: {exc}")
        return 1

    try:
        with run_result.manifest_path.open(encoding="utf-8") as _f:
            manifest = json.load(_f)
    except Exception as exc:
        print(f"[ERRO] Manifest nao pode ser lido: {exc}")
        return 1

    totais = manifest.get("totais", {})
    manuais = manifest.get("lancamentos_manuais_gerenciais", {})
    pacote = manifest.get("pacote_contabil_gerencial", {})

    print()
    print("=" * 60)
    print("FANTOZZI - EXECUTAR LANCAMENTOS MANUAIS GERENCIAIS")
    print("=" * 60)
    print(f"run_id            : {run_result.run_id}")
    print(f"status_run        : {run_result.status_run}")
    print()
    print("MANUAIS GERENCIAIS:")
    print(f"  linhas_validas        : {val_result.linhas_validas}")
    print(f"  lancamentos_aceitos   : {manuais.get('lancamentos_aceitos')}")
    print(f"  linhas_tecnicas       : {manuais.get('linhas_tecnicas_geradas')}")
    print(f"  total_manual          : R$ {manuais.get('total_manual')}")
    print()
    print("TOTAIS RUN:")
    print(f"  xml_processados       : {totais.get('xml_processados')}")
    print(f"  lancamentos_ok        : {totais.get('lancamentos_ok')}")
    print(f"  lancamentos_revisar   : {totais.get('lancamentos_revisar')}")
    print(f"  valor_total_ok        : R$ {totais.get('valor_total_incluido_ok')}")
    print()
    print("VERIFICACOES:")
    avisos_finais: list[str] = []
    if manuais.get("executado"):
        print("  [OK] MANUAL_GERENCIAL aplicado na run")
    else:
        avisos_finais.append("MANUAL_GERENCIAL NAO foi aplicado")
    if pacote.get("executado") and pacote.get("status") == "OK":
        print("  [OK] Pacote contabil gerencial gerado")
        print(f"       -> {pacote.get('diretorio')}")
    else:
        avisos_finais.append(f"Pacote contabil gerencial: {pacote.get('status')}")
    for aviso in run_result.avisos:
        print(f"  [INFO] {aviso}")
    for item in avisos_finais:
        print(f"  [ATENCAO] {item}")
    print()
    print(f"Manifest  : {run_result.manifest_path}")
    print(f"CSV tecnico: {csv_path}")
    print()
    if manifest.get("erros"):
        print("RESULTADO: ERRO")
        for err in manifest["erros"]:
            print(f"  - {err}")
        return 1
    print("RESULTADO: OK")
    return 0


def gerar_pacote_revisao_contabil_command(args: argparse.Namespace) -> int:
    try:
        result = gerar_pacote_revisao_contabil(
            [Path(csv_path) for csv_path in args.csv],
            Path(args.saida_xlsx),
            Path(args.relatorio) if args.relatorio else None,
            limite_top=args.limite_top,
            separador=args.separador,
        )
    except ClassificationReviewPackageError as exc:
        print(f"[ERRO] {exc}")
        return 1
    except Exception:
        print(f"[ERRO] Nao foi possivel gerar pacote de revisao: {args.saida_xlsx}")
        return 1

    for line in review_package_console_lines(result):
        print(line)
    return 0


def gerar_relatorios_gerenciais_ok_provisorios_command(
    args: argparse.Namespace,
) -> int:
    try:
        result = gerar_relatorios_gerenciais_ok_provisorios(
            [Path(entrada) for entrada in args.entrada],
            Path(args.saida_xlsx),
            Path(args.relatorio) if args.relatorio else None,
            [Path(pasta) for pasta in (args.pasta_operacoes_especiais or [])],
        )
    except RelatoriosGerenciaisOkError as exc:
        print(f"[ERRO] {exc}")
        return 1
    except Exception:
        print(
            "[ERRO] Nao foi possivel gerar relatorios gerenciais OK "
            f"provisorios: {args.saida_xlsx}"
        )
        return 1

    for line in relatorios_ok_console_lines(result):
        print(line)
    return 0


def auditar_relatorios_gerenciais_ok_provisorios_command(
    args: argparse.Namespace,
) -> int:
    try:
        result = auditar_relatorios_gerenciais_ok_provisorios(
            [Path(entrada) for entrada in args.entrada],
            Path(args.saida_xlsx),
            Path(args.relatorio) if args.relatorio else None,
            [Path(pasta) for pasta in (args.pasta_operacoes_especiais or [])],
        )
    except AuditoriaRelatoriosGerenciaisOkError as exc:
        print(f"[ERRO] {exc}")
        return 1
    except Exception:
        print(
            "[ERRO] Nao foi possivel auditar relatorios gerenciais OK "
            f"provisorios: {args.saida_xlsx}"
        )
        return 1

    for line in auditoria_relatorios_ok_console_lines(result):
        print(line)
    return 0


def gerar_dre_gerencial_ok_provisoria_command(args: argparse.Namespace) -> int:
    try:
        result = gerar_dre_gerencial_ok_provisoria(
            [Path(entrada) for entrada in args.entrada],
            Path(args.mapeamento_dre),
            Path(args.saida_xlsx),
            Path(args.relatorio) if args.relatorio else None,
            [Path(pasta) for pasta in (args.pasta_operacoes_especiais or [])],
        )
    except DreGerencialProvisoriaError as exc:
        print(f"[ERRO] {exc}")
        return 1
    except Exception:
        print(
            "[ERRO] Nao foi possivel gerar DRE gerencial OK provisoria: "
            f"{args.saida_xlsx}"
        )
        return 1

    for line in dre_provisoria_console_lines(result):
        print(line)
    return 0


def preparar_run_reprocessamento_command(args: argparse.Namespace) -> int:
    try:
        result = preparar_run(
            xml_base=Path(args.xml_base),
            saida_output=Path(args.saida_output),
            saida_reports=Path(args.saida_reports),
            dry_run=args.dry_run,
        )
    except RunManagerError as exc:
        print(f"[ERRO] {exc}")
        return 1
    except Exception as exc:
        print(f"[ERRO] Nao foi possivel preparar a run: {exc}")
        return 1

    if result.dry_run:
        print(f"[DRY-RUN] Run calculada: {result.run_id}")
        print(f"[DRY-RUN] Output da run: {result.output_run_dir}")
        print(f"[DRY-RUN] Reports da run: {result.reports_run_dir}")
        print(f"[DRY-RUN] Manifesto (nao criado): {result.manifest_path}")
        print("[DRY-RUN] Nenhum arquivo foi escrito.")
    else:
        print(f"[OK] Run preparada: {result.run_id}")
        print(f"[OK] Output da run: {result.output_run_dir}")
        print(f"[OK] Reports da run: {result.reports_run_dir}")
        print(f"[OK] Manifesto criado: {result.manifest_path}")
        print(f"[OK] current_run.txt atualizado: {result.current_run_output}")
        print(f"[OK] current_run.txt atualizado: {result.current_run_reports}")
        print("[INFO] Nenhum reprocessamento executado nesta fase.")
    return 0


def reprocessar_importacao_real_command(args: argparse.Namespace) -> int:
    try:
        result = reprocessamento_pipeline.reprocessar_importacao_real_provisoria(
            xml_base=Path(args.xml_base),
            saida_output=Path(args.saida_output),
            saida_reports=Path(args.saida_reports),
            modo=args.modo,
            dry_run=args.dry_run,
            fase9_csv_path=args.fase9_csv,
            lancamentos_manuais_gerenciais_csv_path=(
                args.lancamentos_manuais_gerenciais_csv
            ),
            lancamentos_banca_aprovados_path=args.lancamentos_banca_aprovados,
            gerar_pacote_contabil_gerencial=args.gerar_pacote_contabil_gerencial,
        )
    except reprocessamento_pipeline.ReprocessamentoPipelineError as exc:
        print(f"[ERRO] {exc}")
        return 1
    except RunManagerError as exc:
        print(f"[ERRO] {exc}")
        return 1
    except Exception as exc:
        print(f"[ERRO] Nao foi possivel reprocessar importacao real: {exc}")
        return 1

    if result.dry_run:
        print(f"[DRY-RUN] Run calculada: {result.run_id}")
        print(f"[DRY-RUN] Output da run: {result.output_run_dir}")
        print(f"[DRY-RUN] Reports da run: {result.reports_run_dir}")
        print(f"[DRY-RUN] Manifesto (nao criado): {result.manifest_path}")
        print("[DRY-RUN] Etapas previstas:")
        for etapa in result.etapas_previstas:
            print(f"     - {etapa}")
        print("[DRY-RUN] Nenhum arquivo foi criado.")
        print("[DRY-RUN] Nenhum XML foi importado.")
        print("[DRY-RUN] Nenhum relatorio foi gerado.")
        return 0

    print(f"[OK] Reprocessamento provisorio concluido: {result.run_id}")
    print(f"[OK] Manifesto: {result.manifest_path}")
    if args.fase9_csv:
        print(
            "[INFO] Fase 9 executada em CSV separado; "
            "relatorios permanecem provisorios."
        )
    else:
        print("[INFO] Fase 9 nao executada; relatorios permanecem provisorios.")
    if args.lancamentos_manuais_gerenciais_csv:
        print("[INFO] Lancamentos manuais gerenciais aplicados na run.")
    if args.lancamentos_banca_aprovados:
        print("[INFO] Lancamentos banca aprovados incluidos na run.")
    if args.gerar_pacote_contabil_gerencial:
        print("[INFO] Pacote contabil gerencial provisorio gerado na run.")
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


def _text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


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
