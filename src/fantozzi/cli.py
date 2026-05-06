from __future__ import annotations

import argparse
from pathlib import Path

from fantozzi.allocation import allocate_nfe_to_lancamentos
from fantozzi.chart_of_accounts import ChartOfAccounts
from fantozzi.excel_writer import write_lancamentos_to_excel
from fantozzi.logging_utils import error, info, warning
from fantozzi.nfe_importer import parse_nfe
from fantozzi.product_classifier import ProductClassifier


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PLANO_CONTAS = PROJECT_ROOT / "config" / "plano_contas_gerencial.yaml"
DEFAULT_REGRAS_PRODUTOS = (
    PROJECT_ROOT / "config" / "regras_classificacao_produtos.yaml"
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="fantozzi",
        description="CLI minima do Fantozzi para importar NF-e XML.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    importar_nfe = subparsers.add_parser(
        "importar-nfe",
        help="Importa um arquivo XML NF-e e atualiza um Excel de lancamentos.",
    )
    importar_nfe.add_argument("--xml", required=True, help="Arquivo XML NF-e.")
    importar_nfe.add_argument(
        "--excel", required=True, help="Arquivo Excel .xlsx de saida."
    )
    importar_nfe.add_argument(
        "--plano-contas",
        default=str(DEFAULT_PLANO_CONTAS),
        help="Arquivo YAML do plano de contas gerencial.",
    )
    importar_nfe.add_argument(
        "--regras-produtos",
        default=str(DEFAULT_REGRAS_PRODUTOS),
        help="Arquivo YAML com regras de classificacao de produtos.",
    )
    importar_nfe.set_defaults(func=importar_nfe_command)

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

        chart = ChartOfAccounts.load(plano_contas_path)
        validation_errors = chart.validate()
        if validation_errors:
            for validation_error in validation_errors:
                error(validation_error)
            return 1

        classifier = ProductClassifier(chart).load(regras_produtos_path)
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


def _validate_file_exists(path: Path, label: str) -> None:
    if not path.exists():
        raise FileNotFoundError(f"{label} nao encontrado: {path}")
    if not path.is_file():
        raise ValueError(f"{label} nao e um arquivo: {path}")


if __name__ == "__main__":
    raise SystemExit(main())
