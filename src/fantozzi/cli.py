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
    _add_config_arguments(importar_nfe)
    importar_nfe.set_defaults(func=importar_nfe_command)

    importar_pasta_nfe = subparsers.add_parser(
        "importar-pasta-nfe",
        help="Importa uma pasta de arquivos XML NF-e e atualiza um Excel.",
    )
    importar_pasta_nfe.add_argument(
        "--pasta", required=True, help="Pasta com arquivos XML NF-e."
    )
    importar_pasta_nfe.add_argument(
        "--excel", required=True, help="Arquivo Excel .xlsx de saida."
    )
    _add_config_arguments(importar_pasta_nfe)
    importar_pasta_nfe.set_defaults(func=importar_pasta_nfe_command)

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


if __name__ == "__main__":
    raise SystemExit(main())
