from __future__ import annotations

import csv
import json
from pathlib import Path

import pytest

from fantozzi.banking.payment_rules import (
    STATUS_CLASSIFICADO,
    STATUS_MULTIPLAS_REGRAS,
    STATUS_SEM_REGRA,
    PaymentRulesError,
    classify_payment_receipt,
    classify_payment_receipts_file,
    load_payment_rules,
    normalize_match_text,
    validate_payment_rules,
)
from fantozzi.cli import build_parser, main


def test_match_by_beneficiary_and_reference(tmp_path):
    rules_path = _write_rules(
        tmp_path,
        [
            _rule(
                "REGRA_EMBASA",
                "4.01.012",
                "Agua e Esgoto",
                "06.02",
                beneficiary=["EMBASA"],
                reference=["AGUA LOJA"],
            )
        ],
    )
    rules = load_payment_rules(rules_path)

    result = classify_payment_receipt(
        {
            "beneficiary_name": "Embasa SA",
            "reference_id": "Agua Loja Centro",
        },
        rules,
    )

    assert result.status_classificacao == STATUS_CLASSIFICADO
    assert result.conta_id == "4.01.012"
    assert result.regra_id == "REGRA_EMBASA"


def test_match_by_reference_only(tmp_path):
    rules_path = _write_rules(
        tmp_path,
        [
            _rule(
                "REGRA_REF",
                "6.01.001",
                "Honorarios Contabeis",
                "08.01",
                reference=["CONTABILIDADE"],
            )
        ],
    )
    rules = load_payment_rules(rules_path)

    result = classify_payment_receipt(
        {
            "beneficiary_name": "Fornecedor Generico",
            "reference_id": "Mensalidade Contabilidade",
        },
        rules,
    )

    assert result.status_classificacao == STATUS_CLASSIFICADO
    assert result.conta_id == "6.01.001"


def test_without_rule_goes_to_review(tmp_path):
    rules = load_payment_rules(_write_rules(tmp_path, []))

    result = classify_payment_receipt(
        {"beneficiary_name": "Sem Regra", "reference_id": "Outro"},
        rules,
    )

    assert result.status_classificacao == STATUS_SEM_REGRA
    assert result.motivo == "SEM_REGRA"
    assert result.conta_id == ""


def test_multiple_matching_rules_go_to_review(tmp_path):
    rules_path = _write_rules(
        tmp_path,
        [
            _rule(
                "REGRA_A",
                "7.01.001",
                "Tarifas Bancarias",
                "09.01",
                reference=["TARIFA PIX"],
            ),
            _rule(
                "REGRA_B",
                "7.01.003",
                "Tarifas PIX",
                "09.01",
                reference=["TARIFA PIX"],
            ),
        ],
    )
    rules = load_payment_rules(rules_path)

    result = classify_payment_receipt(
        {"beneficiary_name": "Banco", "reference_id": "Tarifa Pix Pagamento"},
        rules,
    )

    assert result.status_classificacao == STATUS_MULTIPLAS_REGRAS
    assert result.regras_candidatas == ("REGRA_A", "REGRA_B")


def test_multiple_matching_rules_same_account_are_classified(tmp_path):
    rules_path = _write_rules(
        tmp_path,
        [
            _rule(
                "REGRA_PARCSN",
                "5.01.003",
                "Parcelamento Federal / Receita Federal / PGFN / Simples",
                "07.01",
                reference=["PARCSN"],
            ),
            _rule(
                "REGRA_PERTSN",
                "5.01.003",
                "Parcelamento Federal / Receita Federal / PGFN / Simples",
                "07.01",
                reference=["PERTSN"],
            ),
        ],
    )
    rules = load_payment_rules(rules_path)

    result = classify_payment_receipt(
        {
            "beneficiary_name": "DAS-SIMPLES NACIONAL-COD BARRA",
            "reference_id": "PARCSN 2 RATA 28 PERTSN 9101 RATA 94",
        },
        rules,
    )

    assert result.status_classificacao == STATUS_CLASSIFICADO
    assert result.conta_id == "5.01.003"
    assert result.regra_id == "REGRA_PARCSN"
    assert result.regras_candidatas == ("REGRA_PARCSN", "REGRA_PERTSN")


def test_normalization_removes_accents_case_and_punctuation():
    assert normalize_match_text("  Companhia de ÁGUA/esgoto!!! ") == (
        "COMPANHIA DE AGUA ESGOTO"
    )


def test_rule_with_missing_account_fails_validation(tmp_path):
    rules = load_payment_rules(
        _write_rules(
            tmp_path,
            [
                _rule(
                    "REGRA_MISSING",
                    "9.99.999",
                    "Conta Inexistente",
                    "99.01",
                    reference=["TESTE"],
                )
            ],
        )
    )
    plano_path, mapeamento_path = _write_plan_and_mapping(tmp_path)

    with pytest.raises(PaymentRulesError, match="conta_id inexistente"):
        validate_payment_rules(
            rules,
            plano_contas_path=plano_path,
            mapeamento_dre_path=mapeamento_path,
        )


def test_rule_with_divergent_dre_mapping_fails_validation(tmp_path):
    rules = load_payment_rules(
        _write_rules(
            tmp_path,
            [
                _rule(
                    "REGRA_DRE",
                    "4.01.012",
                    "Agua e Esgoto",
                    "99.01",
                    reference=["EMBASA"],
                )
            ],
        )
    )
    plano_path, mapeamento_path = _write_plan_and_mapping(tmp_path)

    with pytest.raises(PaymentRulesError, match="dre_linha_id"):
        validate_payment_rules(
            rules,
            plano_contas_path=plano_path,
            mapeamento_dre_path=mapeamento_path,
        )


def test_cli_classifies_and_writes_expected_csvs(tmp_path):
    plano_path, mapeamento_path = _write_plan_and_mapping(tmp_path)
    rules_path = _write_rules(
        tmp_path,
        [
            _rule(
                "REGRA_EMBASA",
                "4.01.012",
                "Agua e Esgoto",
                "06.02",
                beneficiary=["EMBASA"],
                reference=["AGUA"],
            )
        ],
    )
    comprovantes_path = _write_receipts_csv(
        tmp_path,
        [
            _receipt("RC-1", "Embasa", "Agua mensal"),
            _receipt("RC-2", "Fornecedor", "Sem regra"),
        ],
    )

    result = main(
        [
            "classificar-comprovantes-pagamento",
            "--comprovantes",
            str(comprovantes_path),
            "--regras",
            str(rules_path),
            "--plano-contas",
            str(plano_path),
            "--mapeamento-dre",
            str(mapeamento_path),
            "--saida",
            str(tmp_path / "saida"),
        ]
    )

    assert result == 0
    classified = _read_csv(tmp_path / "saida" / "comprovantes_classificados.csv")
    review = _read_csv(tmp_path / "saida" / "comprovantes_a_revisar.csv")
    manifest = json.loads(
        (tmp_path / "saida" / "manifest_classificacao_comprovantes.json").read_text(
            encoding="utf-8"
        )
    )
    assert classified[0]["payment_id"] == "RC-1"
    assert classified[0]["status_classificacao"] == STATUS_CLASSIFICADO
    assert review[0]["payment_id"] == "RC-2"
    assert review[0]["status_classificacao"] == STATUS_SEM_REGRA
    assert manifest["classification_executed"] is True
    assert manifest["nao_inferir_conta"] is True


def test_file_api_validates_rules_before_output(tmp_path):
    plano_path, mapeamento_path = _write_plan_and_mapping(tmp_path)
    rules_path = _write_rules(
        tmp_path,
        [
            _rule(
                "REGRA_OK",
                "4.01.012",
                "Agua e Esgoto",
                "06.02",
                reference=["AGUA"],
            )
        ],
    )
    comprovantes_path = _write_receipts_csv(
        tmp_path,
        [_receipt("RC-1", "Embasa", "Agua mensal")],
    )

    manifest = classify_payment_receipts_file(
        comprovantes_path=comprovantes_path,
        regras_path=rules_path,
        plano_contas_path=plano_path,
        mapeamento_dre_path=mapeamento_path,
        output_dir=tmp_path / "out",
    )

    assert manifest.classified_count == 1
    assert (tmp_path / "out" / "comprovantes_classificados.csv").exists()


def test_parser_accepts_classificar_comprovantes_pagamento_command():
    parser = build_parser()

    args = parser.parse_args(
        [
            "classificar-comprovantes-pagamento",
            "--comprovantes",
            "payment_receipts_normalized.csv",
            "--saida",
            "output/fase_13h",
        ]
    )

    assert args.command == "classificar-comprovantes-pagamento"
    assert args.comprovantes == "payment_receipts_normalized.csv"
    assert args.saida == "output/fase_13h"


def _rule(
    regra_id: str,
    conta_id: str,
    conta_nome: str,
    dre_linha_id: str,
    *,
    beneficiary: list[str] | None = None,
    reference: list[str] | None = None,
    entra_dre: bool = True,
) -> str:
    beneficiary_items = _yaml_list(beneficiary or [])
    reference_items = _yaml_list(reference or [])
    return f"""- regra_id: {regra_id}
  conta_id: "{conta_id}"
  conta_nome: {conta_nome}
  dre_linha_id: "{dre_linha_id}"
  entra_dre: {str(entra_dre).lower()}
  match:
    modo: TODAS_AS_CONDICOES_NAO_VAZIAS
    beneficiario_favorecido_contem_algum:
{beneficiary_items}
    referencia_empresa_contem_algum:
{reference_items}
"""


def _yaml_list(values: list[str]) -> str:
    if not values:
        return "      []"
    return "\n".join(f"      - {value}" for value in values)


def _write_rules(tmp_path: Path, rules: list[str]) -> Path:
    path = tmp_path / "regras.yaml"
    if not rules:
        path.write_text("regras: []\n", encoding="utf-8")
        return path
    rules_text = "\n".join(rules)
    path.write_text(f"regras:\n{rules_text}\n", encoding="utf-8")
    return path


def _write_plan_and_mapping(tmp_path: Path) -> tuple[Path, Path]:
    plano_path = tmp_path / "plano.yaml"
    plano_path.write_text(
        """contas:
- id: "4.01.012"
  nome: Agua e Esgoto
  grupo: Ocupacao
  tipo: despesa
- id: "6.01.001"
  nome: Honorarios Contabeis
  grupo: Administrativo
  tipo: despesa
- id: "7.01.001"
  nome: Tarifas Bancarias
  grupo: Financeiro
  tipo: despesa
- id: "7.01.003"
  nome: Tarifas PIX
  grupo: Financeiro
  tipo: despesa
""",
        encoding="utf-8",
    )
    mapeamento_path = tmp_path / "mapeamento.yaml"
    mapeamento_path.write_text(
        """mapeamento:
  "4.01.012":
    dre_linha_id: "06.02"
    entra_dre: true
  "6.01.001":
    dre_linha_id: "08.01"
    entra_dre: true
  "7.01.001":
    dre_linha_id: "09.01"
    entra_dre: true
  "7.01.003":
    dre_linha_id: "09.01"
    entra_dre: true
""",
        encoding="utf-8",
    )
    return plano_path, mapeamento_path


def _receipt(payment_id: str, beneficiary: str, reference: str) -> dict[str, str]:
    return {
        "payment_id": payment_id,
        "cnpj": "08884180000123",
        "payment_date": "2026-02-05",
        "amount": "20.00",
        "beneficiary_name": beneficiary,
        "beneficiary_document": "12.345.678/0001-90",
        "payment_type": "PIX",
        "reference_id": reference,
        "status": "Efetuado",
        "source_file": "ConsultaPagamentos.xls",
        "raw_hash": payment_id,
    }


def _write_receipts_csv(tmp_path: Path, rows: list[dict[str, str]]) -> Path:
    path = tmp_path / "payment_receipts_normalized.csv"
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=list(rows[0].keys()), delimiter=";")
        writer.writeheader()
        writer.writerows(rows)
    return path


def _read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file, delimiter=";"))
