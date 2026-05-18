from __future__ import annotations

import csv
import json
import re
import unicodedata
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

import yaml

from fantozzi.csv_utils import sniff_delimiter


STATUS_CLASSIFICADO = "CLASSIFICADO_POR_REGRA"
STATUS_SEM_REGRA = "REVISAR_SEM_REGRA"
STATUS_MULTIPLAS_REGRAS = "REVISAR_MULTIPLAS_REGRAS"

MOTIVO_REGRA_UNICA = "REGRA_UNICA_EXPLICITA"
MOTIVO_SEM_REGRA = "SEM_REGRA"
MOTIVO_MULTIPLAS_REGRAS = "MAIS_DE_UMA_REGRA"

MATCH_MODE = "TODAS_AS_CONDICOES_NAO_VAZIAS"

CLASSIFICATION_HEADERS = [
    "status_classificacao",
    "conta_id",
    "conta_nome",
    "dre_linha_id",
    "entra_dre",
    "regra_id",
    "motivo",
    "regras_candidatas",
]

CONDITION_FIELDS = {
    "beneficiario_favorecido_contem_algum": (
        "beneficiary_name",
        "beneficiario_favorecido",
        "favorecido_beneficiario",
    ),
    "referencia_empresa_contem_algum": (
        "reference_id",
        "referencia_empresa",
    ),
    "tipo_pagamento_contem_algum": (
        "payment_type",
        "tipo_pagamento",
    ),
    "descricao_comprovante_contem_algum": (
        "description_receipt",
        "descricao_comprovante",
    ),
}


class PaymentRulesError(ValueError):
    pass


@dataclass(frozen=True)
class PaymentRule:
    regra_id: str
    conta_id: str
    conta_nome: str
    dre_linha_id: str
    entra_dre: bool
    match: dict[str, tuple[str, ...]]


@dataclass(frozen=True)
class PaymentClassificationResult:
    status_classificacao: str
    conta_id: str
    conta_nome: str
    dre_linha_id: str
    entra_dre: bool | None
    regra_id: str
    motivo: str
    regras_candidatas: tuple[str, ...]


@dataclass(frozen=True)
class PaymentClassificationManifest:
    input_files: dict[str, str]
    output_dir: Path
    receipt_count: int
    classified_count: int
    review_without_rule_count: int
    review_multiple_rules_count: int
    status: str
    generated_files: dict[str, str]


def normalize_match_text(value: object) -> str:
    normalized = unicodedata.normalize("NFKD", "" if value is None else str(value))
    ascii_text = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    text = re.sub(r"[^A-Za-z0-9]+", " ", ascii_text.upper())
    return re.sub(r"\s+", " ", text).strip()


def load_payment_rules(path: Path | str) -> list[PaymentRule]:
    rules_path = Path(path)
    with rules_path.open(encoding="utf-8") as file:
        data = yaml.safe_load(file) or {}
    raw_rules = data.get("regras", [])
    if not isinstance(raw_rules, list):
        raise PaymentRulesError("Campo 'regras' deve ser uma lista.")
    return [_rule_from_data(item, index) for index, item in enumerate(raw_rules, start=1)]


def validate_payment_rules(
    rules: list[PaymentRule],
    *,
    plano_contas_path: Path | str,
    mapeamento_dre_path: Path | str,
) -> None:
    accounts = _load_accounts(Path(plano_contas_path))
    mapping = _load_dre_mapping(Path(mapeamento_dre_path))
    errors: list[str] = []

    for rule in rules:
        account = accounts.get(rule.conta_id)
        if account is None:
            errors.append(f"{rule.regra_id}: conta_id inexistente: {rule.conta_id}")
            continue
        if account.get("nome") and account["nome"] != rule.conta_nome:
            errors.append(
                f"{rule.regra_id}: conta_nome '{rule.conta_nome}' diverge do plano "
                f"('{account['nome']}')."
            )

        dre_entry = mapping.get(rule.conta_id)
        if dre_entry is None:
            errors.append(f"{rule.regra_id}: conta_id sem mapeamento DRE: {rule.conta_id}")
            continue
        mapped_dre = str(dre_entry.get("dre_linha_id", ""))
        if mapped_dre != rule.dre_linha_id:
            errors.append(
                f"{rule.regra_id}: dre_linha_id '{rule.dre_linha_id}' diverge do "
                f"mapeamento atual ('{mapped_dre}')."
            )
        mapped_entra_dre = bool(dre_entry.get("entra_dre", True))
        if mapped_entra_dre != rule.entra_dre:
            errors.append(
                f"{rule.regra_id}: entra_dre '{rule.entra_dre}' diverge do "
                f"mapeamento atual ('{mapped_entra_dre}')."
            )

    if errors:
        raise PaymentRulesError("; ".join(errors))


def classify_payment_receipt(
    receipt: dict[str, str],
    rules: list[PaymentRule],
) -> PaymentClassificationResult:
    candidates = [rule for rule in rules if _matches_rule(receipt, rule)]
    candidate_ids = tuple(rule.regra_id for rule in candidates)
    if not candidates:
        return PaymentClassificationResult(
            status_classificacao=STATUS_SEM_REGRA,
            conta_id="",
            conta_nome="",
            dre_linha_id="",
            entra_dre=None,
            regra_id="",
            motivo=MOTIVO_SEM_REGRA,
            regras_candidatas=(),
        )
    candidate_account_ids = {rule.conta_id for rule in candidates}
    if len(candidate_account_ids) > 1:
        return PaymentClassificationResult(
            status_classificacao=STATUS_MULTIPLAS_REGRAS,
            conta_id="",
            conta_nome="",
            dre_linha_id="",
            entra_dre=None,
            regra_id="",
            motivo=MOTIVO_MULTIPLAS_REGRAS,
            regras_candidatas=candidate_ids,
        )

    rule = candidates[0]
    return PaymentClassificationResult(
        status_classificacao=STATUS_CLASSIFICADO,
        conta_id=rule.conta_id,
        conta_nome=rule.conta_nome,
        dre_linha_id=rule.dre_linha_id,
        entra_dre=rule.entra_dre,
        regra_id=rule.regra_id,
        motivo=MOTIVO_REGRA_UNICA,
        regras_candidatas=candidate_ids,
    )


def classify_payment_receipts_file(
    *,
    comprovantes_path: Path,
    regras_path: Path,
    plano_contas_path: Path,
    mapeamento_dre_path: Path,
    output_dir: Path,
) -> PaymentClassificationManifest:
    rules = load_payment_rules(regras_path)
    validate_payment_rules(
        rules,
        plano_contas_path=plano_contas_path,
        mapeamento_dre_path=mapeamento_dre_path,
    )

    headers, rows = _read_csv(comprovantes_path)
    classified_rows: list[dict[str, object]] = []
    review_rows: list[dict[str, object]] = []

    for row in rows:
        result = classify_payment_receipt(row, rules)
        output_row = {**row, **_result_row(result)}
        if result.status_classificacao == STATUS_CLASSIFICADO:
            classified_rows.append(output_row)
        else:
            review_rows.append(output_row)

    output_dir.mkdir(parents=True, exist_ok=True)
    classified_path = output_dir / "comprovantes_classificados.csv"
    review_path = output_dir / "comprovantes_a_revisar.csv"
    manifest_path = output_dir / "manifest_classificacao_comprovantes.json"
    output_headers = [*headers, *CLASSIFICATION_HEADERS]
    _write_csv(classified_path, output_headers, classified_rows)
    _write_csv(review_path, output_headers, review_rows)

    review_without_rule_count = sum(
        1 for row in review_rows if row["status_classificacao"] == STATUS_SEM_REGRA
    )
    review_multiple_rules_count = sum(
        1 for row in review_rows if row["status_classificacao"] == STATUS_MULTIPLAS_REGRAS
    )
    manifest = PaymentClassificationManifest(
        input_files={
            "payment_receipts_normalized": str(comprovantes_path),
            "regras_classificacao_comprovantes": str(regras_path),
            "plano_contas": str(plano_contas_path),
            "mapeamento_dre": str(mapeamento_dre_path),
        },
        output_dir=output_dir,
        receipt_count=len(rows),
        classified_count=len(classified_rows),
        review_without_rule_count=review_without_rule_count,
        review_multiple_rules_count=review_multiple_rules_count,
        status="OK" if not review_rows else "ATENCAO",
        generated_files={
            "comprovantes_classificados": str(classified_path),
            "comprovantes_a_revisar": str(review_path),
            "manifest": str(manifest_path),
        },
    )
    _write_manifest(manifest)
    return manifest


def _rule_from_data(data: Any, index: int) -> PaymentRule:
    if not isinstance(data, dict):
        raise PaymentRulesError(f"Regra #{index} deve ser um objeto.")

    regra_id = str(data.get("regra_id") or data.get("id") or f"REGRA_{index:03d}")
    conta_id = str(data.get("conta_id") or "").strip()
    conta_nome = str(data.get("conta_nome") or "").strip()
    dre_linha_id = str(data.get("dre_linha_id") or "").strip()
    entra_dre = data.get("entra_dre", True)
    match_data = data.get("match")

    missing = [
        name
        for name, value in (
            ("conta_id", conta_id),
            ("conta_nome", conta_nome),
            ("dre_linha_id", dre_linha_id),
        )
        if not value
    ]
    if missing:
        raise PaymentRulesError(f"{regra_id}: campos obrigatorios ausentes: {missing}.")
    if not isinstance(entra_dre, bool):
        raise PaymentRulesError(f"{regra_id}: entra_dre deve ser booleano.")
    if not isinstance(match_data, dict):
        raise PaymentRulesError(f"{regra_id}: match deve ser um objeto.")
    if match_data.get("modo") != MATCH_MODE:
        raise PaymentRulesError(f"{regra_id}: match.modo deve ser {MATCH_MODE}.")

    match: dict[str, tuple[str, ...]] = {}
    for condition in CONDITION_FIELDS:
        aliases = tuple(
            normalized
            for item in _as_list(match_data.get(condition))
            if (normalized := normalize_match_text(item))
        )
        match[condition] = aliases

    if not any(match.values()):
        raise PaymentRulesError(f"{regra_id}: ao menos uma condicao de match e obrigatoria.")

    return PaymentRule(
        regra_id=regra_id,
        conta_id=conta_id,
        conta_nome=conta_nome,
        dre_linha_id=dre_linha_id,
        entra_dre=entra_dre,
        match=match,
    )


def _matches_rule(receipt: dict[str, str], rule: PaymentRule) -> bool:
    active_conditions = {
        condition: aliases
        for condition, aliases in rule.match.items()
        if aliases
    }
    for condition, aliases in active_conditions.items():
        field_text = normalize_match_text(_first_value(receipt, CONDITION_FIELDS[condition]))
        if not field_text:
            return False
        if not any(alias in field_text for alias in aliases):
            return False
    return True


def _first_value(row: dict[str, str], keys: tuple[str, ...]) -> str:
    lowered = {key.lower(): value for key, value in row.items()}
    for key in keys:
        value = lowered.get(key.lower())
        if value:
            return value
    return ""


def _as_list(value: object) -> list[object]:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


def _load_accounts(path: Path) -> dict[str, dict[str, object]]:
    with path.open(encoding="utf-8") as file:
        data = yaml.safe_load(file) or {}
    return {str(item.get("id")): item for item in data.get("contas", [])}


def _load_dre_mapping(path: Path) -> dict[str, dict[str, object]]:
    with path.open(encoding="utf-8") as file:
        data = yaml.safe_load(file) or {}
    return {str(key): value for key, value in data.get("mapeamento", {}).items()}


def _read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        sample = file.read(4096)
        file.seek(0)
        reader = csv.DictReader(file, delimiter=sniff_delimiter(sample))
        headers = list(reader.fieldnames or [])
        rows = [
            {header: (row.get(header) or "").strip() for header in headers}
            for row in reader
        ]
    return headers, rows


def _write_csv(
    path: Path,
    headers: list[str],
    rows: list[dict[str, object]],
) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=headers, delimiter=";")
        writer.writeheader()
        for row in rows:
            writer.writerow({header: _csv_value(row.get(header, "")) for header in headers})


def _result_row(result: PaymentClassificationResult) -> dict[str, object]:
    data = asdict(result)
    data["entra_dre"] = "" if result.entra_dre is None else str(result.entra_dre).lower()
    data["regras_candidatas"] = "|".join(result.regras_candidatas)
    return data


def _write_manifest(manifest: PaymentClassificationManifest) -> None:
    payload = {
        "status": manifest.status,
        "input_files": manifest.input_files,
        "receipt_count": manifest.receipt_count,
        "classified_count": manifest.classified_count,
        "review_without_rule_count": manifest.review_without_rule_count,
        "review_multiple_rules_count": manifest.review_multiple_rules_count,
        "generated_files": manifest.generated_files,
        "matching_executed": False,
        "classification_executed": True,
        "nao_inferir_conta": True,
        "nao_usar_modelo_para_adivinhar": True,
    }
    path = manifest.output_dir / "manifest_classificacao_comprovantes.json"
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def _csv_value(value: object) -> str:
    if value is None:
        return ""
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)
