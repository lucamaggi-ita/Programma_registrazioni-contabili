from __future__ import annotations

import csv
import json
import subprocess
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from fantozzi.banking.payment_rules import (
    STATUS_CLASSIFICADO,
    STATUS_MULTIPLAS_REGRAS,
    STATUS_SEM_REGRA,
    PaymentRule,
    classify_payment_receipt,
    load_payment_rules,
    normalize_match_text,
    validate_payment_rules,
)
from fantozzi.csv_utils import sniff_delimiter


PROJECT_ROOT = Path(__file__).resolve().parents[3]
RECEIPTS_FILE_NAME = "payment_receipts_normalized.csv"

CLASSIFIED_HEADERS = [
    "cnpj",
    "mes",
    "data_pagamento",
    "beneficiario_normalizado",
    "referencia",
    "valor",
    "arquivo_origem",
    "status_classificacao",
    "conta_id",
    "conta_nome",
    "dre_linha_id",
    "entra_dre",
    "regra_id",
    "motivo_classificacao",
    "regras_candidatas",
    "payment_id",
    "beneficiary_name",
    "beneficiary_document",
    "payment_type",
    "status_pagamento",
    "source_file",
    "raw_hash",
]

REVIEW_HEADERS = [
    "cnpj",
    "mes",
    "data_pagamento",
    "beneficiario_normalizado",
    "referencia",
    "valor",
    "arquivo_origem",
    "status_classificacao",
    "motivo_revisao",
    "regras_candidatas",
    "contas_candidatas",
    "DECISAO_OPERADOR",
    "CONTA_ID_APROVADA",
    "CRIAR_REGRA_FUTURA",
    "ALIAS_O_KEYWORD_APROVADO",
    "OBSERVACOES",
    "payment_id",
    "beneficiary_name",
    "beneficiary_document",
    "payment_type",
    "status_pagamento",
    "source_file",
    "raw_hash",
]

SUMMARY_ACCOUNT_HEADERS = [
    "conta_id",
    "conta_nome",
    "dre_linha_id",
    "entra_dre",
    "quantidade",
    "valor_total",
]

SUMMARY_DRE_HEADERS = [
    "dre_linha_id",
    "entra_dre",
    "quantidade",
    "valor_total",
]

SUMMARY_CNPJ_MONTH_HEADERS = [
    "cnpj",
    "mes",
    "total_comprovantes",
    "total_classificados",
    "total_a_revisar_sem_regra",
    "total_a_revisar_multiplas_regras",
    "valor_total",
    "valor_classificado",
    "valor_a_revisar",
]


class PaymentClassificationRunError(ValueError):
    pass


@dataclass(frozen=True)
class ReceiptSource:
    path: Path
    cnpj: str
    mes: str


@dataclass(frozen=True)
class PaymentClassificationRunManifest:
    run_id: str
    output_dir: Path
    total_comprovantes: int
    total_classificados: int
    total_a_revisar_sem_regra: int
    total_a_revisar_multiplas_regras: int
    total_erros: int
    cnpjs_processados: tuple[str, ...]
    meses_processados: tuple[str, ...]
    generated_files: dict[str, str]


def execute_payment_classification_run(
    *,
    entrada_raiz: Path,
    saida_raiz: Path,
    regras_path: Path,
    plano_contas_path: Path,
    mapeamento_dre_path: Path,
    run_id: str | None = None,
    today: date | None = None,
) -> PaymentClassificationRunManifest:
    sources = discover_payment_receipt_sources(entrada_raiz)
    if not sources:
        raise PaymentClassificationRunError(
            f"Nenhum {RECEIPTS_FILE_NAME} encontrado em {entrada_raiz}."
        )

    rules = load_payment_rules(regras_path)
    validate_payment_rules(
        rules,
        plano_contas_path=plano_contas_path,
        mapeamento_dre_path=mapeamento_dre_path,
    )
    rules_by_id = {rule.regra_id: rule for rule in rules}

    effective_run_id = run_id or _generate_run_id(saida_raiz, today=today)
    output_dir = saida_raiz / "runs" / effective_run_id
    output_dir.mkdir(parents=True, exist_ok=True)

    classified_rows: list[dict[str, object]] = []
    review_rows: list[dict[str, object]] = []
    summary_cnpj_mes: dict[tuple[str, str], dict[str, object]] = {}

    for source in sources:
        _, rows = _read_csv(source.path)
        for row in rows:
            result = classify_payment_receipt(row, rules)
            base = _base_output_row(row, source)
            amount = _parse_decimal(base["valor"])
            _update_cnpj_month_summary(summary_cnpj_mes, source, amount, result.status_classificacao)

            if result.status_classificacao == STATUS_CLASSIFICADO:
                classified_rows.append(
                    {
                        **base,
                        "status_classificacao": result.status_classificacao,
                        "conta_id": result.conta_id,
                        "conta_nome": result.conta_nome,
                        "dre_linha_id": result.dre_linha_id,
                        "entra_dre": str(result.entra_dre).lower(),
                        "regra_id": result.regra_id,
                        "motivo_classificacao": result.motivo,
                        "regras_candidatas": "|".join(result.regras_candidatas),
                    }
                )
            else:
                review_rows.append(
                    {
                        **base,
                        "status_classificacao": result.status_classificacao,
                        "motivo_revisao": result.motivo,
                        "regras_candidatas": "|".join(result.regras_candidatas),
                        "contas_candidatas": _candidate_accounts(
                            result.regras_candidatas,
                            rules_by_id,
                        ),
                        "DECISAO_OPERADOR": "",
                        "CONTA_ID_APROVADA": "",
                        "CRIAR_REGRA_FUTURA": "",
                        "ALIAS_O_KEYWORD_APROVADO": "",
                        "OBSERVACOES": "",
                    }
                )

    summary_by_account = _summarize_classified(
        classified_rows,
        ("conta_id", "conta_nome", "dre_linha_id", "entra_dre"),
        SUMMARY_ACCOUNT_HEADERS,
    )
    summary_by_dre = _summarize_classified(
        classified_rows,
        ("dre_linha_id", "entra_dre"),
        SUMMARY_DRE_HEADERS,
    )
    summary_by_cnpj_month = _summary_cnpj_month_rows(summary_cnpj_mes)

    classified_path = output_dir / "comprovantes_classificados.csv"
    review_path = output_dir / "comprovantes_a_revisar.csv"
    summary_account_path = output_dir / "resumo_por_conta.csv"
    summary_dre_path = output_dir / "resumo_por_dre.csv"
    summary_cnpj_month_path = output_dir / "resumo_por_cnpj_mes.csv"
    manifest_path = output_dir / "manifest_classificacao_comprovantes.json"

    _write_csv(classified_path, CLASSIFIED_HEADERS, classified_rows)
    _write_csv(review_path, REVIEW_HEADERS, review_rows)
    _write_csv(summary_account_path, SUMMARY_ACCOUNT_HEADERS, summary_by_account)
    _write_csv(summary_dre_path, SUMMARY_DRE_HEADERS, summary_by_dre)
    _write_csv(summary_cnpj_month_path, SUMMARY_CNPJ_MONTH_HEADERS, summary_by_cnpj_month)

    generated_files = {
        "manifest": str(manifest_path),
        "comprovantes_classificados": str(classified_path),
        "comprovantes_a_revisar": str(review_path),
        "resumo_por_conta": str(summary_account_path),
        "resumo_por_dre": str(summary_dre_path),
        "resumo_por_cnpj_mes": str(summary_cnpj_month_path),
    }
    total_sem_regra = sum(
        1 for row in review_rows if row["status_classificacao"] == STATUS_SEM_REGRA
    )
    total_multiplas = sum(
        1 for row in review_rows if row["status_classificacao"] == STATUS_MULTIPLAS_REGRAS
    )
    manifest = PaymentClassificationRunManifest(
        run_id=effective_run_id,
        output_dir=output_dir,
        total_comprovantes=len(classified_rows) + len(review_rows),
        total_classificados=len(classified_rows),
        total_a_revisar_sem_regra=total_sem_regra,
        total_a_revisar_multiplas_regras=total_multiplas,
        total_erros=0,
        cnpjs_processados=tuple(sorted({source.cnpj for source in sources})),
        meses_processados=tuple(sorted({source.mes for source in sources})),
        generated_files=generated_files,
    )
    _write_manifest(
        manifest_path,
        manifest,
        fontes_lidas=[source.path for source in sources],
        regras_path=regras_path,
        plano_contas_path=plano_contas_path,
        mapeamento_dre_path=mapeamento_dre_path,
    )
    return manifest


def discover_payment_receipt_sources(entrada_raiz: Path) -> list[ReceiptSource]:
    if entrada_raiz.is_file():
        if entrada_raiz.name != RECEIPTS_FILE_NAME:
            raise PaymentClassificationRunError(
                f"Arquivo de entrada deve chamar {RECEIPTS_FILE_NAME}: {entrada_raiz}"
            )
        return [_source_from_path(entrada_raiz)]

    if not entrada_raiz.exists():
        raise PaymentClassificationRunError(f"Pasta de entrada nao encontrada: {entrada_raiz}")

    sources = [
        _source_from_path(path)
        for path in sorted(entrada_raiz.rglob(RECEIPTS_FILE_NAME), key=lambda item: str(item))
    ]
    return sources


def _source_from_path(path: Path) -> ReceiptSource:
    cnpj = path.parent.name
    mes = _month_from_parts(path.parts) or ""
    return ReceiptSource(path=path, cnpj=cnpj, mes=mes)


def _month_from_parts(parts: tuple[str, ...]) -> str:
    for part in reversed(parts):
        if len(part) == 7 and part[4] == "-" and part[:4].isdigit() and part[5:].isdigit():
            return part
    return ""


def _base_output_row(row: dict[str, str], source: ReceiptSource) -> dict[str, object]:
    payment_date = row.get("payment_date", "")
    mes = source.mes or payment_date[:7]
    beneficiary = row.get("beneficiary_name", "")
    return {
        "cnpj": row.get("cnpj") or source.cnpj,
        "mes": mes,
        "data_pagamento": payment_date,
        "beneficiario_normalizado": normalize_match_text(beneficiary),
        "referencia": row.get("reference_id", ""),
        "valor": row.get("amount", ""),
        "arquivo_origem": str(source.path),
        "payment_id": row.get("payment_id", ""),
        "beneficiary_name": beneficiary,
        "beneficiary_document": row.get("beneficiary_document", ""),
        "payment_type": row.get("payment_type", ""),
        "status_pagamento": row.get("status", ""),
        "source_file": row.get("source_file", ""),
        "raw_hash": row.get("raw_hash", ""),
    }


def _update_cnpj_month_summary(
    summary: dict[tuple[str, str], dict[str, object]],
    source: ReceiptSource,
    amount: Decimal,
    status_classificacao: str,
) -> None:
    key = (source.cnpj, source.mes)
    row = summary.setdefault(
        key,
        {
            "cnpj": source.cnpj,
            "mes": source.mes,
            "total_comprovantes": 0,
            "total_classificados": 0,
            "total_a_revisar_sem_regra": 0,
            "total_a_revisar_multiplas_regras": 0,
            "valor_total": Decimal("0"),
            "valor_classificado": Decimal("0"),
            "valor_a_revisar": Decimal("0"),
        },
    )
    row["total_comprovantes"] = int(row["total_comprovantes"]) + 1
    row["valor_total"] = row["valor_total"] + amount
    if status_classificacao == STATUS_CLASSIFICADO:
        row["total_classificados"] = int(row["total_classificados"]) + 1
        row["valor_classificado"] = row["valor_classificado"] + amount
    else:
        row["valor_a_revisar"] = row["valor_a_revisar"] + amount
        if status_classificacao == STATUS_SEM_REGRA:
            row["total_a_revisar_sem_regra"] = int(row["total_a_revisar_sem_regra"]) + 1
        else:
            row["total_a_revisar_multiplas_regras"] = (
                int(row["total_a_revisar_multiplas_regras"]) + 1
            )


def _summarize_classified(
    rows: list[dict[str, object]],
    key_fields: tuple[str, ...],
    output_headers: list[str],
) -> list[dict[str, object]]:
    summary: dict[tuple[str, ...], dict[str, object]] = {}
    for row in rows:
        key = tuple(str(row.get(field, "")) for field in key_fields)
        target = summary.setdefault(
            key,
            {field: row.get(field, "") for field in key_fields}
            | {"quantidade": 0, "valor_total": Decimal("0")},
        )
        target["quantidade"] = int(target["quantidade"]) + 1
        target["valor_total"] = target["valor_total"] + _parse_decimal(row.get("valor", ""))
    return [
        {header: _summary_value(row.get(header, "")) for header in output_headers}
        for row in sorted(summary.values(), key=lambda item: tuple(str(item.get(f, "")) for f in key_fields))
    ]


def _summary_cnpj_month_rows(
    summary: dict[tuple[str, str], dict[str, object]],
) -> list[dict[str, object]]:
    return [
        {header: _summary_value(row.get(header, "")) for header in SUMMARY_CNPJ_MONTH_HEADERS}
        for row in sorted(summary.values(), key=lambda item: (str(item["mes"]), str(item["cnpj"])))
    ]


def _candidate_accounts(
    regra_ids: tuple[str, ...],
    rules_by_id: dict[str, PaymentRule],
) -> str:
    accounts = {
        rules_by_id[regra_id].conta_id
        for regra_id in regra_ids
        if regra_id in rules_by_id
    }
    return "|".join(sorted(accounts))


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


def _write_csv(path: Path, headers: list[str], rows: list[dict[str, object]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=headers, delimiter=";")
        writer.writeheader()
        for row in rows:
            writer.writerow({header: _csv_value(row.get(header, "")) for header in headers})


def _write_manifest(
    path: Path,
    manifest: PaymentClassificationRunManifest,
    *,
    fontes_lidas: list[Path],
    regras_path: Path,
    plano_contas_path: Path,
    mapeamento_dre_path: Path,
) -> None:
    git = _git_info()
    payload = {
        "run_id": manifest.run_id,
        "data_execucao": datetime.now().isoformat(timespec="seconds"),
        "git_commit": git["commit"],
        "git_dirty": git["dirty"],
        "modo": "provisorio",
        "fase": "13S",
        "fontes_lidas": [str(path) for path in fontes_lidas],
        "cnpjs_processados": list(manifest.cnpjs_processados),
        "meses_processados": list(manifest.meses_processados),
        "total_comprovantes": manifest.total_comprovantes,
        "total_classificados": manifest.total_classificados,
        "total_a_revisar_sem_regra": manifest.total_a_revisar_sem_regra,
        "total_a_revisar_multiplas_regras": manifest.total_a_revisar_multiplas_regras,
        "total_erros": manifest.total_erros,
        "regras_yaml_usado": str(regras_path),
        "plano_contas_usado": str(plano_contas_path),
        "mapeamento_dre_usado": str(mapeamento_dre_path),
        "matching_automatico_executado": False,
        "classification_executed": True,
        "nao_inferir_conta": True,
        "generated_files": manifest.generated_files,
    }
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def _generate_run_id(saida_raiz: Path, today: date | None = None) -> str:
    today = today or date.today()
    date_prefix = today.strftime("%Y-%m-%d")
    prefix = date_prefix + "_"
    max_counter = 0
    runs_dir = saida_raiz / "runs"
    if runs_dir.exists():
        for entry in runs_dir.iterdir():
            if not entry.is_dir() or not entry.name.startswith(prefix):
                continue
            try:
                max_counter = max(max_counter, int(entry.name[len(prefix):]))
            except ValueError:
                pass
    return f"{date_prefix}_{max_counter + 1:03d}"


def _git_info() -> dict[str, object]:
    try:
        commit = subprocess.check_output(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=PROJECT_ROOT,
            text=True,
            stderr=subprocess.DEVNULL,
        ).strip()
        dirty_out = subprocess.check_output(
            ["git", "status", "--porcelain"],
            cwd=PROJECT_ROOT,
            text=True,
            stderr=subprocess.DEVNULL,
        )
        return {"commit": commit, "dirty": bool(dirty_out.strip())}
    except Exception:
        return {"commit": "desconhecido", "dirty": None}


def _parse_decimal(value: object) -> Decimal:
    text = str(value or "").strip()
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    if not text:
        return Decimal("0")
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal("0")


def _summary_value(value: object) -> str:
    if isinstance(value, Decimal):
        return f"{value:.2f}"
    return str(value)


def _csv_value(value: object) -> str:
    if isinstance(value, Decimal):
        return f"{value:.2f}"
    if value is None:
        return ""
    return str(value)
