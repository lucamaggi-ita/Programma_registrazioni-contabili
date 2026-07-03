from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Tuple

import yaml

from fantozzi.chart_of_accounts import ChartOfAccounts
from fantozzi.models import ClassificationResult
from fantozzi.normalization import normalize


class ProductClassifier:
    def __init__(self, chart: ChartOfAccounts) -> None:
        self._chart = chart
        self._rules: List[Dict[str, Any]] = []
        self._supplier_rules: List[Dict[str, Any]] = []
        self._fallback_id: str = "13.01.001"

    def load(self, yaml_path: Path | str) -> ProductClassifier:
        with open(yaml_path, encoding="utf-8") as f:
            data = yaml.safe_load(f)

        self._fallback_id = data.get("conta_fallback_id", self._fallback_id)

        for rule in data.get("regras_por_fornecedor", []):
            supplier_aliases = list(rule.get("fornecedor_aliases", []))
            if rule.get("fornecedor_contem"):
                supplier_aliases.append(rule["fornecedor_contem"])

            product_aliases = list(rule.get("produto_aliases", []))
            if rule.get("produto_normalizado_contem"):
                product_aliases.append(rule["produto_normalizado_contem"])
            if rule.get("produto_contem"):
                product_aliases.append(rule["produto_contem"])

            conta_id = rule.get("conta_id") or _parse_conta_id(rule.get("conta", ""))
            if not conta_id:
                continue

            supplier_norm = [normalize(a) for a in supplier_aliases]
            product_norm = [normalize(a) for a in product_aliases]
            self._supplier_rules.append(
                {
                    "conta_id": conta_id,
                    "fornecedor_norm": [a for a in supplier_norm if a],
                    "produto_norm": [a for a in product_norm if a],
                }
            )

        for rule in data.get("regras", []):
            aliases_norm = [normalize(a) for a in rule.get("aliases", [])]
            evitar_norm = [normalize(a) for a in rule.get("evitar_aliases", [])]
            self._rules.append(
                {
                    "conta_id": rule["conta_id"],
                    "conta_nome": rule.get("conta_nome", ""),
                    "prioridade": rule.get("prioridade", 0),
                    "aliases_norm": [a for a in aliases_norm if a],
                    "evitar_norm": [a for a in evitar_norm if a],
                }
            )

        return self

    def classify(self, desc: str, fornecedor: str | None = None) -> ClassificationResult:
        product_result = self._classify_by_product(desc)
        if product_result.confianca == "OK":
            return product_result

        supplier_result = self._classify_by_supplier(desc, fornecedor)
        return supplier_result or product_result

    def _classify_by_product(self, desc: str) -> ClassificationResult:
        norm_desc = normalize(desc)

        hits: List[Tuple[Dict[str, Any], str]] = []

        for rule in self._rules:
            if any(ev in norm_desc for ev in rule["evitar_norm"]):
                continue
            for alias in rule["aliases_norm"]:
                if _alias_matches(alias, norm_desc):
                    hits.append((rule, alias))

        if not hits:
            conta = self._chart.get_by_id(self._fallback_id)
            return ClassificationResult(conta=conta, confianca="NAO_CLASSIFICADO")

        conta_ids = {h[0]["conta_id"] for h in hits}

        if len(conta_ids) == 1:
            best_rule = max(hits, key=lambda h: h[0]["prioridade"])[0]
            conta = self._chart.get_by_id(best_rule["conta_id"])
            matched = list({h[1] for h in hits})
            return ClassificationResult(conta=conta, confianca="OK", matches=matched)

        best_rule = max(hits, key=lambda h: h[0]["prioridade"])[0]
        conta = self._chart.get_by_id(best_rule["conta_id"])
        matched = list({h[1] for h in hits})
        alertas = [f"Classificação ambígua: contas {conta_ids}"]
        return ClassificationResult(
            conta=conta, confianca="AMBIGUO", matches=matched, alertas=alertas
        )

    def _classify_by_supplier(
        self, desc: str, fornecedor: str | None
    ) -> ClassificationResult | None:
        if not fornecedor:
            return None

        norm_desc = normalize(desc)
        norm_supplier = normalize(fornecedor)

        for rule in self._supplier_rules:
            supplier_match = next(
                (
                    alias
                    for alias in rule["fornecedor_norm"]
                    if _alias_matches(alias, norm_supplier)
                ),
                "",
            )
            if not supplier_match:
                continue

            product_aliases = rule["produto_norm"]
            if product_aliases and not any(
                _alias_matches(alias, norm_desc) for alias in product_aliases
            ):
                continue

            conta = self._chart.get_by_id(rule["conta_id"])
            return ClassificationResult(
                conta=conta,
                confianca="OK",
                matches=[f"fornecedor:{supplier_match}"],
            )

        return None


def _parse_conta_id(conta: str) -> str:
    return conta.split("-", 1)[0].strip() if conta else ""


def _alias_matches(alias: str, text: str) -> bool:
    norm_alias = normalize(alias)
    if not norm_alias:
        return False

    norm_text = normalize(text)
    if len(norm_alias) <= 3:
        return f" {norm_alias} " in f" {norm_text} "

    return norm_alias in norm_text
