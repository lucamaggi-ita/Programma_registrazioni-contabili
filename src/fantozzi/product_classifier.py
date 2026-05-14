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
        self._fallback_id: str = "13.01.001"

    def load(self, yaml_path: Path | str) -> ProductClassifier:
        with open(yaml_path, encoding="utf-8") as f:
            data = yaml.safe_load(f)

        self._fallback_id = data.get("conta_fallback_id", self._fallback_id)

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

    def classify(self, desc: str) -> ClassificationResult:
        norm_desc = normalize(desc)

        hits: List[Tuple[Dict[str, Any], str]] = []

        for rule in self._rules:
            if any(ev in norm_desc for ev in rule["evitar_norm"]):
                continue
            for alias in rule["aliases_norm"]:
                if alias in norm_desc:
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
