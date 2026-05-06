from __future__ import annotations

from pathlib import Path
from typing import Dict, List

import yaml

from fantozzi.models import Conta


REVIEW_ACCOUNT_MAP: Dict[str, str] = {
    "NAO_CLASSIFICADO": "13.01.001",
    "AMBIGUO": "13.01.002",
    "VALOR_INCONSISTENTE": "13.01.006",
    "POSSIVEL_INVESTIMENTO": "13.01.009",
    "PESSOAL_A_REGULARIZAR": "13.01.010",
    "OUTROS": "13.01.999",
}


class ChartOfAccounts:
    def __init__(self) -> None:
        self._accounts: Dict[str, Conta] = {}
        self._ordered: List[Conta] = []  # preserves insertion order for duplicate detection

    @classmethod
    def load(cls, yaml_path: Path | str) -> ChartOfAccounts:
        chart = cls()
        with open(yaml_path, encoding="utf-8") as f:
            data = yaml.safe_load(f)
        for item in data.get("contas", []):
            conta = Conta(
                id=item["id"],
                nome=item["nome"],
                grupo=item["grupo"],
                tipo=item["tipo"],
                exige_revisao=item.get("exige_revisao", False),
            )
            chart._ordered.append(conta)
            chart._accounts[conta.id] = conta
        return chart

    def get_by_id(self, account_id: str) -> Conta:
        if account_id not in self._accounts:
            raise KeyError(f"Conta não encontrada: {account_id}")
        return self._accounts[account_id]

    def get_review_account(self, reason: str) -> Conta:
        account_id = REVIEW_ACCOUNT_MAP.get(reason, REVIEW_ACCOUNT_MAP["OUTROS"])
        return self.get_by_id(account_id)

    def validate(self) -> List[str]:
        errors: List[str] = []
        seen_ids: set[str] = set()

        for conta in self._ordered:
            if conta.id in seen_ids:
                errors.append(f"ID duplicado: {conta.id}")
            seen_ids.add(conta.id)

            for field_name in ("id", "nome", "grupo", "tipo"):
                if not getattr(conta, field_name):
                    errors.append(f"Campo '{field_name}' ausente em conta {conta.id}")

        for reason, review_id in REVIEW_ACCOUNT_MAP.items():
            if review_id not in self._accounts:
                errors.append(
                    f"Conta de revisão obrigatória ausente: {review_id} ({reason})"
                )

        return errors
