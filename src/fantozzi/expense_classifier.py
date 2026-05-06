from __future__ import annotations

from fantozzi.chart_of_accounts import ChartOfAccounts
from fantozzi.product_classifier import ProductClassifier


class ExpenseClassifier(ProductClassifier):
    """Classifies non-NF-e expenses (boletos, bank statements).

    Identical logic to ProductClassifier; loaded from a different YAML
    whose conta_fallback_id is "13.01.999" (Revisar - Outros Casos).
    """

    def __init__(self, chart: ChartOfAccounts) -> None:
        super().__init__(chart)
        self._fallback_id = "13.01.999"
