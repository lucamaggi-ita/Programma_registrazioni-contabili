from __future__ import annotations

import json
from pathlib import Path

from fantozzi.banking.models import BankImportManifest


def write_manifest(manifest: BankImportManifest) -> Path:
    path = manifest.output_dir / "manifest_bank_import.json"
    payload = {
        "cnpj": manifest.cnpj,
        "status": manifest.status,
        "avisos": manifest.avisos,
        "input_files": manifest.input_files,
        "optional_inputs_absent": manifest.optional_inputs_absent,
        "transaction_count": manifest.transaction_count,
        "payment_receipt_count": manifest.payment_receipt_count,
        "balance_audit_count": manifest.balance_audit_count,
        "filtered_rows": manifest.filtered_rows,
        "generated_files": manifest.generated_files,
        "matching_executed": False,
        "classification_executed": False,
    }
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return path
