from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import hashlib
import re
import unicodedata


SALDO_RE = re.compile(
    r"\bSALDO\b|\bSUBTOTAL\b|\bTOTAL\s+DISPON[IÍ]VEL\b",
    flags=re.IGNORECASE,
)


def only_digits(value: object) -> str:
    return re.sub(r"\D+", "", text(value))


def text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_key(value: object) -> str:
    normalized = unicodedata.normalize("NFKD", text(value))
    ascii_text = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", "", ascii_text.lower())


def parse_decimal(value: object) -> Decimal:
    raw = text(value)
    if not raw:
        return Decimal("0")
    raw = raw.replace("R$", "").replace("\xa0", " ").strip()
    raw = re.sub(r"\s+", "", raw)
    if "," in raw and "." in raw:
        raw = raw.replace(".", "").replace(",", ".")
    elif "," in raw:
        raw = raw.replace(",", ".")
    try:
        return Decimal(raw)
    except InvalidOperation as exc:
        raise ValueError(f"Valor decimal invalido: {value!r}") from exc


def parse_date(value: object) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raw = text(value)
    if not raw:
        return None
    for fmt in ("%Y%m%d", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(raw[:10], fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Data invalida: {value!r}")


def stable_hash(*parts: object) -> str:
    payload = "|".join(text(part) for part in parts)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def is_balance_description(value: object) -> bool:
    return bool(SALDO_RE.search(text(value)))
