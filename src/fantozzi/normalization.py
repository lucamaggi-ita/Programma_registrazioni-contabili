import re
import unicodedata


def normalize(text: str) -> str:
    text = text.lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = re.sub(r'[-/.,;:!()"]', " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_cnpj(value: str) -> str:
    """Strip punctuation, keeping only digits (e.g. "42.870.341/0001-68" -> "42870341000168")."""
    return re.sub(r"\D", "", value or "")
