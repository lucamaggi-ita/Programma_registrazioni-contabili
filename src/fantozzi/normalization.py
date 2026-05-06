import re
import unicodedata


def normalize(text: str) -> str:
    text = text.lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = re.sub(r'[-/.,;:!()"]', " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text
