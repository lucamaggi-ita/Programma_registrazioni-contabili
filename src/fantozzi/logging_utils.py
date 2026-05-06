from __future__ import annotations

from datetime import datetime


def info(message: str) -> None:
    _log("INFO", message)


def warning(message: str) -> None:
    _log("WARNING", message)


def error(message: str) -> None:
    _log("ERROR", message)


def _log(level: str, message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] [{level}] {message}")
