from __future__ import annotations

from pathlib import Path

_EXCLUDED_DIR_NAMES: frozenset[str] = frozenset({
    "runs", "archive", "archives", "backup", "backups",
    "old", "tmp", "temp", "__pycache__",
})


class PipelinePathsError(Exception):
    pass


def resolver_arquivos_lancamentos_xlsx(entrada: Path | str) -> tuple[Path, ...]:
    """Retorna xlsx de lancamentos sob entrada, excluindo dirs de infraestrutura.

    Se entrada e um arquivo, valida que e .xlsx e retorna-o.
    Se entrada e um diretorio, recursa nele mas pula subdiretórios cujo nome
    (case-insensitive) esteja em _EXCLUDED_DIR_NAMES.
    O ponto de entrada em si nunca e excluido — passar
    runs/2026-05-06_001/ diretamente funciona corretamente.
    """
    entrada = Path(entrada)
    if not entrada.exists():
        raise PipelinePathsError(f"Entrada nao encontrada: {entrada}")
    if entrada.is_file():
        if entrada.suffix.lower() != ".xlsx":
            raise PipelinePathsError(f"Entrada nao e um Excel .xlsx: {entrada}")
        if entrada.name.startswith("~$"):
            raise PipelinePathsError(f"Arquivo temporario nao permitido: {entrada}")
        return (entrada,)
    if entrada.is_dir():
        found = _find_xlsx_in_dir(entrada)
        if not found:
            raise PipelinePathsError(
                f"Nenhum Excel .xlsx encontrado em: {entrada}"
            )
        return tuple(found)
    raise PipelinePathsError(f"Entrada invalida: {entrada}")


def _find_xlsx_in_dir(directory: Path) -> list[Path]:
    result: list[Path] = []
    try:
        items = sorted(directory.iterdir(), key=lambda p: str(p).lower())
    except PermissionError:
        return result
    for item in items:
        if item.is_file():
            if item.suffix.lower() == ".xlsx" and not item.name.startswith("~$"):
                result.append(item)
        elif item.is_dir() and item.name.lower() not in _EXCLUDED_DIR_NAMES:
            result.extend(_find_xlsx_in_dir(item))
    return result
