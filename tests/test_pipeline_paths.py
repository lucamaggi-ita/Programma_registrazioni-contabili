from pathlib import Path

from openpyxl import Workbook
import pytest

from fantozzi.pipeline_paths import PipelinePathsError, resolver_arquivos_lancamentos_xlsx


def _make_xlsx(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    Workbook().save(path)


# ---------------------------------------------------------------------------
# File singolo
# ---------------------------------------------------------------------------

def test_resolver_arquivo_unico_xlsx(tmp_path):
    xlsx = tmp_path / "lancamentos.xlsx"
    _make_xlsx(xlsx)

    result = resolver_arquivos_lancamentos_xlsx(xlsx)

    assert result == (xlsx,)


def test_resolver_arquivo_nao_xlsx_erro(tmp_path):
    txt = tmp_path / "arquivo.txt"
    txt.write_text("x")

    with pytest.raises(PipelinePathsError, match="nao e um Excel"):
        resolver_arquivos_lancamentos_xlsx(txt)


def test_resolver_arquivo_nao_existe_erro(tmp_path):
    with pytest.raises(PipelinePathsError, match="nao encontrada"):
        resolver_arquivos_lancamentos_xlsx(tmp_path / "inexistente.xlsx")


# ---------------------------------------------------------------------------
# Diretorio
# ---------------------------------------------------------------------------

def test_resolver_diretorio_sem_xlsx_erro(tmp_path):
    subdir = tmp_path / "vazio"
    subdir.mkdir()

    with pytest.raises(PipelinePathsError, match="Nenhum Excel"):
        resolver_arquivos_lancamentos_xlsx(subdir)


def test_resolver_nao_entra_em_runs(tmp_path):
    xlsx_ok = tmp_path / "55746722000181" / "lancamentos.xlsx"
    _make_xlsx(xlsx_ok)
    xlsx_runs = tmp_path / "runs" / "2026-05-06_001" / "55746722000181" / "lancamentos.xlsx"
    _make_xlsx(xlsx_runs)

    result = resolver_arquivos_lancamentos_xlsx(tmp_path)

    assert xlsx_ok in result
    assert xlsx_runs not in result


def test_resolver_nao_entra_em_archive(tmp_path):
    xlsx_ok = tmp_path / "dados" / "lancamentos.xlsx"
    _make_xlsx(xlsx_ok)
    xlsx_arq = tmp_path / "archive" / "lancamentos_old.xlsx"
    _make_xlsx(xlsx_arq)

    result = resolver_arquivos_lancamentos_xlsx(tmp_path)

    assert xlsx_ok in result
    assert xlsx_arq not in result


def test_resolver_nao_entra_em_backup(tmp_path):
    xlsx_ok = tmp_path / "dados" / "lancamentos.xlsx"
    _make_xlsx(xlsx_ok)
    xlsx_bkp = tmp_path / "backup" / "lancamentos_bkp.xlsx"
    _make_xlsx(xlsx_bkp)

    result = resolver_arquivos_lancamentos_xlsx(tmp_path)

    assert xlsx_ok in result
    assert xlsx_bkp not in result


def test_resolver_nao_entra_em_old(tmp_path):
    xlsx_ok = tmp_path / "dados" / "lancamentos.xlsx"
    _make_xlsx(xlsx_ok)
    xlsx_old = tmp_path / "old" / "lancamentos_old.xlsx"
    _make_xlsx(xlsx_old)

    result = resolver_arquivos_lancamentos_xlsx(tmp_path)

    assert xlsx_ok in result
    assert xlsx_old not in result


def test_resolver_run_especifica_le_apenas_aquela_run(tmp_path):
    run1 = tmp_path / "runs" / "2026-05-06_001"
    run2 = tmp_path / "runs" / "2026-05-06_002"
    xlsx1 = run1 / "cnpj" / "lancamentos.xlsx"
    xlsx2 = run2 / "cnpj" / "lancamentos.xlsx"
    _make_xlsx(xlsx1)
    _make_xlsx(xlsx2)

    result = resolver_arquivos_lancamentos_xlsx(run1)

    assert xlsx1 in result
    assert xlsx2 not in result


def test_resolver_duas_runs_nao_mistura_xlsx(tmp_path):
    xlsx_base = tmp_path / "55746722000181" / "lancamentos.xlsx"
    _make_xlsx(xlsx_base)
    xlsx_run1 = tmp_path / "runs" / "2026-05-06_001" / "55746722000181" / "lancamentos.xlsx"
    xlsx_run2 = tmp_path / "runs" / "2026-05-06_002" / "55746722000181" / "lancamentos.xlsx"
    _make_xlsx(xlsx_run1)
    _make_xlsx(xlsx_run2)

    result_base = resolver_arquivos_lancamentos_xlsx(tmp_path)
    assert xlsx_base in result_base
    assert xlsx_run1 not in result_base
    assert xlsx_run2 not in result_base

    result_run1 = resolver_arquivos_lancamentos_xlsx(tmp_path / "runs" / "2026-05-06_001")
    assert xlsx_run1 in result_run1
    assert xlsx_run2 not in result_run1


def test_resolver_retorna_ordem_deterministica(tmp_path):
    xlsx_a = tmp_path / "aaa" / "lancamentos.xlsx"
    xlsx_b = tmp_path / "bbb" / "lancamentos.xlsx"
    _make_xlsx(xlsx_a)
    _make_xlsx(xlsx_b)

    result1 = resolver_arquivos_lancamentos_xlsx(tmp_path)
    result2 = resolver_arquivos_lancamentos_xlsx(tmp_path)

    assert result1 == result2
    assert result1[0] == xlsx_a
    assert result1[1] == xlsx_b


def test_resolver_ignora_arquivos_temporarios_til(tmp_path):
    xlsx_ok = tmp_path / "lancamentos.xlsx"
    _make_xlsx(xlsx_ok)
    xlsx_tmp = tmp_path / "~$lancamentos.xlsx"
    _make_xlsx(xlsx_tmp)

    result = resolver_arquivos_lancamentos_xlsx(tmp_path)

    assert xlsx_ok in result
    assert xlsx_tmp not in result


def test_resolver_nao_exclui_run_especifica_por_causa_do_parent_runs(tmp_path):
    """Passar runs/2026-05-06_001/ diretamente deve funcionar, mesmo que
    o parent se chame 'runs'."""
    run_dir = tmp_path / "runs" / "2026-05-06_001" / "cnpj"
    xlsx = run_dir / "lancamentos.xlsx"
    _make_xlsx(xlsx)

    result = resolver_arquivos_lancamentos_xlsx(tmp_path / "runs" / "2026-05-06_001")

    assert len(result) == 1
    assert result[0] == xlsx


def test_resolver_exclui_dirs_case_insensitive(tmp_path):
    """RUNS, Archive, BACKUP devem ser excluidos independente do case."""
    xlsx_ok = tmp_path / "dados" / "lancamentos.xlsx"
    _make_xlsx(xlsx_ok)
    _make_xlsx(tmp_path / "RUNS" / "run1" / "l.xlsx")
    _make_xlsx(tmp_path / "Archive" / "l.xlsx")
    _make_xlsx(tmp_path / "BACKUP" / "l.xlsx")

    result = resolver_arquivos_lancamentos_xlsx(tmp_path)

    assert result == (xlsx_ok,)


def test_resolver_rejeita_pasta_agregadora_runs(tmp_path):
    """Passar runs/ diretamente deve falhar — misturaria todas as runs."""
    runs_dir = tmp_path / "runs"
    _make_xlsx(runs_dir / "2026-05-06_001" / "cnpj" / "l.xlsx")

    with pytest.raises(ValueError, match="run especifica"):
        resolver_arquivos_lancamentos_xlsx(runs_dir)


def test_resolver_rejeita_pasta_agregadora_runs_case_insensitive(tmp_path):
    """RUNS, Runs devem ser rejeitados da mesma forma que runs."""
    runs_upper = tmp_path / "RUNS"
    _make_xlsx(runs_upper / "2026-05-06_001" / "cnpj" / "l.xlsx")

    with pytest.raises(ValueError, match="run especifica"):
        resolver_arquivos_lancamentos_xlsx(runs_upper)
