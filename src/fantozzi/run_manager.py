from __future__ import annotations

import hashlib
import json
import subprocess
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]

_NF_NS = "http://www.portalfiscal.inf.br/nfe"

STATUS_INICIADA = "INICIADA"
STATUS_DRY_RUN = "DRY_RUN"
STATUS_ABORTADA = "ABORTADA"
STATUS_ERRO = "ERRO"


class RunManagerError(Exception):
    pass


@dataclass(frozen=True)
class RunResult:
    run_id: str
    output_run_dir: Path
    reports_run_dir: Path
    manifest_path: Path
    current_run_output: Path
    current_run_reports: Path
    dry_run: bool


def generate_run_id(
    saida_output: Path,
    saida_reports: Path,
    today: date | None = None,
) -> str:
    """Gera o proximo run_id consultando AMBOS saida_output/runs/ e saida_reports/runs/.

    Evita colisoes e dessincronizacao entre as duas arvores de saida.
    Formato: YYYY-MM-DD_NNN (contador zero-padded com 3 digitos).
    """
    if today is None:
        today = date.today()
    date_str = today.strftime("%Y-%m-%d")
    prefix = date_str + "_"
    max_counter = 0
    for base in (saida_output, saida_reports):
        runs_dir = base / "runs"
        if not runs_dir.exists() or not runs_dir.is_dir():
            continue
        for entry in runs_dir.iterdir():
            if not entry.is_dir() or not entry.name.startswith(prefix):
                continue
            suffix = entry.name[len(prefix):]
            try:
                counter = int(suffix)
                if counter > max_counter:
                    max_counter = counter
            except ValueError:
                pass
    return f"{date_str}_{max_counter + 1:03d}"


def preparar_run(
    xml_base: Path,
    saida_output: Path,
    saida_reports: Path,
    config_paths: list[Path] | None = None,
    dry_run: bool = False,
    today: date | None = None,
    project_root: Path | None = None,
) -> RunResult:
    """Prepara a infraestrutura de uma nova run: dirs, current_run.txt, manifesto.

    Nao executa importacao, DRE, pre-livro nem qualquer reprocessamento.
    Com dry_run=True, calcula run_id e caminhos mas nao escreve nada.
    """
    _project_root = (project_root or PROJECT_ROOT).resolve()
    _validate_paths(xml_base, saida_output, saida_reports, _project_root)

    run_id = generate_run_id(saida_output, saida_reports, today=today)
    output_run_dir = saida_output / "runs" / run_id
    reports_run_dir = saida_reports / "runs" / run_id
    current_run_output = saida_output / "current_run.txt"
    current_run_reports = saida_reports / "current_run.txt"
    manifest_path = reports_run_dir / "manifest_run.json"

    result = RunResult(
        run_id=run_id,
        output_run_dir=output_run_dir,
        reports_run_dir=reports_run_dir,
        manifest_path=manifest_path,
        current_run_output=current_run_output,
        current_run_reports=current_run_reports,
        dry_run=dry_run,
    )

    if dry_run:
        return result

    output_run_dir.mkdir(parents=True, exist_ok=True)
    reports_run_dir.mkdir(parents=True, exist_ok=True)
    saida_output.mkdir(parents=True, exist_ok=True)
    saida_reports.mkdir(parents=True, exist_ok=True)

    current_run_output.write_text(run_id + "\n", encoding="utf-8")
    current_run_reports.write_text(run_id + "\n", encoding="utf-8")

    manifest = _build_manifest(
        run_id=run_id,
        xml_base=xml_base,
        saida_output=saida_output,
        saida_reports=saida_reports,
        config_paths=config_paths or _default_config_paths(),
        dry_run=False,
    )
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    return result


def add_checkpoint(manifest_path: Path, descricao: str) -> None:
    data = json.loads(manifest_path.read_text(encoding="utf-8"))
    data["checkpoints"].append({
        "momento": datetime.now().isoformat(timespec="seconds"),
        "descricao": descricao,
    })
    manifest_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def update_manifest_status(manifest_path: Path, status: str) -> None:
    data = json.loads(manifest_path.read_text(encoding="utf-8"))
    data["status_run"] = status
    manifest_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# ---------------------------------------------------------------------------
# Internos
# ---------------------------------------------------------------------------

def _validate_paths(
    xml_base: Path,
    saida_output: Path,
    saida_reports: Path,
    project_root: Path,
) -> None:
    xml_r = xml_base.resolve()
    out_r = saida_output.resolve()
    rep_r = saida_reports.resolve()

    for path, label in ((xml_r, "xml_base"), (out_r, "saida_output"), (rep_r, "saida_reports")):
        if not _is_subpath_or_equal(path, project_root):
            raise RunManagerError(
                f"{label} esta fora do project root: {path}"
            )

    if _is_subpath_or_equal(xml_r, out_r):
        raise RunManagerError(
            f"xml_base nao pode estar dentro de saida_output: {xml_r}"
        )
    if _is_subpath_or_equal(out_r, xml_r):
        raise RunManagerError(
            f"saida_output nao pode estar dentro de xml_base: {out_r}"
        )
    if _is_subpath_or_equal(xml_r, rep_r):
        raise RunManagerError(
            f"xml_base nao pode estar dentro de saida_reports: {xml_r}"
        )
    if _is_subpath_or_equal(rep_r, xml_r):
        raise RunManagerError(
            f"saida_reports nao pode estar dentro de xml_base: {rep_r}"
        )


def _is_subpath_or_equal(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


def _build_manifest(
    run_id: str,
    xml_base: Path,
    saida_output: Path,
    saida_reports: Path,
    config_paths: list[Path],
    dry_run: bool,
) -> dict:
    status = STATUS_DRY_RUN if dry_run else STATUS_INICIADA
    config_usado = {p.stem: _config_entry(p) for p in config_paths}
    fontes_xml = _collect_xml_fontes(xml_base)
    return {
        "run_id": run_id,
        "data_execucao": datetime.now().isoformat(timespec="seconds"),
        "status_run": status,
        "modo": "provisorio",
        "dry_run": dry_run,
        "git": _get_git_info(),
        "parametros": {
            "xml_base": str(xml_base),
            "saida_output": str(saida_output),
            "saida_reports": str(saida_reports),
            "limpar_saida": False,
        },
        "config_usado": config_usado,
        "fontes_xml": fontes_xml,
        "totais": {
            "xml_processados": None,
            "lancamentos_gerados": None,
            "lancamentos_ok": None,
            "lancamentos_revisar": None,
            "operacoes_especiais": None,
            "valor_total_importado": None,
            "valor_total_incluido_ok": None,
        },
        "classificacao": {
            "fase9_executada": False,
            "regras_aplicadas": 0,
            "candidatos_manter_revisao": None,
            "nota": "Revisao contabil pendente. Nenhuma regra aprovada aplicada.",
        },
        "relatorios_gerados": [],
        "checkpoints": [],
        "avisos": [],
        "erros": [],
    }


def _default_config_paths() -> list[Path]:
    config_dir = PROJECT_ROOT / "config"
    return [
        config_dir / "plano_contas_gerencial.yaml",
        config_dir / "regras_classificacao_produtos.yaml",
        config_dir / "mapeamento_dre.yaml",
    ]


def _config_entry(path: Path) -> dict:
    if not path.exists():
        return {
            "path": str(path),
            "hash_sha256": None,
            "data_modificacao": None,
            "nota": "arquivo nao encontrado",
        }
    mtime = datetime.fromtimestamp(path.stat().st_mtime).isoformat(timespec="seconds")
    return {
        "path": str(path),
        "hash_sha256": _hash_file(path),
        "data_modificacao": mtime,
    }


def _hash_file(path: Path) -> str:
    h = hashlib.sha256()
    h.update(path.read_bytes())
    return h.hexdigest()


def _collect_xml_fontes(xml_base: Path) -> list[dict]:
    fontes: list[dict] = []
    if not xml_base.exists() or not xml_base.is_dir():
        return fontes
    for subdir in sorted(xml_base.iterdir(), key=lambda p: p.name):
        if not subdir.is_dir():
            continue
        cnpj_pasta = subdir.name
        xml_files = sorted(subdir.rglob("*.xml"))
        xml_count = len(xml_files)
        cnpj_real, origem, nota = _resolve_cnpj(cnpj_pasta, xml_files)
        fonte: dict = {
            "pasta": str(subdir),
            "cnpj_pasta": cnpj_pasta,
            "cnpj_real": cnpj_real,
            "origem_cnpj_real": origem,
            "xml_encontrados": xml_count,
        }
        if nota:
            fonte["nota"] = nota
        fontes.append(fonte)
    return fontes


def _resolve_cnpj(
    cnpj_pasta: str,
    xml_files: list[Path],
) -> tuple[str, str, str]:
    for xml_path in xml_files[:3]:
        cnpj = _ler_cnpj_dest_xml(xml_path)
        if cnpj:
            nota = ""
            if cnpj != cnpj_pasta:
                nota = (
                    f"pasta nomeada sem zero inicial — "
                    f"CNPJ real tem {len(cnpj)} digitos"
                )
            return cnpj, "lido_do_xml", nota

    cnpj_padded = cnpj_pasta.zfill(14)
    nota = ""
    if cnpj_padded != cnpj_pasta:
        nota = "CNPJ real inferido por zero-padding — leitura XML nao disponivel"
    return cnpj_padded, "inferido_da_pasta", nota


def _ler_cnpj_dest_xml(xml_path: Path) -> str:
    try:
        root = ET.parse(xml_path).getroot()
        dest = root.find(f".//{{{_NF_NS}}}dest")
        if dest is not None:
            cnpj_elem = dest.find(f"{{{_NF_NS}}}CNPJ")
            if cnpj_elem is not None and cnpj_elem.text:
                return cnpj_elem.text.strip()
    except Exception:
        pass
    return ""


def _get_git_info() -> dict:
    try:
        commit = subprocess.check_output(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=PROJECT_ROOT,
            text=True,
            stderr=subprocess.DEVNULL,
        ).strip()
        branch = subprocess.check_output(
            ["git", "rev-parse", "--abbrev-ref", "HEAD"],
            cwd=PROJECT_ROOT,
            text=True,
            stderr=subprocess.DEVNULL,
        ).strip()
        dirty_out = subprocess.check_output(
            ["git", "status", "--porcelain"],
            cwd=PROJECT_ROOT,
            text=True,
            stderr=subprocess.DEVNULL,
        )
        return {"commit": commit, "branch": branch, "dirty": bool(dirty_out.strip())}
    except Exception:
        return {"commit": "desconhecido", "branch": "desconhecido", "dirty": None}
