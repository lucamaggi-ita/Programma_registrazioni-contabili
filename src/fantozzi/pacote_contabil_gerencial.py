from __future__ import annotations

import json
import shutil
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Mapping

from openpyxl import load_workbook


PACOTE_DIR_NAME = "pacote_contabil_gerencial_provisorio"

INDICE_FILE = "00_INDICE_PACOTE_CONTABIL_GERENCIAL.txt"
DRE_XLSX_FILE = "01_DRE_GERENCIAL_CONSOLIDADO.xlsx"
DRE_TXT_FILE = "01_DRE_GERENCIAL_CONSOLIDADO.txt"
PRE_LIVRO_XLSX_FILE = "02_PRE_LIVRO_DIARIO_GERENCIAL.xlsx"
PRE_LIVRO_AUDIT_FILE = "02_AUDITORIA_PRE_LIVRO_DIARIO_GERENCIAL.txt"
RELATORIOS_XLSX_FILE = "03_RELATORIOS_GERENCIAIS_OK.xlsx"
RELATORIOS_TXT_FILE = "03_RELATORIOS_GERENCIAIS_OK.txt"
COMODATO_AUDIT_FILE = "04_AUDITORIA_COMODATO_FORA_DRE.txt"
BONIFICACAO_AUDIT_FILE = "05_AUDITORIA_BONIFICACAO_RECEBIDA_DRE.txt"
MANUAIS_AUDIT_FILE = "06_AUDITORIA_LANCAMENTOS_MANUAIS_GERENCIAIS.txt"
MANIFEST_FILE = "manifest_pacote_contabil_gerencial.json"

BASE_OK_DRE_SHEET = "Base_OK_DRE"
LANCAMENTOS_OK_SHEET = "Lancamentos_OK"
MANUAL_GERENCIAL = "MANUAL_GERENCIAL"

SOURCE_FILES = {
    DRE_XLSX_FILE: "dre_gerencial_ok_provisoria.xlsx",
    DRE_TXT_FILE: "dre_gerencial_ok_provisoria.txt",
    PRE_LIVRO_XLSX_FILE: "pre_livro_diario_gerencial_teste.xlsx",
    PRE_LIVRO_AUDIT_FILE: "auditoria_pre_livro_diario_gerencial_teste.txt",
    RELATORIOS_XLSX_FILE: "relatorios_gerenciais_ok_provisorios.xlsx",
    RELATORIOS_TXT_FILE: "relatorios_gerenciais_ok_provisorios.txt",
}


class PacoteContabilGerencialError(Exception):
    pass


@dataclass(frozen=True)
class PacoteContabilGerencialResult:
    pacote_dir: Path
    indice_path: Path
    manifest_path: Path
    arquivos_gerados: tuple[Path, ...]
    status: str


def gerar_pacote_contabil_gerencial_provisorio(
    reports_run_dir: Path | str,
) -> PacoteContabilGerencialResult:
    reports_run_dir = Path(reports_run_dir)
    if not reports_run_dir.exists() or not reports_run_dir.is_dir():
        raise PacoteContabilGerencialError(
            f"Diretorio da run nao encontrado: {reports_run_dir}"
        )

    run_manifest = _load_json(reports_run_dir / "manifest_run.json")
    run_id = _text(run_manifest.get("run_id")) or reports_run_dir.name
    pacote_dir = reports_run_dir / PACOTE_DIR_NAME
    pacote_dir.mkdir(parents=True, exist_ok=True)

    files_manifest: dict[str, dict[str, object]] = {}
    generated: list[Path] = []
    for output_name, source_name in SOURCE_FILES.items():
        output_path = pacote_dir / output_name
        source_path = reports_run_dir / source_name
        if source_path.exists() and source_path.is_file():
            shutil.copy2(source_path, output_path)
            status = "OK"
            generated.append(output_path)
        else:
            status = "NAO_GERADO"
        files_manifest[output_name] = {
            "status": status,
            "origem": str(source_path),
            "destino": str(output_path),
        }

    comodato_audit = _audit_comodato(reports_run_dir / SOURCE_FILES[DRE_XLSX_FILE])
    bonificacao_audit = _audit_bonificacao(reports_run_dir / SOURCE_FILES[DRE_XLSX_FILE])
    manuais_audit = _audit_manuais(
        reports_run_dir / SOURCE_FILES[PRE_LIVRO_XLSX_FILE],
        run_manifest.get("lancamentos_manuais_gerenciais", {"executado": False}),
    )

    audit_outputs = {
        COMODATO_AUDIT_FILE: comodato_audit,
        BONIFICACAO_AUDIT_FILE: bonificacao_audit,
        MANUAIS_AUDIT_FILE: manuais_audit,
    }
    for output_name, audit in audit_outputs.items():
        output_path = pacote_dir / output_name
        output_path.write_text(audit["texto"], encoding="utf-8", newline="\n")
        files_manifest[output_name] = {
            "status": audit["status"],
            "destino": str(output_path),
        }
        generated.append(output_path)

    status = (
        "OK"
        if all(
            row["status"] in {"OK", "NAO_APLICAVEL"}
            for row in files_manifest.values()
        )
        else "PARCIAL"
    )
    manifest = _build_manifest(
        run_manifest,
        run_id=run_id,
        status=status,
        files_manifest=files_manifest,
        comodato_audit=comodato_audit,
        bonificacao_audit=bonificacao_audit,
        manuais_audit=manuais_audit,
    )
    manifest_path = pacote_dir / MANIFEST_FILE
    manifest_path.write_text(
        json.dumps(_json_ready(manifest), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    files_manifest[MANIFEST_FILE] = {
        "status": "OK",
        "destino": str(manifest_path),
    }
    generated.append(manifest_path)

    indice_text = _indice_text(manifest, files_manifest)
    indice_path = pacote_dir / INDICE_FILE
    indice_path.write_text(indice_text, encoding="utf-8", newline="\n")
    files_manifest[INDICE_FILE] = {
        "status": "OK",
        "destino": str(indice_path),
    }
    generated.append(indice_path)

    return PacoteContabilGerencialResult(
        pacote_dir=pacote_dir,
        indice_path=indice_path,
        manifest_path=manifest_path,
        arquivos_gerados=tuple(generated),
        status=status,
    )


def _audit_comodato(dre_xlsx_path: Path) -> dict[str, object]:
    rows, read_status = _read_sheet_rows(dre_xlsx_path, BASE_OK_DRE_SHEET)
    if read_status != "OK":
        return {
            "status": read_status,
            "comodato_base_dre": None,
            "valor_total": None,
            "texto": _comodato_text(read_status, None, None),
        }
    comodato_rows = [
        row for row in rows if _text(row.get("tipo_operacao_especial")).upper() == "COMODATO"
    ]
    valor_total = _sum_decimal(comodato_rows, "valor")
    status = "OK" if not comodato_rows else "ATENCAO"
    return {
        "status": status,
        "comodato_base_dre": len(comodato_rows),
        "valor_total": valor_total,
        "texto": _comodato_text(status, len(comodato_rows), valor_total),
    }


def _audit_bonificacao(dre_xlsx_path: Path) -> dict[str, object]:
    rows, read_status = _read_sheet_rows(dre_xlsx_path, BASE_OK_DRE_SHEET)
    if read_status != "OK":
        return {
            "status": read_status,
            "quantidade": None,
            "valor_total": None,
            "texto": _bonificacao_text(read_status, None, None),
        }
    bonificacao_rows = [
        row
        for row in rows
        if _text(row.get("tipo_operacao_especial")).upper() == "BONIFICACAO_RECEBIDA"
        and _text(row.get("conta_id")) == "2.05.003"
        and _text(row.get("tratamento_operacao_especial")).upper() == "REDUCAO_CUSTO"
    ]
    valor_total = _sum_decimal(bonificacao_rows, "valor")
    status = "OK" if bonificacao_rows else "NAO_GERADO"
    return {
        "status": status,
        "quantidade": len(bonificacao_rows),
        "valor_total": valor_total,
        "texto": _bonificacao_text(status, len(bonificacao_rows), valor_total),
    }


def _audit_manuais(pre_livro_xlsx_path: Path, manual_manifest: object) -> dict[str, object]:
    manual = manual_manifest if isinstance(manual_manifest, dict) else {"executado": False}
    if not manual.get("executado"):
        return {
            "status": "NAO_APLICAVEL",
            "executado": False,
            "texto": _manuais_text(manual, "NAO_APLICAVEL", (), 0),
        }

    rows, read_status = _read_sheet_rows(pre_livro_xlsx_path, LANCAMENTOS_OK_SHEET)
    if read_status != "OK":
        return {
            "status": read_status,
            "executado": True,
            "texto": _manuais_text(manual, read_status, (), 0),
        }

    manual_rows = [
        row for row in rows if _text(row.get("origem")).upper() == MANUAL_GERENCIAL
    ]
    tipos = tuple(
        sorted({
            _text(row.get("tipo_lancamento"))
            for row in manual_rows
            if _text(row.get("tipo_lancamento"))
        })
    )
    accepted = int(manual.get("lancamentos_aceitos") or 0)
    expected_rows = accepted * 2
    totalizable = [tipo for tipo in tipos if tipo.endswith("_TOTALIZAVEL")]
    status = "OK" if len(manual_rows) == expected_rows else "ATENCAO"
    return {
        "status": status,
        "executado": True,
        "linhas_pre_livro": len(manual_rows),
        "tipos": tipos,
        "texto": _manuais_text(manual, status, tipos, len(manual_rows), len(totalizable)),
    }


def _comodato_text(status: str, count: int | None, valor_total: Decimal | None) -> str:
    lines = [
        "AUDITORIA COMODATO FORA DRE",
        "===========================",
        "",
        f"status: {status}",
        f"COMODATO_BASE_DRE: {_nullable(count)}",
        f"valor_total_base_dre: {_decimal_or_na(valor_total)}",
        "COMODATO fora DRE: OK" if count == 0 else "COMODATO fora DRE: VERIFICAR",
        "",
        "Regra: COMODATO nao entra no DRE; permanece controle extracontabil.",
    ]
    return "\n".join(lines) + "\n"


def _bonificacao_text(status: str, count: int | None, valor_total: Decimal | None) -> str:
    lines = [
        "AUDITORIA BONIFICACAO_RECEBIDA DRE",
        "==================================",
        "",
        f"status: {status}",
        "conta_id: 2.05.003",
        "conta_nome: Reducao de Custos - Bonificacoes Recebidas",
        "tratamento: REDUCAO_CUSTO",
        "entra_dre: true",
        f"quantidade_lancamentos: {_nullable(count)}",
        f"valor_total: {_decimal_or_na(valor_total)}",
    ]
    return "\n".join(lines) + "\n"


def _manuais_text(
    manual: Mapping[str, object],
    status: str,
    tipos: tuple[str, ...],
    linhas_pre_livro: int,
    totalizaveis: int = 0,
) -> str:
    executed = bool(manual.get("executado"))
    lines = [
        "AUDITORIA LANCAMENTOS MANUAIS GERENCIAIS",
        "========================================",
        "",
        f"status: {status}",
        f"executado: {str(executed).lower()}",
    ]
    if not executed:
        lines.append("NAO_APLICAVEL")
        return "\n".join(lines) + "\n"

    accepted = int(manual.get("lancamentos_aceitos") or 0)
    technical = int(manual.get("linhas_tecnicas_geradas") or 0)
    lines.extend(
        [
            f"input_csv: {_text(manual.get('input_csv'))}",
            f"lancamentos_aceitos: {accepted}",
            f"linhas_tecnicas_geradas: {technical}",
            f"linhas_pre_livro_MANUAL_GERENCIAL: {linhas_pre_livro}",
            f"total_manual: {_text(manual.get('total_manual'))}",
            "tipos_MANUAL_GERENCIAL: " + (", ".join(tipos) if tipos else "-"),
            "pre_livro_partida_dobrada: "
            + ("OK" if linhas_pre_livro == accepted * 2 else "VERIFICAR"),
            "dre_impacta_uma_vez: "
            + ("OK" if totalizaveis <= accepted else "VERIFICAR"),
        ]
    )
    return "\n".join(lines) + "\n"


def _build_manifest(
    run_manifest: Mapping[str, object],
    *,
    run_id: str,
    status: str,
    files_manifest: Mapping[str, Mapping[str, object]],
    comodato_audit: Mapping[str, object],
    bonificacao_audit: Mapping[str, object],
    manuais_audit: Mapping[str, object],
) -> dict[str, object]:
    return {
        "tipo_pacote": "pacote_contabil_gerencial_provisorio",
        "run_id": run_id,
        "fase": "9O",
        "status": status,
        "data_geracao": datetime.now().isoformat(timespec="seconds"),
        "cnpjs_consolidados": _cnpjs_from_manifest(run_manifest),
        "fase9_executada": bool(
            (run_manifest.get("classificacao") or {}).get("fase9_executada")
            if isinstance(run_manifest.get("classificacao"), dict)
            else False
        ),
        "comodato_entra_dre": False,
        "bonificacao_recebida_entra_dre": True,
        "lancamentos_manuais_gerenciais": run_manifest.get(
            "lancamentos_manuais_gerenciais", {"executado": False}
        ),
        "auditorias": {
            "comodato_fora_dre": {
                "status": comodato_audit["status"],
                "comodato_base_dre": comodato_audit.get("comodato_base_dre"),
            },
            "bonificacao_recebida_dre": {
                "status": bonificacao_audit["status"],
                "quantidade": bonificacao_audit.get("quantidade"),
            },
            "lancamentos_manuais_gerenciais": {
                "status": manuais_audit["status"],
                "executado": manuais_audit.get("executado"),
            },
        },
        "arquivos": files_manifest,
        "avisos": [
            "Pacote gerencial provisorio; nao e livro diario definitivo.",
            "DRE e pre-livro consolidados para os 2 CNPJ como uma unica empresa gerencial.",
            "Extrato bancario e comprovantes nao importados nesta fase.",
        ],
    }


def _indice_text(
    manifest: Mapping[str, object],
    files_manifest: Mapping[str, Mapping[str, object]],
) -> str:
    manual = manifest.get("lancamentos_manuais_gerenciais")
    manual_executado = (
        manual.get("executado") if isinstance(manual, dict) else False
    )
    lines = [
        "INDICE - PACOTE CONTABIL GERENCIAL PROVISORIO",
        "============================================",
        "",
        f"run_id: {manifest.get('run_id')}",
        f"fase: {manifest.get('fase')}",
        f"status: {manifest.get('status')}",
        f"data_geracao: {manifest.get('data_geracao')}",
        "cnpjs_consolidados: " + ", ".join(manifest.get("cnpjs_consolidados", [])),
        "",
        "Arquivos do pacote:",
    ]
    for name, row in sorted(files_manifest.items()):
        lines.append(f"- {name}: {row.get('status')}")
    lines.extend(
        [
            "",
            "Entra no DRE:",
            "- Lancamentos OK com conta mapeada DRE.",
            "- BONIFICACAO_RECEBIDA em 2.05.003 / REDUCAO_CUSTO.",
            "- Lancamentos manuais gerenciais somente no lado economico totalizavel.",
            "",
            "Fora DRE:",
            "- COMODATO: fora DRE / controle extracontabil.",
            "- Extrato bancario, comprovantes, OCR e matching bancario: nao importados nesta fase.",
            "",
            f"COMODATO_BASE_DRE: {manifest['auditorias']['comodato_fora_dre'].get('comodato_base_dre')}",
            "COMODATO fora DRE: OK",
            f"BONIFICACAO_RECEBIDA entra DRE: {str(manifest.get('bonificacao_recebida_entra_dre')).lower()}",
            f"lancamentos_manuais_gerenciais executado: {str(bool(manual_executado)).lower()}",
            "",
            "Avisos:",
            "- Pacote provisorio/gerencial; nao usar como livro diario definitivo.",
            "- Nao ha importacao de extrato bancario, comprovantes, OCR ou matching bancario nesta fase.",
        ]
    )
    return "\n".join(lines) + "\n"


def _read_sheet_rows(xlsx_path: Path, sheet_name: str) -> tuple[list[dict[str, object]], str]:
    if not xlsx_path.exists() or not xlsx_path.is_file():
        return [], "NAO_GERADO"
    try:
        workbook = load_workbook(xlsx_path, data_only=True)
    except Exception:
        return [], "NAO_GERADO"
    if sheet_name not in workbook.sheetnames:
        return [], "NAO_GERADO"
    sheet = workbook[sheet_name]
    headers = [_text(cell.value) for cell in sheet[1]]
    rows: list[dict[str, object]] = []
    for values in sheet.iter_rows(min_row=2, values_only=True):
        if not any(_text(value) for value in values):
            continue
        rows.append(
            {
                header: values[index] if index < len(values) else None
                for index, header in enumerate(headers)
            }
        )
    return rows, "OK"


def _load_json(path: Path) -> dict[str, object]:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _cnpjs_from_manifest(manifest: Mapping[str, object]) -> list[str]:
    fontes = manifest.get("fontes_xml")
    if not isinstance(fontes, list):
        return []
    cnpjs = []
    for item in fontes:
        if not isinstance(item, dict):
            continue
        cnpj = _text(item.get("cnpj_real")) or _text(item.get("cnpj_pasta"))
        if cnpj and cnpj not in cnpjs:
            cnpjs.append(cnpj)
    return cnpjs


def _sum_decimal(rows: list[dict[str, object]], field: str) -> Decimal:
    total = Decimal("0.00")
    for row in rows:
        total += _decimal(row.get(field))
    return total.quantize(Decimal("0.01"))


def _decimal(value: object) -> Decimal:
    if value is None or value == "":
        return Decimal("0.00")
    try:
        return Decimal(str(value)).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError):
        return Decimal("0.00")


def _json_ready(value: object) -> object:
    if isinstance(value, Decimal):
        return f"{value.quantize(Decimal('0.01')):.2f}"
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, Mapping):
        return {str(key): _json_ready(item) for key, item in value.items()}
    if isinstance(value, tuple | list):
        return [_json_ready(item) for item in value]
    return value


def _nullable(value: object) -> str:
    if value is None:
        return "NAO_GERADO"
    return str(value)


def _decimal_or_na(value: Decimal | None) -> str:
    if value is None:
        return "NAO_GERADO"
    return f"{value.quantize(Decimal('0.01')):.2f}"


def _text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()
