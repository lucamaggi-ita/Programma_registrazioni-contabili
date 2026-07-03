from __future__ import annotations

import json
from datetime import date
from pathlib import Path

import pytest

import fantozzi.run_manager as run_manager
from fantozzi.run_manager import (
    PROJECT_ROOT,
    RunManagerError,
    add_checkpoint,
    generate_run_id,
    preparar_run,
    update_manifest_status,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_run_dir(base: Path, run_id: str) -> None:
    (base / "runs" / run_id).mkdir(parents=True, exist_ok=True)


def _make_nfe_xml(path: Path, cnpj_dest: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        f'<?xml version="1.0"?>'
        f'<nfeProc xmlns="http://www.portalfiscal.inf.br/nfe">'
        f'<NFe><infNFe Id="NFe1"><dest><CNPJ>{cnpj_dest}</CNPJ></dest></infNFe></NFe>'
        f'</nfeProc>',
        encoding="utf-8",
    )


def _dirs(tmp_path: Path) -> tuple[Path, Path, Path]:
    xml = tmp_path / "xml_base"
    out = tmp_path / "output"
    rep = tmp_path / "reports"
    xml.mkdir()
    return xml, out, rep


# ---------------------------------------------------------------------------
# generate_run_id
# ---------------------------------------------------------------------------

def test_run_id_formato(tmp_path):
    xml, out, rep = _dirs(tmp_path)
    run_id = generate_run_id(out, rep, today=date(2026, 5, 6))
    assert run_id == "2026-05-06_001"


def test_run_id_primeira_run_do_dia(tmp_path):
    xml, out, rep = _dirs(tmp_path)
    # Dirs de runs nao existem ainda
    run_id = generate_run_id(out, rep, today=date(2026, 5, 6))
    assert run_id.endswith("_001")


def test_run_id_incrementa_no_mesmo_dia(tmp_path):
    xml, out, rep = _dirs(tmp_path)
    # output tem run_003, reports tem run_002 — proximo deve ser _004
    _make_run_dir(out, "2026-05-06_003")
    _make_run_dir(rep, "2026-05-06_002")

    run_id = generate_run_id(out, rep, today=date(2026, 5, 6))

    assert run_id == "2026-05-06_004"


def test_run_id_ignora_pastas_invalidas(tmp_path):
    xml, out, rep = _dirs(tmp_path)
    # Pastas com nomes invalidos nao devem interferir
    _make_run_dir(out, "2026-05-06_abc")
    _make_run_dir(out, "nao-e-run")

    run_id = generate_run_id(out, rep, today=date(2026, 5, 6))

    assert run_id == "2026-05-06_001"


def test_run_id_com_data_injetavel(tmp_path):
    xml, out, rep = _dirs(tmp_path)
    _make_run_dir(out, "2026-01-15_005")

    run_id_jan = generate_run_id(out, rep, today=date(2026, 1, 15))
    run_id_may = generate_run_id(out, rep, today=date(2026, 5, 6))

    assert run_id_jan == "2026-01-15_006"
    assert run_id_may == "2026-05-06_001"


# ---------------------------------------------------------------------------
# current_run.txt
# ---------------------------------------------------------------------------

def test_current_run_txt_criado(tmp_path):
    xml, out, rep = _dirs(tmp_path)

    result = preparar_run(xml, out, rep, today=date(2026, 5, 6), project_root=tmp_path)

    assert result.current_run_output.exists()
    assert result.current_run_reports.exists()


def test_current_run_txt_atualizado_na_segunda_run(tmp_path):
    xml, out, rep = _dirs(tmp_path)

    preparar_run(xml, out, rep, today=date(2026, 5, 6), project_root=tmp_path)
    result2 = preparar_run(xml, out, rep, today=date(2026, 5, 6), project_root=tmp_path)

    content_out = result2.current_run_output.read_text(encoding="utf-8").strip()
    content_rep = result2.current_run_reports.read_text(encoding="utf-8").strip()
    assert content_out == result2.run_id
    assert content_rep == result2.run_id
    assert result2.run_id == "2026-05-06_002"


def test_current_run_txt_contem_apenas_run_id(tmp_path):
    xml, out, rep = _dirs(tmp_path)

    result = preparar_run(xml, out, rep, today=date(2026, 5, 6), project_root=tmp_path)

    content_out = result.current_run_output.read_text(encoding="utf-8").strip()
    content_rep = result.current_run_reports.read_text(encoding="utf-8").strip()
    assert content_out == "2026-05-06_001"
    assert content_rep == "2026-05-06_001"
    assert "\n" not in content_out
    assert "\n" not in content_rep


def test_preparar_run_dry_run_nao_escreve_infraestrutura(tmp_path):
    xml, out, rep = _dirs(tmp_path)

    result = preparar_run(
        xml,
        out,
        rep,
        dry_run=True,
        today=date(2026, 5, 6),
        project_root=tmp_path,
    )

    assert result.dry_run is True
    assert result.run_id == "2026-05-06_001"
    assert not result.output_run_dir.exists()
    assert not result.reports_run_dir.exists()
    assert not result.current_run_output.exists()
    assert not result.current_run_reports.exists()
    assert not result.manifest_path.exists()


# ---------------------------------------------------------------------------
# manifest_run.json
# ---------------------------------------------------------------------------

def test_manifesto_criado_com_campos_obrigatorios(tmp_path):
    xml, out, rep = _dirs(tmp_path)

    result = preparar_run(xml, out, rep, today=date(2026, 5, 6), project_root=tmp_path)

    assert result.manifest_path.exists()
    data = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    for campo in (
        "run_id", "data_execucao", "status_run", "modo", "dry_run",
        "git", "parametros", "config_usado", "fontes_xml", "totais",
        "classificacao", "relatorios_gerados", "checkpoints", "avisos", "erros",
    ):
        assert campo in data, f"campo obrigatorio ausente: {campo}"
    assert data["run_id"] == "2026-05-06_001"
    assert data["status_run"] == "INICIADA"
    assert data["dry_run"] is False
    # manifesto esta APENAS em reports, nao em output
    assert not (result.output_run_dir / "manifest_run.json").exists()


def test_manifesto_registra_hash_config(tmp_path):
    xml, out, rep = _dirs(tmp_path)
    config_dir = tmp_path / "config"
    config_dir.mkdir()
    cfg = config_dir / "plano.yaml"
    cfg.write_text("contas: []", encoding="utf-8")

    result = preparar_run(
        xml, out, rep,
        config_paths=[cfg],
        today=date(2026, 5, 6),
        project_root=tmp_path,
    )

    data = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    entry = data["config_usado"]["plano"]
    assert entry["hash_sha256"] is not None
    assert len(entry["hash_sha256"]) == 64
    assert entry["data_modificacao"] is not None


def test_manifesto_registra_regras_comprovantes_no_config_default(
    tmp_path, monkeypatch
):
    xml, out, rep = _dirs(tmp_path)
    config_dir = tmp_path / "config"
    config_dir.mkdir()
    for file_name in (
        "plano_contas_gerencial.yaml",
        "regras_classificacao_produtos.yaml",
        "regras_classificacao_comprovantes.yaml",
        "mapeamento_dre.yaml",
    ):
        (config_dir / file_name).write_text("regras: []\n", encoding="utf-8")
    regras_path = config_dir / "regras_classificacao_comprovantes.yaml"

    monkeypatch.setattr(run_manager, "PROJECT_ROOT", tmp_path)

    result = preparar_run(
        xml,
        out,
        rep,
        today=date(2026, 5, 6),
        project_root=tmp_path,
    )

    data = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    entry = data["config_usado"]["regras_classificacao_comprovantes"]
    assert entry["path"] == str(regras_path)
    assert entry["hash_sha256"] is not None
    assert len(entry["hash_sha256"]) == 64
    assert entry["data_modificacao"] is not None


def test_manifesto_checkpoint_adicionado(tmp_path):
    xml, out, rep = _dirs(tmp_path)
    result = preparar_run(xml, out, rep, today=date(2026, 5, 6), project_root=tmp_path)

    add_checkpoint(result.manifest_path, "importacao concluida")

    data = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert len(data["checkpoints"]) == 1
    assert data["checkpoints"][0]["descricao"] == "importacao concluida"
    assert "momento" in data["checkpoints"][0]


def test_manifesto_status_atualizado(tmp_path):
    xml, out, rep = _dirs(tmp_path)
    result = preparar_run(xml, out, rep, today=date(2026, 5, 6), project_root=tmp_path)
    data = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert data["status_run"] == "INICIADA"

    update_manifest_status(result.manifest_path, "PROVISORIO_CONCLUIDO")

    data2 = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert data2["status_run"] == "PROVISORIO_CONCLUIDO"


def test_manifesto_cnpj_pasta_vs_cnpj_real(tmp_path):
    xml_base = tmp_path / "xml_base"
    xml_base.mkdir()
    # Pasta nomeada sem zero inicial, CNPJ real tem 14 digitos
    xml_dir = xml_base / "8884180000123" / "xml"
    xml_dir.mkdir(parents=True)
    _make_nfe_xml(xml_dir / "nota.xml", "08884180000123")

    out = tmp_path / "output"
    rep = tmp_path / "reports"
    result = preparar_run(
        xml_base, out, rep, today=date(2026, 5, 6), project_root=tmp_path
    )

    data = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    fonte = data["fontes_xml"][0]
    assert fonte["cnpj_pasta"] == "8884180000123"
    assert fonte["cnpj_real"] == "08884180000123"
    assert fonte["origem_cnpj_real"] == "lido_do_xml"
    assert "nota" in fonte


def test_manifesto_cnpj_real_lido_de_xml_apos_primeiros_invalidos(tmp_path):
    """Os 3 primeiros XML sem CNPJ legivel nao devem impedir a leitura do 4o."""
    xml_base = tmp_path / "xml_base"
    xml_base.mkdir()
    xml_dir = xml_base / "8884180000123" / "xml"
    xml_dir.mkdir(parents=True)

    # XML invalido (malformado)
    (xml_dir / "nota_01.xml").write_text("nao e xml valido", encoding="utf-8")
    # XML sem elemento dest
    (xml_dir / "nota_02.xml").write_text(
        '<?xml version="1.0"?>'
        '<nfeProc xmlns="http://www.portalfiscal.inf.br/nfe">'
        '<NFe><infNFe Id="NFe1"></infNFe></NFe>'
        '</nfeProc>',
        encoding="utf-8",
    )
    # XML com dest mas sem CNPJ (apenas CPF)
    (xml_dir / "nota_03.xml").write_text(
        '<?xml version="1.0"?>'
        '<nfeProc xmlns="http://www.portalfiscal.inf.br/nfe">'
        '<NFe><infNFe Id="NFe1"><dest><CPF>12345678901</CPF></dest></infNFe></NFe>'
        '</nfeProc>',
        encoding="utf-8",
    )
    # XML valido com CNPJ real — deve ser encontrado
    _make_nfe_xml(xml_dir / "nota_04.xml", "08884180000123")

    out = tmp_path / "output"
    rep = tmp_path / "reports"
    result = preparar_run(
        xml_base, out, rep, today=date(2026, 5, 6), project_root=tmp_path
    )

    data = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    fonte = data["fontes_xml"][0]
    assert fonte["cnpj_pasta"] == "8884180000123"
    assert fonte["cnpj_real"] == "08884180000123"
    assert fonte["origem_cnpj_real"] == "lido_do_xml"


def test_manifesto_cnpj_real_inferido_da_pasta_quando_xml_sem_cnpj(tmp_path):
    """Quando nenhum XML fornece CNPJ, usa zero-padding com origem inferido_da_pasta."""
    xml_base = tmp_path / "xml_base"
    xml_base.mkdir()
    xml_dir = xml_base / "8884180000123" / "xml"
    xml_dir.mkdir(parents=True)
    # XML sem CNPJ do dest
    (xml_dir / "nota.xml").write_text(
        '<?xml version="1.0"?>'
        '<nfeProc xmlns="http://www.portalfiscal.inf.br/nfe">'
        '<NFe><infNFe Id="NFe1"></infNFe></NFe>'
        '</nfeProc>',
        encoding="utf-8",
    )

    out = tmp_path / "output"
    rep = tmp_path / "reports"
    result = preparar_run(
        xml_base, out, rep, today=date(2026, 5, 6), project_root=tmp_path
    )

    data = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    fonte = data["fontes_xml"][0]
    assert fonte["cnpj_pasta"] == "8884180000123"
    assert fonte["cnpj_real"] == "08884180000123"
    assert fonte["origem_cnpj_real"] == "inferido_da_pasta"


# ---------------------------------------------------------------------------
# Validacao de paths
# ---------------------------------------------------------------------------

def test_validate_paths_rejeita_xml_dentro_saida(tmp_path):
    out = tmp_path / "output"
    rep = tmp_path / "reports"
    xml_dentro_saida = out / "xml_base"
    xml_dentro_saida.mkdir(parents=True)

    with pytest.raises(RunManagerError, match="xml_base nao pode estar dentro de saida_output"):
        preparar_run(xml_dentro_saida, out, rep, project_root=tmp_path)


def test_validate_paths_rejeita_saida_dentro_xml(tmp_path):
    xml = tmp_path / "xml_base"
    xml.mkdir()
    saida_dentro_xml = xml / "output"

    with pytest.raises(RunManagerError, match="saida_output nao pode estar dentro de xml_base"):
        preparar_run(xml, saida_dentro_xml, tmp_path / "reports", project_root=tmp_path)


def test_validate_paths_rejeita_path_fora_project_root(tmp_path):
    project = tmp_path / "project"
    project.mkdir()
    xml_fora = tmp_path / "xml_externo"
    xml_fora.mkdir()
    out = project / "output"
    rep = project / "reports"

    with pytest.raises(RunManagerError, match="fora do project root"):
        preparar_run(xml_fora, out, rep, project_root=project)
