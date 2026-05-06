from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path
import re

import yaml

from fantozzi.chart_of_accounts import ChartOfAccounts
from fantozzi.classification_review_validation import (
    ALLOWED_ACTIONS,
    ClassificationReviewValidationError,
    validate_classification_review_csv,
)
from fantozzi.normalization import normalize


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PLANO_CONTAS = PROJECT_ROOT / "config" / "plano_contas_gerencial.yaml"
DEFAULT_REGRAS_PRODUTOS = (
    PROJECT_ROOT / "config" / "regras_classificacao_produtos.yaml"
)

PROPOSAL_ACTION_TYPES = {
    "CLASSIFICAR": "regra",
    "CRIAR_ALIAS": "alias",
    "CRIAR_REGRA": "regra",
}

ACOES_PROPOSTA_REAL = (
    "ANALISAR_COM_CONTADOR",
    "CLASSIFICAR",
    "CRIAR_ALIAS",
    "CRIAR_REGRA",
    "IGNORAR",
    "MANTER_REVISAO",
)

CSV_REVISAO_REAL_HEADERS = (
    "tipo_candidato",
    "descricao_normalizada",
    "descricao_exemplo",
    "ocorrencias",
    "fornecedores",
    "documentos",
    "valor_total",
    "conta_atual_id",
    "conta_atual_nome",
    "status_revisao",
    "alertas",
    "acao_manual",
    "conta_id_sugerida",
    "conta_nome_sugerida",
    "palavras_chave_sugeridas",
    "observacao_manual",
)


class ClassificationRuleProposalError(Exception):
    pass


@dataclass(frozen=True)
class PropostaClassificacao:
    tipo: str
    termo: str
    termo_normalizado: str
    conta_id: str
    conta_nome: str
    acao_manual: str
    origem: str = "revisao_manual_csv"
    observacao: str = "Revisar manualmente antes de aplicar"
    alertas: tuple[str, ...] = ()


@dataclass(frozen=True)
class ResultadoPropostaClassificacao:
    csv_path: Path
    plano_contas_path: Path
    regras_produtos_path: Path | None
    total_linhas: int
    linhas_pendentes: int
    linhas_validas: int
    linhas_invalidas: int
    linhas_validas_sem_proposta: int
    contas_desconhecidas: int
    acoes_desconhecidas: int
    duplicidades_potenciais: int
    propostas: tuple[PropostaClassificacao, ...]
    alertas: tuple[str, ...]
    status: str


@dataclass(frozen=True)
class ResultadoPropostaRegrasClassificacaoReal:
    yaml_path: Path
    relatorio_path: Path | None
    csv_paths: tuple[Path, ...]
    plano_contas_path: Path
    total_linhas_lidas: int
    total_propostas_classificar: int
    total_propostas_alias: int
    total_propostas_regras: int
    total_pendencias_contador: int
    total_pendencias_revisao: int
    total_ignoradas: int
    total_avisos: int
    total_erros: int
    acoes: dict[str, int]
    avisos: tuple[str, ...]
    erros: tuple[str, ...]
    status: str


def gerar_proposta_classificacao(
    csv_path: Path | str,
    plano_contas_path: Path | str | None = None,
    regras_produtos_path: Path | str | None = DEFAULT_REGRAS_PRODUTOS,
) -> ResultadoPropostaClassificacao:
    csv_path = Path(csv_path)
    plano_contas_path = Path(plano_contas_path or DEFAULT_PLANO_CONTAS)
    regras_produtos_path = (
        Path(regras_produtos_path) if regras_produtos_path is not None else None
    )

    try:
        validation_result = validate_classification_review_csv(
            csv_path, plano_contas_path
        )
    except ClassificationReviewValidationError as exc:
        raise ClassificationRuleProposalError(str(exc)) from exc

    chart = _load_chart(plano_contas_path)
    existing_terms = _load_existing_rule_terms(regras_produtos_path)
    rows = _read_csv_rows(csv_path)

    propostas: list[PropostaClassificacao] = []
    alertas_gerais: list[str] = []
    propostas_by_term: dict[str, set[str]] = {}
    accounts_by_description: dict[str, set[str]] = {}
    duplicidades_potenciais = 0
    linhas_validas_sem_proposta = 0

    for index, row in enumerate(rows, start=2):
        row_status = _row_status(row, chart)
        if row_status != "VALIDA":
            continue

        acao_manual = _normalize_action(row.get("acao_manual"))
        tipo = PROPOSAL_ACTION_TYPES.get(acao_manual)
        if not tipo:
            linhas_validas_sem_proposta += 1
            alertas_gerais.append(
                f"Linha {index}: acao_manual valida sem proposta YAML ({acao_manual})."
            )
            continue

        conta_id = _normalize_account_id(_text(row.get("conta_id_sugerida")))
        conta = chart.get_by_id(conta_id)
        row_alertas = []
        description_key = normalize(
            _text(row.get("descricao_normalizada"))
            or _text(row.get("descricao_exemplo"))
        )
        if description_key:
            known_accounts = accounts_by_description.setdefault(description_key, set())
            if known_accounts and conta_id not in known_accounts:
                row_alertas.append("CONTA_CONFLITANTE")
                duplicidades_potenciais += 1
            known_accounts.add(conta_id)

        terms = _proposal_terms(row)
        if not terms:
            linhas_validas_sem_proposta += 1
            alertas_gerais.append(
                f"Linha {index}: nenhum termo informado para proposta."
            )
            continue

        for term in terms:
            normalized_term = normalize(term)
            alertas = list(row_alertas)
            if normalized_term in existing_terms:
                alertas.append("TERMO_JA_EXISTE_EM_REGRAS")
            if normalized_term in propostas_by_term:
                alertas.append("TERMO_DUPLICADO_NA_PROPOSTA")
                if conta_id not in propostas_by_term[normalized_term]:
                    alertas.append("CONTA_CONFLITANTE")
                duplicidades_potenciais += 1
            propostas_by_term.setdefault(normalized_term, set()).add(conta_id)

            propostas.append(
                PropostaClassificacao(
                    tipo=tipo,
                    termo=term,
                    termo_normalizado=normalized_term,
                    conta_id=conta.id,
                    conta_nome=conta.nome,
                    acao_manual=acao_manual,
                    alertas=tuple(dict.fromkeys(alertas)),
                )
            )

    status = (
        "PROPOSTA GERADA PARA REVISAO MANUAL"
        if validation_result.invalid_rows == 0
        else "NAO FOI POSSIVEL GERAR PROPOSTA COMPLETA"
    )

    return ResultadoPropostaClassificacao(
        csv_path=csv_path,
        plano_contas_path=plano_contas_path,
        regras_produtos_path=regras_produtos_path,
        total_linhas=validation_result.total_rows,
        linhas_pendentes=validation_result.pending_rows,
        linhas_validas=validation_result.valid_rows,
        linhas_invalidas=validation_result.invalid_rows,
        linhas_validas_sem_proposta=linhas_validas_sem_proposta,
        contas_desconhecidas=validation_result.unknown_accounts,
        acoes_desconhecidas=validation_result.unknown_actions,
        duplicidades_potenciais=duplicidades_potenciais,
        propostas=tuple(propostas),
        alertas=tuple(alertas_gerais),
        status=status,
    )


def gerar_proposta_regras_classificacao_real(
    csv_paths: list[Path] | tuple[Path, ...],
    output_yaml_path: Path,
    plano_contas_path: Path | None = None,
    relatorio_path: Path | None = None,
) -> ResultadoPropostaRegrasClassificacaoReal:
    csv_paths = [Path(path) for path in csv_paths]
    output_yaml_path = Path(output_yaml_path)
    plano_contas_path = Path(plano_contas_path or DEFAULT_PLANO_CONTAS)
    relatorio_path = Path(relatorio_path) if relatorio_path else None

    chart = _load_chart(plano_contas_path)
    proposal, result = _build_real_proposal(
        csv_paths,
        output_yaml_path,
        plano_contas_path,
        relatorio_path,
        chart,
    )

    if result.erros:
        if relatorio_path:
            salvar_relatorio_proposta(
                relatorio_path,
                gerar_relatorio_proposta_regras_classificacao_real(result),
            )
        raise ClassificationRuleProposalError("; ".join(result.erros))

    output_yaml_path.parent.mkdir(parents=True, exist_ok=True)
    output_yaml_path.write_text(
        yaml.safe_dump(proposal, sort_keys=False, allow_unicode=True),
        encoding="utf-8",
        newline="\n",
    )

    if relatorio_path:
        salvar_relatorio_proposta(
            relatorio_path,
            gerar_relatorio_proposta_regras_classificacao_real(result),
        )

    return result


def gerar_relatorio_proposta_regras_classificacao_real(
    resultado: ResultadoPropostaRegrasClassificacaoReal,
) -> str:
    linhas = [
        "RESUMO DA PROPOSTA DE REGRAS DE CLASSIFICACAO",
        "============================================",
        "",
        "Arquivos CSV:",
    ]
    linhas.extend(f"- {csv_path}" for csv_path in resultado.csv_paths)
    linhas.extend(
        [
            f"Arquivo YAML gerado: {resultado.yaml_path}",
            f"Plano de contas: {resultado.plano_contas_path}",
            "",
            f"Linhas lidas: {resultado.total_linhas_lidas}",
            "",
            "Acoes:",
        ]
    )
    for acao in ACOES_PROPOSTA_REAL:
        linhas.append(f"- {acao}: {resultado.acoes.get(acao, 0)}")

    linhas.extend(
        [
            "",
            "Propostas geradas:",
            f"- Classificacoes diretas: {resultado.total_propostas_classificar}",
            f"- Aliases: {resultado.total_propostas_alias}",
            f"- Regras: {resultado.total_propostas_regras}",
            "",
            "Pendencias:",
            f"- Contador: {resultado.total_pendencias_contador}",
            f"- Revisao futura: {resultado.total_pendencias_revisao}",
            "",
            f"Ignoradas: {resultado.total_ignoradas}",
            f"Avisos: {resultado.total_avisos}",
            f"Erros: {resultado.total_erros}",
            "",
            f"Resultado: {resultado.status}",
        ]
    )
    if resultado.avisos:
        linhas.extend(["", "Avisos:", "----------------------------------------"])
        linhas.extend(f"- {aviso}" for aviso in resultado.avisos)
    if resultado.erros:
        linhas.extend(["", "Erros:", "----------------------------------------"])
        linhas.extend(f"- {erro}" for erro in resultado.erros)
    linhas.extend(
        [
            "",
            "Seguranca:",
            "----------------------------------------",
            "- Proposta gerada em arquivo separado.",
            "- Nenhum YAML de config foi alterado.",
            "- Nenhum CSV de revisao foi reescrito.",
            "- Nenhum Excel de lancamentos foi alterado.",
        ]
    )
    return "\n".join(linhas) + "\n"


def console_lines_real(
    resultado: ResultadoPropostaRegrasClassificacaoReal,
) -> list[str]:
    return [
        "[OK] Proposta real de classificacao gerada para revisao humana.",
        f"     Linhas lidas: {resultado.total_linhas_lidas}",
        f"     Classificacoes diretas: {resultado.total_propostas_classificar}",
        f"     Aliases: {resultado.total_propostas_alias}",
        f"     Regras: {resultado.total_propostas_regras}",
        f"     Pendencias contador: {resultado.total_pendencias_contador}",
        f"     Pendencias revisao: {resultado.total_pendencias_revisao}",
        f"     YAML gerado: {resultado.yaml_path}",
    ] + (
        [f"     Relatorio: {resultado.relatorio_path}"]
        if resultado.relatorio_path
        else []
    )


def escrever_proposta_yaml(
    resultado: ResultadoPropostaClassificacao,
    saida_yaml: Path | str,
) -> None:
    saida_yaml = Path(saida_yaml)
    saida_yaml.parent.mkdir(parents=True, exist_ok=True)
    data = {
        "propostas": [
            {
                "tipo": proposta.tipo,
                "termo": proposta.termo,
                "termo_normalizado": proposta.termo_normalizado,
                "conta_id": proposta.conta_id,
                "conta_nome": proposta.conta_nome,
                "acao_manual": proposta.acao_manual,
                "origem": proposta.origem,
                "observacao": proposta.observacao,
                "alertas": list(proposta.alertas),
            }
            for proposta in resultado.propostas
        ]
    }
    header = (
        "# PROPOSTA GERADA AUTOMATICAMENTE - NAO APLICADA\n"
        "# Revisar manualmente antes de copiar para "
        "config/regras_classificacao_produtos.yaml\n"
        f"# Origem: {resultado.csv_path}\n\n"
    )
    content = yaml.safe_dump(data, sort_keys=False, allow_unicode=True)
    saida_yaml.write_text(header + content, encoding="utf-8", newline="\n")


def gerar_relatorio_proposta_classificacao(
    resultado: ResultadoPropostaClassificacao,
    csv_path: Path | str | None = None,
    saida_yaml: Path | str | None = None,
) -> str:
    csv_path = Path(csv_path) if csv_path else resultado.csv_path
    regras_path = resultado.regras_produtos_path
    linhas = [
        "FANTOZZI - PROPOSTA DE REGRAS DE CLASSIFICACAO",
        "================================================",
        f"Arquivo CSV: {csv_path}",
        f"Plano de contas: {resultado.plano_contas_path}",
        f"Regras consultadas: {regras_path if regras_path else '-'}",
    ]
    if saida_yaml:
        linhas.append(f"Arquivo YAML proposto: {Path(saida_yaml)}")
    linhas.extend(
        [
            "",
            "Resumo",
            "----------------------------------------",
            f"Total de linhas no CSV: {resultado.total_linhas}",
            f"Linhas pendentes ignoradas: {resultado.linhas_pendentes}",
            f"Linhas validas: {resultado.linhas_validas}",
            f"Linhas invalidas ignoradas: {resultado.linhas_invalidas}",
            f"Linhas validas sem proposta YAML: {resultado.linhas_validas_sem_proposta}",
            f"Propostas geradas: {len(resultado.propostas)}",
            f"Propostas de regra: {_count_by_type(resultado, 'regra')}",
            f"Propostas de alias: {_count_by_type(resultado, 'alias')}",
            f"Contas desconhecidas: {resultado.contas_desconhecidas}",
            f"Acoes desconhecidas: {resultado.acoes_desconhecidas}",
            f"Duplicidades potenciais: {resultado.duplicidades_potenciais}",
            "",
            f"Status: {resultado.status}",
        ]
    )

    propostas_com_alertas = [
        proposta for proposta in resultado.propostas if proposta.alertas
    ]
    if propostas_com_alertas:
        linhas.extend(["", "PROPOSTAS COM ALERTAS", "----------------------------------------"])
        for proposta in propostas_com_alertas:
            linhas.extend(
                [
                    f"Termo: {proposta.termo}",
                    f"Conta sugerida: {proposta.conta_id} - {proposta.conta_nome}",
                    f"Alertas: {', '.join(proposta.alertas)}",
                    "",
                ]
            )

    if resultado.alertas:
        linhas.extend(["", "OBSERVACOES", "----------------------------------------"])
        linhas.extend(f"- {alerta}" for alerta in resultado.alertas)

    linhas.extend(
        [
            "",
            "Seguranca",
            "----------------------------------------",
            "- Proposta gerada em arquivo separado.",
            "- Nenhum YAML de config foi alterado.",
            "- Nenhum Excel foi alterado.",
            "- O CSV de revisao nao foi reescrito.",
        ]
    )
    return "\n".join(linhas) + "\n"


def salvar_relatorio_proposta(report_path: Path | str, text: str) -> None:
    report_path = Path(report_path)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(text, encoding="utf-8", newline="\n")


def console_lines(
    resultado: ResultadoPropostaClassificacao,
    *,
    saida_yaml: Path | str | None = None,
    relatorio: Path | str | None = None,
) -> list[str]:
    lines = [
        "[OK] Proposta de classificacao gerada para revisao manual.",
        f"     Propostas geradas: {len(resultado.propostas)}",
        f"     Linhas pendentes ignoradas: {resultado.linhas_pendentes}",
        f"     Linhas invalidas ignoradas: {resultado.linhas_invalidas}",
    ]
    if saida_yaml:
        lines.append(f"     YAML proposto: {Path(saida_yaml)}")
    if relatorio:
        lines.append(f"     Relatorio: {Path(relatorio)}")
    return lines


def _build_real_proposal(
    csv_paths: list[Path],
    output_yaml_path: Path,
    plano_contas_path: Path,
    relatorio_path: Path | None,
    chart: ChartOfAccounts,
) -> tuple[dict[str, object], ResultadoPropostaRegrasClassificacaoReal]:
    avisos: list[str] = []
    erros: list[str] = []
    propostas_classificacao: list[dict[str, object]] = []
    propostas_alias: list[dict[str, object]] = []
    propostas_regras: list[dict[str, object]] = []
    pendencias_contador: list[dict[str, object]] = []
    pendencias_revisao: list[dict[str, object]] = []
    acoes = {acao: 0 for acao in ACOES_PROPOSTA_REAL}
    total_linhas_lidas = 0
    duplicate_keys: set[tuple[str, str, str]] = set()

    if not csv_paths:
        erros.append("[ERRO] Informe ao menos um arquivo CSV de revisao.")

    for csv_path in csv_paths:
        rows = _read_real_csv_rows(csv_path, erros)
        for line_number, row in rows:
            total_linhas_lidas += 1
            acao = _normalize_action(row.get("acao_manual"))
            if not acao:
                avisos.append(f"Linha {line_number}: acao_manual vazia; linha ignorada.")
                continue
            if acao not in ACOES_PROPOSTA_REAL:
                erros.append(
                    f"[ERRO] Linha {line_number}: acao_manual invalida: {acao}"
                )
                continue

            acoes[acao] += 1
            suggested_account_id = _normalize_account_id(
                _text(row.get("conta_id_sugerida"))
            )
            current_account_id = _normalize_account_id(_text(row.get("conta_atual_id")))
            if suggested_account_id:
                _validate_account_exists(
                    chart, suggested_account_id, line_number, "conta_id_sugerida", erros
                )

            if acao == "CLASSIFICAR":
                if not suggested_account_id:
                    erros.append(
                        "[ERRO] Linha "
                        f"{line_number}: conta_id_sugerida obrigatoria para CLASSIFICAR."
                    )
                    continue
                conta = _get_account_or_none(chart, suggested_account_id)
                if conta is None:
                    continue
                keywords = _keywords_or_description(row, line_number, avisos)
                propostas_classificacao.append(
                    {
                        **_base_real_row(row),
                        "conta_id_sugerida": conta.id,
                        "conta_nome_sugerida": _text(
                            row.get("conta_nome_sugerida")
                        )
                        or conta.nome,
                        "palavras_chave_sugeridas": keywords,
                    }
                )
                _warn_duplicate_real(
                    duplicate_keys,
                    row,
                    conta.id,
                    acao,
                    line_number,
                    avisos,
                )
            elif acao == "CRIAR_ALIAS":
                destination_id = suggested_account_id or current_account_id
                if not destination_id:
                    avisos.append(
                        "Linha "
                        f"{line_number}: CRIAR_ALIAS sem conta_id_sugerida "
                        "e sem conta_atual_id; proposta nao gerada."
                    )
                    continue
                _validate_account_exists(
                    chart, destination_id, line_number, "conta_id_destino", erros
                )
                conta = _get_account_or_none(chart, destination_id)
                if conta is None:
                    continue
                aliases = _keywords_or_description(row, line_number, avisos)
                propostas_alias.append(
                    {
                        "descricao_normalizada": _text(
                            row.get("descricao_normalizada")
                        ),
                        "conta_id_destino": conta.id,
                        "conta_nome_destino": _text(row.get("conta_nome_sugerida"))
                        or _text(row.get("conta_atual_nome"))
                        or conta.nome,
                        "alias_sugeridos": aliases,
                        "origem": "CRIAR_ALIAS",
                        "observacao_manual": _text(row.get("observacao_manual")),
                    }
                )
                _warn_duplicate_real(
                    duplicate_keys,
                    row,
                    conta.id,
                    acao,
                    line_number,
                    avisos,
                )
            elif acao == "CRIAR_REGRA":
                if not suggested_account_id:
                    erros.append(
                        "[ERRO] Linha "
                        f"{line_number}: conta_id_sugerida obrigatoria para CRIAR_REGRA."
                    )
                    continue
                conta = _get_account_or_none(chart, suggested_account_id)
                if conta is None:
                    continue
                keywords = _split_keywords(row.get("palavras_chave_sugeridas"))
                if not keywords:
                    fallback = _text(row.get("descricao_normalizada"))
                    if fallback:
                        keywords = [fallback]
                        avisos.append(
                            "Linha "
                            f"{line_number}: CRIAR_REGRA sem "
                            "palavras_chave_sugeridas; usada descricao_normalizada."
                        )
                    else:
                        erros.append(
                            "[ERRO] Linha "
                            f"{line_number}: CRIAR_REGRA sem palavras-chave "
                            "ou descricao_normalizada."
                        )
                        continue
                propostas_regras.append(
                    {
                        "regra": keywords[0],
                        "conta_id_destino": conta.id,
                        "conta_nome_destino": _text(row.get("conta_nome_sugerida"))
                        or conta.nome,
                        "origem": "CRIAR_REGRA",
                        "criterios": {"palavras_chave": keywords},
                        "observacao_manual": _text(row.get("observacao_manual")),
                    }
                )
                _warn_duplicate_real(
                    duplicate_keys,
                    row,
                    conta.id,
                    acao,
                    line_number,
                    avisos,
                )
            elif acao == "ANALISAR_COM_CONTADOR":
                pendencias_contador.append(
                    {
                        **_base_pending_row(row),
                        "motivo": "ANALISAR_COM_CONTADOR",
                    }
                )
            elif acao == "MANTER_REVISAO":
                pendencias_revisao.append(
                    {
                        **_base_pending_row(row),
                        "motivo": "MANTER_REVISAO",
                    }
                )

    status = "ERRO" if erros else "OK"
    proposal: dict[str, object] = {
        "metadata": {
            "tipo": "proposta_regras_classificacao",
            "origem": "revisao_classificacao_revisada",
            "gerado_por": "fantozzi",
            "fase": "6C",
            "observacao": (
                "Proposta gerada automaticamente para revisao humana. "
                "Nao aplicada em config/*.yaml."
            ),
        },
        "resumo": {
            "total_linhas_lidas": total_linhas_lidas,
            "total_propostas_classificar": len(propostas_classificacao),
            "total_propostas_alias": len(propostas_alias),
            "total_propostas_regras": len(propostas_regras),
            "total_pendencias_contador": len(pendencias_contador),
            "total_pendencias_revisao": len(pendencias_revisao),
            "total_ignoradas": acoes["IGNORAR"],
            "total_avisos": len(avisos),
        },
        "propostas_classificacao": propostas_classificacao,
        "propostas_alias": propostas_alias,
        "propostas_regras": propostas_regras,
        "pendencias_contador": pendencias_contador,
        "pendencias_revisao": pendencias_revisao,
        "avisos": avisos,
    }
    result = ResultadoPropostaRegrasClassificacaoReal(
        yaml_path=output_yaml_path,
        relatorio_path=relatorio_path,
        csv_paths=tuple(csv_paths),
        plano_contas_path=plano_contas_path,
        total_linhas_lidas=total_linhas_lidas,
        total_propostas_classificar=len(propostas_classificacao),
        total_propostas_alias=len(propostas_alias),
        total_propostas_regras=len(propostas_regras),
        total_pendencias_contador=len(pendencias_contador),
        total_pendencias_revisao=len(pendencias_revisao),
        total_ignoradas=acoes["IGNORAR"],
        total_avisos=len(avisos),
        total_erros=len(erros),
        acoes=acoes,
        avisos=tuple(avisos),
        erros=tuple(erros),
        status=status,
    )
    return proposal, result


def _read_real_csv_rows(
    csv_path: Path,
    erros: list[str],
) -> list[tuple[int, dict[str, str]]]:
    if not csv_path.exists():
        erros.append(f"[ERRO] CSV de revisao nao encontrado: {csv_path}")
        return []
    if not csv_path.is_file():
        erros.append(f"[ERRO] CSV de revisao nao e um arquivo: {csv_path}")
        return []
    try:
        with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
            reader = csv.DictReader(file, delimiter=";")
            fieldnames = tuple(reader.fieldnames or ())
            missing_headers = [
                header for header in CSV_REVISAO_REAL_HEADERS if header not in fieldnames
            ]
            if missing_headers:
                erros.append(
                    "[ERRO] CSV de revisao sem colunas obrigatorias: "
                    + ", ".join(missing_headers)
                )
                return []
            return [(index, dict(row)) for index, row in enumerate(reader, start=2)]
    except UnicodeDecodeError as exc:
        erros.append(f"[ERRO] CSV de revisao nao pode ser lido em UTF-8: {csv_path}")
        return []
    except OSError as exc:
        erros.append(f"[ERRO] CSV de revisao nao pode ser lido: {csv_path}")
        return []


def _validate_account_exists(
    chart: ChartOfAccounts,
    account_id: str,
    line_number: int,
    field_name: str,
    erros: list[str],
) -> None:
    try:
        chart.get_by_id(account_id)
    except KeyError:
        erros.append(
            f"[ERRO] Linha {line_number}: {field_name} nao existe "
            f"no plano de contas: {account_id}"
        )


def _get_account_or_none(chart: ChartOfAccounts, account_id: str):
    try:
        return chart.get_by_id(account_id)
    except KeyError:
        return None


def _base_real_row(row: dict[str, str]) -> dict[str, object]:
    return {
        "descricao_normalizada": _text(row.get("descricao_normalizada")),
        "descricao_exemplo": _text(row.get("descricao_exemplo")),
        "ocorrencias": _parse_int_or_text(row.get("ocorrencias")),
        "valor_total": _text(row.get("valor_total")),
        "fornecedores": _text(row.get("fornecedores")),
        "documentos": _text(row.get("documentos")),
        "observacao_manual": _text(row.get("observacao_manual")),
    }


def _base_pending_row(row: dict[str, str]) -> dict[str, object]:
    return {
        "descricao_normalizada": _text(row.get("descricao_normalizada")),
        "descricao_exemplo": _text(row.get("descricao_exemplo")),
        "alertas": _text(row.get("alertas")),
        "observacao_manual": _text(row.get("observacao_manual")),
        "fornecedores": _text(row.get("fornecedores")),
        "documentos": _text(row.get("documentos")),
    }


def _keywords_or_description(
    row: dict[str, str],
    line_number: int,
    avisos: list[str],
) -> list[str]:
    keywords = _split_keywords(row.get("palavras_chave_sugeridas"))
    if keywords:
        return keywords
    fallback = _text(row.get("descricao_normalizada")) or _text(
        row.get("descricao_exemplo")
    )
    if fallback:
        return [fallback]
    avisos.append(
        f"Linha {line_number}: sem palavras_chave_sugeridas ou descricao usavel."
    )
    return []


def _split_keywords(value: object) -> list[str]:
    raw = _text(value)
    if not raw:
        return []
    parts = [part.strip() for part in re.split(r"[,;|\r\n]+", raw)]
    return _unique_preserving_order([part for part in parts if part])


def _warn_duplicate_real(
    duplicate_keys: set[tuple[str, str, str]],
    row: dict[str, str],
    account_id: str,
    action: str,
    line_number: int,
    avisos: list[str],
) -> None:
    key = (
        normalize(_text(row.get("descricao_normalizada"))),
        account_id,
        action,
    )
    if not key[0]:
        return
    if key in duplicate_keys:
        avisos.append(
            "Linha "
            f"{line_number}: proposta duplicada para descricao_normalizada "
            f"'{_text(row.get('descricao_normalizada'))}', conta {account_id} "
            f"e acao {action}."
        )
    duplicate_keys.add(key)


def _parse_int_or_text(value: object) -> int | str:
    text = _text(value)
    if not text:
        return ""
    try:
        return int(text)
    except ValueError:
        return text


def _load_chart(plano_contas_path: Path) -> ChartOfAccounts:
    try:
        chart = ChartOfAccounts.load(plano_contas_path)
    except Exception as exc:
        raise ClassificationRuleProposalError(
            f"Plano de contas nao pode ser lido: {plano_contas_path}"
        ) from exc
    errors = chart.validate()
    if errors:
        raise ClassificationRuleProposalError(
            "Plano de contas invalido: " + "; ".join(errors)
        )
    return chart


def _read_csv_rows(csv_path: Path) -> list[dict[str, str]]:
    try:
        with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
            reader = csv.DictReader(file, delimiter=";")
            return [dict(row) for row in reader]
    except UnicodeDecodeError as exc:
        raise ClassificationRuleProposalError(
            f"CSV de revisao nao pode ser lido em UTF-8: {csv_path}"
        ) from exc


def _row_status(row: dict[str, str], chart: ChartOfAccounts) -> str:
    raw_account_id = _text(row.get("conta_id_sugerida"))
    raw_action = _text(row.get("acao_manual"))
    if not raw_account_id or not raw_action:
        return "PENDENTE"

    account_id = _normalize_account_id(raw_account_id)
    action = _normalize_action(raw_action)
    try:
        chart.get_by_id(account_id)
    except KeyError:
        return "INVALIDA"
    if action not in ALLOWED_ACTIONS:
        return "INVALIDA"
    return "VALIDA"


def _proposal_terms(row: dict[str, str]) -> list[str]:
    raw_keywords = _text(row.get("palavras_chave_sugeridas"))
    if raw_keywords:
        terms = [
            term.strip()
            for piece in raw_keywords.splitlines()
            for term in piece.split(",")
            if term.strip()
        ]
        if terms:
            return _unique_preserving_order(terms)

    fallback = _text(row.get("descricao_normalizada")) or _text(
        row.get("descricao_exemplo")
    )
    return [fallback] if fallback else []


def _load_existing_rule_terms(regras_produtos_path: Path | None) -> set[str]:
    if regras_produtos_path is None or not regras_produtos_path.exists():
        return set()
    try:
        data = yaml.safe_load(regras_produtos_path.read_text(encoding="utf-8")) or {}
    except Exception as exc:
        raise ClassificationRuleProposalError(
            f"Regras de produtos nao podem ser lidas: {regras_produtos_path}"
        ) from exc

    terms: set[str] = set()
    for rule in data.get("regras", []):
        for alias in rule.get("aliases", []) or []:
            terms.add(normalize(str(alias)))
        for alias in rule.get("evitar_aliases", []) or []:
            terms.add(normalize(str(alias)))
    for rule in data.get("regras_por_fornecedor", []):
        term = rule.get("produto_normalizado_contem")
        if term:
            terms.add(normalize(str(term)))
    return {term for term in terms if term}


def _count_by_type(resultado: ResultadoPropostaClassificacao, tipo: str) -> int:
    return sum(1 for proposta in resultado.propostas if proposta.tipo == tipo)


def _normalize_action(value: object) -> str:
    return _text(value).upper()


def _normalize_account_id(value: str) -> str:
    return "".join(value.split())


def _text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _unique_preserving_order(values: list[str]) -> list[str]:
    seen = set()
    result = []
    for value in values:
        normalized = normalize(value)
        if normalized and normalized not in seen:
            result.append(value)
            seen.add(normalized)
    return result
