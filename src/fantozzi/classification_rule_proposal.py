from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

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
