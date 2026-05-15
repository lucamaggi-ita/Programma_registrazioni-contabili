from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path
from types import MappingProxyType
from typing import Iterable, Mapping

from fantozzi.classification_review_workbook import GENERIC_KEYWORDS
from fantozzi.normalization import normalize


PROMOTABLE_ACTIONS = frozenset({"CLASSIFICAR", "CRIAR_ALIAS", "CRIAR_REGRA"})
BLOCKED_ACTIONS = frozenset({"MANTER_REVISAO", "ANALISAR_COM_CONTADOR"})
IGNORED_ACTIONS = frozenset({"IGNORAR"})

DEFAULT_BLOCKED_NF_NUMBERS = frozenset({"359680"})

EXPLICIT_FASE9_GENERIC_KEYWORDS = frozenset(
    {
        "PAPEL",
        "MOLHO",
        "SACO",
        "CABO",
        "FIO",
        "LIMAO",
        "MARACUJA",
        "FACA",
        "PALETA",
        "IMPRESSORA",
        "ESPELHO",
    }
)


class Fase9DecisionReason(StrEnum):
    PROMOTION_CANDIDATE = "PROMOTION_CANDIDATE"
    ACAO_NAO_PROMOVE = "ACAO_NAO_PROMOVE"
    MANTER_REVISAO = "MANTER_REVISAO"
    ANALISAR_COM_CONTADOR = "ANALISAR_COM_CONTADOR"
    IGNORAR = "IGNORAR"
    CONTA_INVALIDA = "CONTA_INVALIDA"
    CONTA_AUSENTE = "CONTA_AUSENTE"
    CONTA_NOME_INCOERENTE = "CONTA_NOME_INCOERENTE"
    PALAVRAS_GENERICAS = "PALAVRAS_GENERICAS"
    PALAVRAS_AUSENTES = "PALAVRAS_AUSENTES"
    ALIAS_GENERICO = "ALIAS_GENERICO"
    ALERTA_NAO_RESOLVIDO = "ALERTA_NAO_RESOLVIDO"
    COMODATO_FORA_DRE = "COMODATO_FORA_DRE"
    CFOP_1949_CONTROLE_MANUAL = "CFOP_1949_CONTROLE_MANUAL"
    NF_359680_CONTROLE_MANUAL = "NF_359680_CONTROLE_MANUAL"
    JA_OK = "JA_OK"
    STATUS_NAO_REVISAR = "STATUS_NAO_REVISAR"


@dataclass(frozen=True)
class Fase9Decision:
    line_number: int
    action: str
    description: str
    descricao_normalizada: str
    supplier: str
    document: str
    account_id: str
    account_name: str
    keywords: tuple[str, ...]
    status_revisao: str
    alertas: str
    motivo_revisao: str
    tipo_operacao_especial: str
    tratamento_operacao_especial: str
    cfop: str
    numero_nota: str
    chave_nfe: str
    cnpj: str
    valor: str
    source_row: tuple[tuple[str, str], ...] = ()


@dataclass(frozen=True)
class Fase9PromotionCandidate:
    decision: Fase9Decision
    reason: Fase9DecisionReason
    action: str
    account_id: str
    account_name: str
    description: str
    rule_terms: tuple[str, ...]


@dataclass(frozen=True)
class Fase9BlockedDecision:
    decision: Fase9Decision
    reason: Fase9DecisionReason
    action: str
    details: str = ""


@dataclass(frozen=True)
class Fase9DecisionResult:
    promotions: tuple[Fase9PromotionCandidate, ...]
    blocked: tuple[Fase9BlockedDecision, ...]
    ignored: tuple[Fase9BlockedDecision, ...]


def load_fase9_decisions_from_csv(csv_path: Path | str) -> tuple[Fase9Decision, ...]:
    """Le revisao_classificacao_revisada.csv sem alterar o arquivo."""
    csv_path = Path(csv_path)
    with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
        sample = file.read(4096)
        file.seek(0)
        dialect = csv.Sniffer().sniff(sample, delimiters=";,") if sample else csv.excel
        reader = csv.DictReader(file, dialect=dialect)
        return tuple(
            _decision_from_row(_clean_row(row), line_number)
            for line_number, row in enumerate(reader, start=2)
        )


def evaluate_fase9_decisions(
    decisions: Iterable[Fase9Decision],
    *,
    valid_account_ids: set[str] | frozenset[str],
    generic_keywords: set[str] | frozenset[str] | None = None,
    blocked_nf_numbers: set[str] | frozenset[str] | None = None,
    account_names_by_id: Mapping[str, str] | None = None,
) -> Fase9DecisionResult:
    """Calcula promocoes candidatas e bloqueios, sem escrever arquivos."""
    normalized_accounts = frozenset(_normalize_account_id(item) for item in valid_account_ids)
    normalized_generic_keywords = _normalized_generic_keywords(generic_keywords)
    normalized_blocked_nfs = frozenset(
        _text(item) for item in (blocked_nf_numbers or DEFAULT_BLOCKED_NF_NUMBERS)
    )
    account_names = MappingProxyType(dict(account_names_by_id or {}))

    promotions: list[Fase9PromotionCandidate] = []
    blocked: list[Fase9BlockedDecision] = []
    ignored: list[Fase9BlockedDecision] = []

    for decision in tuple(decisions):
        block_reason = _blocking_reason(
            decision,
            valid_account_ids=normalized_accounts,
            generic_keywords=normalized_generic_keywords,
            blocked_nf_numbers=normalized_blocked_nfs,
            account_names_by_id=account_names,
        )
        if block_reason is None:
            promotions.append(_promotion_from_decision(decision))
            continue
        blocked_decision = Fase9BlockedDecision(
            decision=decision,
            reason=block_reason,
            action=decision.action,
            details=_blocked_details(decision, block_reason),
        )
        if block_reason is Fase9DecisionReason.IGNORAR:
            ignored.append(blocked_decision)
        else:
            blocked.append(blocked_decision)

    return Fase9DecisionResult(
        promotions=tuple(promotions),
        blocked=tuple(blocked),
        ignored=tuple(ignored),
    )


def _blocking_reason(
    decision: Fase9Decision,
    *,
    valid_account_ids: frozenset[str],
    generic_keywords: frozenset[str],
    blocked_nf_numbers: frozenset[str],
    account_names_by_id: Mapping[str, str],
) -> Fase9DecisionReason | None:
    if decision.numero_nota in blocked_nf_numbers or decision.document in blocked_nf_numbers:
        return Fase9DecisionReason.NF_359680_CONTROLE_MANUAL
    if decision.cfop == "1949":
        return Fase9DecisionReason.CFOP_1949_CONTROLE_MANUAL
    if (
        decision.tipo_operacao_especial == "COMODATO"
        or decision.tratamento_operacao_especial == "FORA_DRE"
    ):
        return Fase9DecisionReason.COMODATO_FORA_DRE
    if decision.status_revisao == "OK":
        return Fase9DecisionReason.JA_OK
    if decision.status_revisao != "REVISAR":
        return Fase9DecisionReason.STATUS_NAO_REVISAR
    if _has_unresolved_alert(decision):
        return Fase9DecisionReason.ALERTA_NAO_RESOLVIDO
    if decision.action == "MANTER_REVISAO":
        return Fase9DecisionReason.MANTER_REVISAO
    if decision.action == "ANALISAR_COM_CONTADOR":
        return Fase9DecisionReason.ANALISAR_COM_CONTADOR
    if decision.action == "IGNORAR":
        return Fase9DecisionReason.IGNORAR
    if decision.action not in PROMOTABLE_ACTIONS:
        return Fase9DecisionReason.ACAO_NAO_PROMOVE
    if not decision.account_id:
        return Fase9DecisionReason.CONTA_AUSENTE
    if decision.account_id not in valid_account_ids:
        return Fase9DecisionReason.CONTA_INVALIDA
    if _has_incoherent_account_name(decision, account_names_by_id):
        return Fase9DecisionReason.CONTA_NOME_INCOERENTE
    if decision.action == "CRIAR_REGRA":
        if not decision.keywords:
            return Fase9DecisionReason.PALAVRAS_AUSENTES
        if _all_terms_generic(decision.keywords, generic_keywords):
            return Fase9DecisionReason.PALAVRAS_GENERICAS
    if decision.action == "CRIAR_ALIAS" and _is_generic_alias(decision, generic_keywords):
        return Fase9DecisionReason.ALIAS_GENERICO
    return None


def _promotion_from_decision(decision: Fase9Decision) -> Fase9PromotionCandidate:
    rule_terms = decision.keywords
    if decision.action in {"CLASSIFICAR", "CRIAR_ALIAS"} and not rule_terms:
        rule_terms = (decision.descricao_normalizada or decision.description,)
    return Fase9PromotionCandidate(
        decision=decision,
        reason=Fase9DecisionReason.PROMOTION_CANDIDATE,
        action=decision.action,
        account_id=decision.account_id,
        account_name=decision.account_name,
        description=decision.description,
        rule_terms=tuple(term for term in rule_terms if term),
    )


def _decision_from_row(row: dict[str, str], line_number: int) -> Fase9Decision:
    description = _first_value(row, "descricao", "descricao_exemplo", "descricao_normalizada")
    normalized_description = _first_value(row, "descricao_normalizada", "descricao")
    numero_nota = _first_value(row, "numero_nota", "nota_fiscal", "documentos")
    document = _first_value(row, "documentos", "numero_nota", "nota_fiscal")
    return Fase9Decision(
        line_number=line_number,
        action=_normalize_action(_first_value(row, "acao_manual", "acao")),
        description=description,
        descricao_normalizada=normalized_description,
        supplier=_first_value(row, "fornecedor", "fornecedores"),
        document=document,
        account_id=_normalize_account_id(_first_value(row, "conta_id_sugerida")),
        account_name=_first_value(row, "conta_nome_sugerida"),
        keywords=_split_keywords(_first_value(row, "palavras_chave_sugeridas")),
        status_revisao=_normalize_action(_first_value(row, "status_revisao")),
        alertas=_first_value(row, "alertas"),
        motivo_revisao=_first_value(row, "motivo_revisao"),
        tipo_operacao_especial=_normalize_action(
            _first_value(row, "tipo_operacao_especial")
        ),
        tratamento_operacao_especial=_normalize_action(
            _first_value(row, "tratamento_operacao_especial")
        ),
        cfop=_first_value(row, "cfop"),
        numero_nota=numero_nota,
        chave_nfe=_first_value(row, "chave_nfe"),
        cnpj=_first_value(row, "cnpj"),
        valor=_first_value(row, "valor", "valor_total"),
        source_row=tuple(sorted(row.items())),
    )


def _clean_row(row: Mapping[str, object]) -> dict[str, str]:
    return {_text(key): _text(value) for key, value in row.items() if key is not None}


def _first_value(row: Mapping[str, str], *keys: str) -> str:
    for key in keys:
        value = _text(row.get(key))
        if value:
            return value
    return ""


def _has_unresolved_alert(decision: Fase9Decision) -> bool:
    return bool(_text(decision.alertas) or _text(decision.motivo_revisao))


def _has_incoherent_account_name(
    decision: Fase9Decision,
    account_names_by_id: Mapping[str, str],
) -> bool:
    if not decision.account_name or not account_names_by_id:
        return False
    expected_name = account_names_by_id.get(decision.account_id)
    if not expected_name:
        return False
    return normalize(expected_name) != normalize(decision.account_name)


def _all_terms_generic(terms: tuple[str, ...], generic_keywords: frozenset[str]) -> bool:
    return all(_is_generic_term(term, generic_keywords) for term in terms)


def _is_generic_alias(
    decision: Fase9Decision,
    generic_keywords: frozenset[str],
) -> bool:
    source = decision.descricao_normalizada or decision.description
    terms = _split_alias_terms(source)
    return bool(terms) and all(_is_generic_term(term, generic_keywords) for term in terms)


def _is_generic_term(term: str, generic_keywords: frozenset[str]) -> bool:
    normalized = normalize(term).upper()
    compact = "".join(normalized.split())
    return (
        not compact
        or normalized in generic_keywords
        or compact in generic_keywords
        or len(compact) < 3
        or compact.isnumeric()
    )


def _split_keywords(value: str) -> tuple[str, ...]:
    parts = re.split(r"[,;|\r\n]+", _text(value))
    return tuple(_unique_preserving_order(part.strip() for part in parts if part.strip()))


def _split_alias_terms(value: str) -> tuple[str, ...]:
    return tuple(
        _unique_preserving_order(
            part.strip() for part in re.split(r"[,;|\r\n\s]+", _text(value)) if part.strip()
        )
    )


def _unique_preserving_order(values: Iterable[str]) -> tuple[str, ...]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        key = normalize(value)
        if key and key not in seen:
            seen.add(key)
            result.append(value)
    return tuple(result)


def _normalized_generic_keywords(
    generic_keywords: set[str] | frozenset[str] | None,
) -> frozenset[str]:
    source = set(GENERIC_KEYWORDS) | set(EXPLICIT_FASE9_GENERIC_KEYWORDS)
    if generic_keywords is not None:
        source |= set(generic_keywords)
    return frozenset(normalize(item).upper() for item in source if _text(item))


def _blocked_details(decision: Fase9Decision, reason: Fase9DecisionReason) -> str:
    if reason is Fase9DecisionReason.CONTA_INVALIDA:
        return decision.account_id
    if reason is Fase9DecisionReason.CONTA_NOME_INCOERENTE:
        return decision.account_name
    if reason is Fase9DecisionReason.ALERTA_NAO_RESOLVIDO:
        return decision.alertas or decision.motivo_revisao
    if reason is Fase9DecisionReason.CFOP_1949_CONTROLE_MANUAL:
        return decision.cfop
    if reason is Fase9DecisionReason.NF_359680_CONTROLE_MANUAL:
        return decision.numero_nota or decision.document
    if reason in {
        Fase9DecisionReason.PALAVRAS_GENERICAS,
        Fase9DecisionReason.ALIAS_GENERICO,
    }:
        return "; ".join(decision.keywords) or decision.descricao_normalizada
    return ""


def _normalize_action(value: str) -> str:
    return "".join(_text(value).upper().split())


def _normalize_account_id(value: str) -> str:
    return "".join(_text(value).split())


def _text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()
