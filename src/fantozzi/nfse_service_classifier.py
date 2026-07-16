from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
import hashlib
from pathlib import Path
from typing import Iterable

import yaml

from fantozzi.chart_of_accounts import ChartOfAccounts
from fantozzi.models import Conta, Lancamento
from fantozzi.nfse_importer import NFSE, parse_nfse
from fantozzi.normalization import normalize, normalize_cnpj


ORIGEM_XML_NFSE = "XML_NFSE"
STATUS_CLASSIFICADO_XML_NFSE = "CLASSIFICADO_XML_NFSE"
STATUS_REVISAR_SEM_REGRA_NFSE = "REVISAR_SEM_REGRA_NFSE"
STATUS_REVISAR_MULTIPLAS_REGRAS_NFSE = "REVISAR_MULTIPLAS_REGRAS_NFSE"
STATUS_DUPLICADO_XML_NFSE = "DUPLICADO_XML_NFSE"
STATUS_REVISAR_CHAVE_NFSE_INSUFICIENTE = "REVISAR_CHAVE_NFSE_INSUFICIENTE"
STATUS_COMPROVANTE_SOLO_RECONCILIACAO = "COMPROVANTE_SOLO_RECONCILIACAO"
STATUS_REVISAR_SEM_XML_AUTORIZZATO = "REVISAR_SEM_XML_AUTORIZZATO"
STATUS_REVISAR_FORNECEDOR_XML_ONLY_BLOCCATO = "REVISAR_FORNECEDOR_XML_ONLY_BLOCCATO"

GENERIC_DISCRIMINATION_TERMS = {
    "SERVICO",
    "SERVIÇO",
    "CONSULTORIA",
    "MANUTENCAO",
    "MANUTENÇÃO",
    "ASSESSORIA",
    "HONORARIOS",
    "HONORÁRIOS",
}


class NfseServiceRulesError(ValueError):
    pass


@dataclass(frozen=True)
class DreInfo:
    linha_id: str
    linha_nome: str
    grupo: str
    entra_dre: bool


@dataclass(frozen=True)
class CostCenter:
    id: str
    nome: str


@dataclass(frozen=True)
class NfseServiceRule:
    regra_id: str
    cnpj_prestador: str
    cnpj_tomador: str
    discriminacao_contem: tuple[str, ...]
    conta: Conta
    dre: DreInfo
    centro_custo: CostCenter
    codigo_verificacao: str = ""
    xml_only: bool = True
    precedencia: str = "XML"
    assenza_xml: str = STATUS_REVISAR_SEM_XML_AUTORIZZATO
    fallback_comprovante: str = "VIETATO"
    diferenca_irrisoria_max: Decimal = Decimal("0")
    observacao: str = ""


@dataclass(frozen=True)
class XmlOnlyPolicy:
    policy_id: str
    cnpj_prestador: str
    fornecedor: str
    escopo: str
    assenza_xml: str
    fallback_comprovante: str


@dataclass(frozen=True)
class NfseServiceRuleSet:
    rules: tuple[NfseServiceRule, ...]
    policies: tuple[XmlOnlyPolicy, ...] = ()

    def policies_for_receipt(self, receipt: dict[str, object]) -> tuple[XmlOnlyPolicy, ...]:
        doc = normalize_cnpj(str(_first_value(receipt, "beneficiary_document", "cnpj_beneficiario")))
        name = normalize(str(_first_value(receipt, "beneficiary_name", "beneficiario_normalizado")))
        matches: list[XmlOnlyPolicy] = []
        for policy in self.policies:
            if doc and doc == policy.cnpj_prestador:
                matches.append(policy)
                continue
            if name and normalize(policy.fornecedor) in name:
                matches.append(policy)
        return tuple(matches)


@dataclass(frozen=True)
class NfseServiceClassification:
    status: str
    nfse_key: str
    regra_id: str = ""
    motivo: str = ""
    regras_candidatas: tuple[str, ...] = ()
    lancamento: Lancamento | None = None


@dataclass(frozen=True)
class NfseServiceBatchResult:
    lancamentos: tuple[Lancamento, ...]
    diagnostics: tuple[NfseServiceClassification, ...]


@dataclass(frozen=True)
class XmlOnlyReceiptDiagnostic:
    status: str
    policy_id: str
    cnpj_prestador: str
    motivo: str
    comprovante_valor: Decimal | None = None
    xml_valor: Decimal | None = None
    diferenca: Decimal | None = None
    diferenca_efeito_contabil: str = "NENHUM"
    nfse_key: str = ""
    regra_id: str = ""


def load_nfse_service_rules(
    rules_path: Path | str,
    plano_contas_path: Path | str,
    mapeamento_dre_path: Path | str,
) -> NfseServiceRuleSet:
    rules_path = Path(rules_path)
    with rules_path.open(encoding="utf-8") as file:
        data = yaml.safe_load(file) or {}

    chart = ChartOfAccounts.load(plano_contas_path)
    mapeamento_data = _load_yaml_mapping(Path(mapeamento_dre_path))
    dre_linhas = _dre_linhas_by_id(mapeamento_data)
    conta_mapping = mapeamento_data.get("mapeamento") or {}
    centros = _cost_centers_by_id(Path(plano_contas_path))

    rules = tuple(
        _parse_rule(
            raw_rule,
            index,
            chart=chart,
            conta_mapping=conta_mapping,
            dre_linhas=dre_linhas,
            centros=centros,
        )
        for index, raw_rule in enumerate(data.get("regras") or [], start=1)
    )
    policies = tuple(
        _parse_policy(raw_policy, index)
        for index, raw_policy in enumerate(data.get("politicas_xml_only") or [], start=1)
    )
    rule_set = NfseServiceRuleSet(rules=rules, policies=policies)
    _validate_rule_set(rule_set)
    return rule_set


def classify_nfse_service(
    nfse: NFSE,
    rule_set: NfseServiceRuleSet | Iterable[NfseServiceRule],
) -> NfseServiceClassification:
    rules = tuple(rule_set.rules if isinstance(rule_set, NfseServiceRuleSet) else rule_set)
    nfse_key = nfse_dedup_key(nfse)
    if not nfse_key:
        return NfseServiceClassification(
            status=STATUS_REVISAR_CHAVE_NFSE_INSUFICIENTE,
            nfse_key="",
            motivo="CHAVE_NFSE_INSUFICIENTE",
        )
    candidates = tuple(rule for rule in rules if _matches_rule(nfse, rule))
    candidate_ids = tuple(rule.regra_id for rule in candidates)

    if not candidates:
        return NfseServiceClassification(
            status=STATUS_REVISAR_SEM_REGRA_NFSE,
            nfse_key=nfse_key,
            motivo="SEM_REGRA_NFSE_SERVICO",
        )
    if len(candidates) > 1:
        return NfseServiceClassification(
            status=STATUS_REVISAR_MULTIPLAS_REGRAS_NFSE,
            nfse_key=nfse_key,
            motivo="MAIS_DE_UMA_REGRA_NFSE_SERVICO",
            regras_candidatas=candidate_ids,
        )

    rule = candidates[0]
    lancamento = _lancamento_from_nfse(nfse, rule, nfse_key)
    return NfseServiceClassification(
        status=STATUS_CLASSIFICADO_XML_NFSE,
        nfse_key=nfse_key,
        regra_id=rule.regra_id,
        motivo="REGRA_NFSE_SERVICO_XML_ONLY",
        regras_candidatas=candidate_ids,
        lancamento=lancamento,
    )


def classify_nfse_service_file(
    xml_path: Path | str,
    rule_set: NfseServiceRuleSet | Iterable[NfseServiceRule],
) -> NfseServiceClassification:
    return classify_nfse_service(parse_nfse(xml_path), rule_set)


def classify_nfse_service_files(
    xml_paths: Iterable[Path | str],
    rule_set: NfseServiceRuleSet | Iterable[NfseServiceRule],
) -> NfseServiceBatchResult:
    diagnostics: list[NfseServiceClassification] = []
    lancamentos: list[Lancamento] = []
    seen: set[str] = set()

    for xml_path in xml_paths:
        result = classify_nfse_service_file(xml_path, rule_set)
        if result.status == STATUS_CLASSIFICADO_XML_NFSE and result.nfse_key in seen:
            diagnostics.append(
                NfseServiceClassification(
                    status=STATUS_DUPLICADO_XML_NFSE,
                    nfse_key=result.nfse_key,
                    regra_id=result.regra_id,
                    motivo="NFSE_JA_CLASSIFICADA_NESTE_LOTE",
                    regras_candidatas=result.regras_candidatas,
                )
            )
            continue
        diagnostics.append(result)
        if result.lancamento is not None:
            seen.add(result.nfse_key)
            lancamentos.append(result.lancamento)

    return NfseServiceBatchResult(lancamentos=tuple(lancamentos), diagnostics=tuple(diagnostics))


def reconcile_xml_only_receipt(
    receipt: dict[str, object],
    classifications: Iterable[NfseServiceClassification],
    rule_set: NfseServiceRuleSet,
) -> XmlOnlyReceiptDiagnostic:
    policies = rule_set.policies_for_receipt(receipt)
    if not policies:
        return XmlOnlyReceiptDiagnostic(
            status=STATUS_REVISAR_SEM_XML_AUTORIZZATO,
            policy_id="",
            cnpj_prestador="",
            motivo="FORNECEDOR_NAO_XML_ONLY",
        )

    policy = policies[0]
    if policy.escopo == "BLOCCATO_MULTI_XML_005E_B":
        return XmlOnlyReceiptDiagnostic(
            status=STATUS_REVISAR_FORNECEDOR_XML_ONLY_BLOCCATO,
            policy_id=policy.policy_id,
            cnpj_prestador=policy.cnpj_prestador,
            motivo="POLICY_XML_ONLY_BLOCCATA_005E_B",
        )

    xml_matches = [
        result
        for result in classifications
        if result.lancamento is not None
        and normalize_cnpj(result.lancamento.cnpj_cpf_favorecido) == policy.cnpj_prestador
    ]
    if not xml_matches:
        return XmlOnlyReceiptDiagnostic(
            status=policy.assenza_xml,
            policy_id=policy.policy_id,
            cnpj_prestador=policy.cnpj_prestador,
            motivo="XML_CLASSIFICABILE_ASSENTE",
        )

    best = xml_matches[0]
    xml_valor = best.lancamento.valor if best.lancamento else None
    comprovante_valor = _receipt_amount(receipt)
    diferenca = (
        abs(xml_valor - comprovante_valor)
        if xml_valor is not None and comprovante_valor is not None
        else None
    )
    return XmlOnlyReceiptDiagnostic(
        status=STATUS_COMPROVANTE_SOLO_RECONCILIACAO,
        policy_id=policy.policy_id,
        cnpj_prestador=policy.cnpj_prestador,
        motivo="XML_CLASSIFICADO_PREVALE_COMPROVANTE_SOLO_RECONCILIACAO",
        comprovante_valor=comprovante_valor,
        xml_valor=xml_valor,
        diferenca=diferenca,
        diferenca_efeito_contabil="NENHUM",
        nfse_key=best.nfse_key,
        regra_id=best.regra_id,
    )


def nfse_dedup_key(nfse: NFSE) -> str:
    if nfse.chave_acesso:
        return f"NFSE:CHAVE:{nfse.chave_acesso}"

    value = _nfse_value(nfse)
    parts = (
        nfse.cnpj_prestador,
        nfse.cnpj_tomador,
        nfse.numero,
        nfse.data_emissao.isoformat() if nfse.data_emissao else "",
        f"{value:.2f}" if value else "",
    )
    if any(not part for part in parts):
        return ""
    return "NFSE:COMPOSTA:" + "|".join(parts)


def _parse_rule(
    data: object,
    index: int,
    *,
    chart: ChartOfAccounts,
    conta_mapping: dict[str, object],
    dre_linhas: dict[str, DreInfo],
    centros: dict[str, CostCenter],
) -> NfseServiceRule:
    if not isinstance(data, dict):
        raise NfseServiceRulesError(f"Regra NFSE #{index} deve ser um objeto.")

    regra_id = str(data.get("regra_id") or f"NFSE_SERVICO_{index:03d}").strip()
    conta_id = str(data.get("conta_id") or "").strip()
    dre_linha_id = str(data.get("dre_linha_id") or "").strip()
    centro_custo_id = str(data.get("centro_custo_id") or "").strip()
    aliases = tuple(str(item).strip() for item in data.get("discriminacao_contem") or [] if str(item).strip())

    if not conta_id or not dre_linha_id or not centro_custo_id:
        raise NfseServiceRulesError(f"{regra_id}: conta_id, dre_linha_id e centro_custo_id sao obrigatorios.")
    if str(data.get("tipo_documento") or "").upper() != "NFSE":
        raise NfseServiceRulesError(f"{regra_id}: tipo_documento deve ser NFSE.")
    if not aliases:
        raise NfseServiceRulesError(f"{regra_id}: discriminacao_contem deve ter ao menos um termo.")
    _validate_aliases(regra_id, aliases)
    if str(data.get("precedencia") or "").upper() != "XML":
        raise NfseServiceRulesError(f"{regra_id}: precedencia deve ser XML.")
    if data.get("xml_only") is not True:
        raise NfseServiceRulesError(f"{regra_id}: xml_only deve ser true.")
    if str(data.get("fallback_comprovante") or "").upper() != "VIETATO":
        raise NfseServiceRulesError(f"{regra_id}: fallback_comprovante deve ser VIETATO.")

    conta = chart.get_by_id(conta_id)
    mapping = conta_mapping.get(conta_id)
    if not isinstance(mapping, dict):
        raise NfseServiceRulesError(f"{regra_id}: conta_id sem mapeamento DRE: {conta_id}.")
    mapped_dre = str(mapping.get("dre_linha_id") or "").strip()
    if mapped_dre != dre_linha_id:
        raise NfseServiceRulesError(
            f"{regra_id}: dre_linha_id {dre_linha_id} diverge do mapeamento atual {mapped_dre}."
        )
    if dre_linha_id not in dre_linhas:
        raise NfseServiceRulesError(f"{regra_id}: linha DRE inexistente: {dre_linha_id}.")
    if centro_custo_id not in centros:
        raise NfseServiceRulesError(f"{regra_id}: centro_custo_id inexistente: {centro_custo_id}.")

    reconciliacao = data.get("reconciliacao") if isinstance(data.get("reconciliacao"), dict) else {}
    return NfseServiceRule(
        regra_id=regra_id,
        cnpj_prestador=normalize_cnpj(str(data.get("cnpj_prestador") or "")),
        cnpj_tomador=normalize_cnpj(str(data.get("cnpj_tomador") or "")),
        discriminacao_contem=aliases,
        conta=conta,
        dre=dre_linhas[dre_linha_id],
        centro_custo=centros[centro_custo_id],
        codigo_verificacao=str(data.get("codigo_verificacao") or "").strip(),
        assenza_xml=str(data.get("assenza_xml") or STATUS_REVISAR_SEM_XML_AUTORIZZATO),
        diferenca_irrisoria_max=_decimal_or_zero(reconciliacao.get("diferenca_irrisoria_max")),
        observacao=str(data.get("observacao") or "").strip(),
    )


def _validate_rule_set(rule_set: NfseServiceRuleSet) -> None:
    rule_ids: set[str] = set()
    policy_ids: set[str] = set()
    policy_cnpjs: set[str] = set()
    for rule in rule_set.rules:
        if rule.regra_id in rule_ids:
            raise NfseServiceRulesError(f"regra_id duplicado: {rule.regra_id}")
        rule_ids.add(rule.regra_id)
    for policy in rule_set.policies:
        if policy.policy_id in policy_ids:
            raise NfseServiceRulesError(f"policy_id duplicado: {policy.policy_id}")
        if policy.cnpj_prestador in policy_cnpjs:
            raise NfseServiceRulesError(f"politica XML-only duplicada para CNPJ: {policy.cnpj_prestador}")
        policy_ids.add(policy.policy_id)
        policy_cnpjs.add(policy.cnpj_prestador)


def _validate_aliases(regra_id: str, aliases: tuple[str, ...]) -> None:
    normalized = tuple(normalize(alias) for alias in aliases)
    for alias in normalized:
        if len(alias) < 4 or alias in {normalize(term) for term in GENERIC_DISCRIMINATION_TERMS}:
            raise NfseServiceRulesError(f"{regra_id}: discriminacao_contem generica demais: {alias}.")


def _parse_policy(data: object, index: int) -> XmlOnlyPolicy:
    if not isinstance(data, dict):
        raise NfseServiceRulesError(f"Politica XML-only #{index} deve ser um objeto.")
    policy_id = str(data.get("policy_id") or f"XML_ONLY_{index:03d}").strip()
    cnpj = normalize_cnpj(str(data.get("cnpj_prestador") or ""))
    if not cnpj:
        raise NfseServiceRulesError(f"{policy_id}: cnpj_prestador obrigatorio.")
    fallback = str(data.get("fallback_comprovante") or "").strip().upper()
    if fallback != "VIETATO":
        raise NfseServiceRulesError(f"{policy_id}: fallback_comprovante deve essere VIETATO.")
    return XmlOnlyPolicy(
        policy_id=policy_id,
        cnpj_prestador=cnpj,
        fornecedor=str(data.get("fornecedor") or "").strip(),
        escopo=str(data.get("escopo") or "").strip(),
        assenza_xml=str(data.get("assenza_xml") or STATUS_REVISAR_SEM_XML_AUTORIZZATO),
        fallback_comprovante=fallback,
    )


def _matches_rule(nfse: NFSE, rule: NfseServiceRule) -> bool:
    if nfse.tipo_documento != "NFSE":
        return False
    if nfse.cnpj_prestador != rule.cnpj_prestador:
        return False
    if nfse.cnpj_tomador != rule.cnpj_tomador:
        return False
    if rule.codigo_verificacao and nfse.codigo_verificacao != rule.codigo_verificacao:
        return False
    discriminacao = normalize(nfse.discriminacao)
    return any(normalize(alias) in discriminacao for alias in rule.discriminacao_contem)


def _lancamento_from_nfse(nfse: NFSE, rule: NfseServiceRule, nfse_key: str) -> Lancamento:
    value = _nfse_value(nfse)
    lancamento = Lancamento(
        data_documento=nfse.data_emissao,
        competencia=(nfse.competencia or nfse.data_emissao).strftime("%Y-%m"),
        cnpj_empresa=nfse.cnpj_tomador,
        nome_empresa=nfse.razao_social_tomador,
        fornecedor_ou_favorecido=nfse.razao_social_prestador,
        cnpj_cpf_favorecido=nfse.cnpj_prestador,
        documento=nfse.numero,
        tipo_documento="NFSE",
        conta_id=rule.conta.id,
        conta_nome=rule.conta.nome,
        grupo_conta=rule.conta.grupo,
        descricao_itens=nfse.discriminacao,
        valor=value,
        origem=ORIGEM_XML_NFSE,
        arquivo_origem=nfse.arquivo_origem,
        chave_origem=nfse_key,
        status_revisao="OK",
        alertas="",
        observacoes=(
            f"regra_nfse={rule.regra_id}; fonte=XML_NFSE; "
            "comprovante=somente_reconciliacao; fallback_comprovante=VIETATO"
        ),
        centro_custo=f"{rule.centro_custo.id} - {rule.centro_custo.nome}",
        tipo_lancamento=rule.conta.tipo,
        dre_linha_id=rule.dre.linha_id,
        dre_linha_nome=rule.dre.linha_nome,
    )
    lancamento.hash_registro = _hash_lancamento(lancamento)
    return lancamento


def _nfse_value(nfse: NFSE) -> Decimal:
    return nfse.valor_liquido or nfse.valor_servicos


def _hash_lancamento(lancamento: Lancamento) -> str:
    payload = "|".join(
        (
            lancamento.origem,
            lancamento.chave_origem,
            lancamento.conta_id,
            f"{lancamento.valor:.2f}",
        )
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _load_yaml_mapping(path: Path) -> dict[str, object]:
    with path.open(encoding="utf-8") as file:
        data = yaml.safe_load(file) or {}
    if not isinstance(data, dict):
        raise NfseServiceRulesError(f"YAML invalido: {path}")
    return data


def _dre_linhas_by_id(data: dict[str, object]) -> dict[str, DreInfo]:
    lines: dict[str, DreInfo] = {}
    items = data.get("linhas_dre") or data.get("linhas") or []
    for item in items:
        if not isinstance(item, dict):
            continue
        line_id = str(item.get("id") or "").strip()
        if not line_id:
            continue
        lines[line_id] = DreInfo(
            linha_id=line_id,
            linha_nome=str(item.get("nome") or "").strip(),
            grupo=str(item.get("grupo") or "").strip(),
            entra_dre=not bool(item.get("calculada", False)),
        )
    return lines


def _cost_centers_by_id(plano_path: Path) -> dict[str, CostCenter]:
    with plano_path.open(encoding="utf-8") as file:
        data = yaml.safe_load(file) or {}
    centers: dict[str, CostCenter] = {}
    for item in data.get("centros_custo") or []:
        if not isinstance(item, dict):
            continue
        center_id = str(item.get("id") or "").strip()
        if center_id:
            centers[center_id] = CostCenter(id=center_id, nome=str(item.get("nome") or "").strip())
    return centers


def _decimal_or_zero(value: object) -> Decimal:
    try:
        return Decimal(str(value or "0"))
    except Exception:
        return Decimal("0")


def _receipt_amount(receipt: dict[str, object]) -> Decimal | None:
    value = _first_value(receipt, "amount", "valor", "valor_comprovante")
    if value in (None, ""):
        return None
    text = str(value).replace(".", "").replace(",", ".") if "," in str(value) else str(value)
    try:
        return Decimal(text)
    except Exception:
        return None


def _first_value(data: dict[str, object], *keys: str) -> object:
    for key in keys:
        value = data.get(key)
        if value not in (None, ""):
            return value
    return ""
