from __future__ import annotations

import csv
import hashlib
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable, Mapping

import yaml

from fantozzi.chart_of_accounts import ChartOfAccounts
from fantozzi.excel_writer import write_lancamentos_to_excel
from fantozzi.models import Lancamento


FONTE_MANUAL_GERENCIAL = "MANUAL_GERENCIAL"
IDENTIFICADOR_MANUAL_GERENCIAL = "MANUAL_GERENCIAL"
NOME_EMPRESA_MANUAL_GERENCIAL = "Fantozzi Consolidado Gerencial"

CSV_REQUIRED_HEADERS = (
    "data",
    "descricao_da_despesa",
    "conta_debito_id",
    "conta_debito_nome",
    "conta_credito_id",
    "conta_credito_nome",
    "valor",
)
CSV_OPTIONAL_HEADERS = ("observacao",)
CSV_APLICADO_HEADERS = (
    *CSV_REQUIRED_HEADERS,
    *CSV_OPTIONAL_HEADERS,
    "fonte",
    "status",
    "tipo_lancamento",
    "chave_deduplicacao",
    "linha_origem_csv",
)


class LancamentosManuaisGerenciaisError(Exception):
    pass


@dataclass(frozen=True)
class LancamentoManualGerencial:
    data: date
    descricao_da_despesa: str
    conta_debito_id: str
    conta_debito_nome: str
    conta_credito_id: str
    conta_credito_nome: str
    valor: Decimal
    observacao: str
    chave_deduplicacao: str
    linha_origem_csv: int
    conta_dre_mapeada_count: int
    conta_debito_totaliza_gerencial: bool
    conta_credito_totaliza_gerencial: bool


@dataclass(frozen=True)
class LancamentosManuaisGerenciaisResult:
    input_csv: Path
    output_xlsx: Path
    output_csv: Path
    auditoria_txt: Path
    linhas_lidas: int
    lancamentos_aceitos: int
    linhas_tecnicas_geradas: int
    total_manual: Decimal
    fonte: str = FONTE_MANUAL_GERENCIAL


def carregar_lancamentos_manuais_gerenciais(
    csv_path: Path | str,
    plano_contas_path: Path | str,
    mapeamento_dre_path: Path | str,
) -> tuple[LancamentoManualGerencial, ...]:
    csv_path = Path(csv_path)
    _validate_csv_file(csv_path)
    chart = _load_chart(Path(plano_contas_path))
    dre_account_ids = _load_dre_account_ids(Path(mapeamento_dre_path))

    headers, raw_rows = _read_csv_rows(csv_path)
    _validate_headers(headers)
    if not raw_rows:
        raise LancamentosManuaisGerenciaisError("CSV sem lancamentos validos.")

    lancamentos: list[LancamentoManualGerencial] = []
    seen_keys: dict[str, int] = {}
    for line_number, raw_row in raw_rows:
        if _is_empty_row(raw_row):
            raise LancamentosManuaisGerenciaisError(
                f"Linha {line_number}: linha vazia encontrada no CSV."
            )
        lancamento = _parse_row(
            raw_row,
            line_number=line_number,
            chart=chart,
            dre_account_ids=dre_account_ids,
        )
        previous_line = seen_keys.get(lancamento.chave_deduplicacao)
        if previous_line is not None:
            raise LancamentosManuaisGerenciaisError(
                "Duplicata no CSV manual: "
                f"linhas {previous_line} e {line_number} possuem a mesma chave."
            )
        seen_keys[lancamento.chave_deduplicacao] = line_number
        lancamentos.append(lancamento)

    return tuple(lancamentos)


def converter_lancamentos_manuais_para_lancamentos(
    lancamentos_manuais: Iterable[LancamentoManualGerencial],
    plano_contas_path: Path | str,
    csv_path: Path | str,
) -> tuple[Lancamento, ...]:
    chart = _load_chart(Path(plano_contas_path))
    csv_path = Path(csv_path)
    output: list[Lancamento] = []
    for manual in lancamentos_manuais:
        output.append(_technical_lancamento(manual, chart, csv_path, side="DEBITO"))
        output.append(_technical_lancamento(manual, chart, csv_path, side="CREDITO"))
    return tuple(output)


def processar_lancamentos_manuais_gerenciais(
    csv_path: Path | str,
    plano_contas_path: Path | str,
    mapeamento_dre_path: Path | str,
    output_xlsx_path: Path | str,
    output_csv_path: Path | str,
    auditoria_txt_path: Path | str,
) -> LancamentosManuaisGerenciaisResult:
    csv_path = Path(csv_path)
    output_xlsx_path = Path(output_xlsx_path)
    output_csv_path = Path(output_csv_path)
    auditoria_txt_path = Path(auditoria_txt_path)

    manuais = carregar_lancamentos_manuais_gerenciais(
        csv_path,
        plano_contas_path,
        mapeamento_dre_path,
    )
    tecnicos = converter_lancamentos_manuais_para_lancamentos(
        manuais,
        plano_contas_path,
        csv_path,
    )

    write_lancamentos_to_excel(output_xlsx_path, tecnicos)
    _write_csv_aplicado(output_csv_path, manuais)
    _write_auditoria(auditoria_txt_path, csv_path, manuais)

    return LancamentosManuaisGerenciaisResult(
        input_csv=csv_path,
        output_xlsx=output_xlsx_path,
        output_csv=output_csv_path,
        auditoria_txt=auditoria_txt_path,
        linhas_lidas=len(manuais),
        lancamentos_aceitos=len(manuais),
        linhas_tecnicas_geradas=len(tecnicos),
        total_manual=sum((item.valor for item in manuais), Decimal("0.00")).quantize(
            Decimal("0.01")
        ),
    )


def _validate_csv_file(csv_path: Path) -> None:
    if not csv_path.exists():
        raise LancamentosManuaisGerenciaisError(
            f"CSV de lancamentos manuais nao encontrado: {csv_path}"
        )
    if not csv_path.is_file():
        raise LancamentosManuaisGerenciaisError(
            f"CSV de lancamentos manuais nao e um arquivo: {csv_path}"
        )


def _load_chart(plano_contas_path: Path) -> ChartOfAccounts:
    try:
        chart = ChartOfAccounts.load(plano_contas_path)
    except Exception as exc:
        raise LancamentosManuaisGerenciaisError(
            f"Plano de contas nao pode ser lido: {plano_contas_path}"
        ) from exc
    errors = chart.validate()
    if errors:
        raise LancamentosManuaisGerenciaisError(
            "Plano de contas invalido: " + "; ".join(errors)
        )
    return chart


def _load_dre_account_ids(mapeamento_dre_path: Path) -> frozenset[str]:
    if not mapeamento_dre_path.exists():
        raise LancamentosManuaisGerenciaisError(
            f"Mapeamento DRE nao encontrado: {mapeamento_dre_path}"
        )
    try:
        with mapeamento_dre_path.open("r", encoding="utf-8") as file:
            data = yaml.safe_load(file) or {}
    except Exception as exc:
        raise LancamentosManuaisGerenciaisError(
            f"Mapeamento DRE nao pode ser lido: {mapeamento_dre_path}"
        ) from exc
    mapeamento = data.get("mapeamento", {}) or {}
    if not isinstance(mapeamento, dict):
        raise LancamentosManuaisGerenciaisError(
            "Mapeamento DRE invalido: campo mapeamento."
        )
    account_ids: set[str] = set()
    for account_id, row in mapeamento.items():
        if not isinstance(row, dict):
            continue
        if row.get("entra_dre", True) and _text(row.get("dre_linha_id")):
            account_ids.add(_text(account_id))
    return frozenset(account_ids)


def _read_csv_rows(
    csv_path: Path,
) -> tuple[tuple[str, ...], tuple[tuple[int, dict[str, str]], ...]]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.reader(file, delimiter=";", quotechar='"', doublequote=True)
        try:
            headers = tuple(_text(header) for header in next(reader))
        except StopIteration:
            return (), ()
        rows = []
        for line_number, values in enumerate(reader, start=2):
            row = {
                header: _text(values[index]) if index < len(values) else ""
                for index, header in enumerate(headers)
            }
            rows.append((line_number, row))
    return headers, tuple(rows)


def _validate_headers(headers: tuple[str, ...]) -> None:
    if not headers:
        raise LancamentosManuaisGerenciaisError("CSV sem cabecalho.")
    duplicates = sorted(
        {header for header in headers if headers.count(header) > 1 and header}
    )
    if duplicates:
        raise LancamentosManuaisGerenciaisError(
            "Colunas duplicadas no CSV: " + ", ".join(duplicates)
        )
    missing = [header for header in CSV_REQUIRED_HEADERS if header not in headers]
    if missing:
        raise LancamentosManuaisGerenciaisError(
            "Colunas obrigatorias ausentes no CSV: " + ", ".join(missing)
        )


def _parse_row(
    raw_row: Mapping[str, str],
    *,
    line_number: int,
    chart: ChartOfAccounts,
    dre_account_ids: frozenset[str],
) -> LancamentoManualGerencial:
    data = _parse_date(_required(raw_row, "data", line_number), line_number)
    descricao = _required(raw_row, "descricao_da_despesa", line_number)
    conta_debito_id = _required(raw_row, "conta_debito_id", line_number)
    conta_debito_nome = _required(raw_row, "conta_debito_nome", line_number)
    conta_credito_id = _required(raw_row, "conta_credito_id", line_number)
    conta_credito_nome = _required(raw_row, "conta_credito_nome", line_number)
    valor = _parse_positive_decimal(_required(raw_row, "valor", line_number), line_number)

    if conta_debito_id == conta_credito_id:
        raise LancamentosManuaisGerenciaisError(
            f"Linha {line_number}: conta_debito_id igual a conta_credito_id."
        )

    _validate_account(chart, conta_debito_id, conta_debito_nome, "debito", line_number)
    _validate_account(
        chart,
        conta_credito_id,
        conta_credito_nome,
        "credito",
        line_number,
    )

    conta_debito_totaliza = conta_debito_id in dre_account_ids
    conta_credito_totaliza = conta_credito_id in dre_account_ids
    conta_dre_count = int(conta_debito_totaliza) + int(conta_credito_totaliza)
    if conta_dre_count > 1:
        raise LancamentosManuaisGerenciaisError(
            f"Linha {line_number}: debito e credito possuem contas mapeadas na DRE; "
            "a v1 aceita no maximo uma conta de resultado por lancamento manual."
        )

    observacao = _text(raw_row.get("observacao"))
    chave = _dedupe_key(
        data=data,
        descricao=descricao,
        conta_debito_id=conta_debito_id,
        conta_credito_id=conta_credito_id,
        valor=valor,
    )
    return LancamentoManualGerencial(
        data=data,
        descricao_da_despesa=descricao,
        conta_debito_id=conta_debito_id,
        conta_debito_nome=conta_debito_nome,
        conta_credito_id=conta_credito_id,
        conta_credito_nome=conta_credito_nome,
        valor=valor,
        observacao=observacao,
        chave_deduplicacao=chave,
        linha_origem_csv=line_number,
        conta_dre_mapeada_count=conta_dre_count,
        conta_debito_totaliza_gerencial=conta_debito_totaliza,
        conta_credito_totaliza_gerencial=conta_credito_totaliza,
    )


def _technical_lancamento(
    manual: LancamentoManualGerencial,
    chart: ChartOfAccounts,
    csv_path: Path,
    *,
    side: str,
) -> Lancamento:
    is_debito = side == "DEBITO"
    conta_id = manual.conta_debito_id if is_debito else manual.conta_credito_id
    conta_nome = manual.conta_debito_nome if is_debito else manual.conta_credito_nome
    contraparte_id = manual.conta_credito_id if is_debito else manual.conta_debito_id
    contraparte_nome = manual.conta_credito_nome if is_debito else manual.conta_debito_nome
    conta = chart.get_by_id(conta_id)
    chave_origem = f"{manual.chave_deduplicacao}:{'D' if is_debito else 'C'}"
    totaliza = (
        manual.conta_debito_totaliza_gerencial
        if is_debito
        else manual.conta_credito_totaliza_gerencial
    )
    tipo_lancamento = f"MANUAL_GERENCIAL_{side}"
    if totaliza:
        tipo_lancamento += "_TOTALIZAVEL"
    observacoes = (
        f"Lancamento manual gerencial linha {manual.linha_origem_csv}; "
        f"partida={side}; contrapartida={contraparte_id} - {contraparte_nome}; "
        f"manual_totaliza_gerencial={'sim' if totaliza else 'nao'}"
    )
    if manual.observacao:
        observacoes += f"; observacao={manual.observacao}"
    return Lancamento(
        data_documento=manual.data,
        competencia=manual.data.strftime("%Y-%m"),
        cnpj_empresa=IDENTIFICADOR_MANUAL_GERENCIAL,
        nome_empresa=NOME_EMPRESA_MANUAL_GERENCIAL,
        fornecedor_ou_favorecido=manual.descricao_da_despesa,
        cnpj_cpf_favorecido="",
        documento="",
        tipo_documento=FONTE_MANUAL_GERENCIAL,
        conta_id=conta_id,
        conta_nome=conta_nome,
        grupo_conta=conta.grupo,
        descricao_itens=manual.descricao_da_despesa,
        valor=manual.valor,
        origem=FONTE_MANUAL_GERENCIAL,
        arquivo_origem=str(csv_path),
        chave_origem=chave_origem,
        hash_registro=hashlib.sha256(chave_origem.encode("utf-8")).hexdigest(),
        status_revisao="OK",
        alertas="",
        observacoes=observacoes,
        centro_custo="",
        tipo_lancamento=tipo_lancamento,
    )


def _write_csv_aplicado(
    output_path: Path,
    lancamentos: tuple[LancamentoManualGerencial, ...],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=CSV_APLICADO_HEADERS, delimiter=";")
        writer.writeheader()
        for item in lancamentos:
            writer.writerow(
                {
                    "data": item.data.isoformat(),
                    "descricao_da_despesa": item.descricao_da_despesa,
                    "conta_debito_id": item.conta_debito_id,
                    "conta_debito_nome": item.conta_debito_nome,
                    "conta_credito_id": item.conta_credito_id,
                    "conta_credito_nome": item.conta_credito_nome,
                    "valor": _decimal_text(item.valor),
                    "observacao": item.observacao,
                    "fonte": FONTE_MANUAL_GERENCIAL,
                    "status": "OK",
                    "tipo_lancamento": FONTE_MANUAL_GERENCIAL,
                    "chave_deduplicacao": item.chave_deduplicacao,
                    "linha_origem_csv": item.linha_origem_csv,
                }
            )


def _write_auditoria(
    output_path: Path,
    input_csv: Path,
    lancamentos: tuple[LancamentoManualGerencial, ...],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    total = sum((item.valor for item in lancamentos), Decimal("0.00")).quantize(
        Decimal("0.01")
    )
    debito_totals = _totals_by_account(
        (item.conta_debito_id, item.valor) for item in lancamentos
    )
    credito_totals = _totals_by_account(
        (item.conta_credito_id, item.valor) for item in lancamentos
    )
    lines = [
        "FANTOZZI - AUDITORIA LANCAMENTOS MANUAIS GERENCIAIS",
        "===================================================",
        "",
        f"arquivo_csv_origem: {input_csv}",
        f"linhas_lidas: {len(lancamentos)}",
        f"lancamentos_aceitos: {len(lancamentos)}",
        f"linhas_tecnicas_geradas: {len(lancamentos) * 2}",
        f"total_geral_manual: {_decimal_text(total)}",
        f"fonte: {FONTE_MANUAL_GERENCIAL}",
        "",
        "total_por_conta_debito:",
        *_account_lines(debito_totals),
        "",
        "total_por_conta_credito:",
        *_account_lines(credito_totals),
        "",
        "resultado: OK",
    ]
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")


def _totals_by_account(
    values: Iterable[tuple[str, Decimal]],
) -> dict[str, Decimal]:
    totals: dict[str, Decimal] = {}
    for account_id, value in values:
        totals[account_id] = (
            totals.get(account_id, Decimal("0.00")) + value
        ).quantize(Decimal("0.01"))
    return dict(sorted(totals.items()))


def _account_lines(totals: Mapping[str, Decimal]) -> list[str]:
    if not totals:
        return ["-"]
    return [f"- {account_id}: {_decimal_text(value)}" for account_id, value in totals.items()]


def _validate_account(
    chart: ChartOfAccounts,
    account_id: str,
    account_name: str,
    side: str,
    line_number: int,
) -> None:
    try:
        conta = chart.get_by_id(account_id)
    except KeyError as exc:
        raise LancamentosManuaisGerenciaisError(
            f"Linha {line_number}: conta_{side}_id inexistente no plano de contas: {account_id}"
        ) from exc
    if _normalize_text(conta.nome) != _normalize_text(account_name):
        raise LancamentosManuaisGerenciaisError(
            f"Linha {line_number}: conta_{side}_nome incompativel com "
            f"{account_id}. Esperado: {conta.nome}"
        )


def _parse_date(value: str, line_number: int) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise LancamentosManuaisGerenciaisError(
            f"Linha {line_number}: data invalida, use YYYY-MM-DD."
        ) from exc


def _parse_positive_decimal(value: str, line_number: int) -> Decimal:
    text = value.strip()
    if text.startswith("-"):
        raise LancamentosManuaisGerenciaisError(
            f"Linha {line_number}: valor deve ser positivo."
        )
    normalized = _normalize_decimal_text(text)
    try:
        decimal = Decimal(normalized).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError) as exc:
        raise LancamentosManuaisGerenciaisError(
            f"Linha {line_number}: valor impossivel de parsear."
        ) from exc
    if decimal <= 0:
        raise LancamentosManuaisGerenciaisError(
            f"Linha {line_number}: valor deve ser maior que zero."
        )
    return decimal


def _normalize_decimal_text(value: str) -> str:
    text = value.replace("R$", "").replace(" ", "").strip()
    if "," in text and "." in text:
        return text.replace(".", "").replace(",", ".")
    if "," in text:
        return text.replace(",", ".")
    return text


def _required(row: Mapping[str, str], header: str, line_number: int) -> str:
    value = _text(row.get(header))
    if not value:
        raise LancamentosManuaisGerenciaisError(
            f"Linha {line_number}: campo obrigatorio vazio: {header}"
        )
    return value


def _dedupe_key(
    *,
    data: date,
    descricao: str,
    conta_debito_id: str,
    conta_credito_id: str,
    valor: Decimal,
) -> str:
    payload = "|".join(
        [
            data.isoformat(),
            _normalize_text(descricao),
            conta_debito_id,
            conta_credito_id,
            _decimal_text(valor),
        ]
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _is_empty_row(row: Mapping[str, str]) -> bool:
    return not any(_text(value) for value in row.values())


def _decimal_text(value: Decimal) -> str:
    return f"{value.quantize(Decimal('0.01')):.2f}"


def _normalize_text(value: str) -> str:
    return " ".join(_text(value).casefold().split())


def _text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()
