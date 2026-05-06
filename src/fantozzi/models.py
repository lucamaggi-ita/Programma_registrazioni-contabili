from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from typing import List, Optional


@dataclass
class Conta:
    id: str
    nome: str
    grupo: str
    tipo: str
    exige_revisao: bool = False


@dataclass
class Duplicata:
    numero: str
    vencimento: Optional[date]
    valor: Decimal


@dataclass
class ProdutoNFE:
    n_item: int
    codigo: str
    descricao: str
    ncm: str
    cfop: str
    unidade: str
    quantidade: Decimal
    valor_unitario: Decimal
    valor_total: Decimal


@dataclass
class NFE:
    chave: str
    numero: str
    serie: str
    data_emissao: date
    cnpj_emitente: str
    nome_emitente: str
    cnpj_destinatario: str
    nome_destinatario: str
    valor_total: Decimal
    produtos: List[ProdutoNFE] = field(default_factory=list)
    duplicatas: List[Duplicata] = field(default_factory=list)
    tem_duplicatas: bool = False
    arquivo_origem: str = ""


@dataclass
class ClassificationResult:
    conta: Conta
    confianca: str  # "OK" | "AMBIGUO" | "NAO_CLASSIFICADO"
    alertas: List[str] = field(default_factory=list)
    matches: List[str] = field(default_factory=list)


@dataclass
class Lancamento:
    data_documento: date
    competencia: str  # "YYYY-MM"
    cnpj_empresa: str
    nome_empresa: str
    fornecedor_ou_favorecido: str
    cnpj_cpf_favorecido: str
    documento: str
    tipo_documento: str
    conta_id: str
    conta_nome: str
    grupo_conta: str
    descricao_itens: str
    valor: Decimal
    duplicatas: List[Duplicata] = field(default_factory=list)
    origem: str = "XML_NFE"  # "XML_NFE" | "CSV" | "MANUAL"
    arquivo_origem: str = ""
    chave_origem: str = ""
    hash_registro: str = ""
    status_revisao: str = "OK"
    alertas: str = ""
    observacoes: str = ""
    centro_custo: str = ""
    tipo_lancamento: str = "custo"
    dre_linha_id: str = ""
    dre_linha_nome: str = ""
    data_pagamento: Optional[date] = None
