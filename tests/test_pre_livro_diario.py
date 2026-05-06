from datetime import date
from decimal import Decimal
from pathlib import Path
from shutil import copyfile

from openpyxl import load_workbook
import pytest

from fantozzi.excel_writer import write_lancamentos_to_excel
from fantozzi.models import Duplicata, Lancamento
from fantozzi.pre_livro_diario import (
    INSTRUCOES_SHEET,
    LANCAMENTOS_OK_SHEET,
    LANCAMENTOS_REVISAR_SHEET,
    OPERACOES_ESPECIAIS_SHEET,
    RESUMO_CONTA_SHEET,
    RESUMO_MENSAL_SHEET,
    PreLivroDiarioError,
    gerar_pre_livro_diario_teste,
)


ROOT = Path(__file__).resolve().parents[1]
SAMPLE_XML = ROOT / "samples" / "nfe_exemplo" / "nfe_oesa_12566_20260505.xml"


def test_gera_pre_livro_diario_com_abas_e_resumos(tmp_path):
    excel_path = tmp_path / "lancamentos.xlsx"
    output_path = tmp_path / "reports" / "pre_livro.xlsx"
    write_lancamentos_to_excel(
        excel_path,
        [
            _lancamento("hash-ok", conta_id="2.01.004", valor=Decimal("10.00")),
            _lancamento(
                "hash-revisar",
                conta_id="13.01.001",
                conta_nome="Revisar - Produto Nao Classificado",
                grupo_conta="Contas de Revisao",
                valor=Decimal("7.50"),
                status_revisao="REVISAR",
                alertas="CLASSIFICACAO_NAO_ENCONTRADA",
            ),
        ],
    )
    original_excel = excel_path.read_bytes()

    result = gerar_pre_livro_diario_teste([excel_path], output_path)

    workbook = load_workbook(output_path, data_only=True)
    assert result.lancamentos_ok == 1
    assert result.lancamentos_revisar == 1
    assert result.operacoes_especiais == 0
    assert workbook.sheetnames == [
        LANCAMENTOS_OK_SHEET,
        LANCAMENTOS_REVISAR_SHEET,
        OPERACOES_ESPECIAIS_SHEET,
        RESUMO_CONTA_SHEET,
        RESUMO_MENSAL_SHEET,
        INSTRUCOES_SHEET,
    ]
    assert workbook[LANCAMENTOS_OK_SHEET].max_row == 2
    assert workbook[LANCAMENTOS_REVISAR_SHEET].max_row == 2
    assert workbook[RESUMO_CONTA_SHEET].max_row == 3
    assert workbook[RESUMO_MENSAL_SHEET].max_row == 2
    assert "nao um livro diario definitivo" in workbook[INSTRUCOES_SHEET]["A3"].value
    assert excel_path.read_bytes() == original_excel


def test_inclui_operacoes_especiais_quando_pasta_xml_e_informada(tmp_path):
    excel_path = tmp_path / "lancamentos.xlsx"
    output_path = tmp_path / "reports" / "pre_livro.xlsx"
    xml_dir = _xml_dir_with_special_operation(tmp_path)
    write_lancamentos_to_excel(excel_path, [_lancamento("hash-ok")])

    result = gerar_pre_livro_diario_teste([excel_path], output_path, [xml_dir])

    sheet = load_workbook(output_path, data_only=True)[OPERACOES_ESPECIAIS_SHEET]
    headers = [cell.value for cell in sheet[1]]
    row = dict(zip(headers, next(sheet.iter_rows(min_row=2, values_only=True))))
    assert result.operacoes_especiais == 1
    assert row["tipo_operacao_especial"] == "COMODATO"
    assert row["destino_gerencial_sugerido"] == "FORA_DRE"
    assert row["status_revisao_sugerido"] == "ANALISAR_COM_CONTADOR"


def test_retorna_erro_quando_excel_nao_existe(tmp_path):
    output_path = tmp_path / "reports" / "pre_livro.xlsx"

    with pytest.raises(PreLivroDiarioError) as exc_info:
        gerar_pre_livro_diario_teste([tmp_path / "inexistente.xlsx"], output_path)

    assert "Excel de lancamentos nao encontrado" in str(exc_info.value)
    assert not output_path.exists()


def _xml_dir_with_special_operation(tmp_path: Path) -> Path:
    xml_dir = tmp_path / "nfe_xml_especial"
    xml_dir.mkdir()
    copyfile(SAMPLE_XML, xml_dir / "comodato.xml")
    xml_path = xml_dir / "comodato.xml"
    text = xml_path.read_text(encoding="utf-8")
    text = text.replace("<natOp>VENDA DE MERCADORIAS</natOp>", "<natOp>Comodato</natOp>")
    text = text.replace("<CFOP>5102</CFOP>", "<CFOP>5908</CFOP>", 1)
    xml_path.write_text(text, encoding="utf-8")
    return xml_dir


def _lancamento(
    hash_registro: str,
    *,
    conta_id: str = "2.01.004",
    conta_nome: str = "Refrigerantes",
    grupo_conta: str = "Custos de Mercadorias e Insumos",
    valor: Decimal = Decimal("10.00"),
    status_revisao: str = "OK",
    alertas: str = "",
) -> Lancamento:
    return Lancamento(
        data_documento=date(2026, 5, 5),
        competencia="2026-05",
        cnpj_empresa="55746722000181",
        nome_empresa="TRATTORIA DO SOSSEGO RESTAURANTE LTDA",
        fornecedor_ou_favorecido="OESA COMERCIO E REPRESENTACOES S/A",
        cnpj_cpf_favorecido="12345678000195",
        documento="12566",
        tipo_documento="XML_NFE",
        conta_id=conta_id,
        conta_nome=conta_nome,
        grupo_conta=grupo_conta,
        descricao_itens="COCA-COLA KS 290ML",
        valor=valor,
        duplicatas=[
            Duplicata(
                numero="001",
                vencimento=date(2026, 5, 20),
                valor=valor,
            )
        ],
        origem="XML_NFE",
        arquivo_origem="nfe.xml",
        chave_origem="NFE123",
        hash_registro=hash_registro,
        status_revisao=status_revisao,
        alertas=alertas,
        tipo_lancamento="custo",
    )
