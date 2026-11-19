"""
teste_import_excecoes.py
========================
Importação da planilha de revisão de exceções preenchida (dados 100% fictícios):
ida e volta com o export, planilha filtrada, decisão inválida, linha não encontrada, linha em branco,
aba/coluna ausente (ArquivoInvalidoError) e sobrescrita de decisão existente.
"""

from __future__ import annotations

import io

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

from importacao_excecoes import importar_decisoes_excecoes_xlsx
from pipeline import exportar_fila_excecoes_xlsx
from validacao import ArquivoInvalidoError

OPCOES = ["", "Falta injustificada", "Falta justificada", "Erro de relógio — abonar", "Atestado médico"]
COLUNAS = ["nome_colaborador", "razao_social", "data", "situacao", "motivo_validacao", "classificacao", "observacao"]


def _fila() -> pd.DataFrame:
    # índice propositalmente fora de ordem/posição: a aplicação precisa ser por índice
    linhas = [
        ["Ana Souza", "Empresa A", "2026-09-02", "Provável falta real", "Sem batidas", "", ""],
        ["Ana Souza", "Empresa A", "2026-09-12", "Intervalo curto", "Almoço curto", "", ""],
        ["Zélia Ñandú", "Empresa B", "2026-09-10", "Batida incompleta", "Saída sem retorno", "", ""],
        ["Bruno Lima", "Empresa B", "2026-09-15", "Provável falta real", "Sem batidas", "", ""],
    ]
    df = pd.DataFrame(linhas, columns=COLUNAS)
    df["data"] = pd.to_datetime(df["data"])
    df.index = [10, 7, 3, 21]
    return df


def _exportar(tmp_path, fila: pd.DataFrame, nome: str = "planilha.xlsx") -> str:
    caminho = tmp_path / nome
    caminho.write_bytes(exportar_fila_excecoes_xlsx(fila, "09/2026", OPCOES, "01/09/2026 a 30/09/2026"))
    return str(caminho)


def _preencher(caminho: str, preenchimento: dict[tuple[str, str], tuple[str, str]]) -> None:
    """Escreve Decisão/Observação na linha com (Colaborador, data dd/mm/aaaa)."""
    wb = load_workbook(caminho)
    ws = wb["Exceções"]
    for linha in range(2, ws.max_row + 1):
        chave = (ws.cell(linha, 1).value, ws.cell(linha, 3).value.strftime("%d/%m/%Y"))
        if chave in preenchimento:
            decisao, obs = preenchimento[chave]
            ws.cell(linha, 6).value = decisao
            ws.cell(linha, 8).value = obs
    wb.save(caminho)


def test_ida_e_volta_aplica_por_indice(tmp_path):
    fila = _fila()
    caminho = _exportar(tmp_path, fila)
    _preencher(caminho, {
        ("Ana Souza", "02/09/2026"): ("Falta injustificada", ""),
        ("Ana Souza", "12/09/2026"): ("  erro de RELÓGIO — abonar ", "relógio travado"),
        ("Zélia Ñandú", "10/09/2026"): ("Atestado médico", "atestado de 1 dia"),
    })
    res = importar_decisoes_excecoes_xlsx(caminho, fila, OPCOES)
    assert res.aplicadas == 3 and res.total_problemas == 0
    assert res.decisoes.loc[10, "classificacao"] == "Falta injustificada"
    assert res.decisoes.loc[7, "classificacao"] == "Erro de relógio — abonar"
    assert res.decisoes.loc[7, "observacao"] == "relógio travado"
    assert res.decisoes.loc[3, "classificacao"] == "Atestado médico"
    assert res.decisoes.loc[3, "observacao"] == "atestado de 1 dia"
    assert 21 not in res.decisoes.index  # em branco na planilha → intocada
    assert res.em_branco == 1


def test_planilha_filtrada_so_toca_as_linhas_dela(tmp_path):
    fila = _fila()
    caminho = _exportar(tmp_path, fila[fila["razao_social"] == "Empresa B"])
    _preencher(caminho, {("Zélia Ñandú", "10/09/2026"): ("Falta justificada", "")})
    res = importar_decisoes_excecoes_xlsx(caminho, fila, OPCOES)
    assert list(res.decisoes.index) == [3] and res.aplicadas == 1 and res.total_problemas == 0


def test_decisao_invalida_e_relatada_e_nao_aplicada(tmp_path):
    fila = _fila()
    caminho = _exportar(tmp_path, fila)
    _preencher(caminho, {("Ana Souza", "02/09/2026"): ("Talvez justificada", "")})
    res = importar_decisoes_excecoes_xlsx(caminho, fila, OPCOES)
    assert res.aplicadas == 0 and len(res.invalidas) == 1
    assert res.invalidas[0]["colaborador"] == "Ana Souza" and res.invalidas[0]["data"] == "02/09/2026"
    assert "Talvez justificada" in res.invalidas[0]["motivo"]


def test_linha_nao_encontrada_na_fila(tmp_path):
    fila = _fila()
    caminho = _exportar(tmp_path, fila)
    # a planilha tem uma linha que não existe na fila atual (ex.: fila foi reprocessada)
    res = importar_decisoes_excecoes_xlsx(caminho, fila.drop(index=10), OPCOES)
    assert res.aplicadas == 0  # nada preenchido ainda
    _preencher(caminho, {("Ana Souza", "02/09/2026"): ("Falta justificada", "")})
    res = importar_decisoes_excecoes_xlsx(caminho, fila.drop(index=10), OPCOES)
    assert res.aplicadas == 0 and len(res.nao_encontradas) == 1
    assert res.nao_encontradas[0]["colaborador"] == "Ana Souza"


def test_nomes_com_caixa_acento_e_espacos_diferentes_casam(tmp_path):
    fila = _fila()
    caminho = _exportar(tmp_path, fila)
    wb = load_workbook(caminho)
    ws = wb["Exceções"]
    for linha in range(2, ws.max_row + 1):
        if ws.cell(linha, 1).value == "Zélia Ñandú":
            ws.cell(linha, 1).value = "  ZELIA   NANDU "
            ws.cell(linha, 3).value = "10/09/2026"  # data como texto
            ws.cell(linha, 6).value = "Falta justificada"
    wb.save(caminho)
    res = importar_decisoes_excecoes_xlsx(caminho, fila, OPCOES)
    assert list(res.decisoes.index) == [3] and res.total_problemas == 0


def test_linhas_em_branco_nao_alteram_nada(tmp_path):
    fila = _fila()
    sem_decisoes = fila.assign(classificacao="", observacao="")
    res = importar_decisoes_excecoes_xlsx(_exportar(tmp_path, sem_decisoes), fila, OPCOES)
    assert res.aplicadas == 0 and res.decisoes.empty and res.em_branco == 4
    assert res.total_problemas == 0
    assert (fila["classificacao"] == "").all()  # a fila nunca é alterada pela leitura


def test_duplicadas_identicas_aplicam_uma_vez_e_divergentes_nao(tmp_path):
    fila = _fila()
    caminho = _exportar(tmp_path, fila)
    wb = load_workbook(caminho)
    ws = wb["Exceções"]
    nova = ws.max_row + 1
    for col in range(1, 9):
        ws.cell(nova, col).value = ws.cell(2, col).value
    ws.cell(2, 6).value = "Falta justificada"
    ws.cell(nova, 6).value = "Falta justificada"
    wb.save(caminho)
    res = importar_decisoes_excecoes_xlsx(caminho, fila, OPCOES)
    assert len(res.duplicadas) == 1 and "idênticas" in res.duplicadas[0]["motivo"]

    wb = load_workbook(caminho)
    wb["Exceções"].cell(nova, 6).value = "Falta injustificada"
    wb.save(caminho)
    res = importar_decisoes_excecoes_xlsx(caminho, fila, OPCOES)
    assert len(res.duplicadas) == 1 and "nenhuma foi aplicada" in res.duplicadas[0]["motivo"]


def test_sobrescreve_decisao_existente_e_conta(tmp_path):
    fila = _fila()
    caminho = _exportar(tmp_path, fila)
    fila.loc[21, ["classificacao", "observacao"]] = ["Falta justificada", "já marcada"]  # decisão na tela
    _preencher(caminho, {
        ("Bruno Lima", "15/09/2026"): ("Falta injustificada", "mudou"),  # diverge da tela
    })
    res = importar_decisoes_excecoes_xlsx(caminho, fila, OPCOES)
    assert res.aplicadas == 1 and res.sobrescritas == 1
    assert res.decisoes.loc[21, "classificacao"] == "Falta injustificada"
    assert res.decisoes.loc[21, "observacao"] == "mudou"
    # mesma decisão que já está na tela não conta como sobrescrita
    _preencher(caminho, {("Bruno Lima", "15/09/2026"): ("Falta justificada", "")})
    assert importar_decisoes_excecoes_xlsx(caminho, fila, OPCOES).sobrescritas == 0


def test_aba_ausente_levanta_erro_amigavel(tmp_path):
    wb = Workbook()
    wb.active.title = "Outra"
    caminho = tmp_path / "sem_aba.xlsx"
    wb.save(caminho)
    with pytest.raises(ArquivoInvalidoError, match="aba"):
        importar_decisoes_excecoes_xlsx(str(caminho), _fila(), OPCOES)


def test_coluna_ausente_levanta_erro_amigavel(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Exceções"
    ws.append(["Colaborador", "Empresa", "Data"])
    ws.append(["Ana Souza", "Empresa A", "02/09/2026"])
    caminho = tmp_path / "sem_coluna.xlsx"
    wb.save(caminho)
    with pytest.raises(ArquivoInvalidoError, match="Decisão"):
        importar_decisoes_excecoes_xlsx(str(caminho), _fila(), OPCOES)


def test_arquivo_corrompido_nao_vaza_traceback(tmp_path):
    caminho = tmp_path / "lixo.xlsx"
    caminho.write_bytes(b"isto nao e um xlsx")
    with pytest.raises(ArquivoInvalidoError):
        importar_decisoes_excecoes_xlsx(str(caminho), _fila(), OPCOES)
