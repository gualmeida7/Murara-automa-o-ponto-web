"""Tela de cadastro editável (app_web.py) com streamlit.testing.v1.AppTest, banco SQLite temporário e dados fictícios."""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest
import streamlit as st
from streamlit.testing.v1 import AppTest

import db_sessao
from cadastro_editor import cadastro_para_xlsx_bytes, normalizar_para_edicao

APP = str(Path(__file__).parent.parent / "app_web.py")
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _linha(matricula: str, **extra) -> dict:
    base = dict(
        pis=matricula.zfill(11), matricula=matricula, nome=f"COLAB {matricula}", codigo_empresa="001",
        razao_social="EMPRESA A LTDA", entrada="07:30", saida_almoco="12:00", retorno_almoco="13:00", saida="17:30",
        dias_trabalho="0,1,2,3,4", saida_sexta="16:30", admissao="01/03/2020", cnpj="17.155.411/0001-59",
        centro_custo="07", valor_va=1000.0,
    )
    base.update(extra)
    return base


def _xlsx(*linhas: dict) -> bytes:
    return cadastro_para_xlsx_bytes(normalizar_para_edicao(pd.DataFrame(list(linhas))))


AFD = "".join(
    f"{n:09d}3{'00000000001':>12}20261005{hm}\n" for n, hm in enumerate(["0730", "1200", "1300", "1730"], 1)
).encode("latin-1")


@pytest.fixture(autouse=True)
def banco_temporario(tmp_path, monkeypatch):
    monkeypatch.setattr(db_sessao, "DB_PATH", tmp_path / "app.db")


def _enviar_e_confirmar(xlsx: bytes) -> AppTest:
    at = AppTest.from_file(APP, default_timeout=60).run()
    at.file_uploader[0].set_value([("cadastro.xlsx", xlsx, MIME_XLSX), ("ponto.txt", AFD, "text/plain")]).run()
    _botao(at, "Confirmar arquivos e continuar").click().run()
    assert not at.exception
    return at


def _botao(at: AppTest, rotulo: str):
    return next(b for b in at.button if b.label == rotulo)


def _tem_download(at: AppTest) -> bool:
    return any("Baixar cadastro corrigido" in e.proto.label for e in at.get("download_button"))


def _textos_de_erro(at: AppTest) -> str:
    return " ".join(e.value for e in at.error)


def _editar(at: AppTest, linhas_editadas: dict, monkeypatch) -> None:
    """
    O AppTest não repassa o estado de um data_editor entre reruns (st.rerun da coluna "Situação" o perderia),
    então, só aqui, o rerun é desligado para o estado editado ficar até o fim da execução.
    """
    monkeypatch.setattr(st, "rerun", lambda *a, **k: None)
    _injetar_edicao(at, linhas_editadas)
    at.run()
    assert not at.exception


def _injetar_edicao(at: AppTest, linhas_editadas: dict) -> None:
    chave = f"editor_cadastro_{at.session_state['cadastro_versao_editor']}"
    at.session_state[chave] = {"edited_rows": linhas_editadas, "added_rows": [], "deleted_rows": []}


def test_envio_mostra_a_tabela_e_bloqueia_ate_corrigir_a_linha_incompleta(monkeypatch):
    incompleta = _linha("14", razao_social=None, entrada=None, saida_almoco=None, retorno_almoco=None, saida=None)
    at = _enviar_e_confirmar(_xlsx(_linha("1"), incompleta))

    assert any("situacao" in d.value.columns for d in at.dataframe)      # a tabela editável (com Situação) apareceu
    assert "5 erro(s) impedem o fechamento" in _textos_de_erro(at)
    painel = next(d.value for d in at.dataframe if "Problema" in d.value.columns)   # painel de erros cita a matrícula
    assert set(painel["Matrícula"]) == {"14"} and {"razao_social", "entrada", "saida"} <= set(painel["Campo"])
    assert _botao(at, "Validar dados").disabled                           # etapas seguintes travadas
    assert _tem_download(at)

    _editar(at, {1: {"razao_social": "EMPRESA A LTDA", "entrada": "07:30", "saida_almoco": "12:00",
                     "retorno_almoco": "13:00", "saida": "17:30"}}, monkeypatch)
    assert not at.error and any("Nenhum erro bloqueante" in s.value for s in at.success)
    assert not _botao(at, "Validar dados").disabled
    assert _tem_download(at)


def test_validar_usa_o_cadastro_editado_com_tipos_do_pipeline(monkeypatch):
    at = _enviar_e_confirmar(_xlsx(_linha("1")))
    edicao = {0: {"va_fixo": True, "valor_va": 1234.5}}
    _editar(at, edicao, monkeypatch)
    at.session_state["competencia_input"] = "10/2026"
    _injetar_edicao(at, edicao)    # o AppTest não carrega o estado do editor de uma execução para a outra
    _botao(at, "Validar dados").click().run()
    assert not at.exception, at.exception
    cad = at.session_state["cadastro_df"]
    assert bool(cad.at[0, "va_fixo"]) and cad.at[0, "va_fixo"].__class__ is not str
    assert cad.at[0, "valor_va"] == 1234.5 and cad["pis"].tolist() == ["00000000001"]


def test_salvar_e_reenviar_planilha_diferente_mostra_diferencas_sem_sobrescrever():
    at = _enviar_e_confirmar(_xlsx(_linha("1"), _linha("2")))
    _botao(at, "Salvar cadastro").click().run()
    assert not at.exception and any("Cadastro salvo" in s.value for s in at.success)
    assert db_sessao.carregar_cadastro_atual()["total_linhas"] == 2

    # novo envio: matrícula 2 mudou a entrada, matrícula 3 é nova, matrícula 1 sumiu
    at2 = _enviar_e_confirmar(_xlsx(_linha("2", entrada="08:00"), _linha("3")))
    texto = " ".join(m.value for m in at2.markdown)
    assert "difere do cadastro salvo" in texto and "3 diferença(s)" in texto
    # o cadastro de trabalho continua sendo o do sistema (1 e 2), com a entrada original
    trabalho = at2.session_state["cadastro_trabalho"].set_index("matricula")
    assert list(trabalho.index) == ["1", "2"] and trabalho.at["2", "entrada"] == "07:30"


def test_tela_mostra_conferencia_e_aviso_de_planilha_sem_cores():
    at = _enviar_e_confirmar(_xlsx(_linha("1", beneficio="Alelo Cartão", cooper=200.0, admissao="2026-11-16"),
                                   _linha("2", codigo_empresa="002", beneficio="Ticket")))
    texto = " ".join(m.value for m in at.markdown)
    assert "Conferência da leitura do cadastro" in texto and "Empresa 001: 1 com Alelo Cartão" in texto
    assert "Empresa 002: 1 com Ticket" in texto and "Com Cooper: 1" in texto
    assert "matrícula 1: 16/11/2026" in texto
    assert any("A planilha chegou sem cores; marque abaixo" in i.value for i in at.info)
