"""Consignado com identificador fora do cadastro (ex.: colaborador desligado): descarta e avisa, sem bloquear.

A tolerância vale só para o extrato de consignados; banco de horas continua estrito.
"""
from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import pytest
from streamlit.testing.v1 import AppTest

import db_sessao
import pipeline
from pipeline import aviso_consignados_ignorados, carregar_consignados
from validacao import ArquivoInvalidoError, normalizar_identificador_colaborador

APP = str(Path(__file__).parent.parent / "app_web.py")


def _cadastro() -> pd.DataFrame:
    return pd.DataFrame({"pis": ["00000000096", "00000000097"], "matricula": ["96", "97"]})


def _extrato(caminho, linhas, coluna_valor="valorParcela") -> str:
    pd.DataFrame(linhas, columns=["matricula", coluna_valor]).to_excel(caminho, index=False)
    return str(caminho)


def test_id_fora_do_cadastro_e_descartado_e_o_resto_soma(tmp_path):
    arq = _extrato(tmp_path / "ext.xlsx", [("96", 100.0), ("99", 50.0), ("99", 20.0), ("97", 10.0), ("96", 5.0)])
    res = carregar_consignados(arq, cadastro=_cadastro())
    assert dict(zip(res["pis"], res["valor_consignado"])) == {"00000000096": 105.0, "00000000097": 10.0}
    assert res.attrs["consignados_ignorados"] == ["99"]


def test_zeros_a_esquerda_sao_reconhecidos(tmp_path):
    arq = _extrato(tmp_path / "ext.xlsx", [("096", 100.0), ("099", 50.0)])
    res = carregar_consignados(arq, cadastro=_cadastro())
    assert dict(zip(res["pis"], res["valor_consignado"])) == {"00000000096": 100.0}
    assert res.attrs["consignados_ignorados"] == ["99"]   # o leitor já normaliza "099" para "99"


def test_extrato_sem_orfao_devolve_lista_vazia(tmp_path):
    arq = _extrato(tmp_path / "ext.xlsx", [("96", 100.0), ("97", 10.0)])
    res = carregar_consignados(arq, cadastro=_cadastro())
    assert res.attrs["consignados_ignorados"] == []
    assert aviso_consignados_ignorados(res) is None


def test_ignorados_de_varios_arquivos_sem_repeticao_e_ordenados(tmp_path):
    a = _extrato(tmp_path / "a.xlsx", [("99", 1.0), ("96", 1.0)])
    b = _extrato(tmp_path / "b.xlsx", [("88", 1.0), ("99", 2.0)])
    res = carregar_consignados([a, b], cadastro=_cadastro())
    assert res.attrs["consignados_ignorados"] == ["88", "99"]


def test_sem_cadastro_nada_e_descartado(tmp_path):
    arq = _extrato(tmp_path / "ext.xlsx", [("96", 100.0), ("99", 50.0)])
    res = carregar_consignados(arq)
    assert set(res["pis"]) == {"96", "99"}
    assert res.attrs["consignados_ignorados"] == []


def test_extrato_sem_coluna_de_valor_continua_erro(tmp_path):
    arq = _extrato(tmp_path / "ext.xlsx", [("96", 100.0)], coluna_valor="outra_coisa")
    with pytest.raises(ArquivoInvalidoError, match="coluna"):
        carregar_consignados(arq, cadastro=_cadastro())


def test_texto_do_aviso():
    df = pd.DataFrame({"pis": [], "valor_consignado": []})
    df.attrs["consignados_ignorados"] = ["99"]
    assert aviso_consignados_ignorados(df) == (
        "Consignado ignorado para a(s) matrícula(s) 99: não consta(m) no cadastro "
        "(colaborador desligado ou matrícula digitada errada). Confira antes de fechar."
    )


def test_banco_de_horas_continua_estrito():
    df = pd.DataFrame({"pis": ["96", "99"]})
    with pytest.raises(ArquivoInvalidoError, match="99"):
        normalizar_identificador_colaborador(df, "pis", _cadastro(), "Banco de Horas do Secullum", "x.xlsx")


def test_aviso_aparece_na_tela(tmp_path, monkeypatch):
    """Validar dados com consignado de matrícula fora do cadastro mostra o aviso e não bloqueia."""
    from cadastro_editor import cadastro_para_xlsx_bytes, normalizar_para_edicao

    monkeypatch.setattr(db_sessao, "DB_PATH", tmp_path / "app.db")
    monkeypatch.setattr(pipeline, "PREENCHER_CONSIGNADO", True)   # a tela só lê o extrato com a flag ligada
    linha = dict(
        pis="00000000096", matricula="96", nome="COLAB 96", codigo_empresa="001", razao_social="EMPRESA A LTDA",
        entrada="07:30", saida_almoco="12:00", retorno_almoco="13:00", saida="17:30", dias_trabalho="0,1,2,3,4",
        saida_sexta="16:30", admissao="01/03/2020", cnpj="17.155.411/0001-59", centro_custo="07", valor_va=1000.0,
    )
    cad = cadastro_para_xlsx_bytes(normalizar_para_edicao(pd.DataFrame([linha])))
    buf = io.BytesIO()
    pd.DataFrame([("96", 100.0), ("99", 50.0)], columns=["matricula", "valorParcela"]).to_excel(buf, index=False)
    afd = "".join(
        f"{n:09d}3{'00000000096':>12}20261005{hm}\n" for n, hm in enumerate(["0730", "1200", "1300", "1730"], 1)
    ).encode("latin-1")
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    at = AppTest.from_file(APP, default_timeout=60).run()
    at.file_uploader[0].set_value([
        ("cadastro.xlsx", cad, mime), ("ponto.txt", afd, "text/plain"),
        ("17155411000159-202610v1.4.xlsx", buf.getvalue(), mime),
    ]).run()
    next(b for b in at.button if b.label == "Confirmar arquivos e continuar").click().run()
    assert not at.exception
    at.text_input(key="data_inicio_input").set_value("26/09/2026")
    at.text_input(key="data_fim_input").set_value("25/10/2026").run()
    next(b for b in at.button if b.label == "Validar dados").click().run()
    assert not at.exception
    assert not at.error, [e.value for e in at.error]
    avisos = " ".join(w.value for w in at.warning) + " " + " ".join(m.value for m in at.markdown)
    assert "Consignado ignorado para a(s) matrícula(s) 99" in avisos
