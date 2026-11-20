"""Cadastro editado: versões no SQLite (migração aditiva), reenvio da planilha e painel de diferenças (dados fictícios)."""
from __future__ import annotations

import pandas as pd
import pytest

import db_sessao
from cadastro_editor import (
    aplicar_aceites, cadastro_de_json, cadastro_para_json, comparar_cadastros, hash_do_cadastro,
    normalizar_para_edicao, reconciliar_com_salvo,
)


@pytest.fixture(autouse=True)
def banco_temporario(tmp_path, monkeypatch):
    monkeypatch.setattr(db_sessao, "DB_PATH", tmp_path / "teste.db")
    db_sessao.inicializar_banco()


def _cad(*matriculas: str, alteracoes: dict | None = None) -> pd.DataFrame:
    linhas = [dict(
        pis=m.zfill(11), matricula=m, nome=f"COLAB {m}", codigo_empresa="001", razao_social="EMPRESA A LTDA",
        entrada="07:30", saida_almoco="12:00", retorno_almoco="13:00", saida="17:30", dias_trabalho="0,1,2,3,4",
        valor_va=1000.0, cooper=200.0, admissao="01/03/2020", centro_custo="07",
    ) for m in matriculas]
    df = normalizar_para_edicao(pd.DataFrame(linhas))
    for (mat, campo), valor in (alteracoes or {}).items():
        df.loc[df["matricula"] == mat, campo] = valor
    return df


def test_migracao_e_idempotente_e_guarda_versoes():
    db_sessao.inicializar_banco()   # segunda chamada nao pode falhar nem apagar nada
    df = _cad("1", "2")
    id1 = db_sessao.salvar_cadastro(cadastro_para_json(df), "planilha.xlsx", hash_do_cadastro(df))
    db_sessao.inicializar_banco()
    atual = db_sessao.carregar_cadastro_atual()
    assert atual["id"] == id1 and atual["nome_arquivo"] == "planilha.xlsx" and atual["total_linhas"] == 2
    volta = cadastro_de_json(atual["linhas_json"])
    assert list(volta["matricula"]) == ["1", "2"] and volta.at[0, "valor_va"] == 1000.0
    assert str(volta.at[0, "admissao"]) == "2020-03-01" and volta.at[0, "centro_custo"] == "07"


def test_salvar_o_mesmo_conteudo_nao_cria_nova_versao_e_mudanca_cria():
    df = _cad("1")
    id1 = db_sessao.salvar_cadastro(cadastro_para_json(df), "a.xlsx", hash_do_cadastro(df))
    assert db_sessao.salvar_cadastro(cadastro_para_json(df), "b.xlsx", hash_do_cadastro(df)) == id1
    df2 = _cad("1", alteracoes={("1", "entrada"): "08:00"})
    id2 = db_sessao.salvar_cadastro(cadastro_para_json(df2), "edicao no sistema", hash_do_cadastro(df2))
    assert id2 != id1
    assert db_sessao.carregar_cadastro_atual()["id"] == id2
    assert len(db_sessao.listar_versoes_cadastro()) == 2


def test_hash_nao_depende_da_ordem_das_linhas():
    assert hash_do_cadastro(_cad("1", "2")) == hash_do_cadastro(_cad("2", "1"))


def test_reenvio_mostra_novos_removidos_e_alterados_e_o_sistema_prevalece():
    salvo = _cad("1", "2", alteracoes={("1", "observacao"): "ajustado no sistema"})
    enviado = _cad("2", "3", alteracoes={("2", "entrada"): "08:00"})
    trabalho, difs = reconciliar_com_salvo(enviado, salvo)
    assert list(trabalho["matricula"]) == ["1", "2"]    # padrao: o cadastro do sistema e' a verdade
    por_tipo = {(d.tipo, d.matricula, d.campo) for d in difs}
    assert ("novo", "3", "") in por_tipo and ("removido", "1", "") in por_tipo
    assert ("alterado", "2", "entrada") in por_tipo
    alterada = next(d for d in difs if d.campo == "entrada")
    assert (alterada.valor_salvo, alterada.valor_novo) == ("07:30", "08:00")


def test_sem_cadastro_salvo_a_planilha_vira_o_cadastro_sem_diferencas():
    enviado = _cad("1")
    trabalho, difs = reconciliar_com_salvo(enviado, None)
    assert difs == [] and list(trabalho["matricula"]) == ["1"]


def test_aceitar_aplica_so_o_que_foi_marcado():
    salvo = _cad("1", "2", alteracoes={("1", "observacao"): "ajustado no sistema"})
    enviado = _cad("2", "3", alteracoes={("2", "entrada"): "08:00"})
    difs = comparar_cadastros(salvo, enviado)
    resultado = aplicar_aceites(salvo, enviado, {("2", "entrada"), ("3", "")})
    r = resultado.set_index("matricula")
    assert r.at["2", "entrada"] == "08:00"            # aceito
    assert "3" in r.index                              # novo aceito
    assert "1" in r.index and r.at["1", "observacao"] == "ajustado no sistema"   # remocao nao aceita
    # aceitar a remocao tira a linha
    assert "1" not in aplicar_aceites(salvo, enviado, {("1", "")})["matricula"].tolist()
    assert len(difs) == 3
