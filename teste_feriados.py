"""
teste_feriados.py
=================
Testes do calendario de feriados (feriados.py / gerar_feriados.py / feriados.csv)
e de como o feriado entra na apuracao do ponto.

Rodar com pytest (`pytest teste_feriados.py`) ou direto (`python teste_feriados.py`).
Se a pasta "Dados e Docs Reais" existir (ela e' ignorada pelo git), roda tambem o
cenario de 09/2026 com os arquivos reais.
"""

from __future__ import annotations

import tempfile
from datetime import date, datetime, time
from pathlib import Path

import pandas as pd

import gerar_feriados as gerador
from business_rules import Jornada, apurar_dia, apurar_periodo, consolidar_mes
from feriados import (
    avisos_de_cobertura, carregar_calendario, carregar_feriados, converter_datas, descrever_feriados,
)
from validacao import ArquivoInvalidoError

JORNADA = Jornada(
    entrada=time(8, 0), saida_almoco=time(12, 0), retorno_almoco=time(13, 0), saida=time(17, 0),
    dias_trabalho={0, 1, 2, 3, 4},
)
PIS = "00000000099"
FERIADOS_2026 = {
    date(2026, 1, 1), date(2026, 4, 3), date(2026, 4, 21), date(2026, 5, 1), date(2026, 9, 7),
    date(2026, 10, 12), date(2026, 11, 2), date(2026, 11, 15), date(2026, 11, 20), date(2026, 12, 25),
}


def _csv_temporario(linhas: list[str]) -> str:
    pasta = tempfile.mkdtemp(prefix="feriados_teste_")
    caminho = Path(pasta) / "feriados_teste.csv"
    caminho.write_text("data;nome;escopo;vale_para_empresa;situacao\n" + "\n".join(linhas) + "\n", encoding="utf-8-sig")
    return str(caminho)


# ---------------------------------------------------------------------------
# Regras de calculo (gerar_feriados.py)
# ---------------------------------------------------------------------------
def test_pascoa_em_anos_conhecidos():
    esperado = {
        2008: date(2008, 3, 23), 2024: date(2024, 3, 31), 2025: date(2025, 4, 20), 2026: date(2026, 4, 5),
        2027: date(2027, 3, 28), 2028: date(2028, 4, 16), 2038: date(2038, 4, 25),
    }
    for ano, data in esperado.items():
        assert gerador.pascoa(ano) == data, ano


def test_feriados_nacionais_2026():
    datas = {date.fromisoformat(l["data"]) for l in gerador.feriados_nacionais(2026)}
    assert datas == FERIADOS_2026


def test_consciencia_negra_so_a_partir_de_2024():
    nomes_2023 = {l["nome"] for l in gerador.feriados_nacionais(2023)}
    nomes_2024 = {l["nome"] for l in gerador.feriados_nacionais(2024)}
    assert not any("Consciência" in n for n in nomes_2023)
    assert any("Consciência" in n for n in nomes_2024)


def test_pontos_facultativos_2026():
    datas = {l["nome"]: date.fromisoformat(l["data"]) for l in gerador.pontos_facultativos(2026)}
    assert datas["Segunda-feira de Carnaval"] == date(2026, 2, 16)
    assert datas["Terça-feira de Carnaval"] == date(2026, 2, 17)
    assert datas["Corpus Christi"] == date(2026, 6, 4)
    assert datas["Véspera de Natal"] == date(2026, 12, 24)
    assert datas["Véspera de Ano Novo"] == date(2026, 12, 31)


def test_cianorte_usa_o_confirmado_e_preve_os_demais_anos():
    c2026 = {(date.fromisoformat(l["data"]), l["situacao"]) for l in gerador.feriados_cianorte(2026)}
    assert c2026 == {(date(2026, 5, 13), "confirmado"), (date(2026, 7, 26), "confirmado")}
    c2030 = {(date.fromisoformat(l["data"]), l["situacao"]) for l in gerador.feriados_cianorte(2030)}
    assert c2030 == {(date(2030, 5, 13), "previsto"), (date(2030, 7, 26), "previsto")}
    assert date(2025, 7, 28) in {date.fromisoformat(l["data"]) for l in gerador.feriados_cianorte(2025)}  # transferido


# ---------------------------------------------------------------------------
# O arquivo feriados.csv que acompanha o sistema
# ---------------------------------------------------------------------------
def test_csv_cobre_2025_a_2199_sem_duplicatas():
    cal = carregar_calendario()
    assert cal["data"].dt.year.min() == 2025 and cal["data"].dt.year.max() == 2199
    assert not cal.duplicated(["data", "escopo", "nome"]).any()
    for ano in (2025, 2026, 2100, 2199):
        nacionais = cal[(cal["escopo"] == "nacional") & (cal["data"].dt.year == ano)]
        assert len(nacionais) == 10, ano


def test_csv_de_2026_tem_os_feriados_esperados():
    feriados = carregar_feriados()
    do_ano = {d for d in feriados if d.year == 2026}
    assert do_ano == FERIADOS_2026 | {date(2026, 5, 13), date(2026, 7, 26)}  # + Cianorte
    assert date(2026, 12, 19) not in feriados          # estadual: nao e' feriado civil (padrao: nao vale)
    assert date(2026, 6, 4) not in feriados            # Corpus Christi: facultativo
    assert date(2026, 6, 4) in carregar_feriados(incluir_facultativos=True)


# ---------------------------------------------------------------------------
# Leitura e validacao do arquivo
# ---------------------------------------------------------------------------
def test_datas_iso_nao_invertem_dia_e_mes():
    # Regressao: com dayfirst=True num parser flexivel, '2026-05-01' virava 5 de janeiro.
    lidas = converter_datas(pd.Series(["2026-05-01", "01/05/2026", "2026-09-07 00:00:00", "07/09/2026"]))
    assert list(lidas.dt.date) == [date(2026, 5, 1), date(2026, 5, 1), date(2026, 9, 7), date(2026, 9, 7)]
    assert converter_datas(pd.Series(["31/02/2026", "abc"])).isna().all()


def test_vale_para_empresa_padrao_e_sobrescrita_por_linha():
    caminho = _csv_temporario([
        "2026-09-07;Independência;nacional;;lei",            # vazio + nacional  -> vale
        "2026-02-16;Carnaval;facultativo;;previsto",         # vazio + facultativo -> nao vale
        "2026-06-04;Corpus Christi;facultativo;sim;previsto",  # explicito sim
        "2026-10-12;Aparecida;nacional;não;lei",             # explicito nao (com acento)
        "19/12/2026;Emancipação PR;estadual;;previsto",      # estadual: nao vale por padrao (data dd/mm/aaaa)
    ])
    assert set(carregar_feriados(caminho)) == {date(2026, 9, 7), date(2026, 6, 4)}
    assert set(carregar_feriados(caminho, incluir_facultativos=True)) == {date(2026, 9, 7), date(2026, 6, 4), date(2026, 2, 16)}


def test_arquivo_invalido_aponta_a_linha_do_problema():
    casos = {
        "data": ["2026-09-07;A;nacional;;lei", "31/02/2026;B;nacional;;lei"],
        "escopo": ["2026-09-07;A;nacional;;lei", "2026-09-08;B;planeta;;lei"],
        "vale_para_empresa": ["2026-09-07;A;nacional;talvez;lei"],
    }
    for coluna, linhas in casos.items():
        try:
            carregar_feriados(_csv_temporario(linhas))
        except ArquivoInvalidoError as e:
            assert "linha" in str(e) and coluna in str(e), (coluna, str(e))
        else:
            raise AssertionError(f"deveria recusar {coluna} invalido")


def test_arquivo_inexistente_da_erro_amigavel():
    try:
        carregar_feriados(Path(tempfile.mkdtemp()) / "nao_existe.csv")
    except ArquivoInvalidoError as e:
        assert "não foi encontrado" in str(e)
    else:
        raise AssertionError("deveria falhar")


def test_avisos_de_cobertura():
    assert avisos_de_cobertura(date(2026, 8, 26), date(2026, 9, 25)) == []              # ano confirmado
    assert avisos_de_cobertura(date(2026, 4, 26), date(2026, 5, 25)) == []              # 13/05/2026 confirmado
    previsao = avisos_de_cobertura(date(2027, 4, 26), date(2027, 5, 25))                # 13/05/2027 e' previsao
    assert len(previsao) == 1 and "13/05/2027" in previsao[0] and "previsão" in previsao[0]
    sem_ano = avisos_de_cobertura(date(2200, 1, 1), date(2200, 1, 31))
    assert len(sem_ano) == 1 and "2200" in sem_ano[0]


def test_descrever_feriados_do_periodo():
    feriados = carregar_feriados()
    assert descrever_feriados(feriados, date(2026, 8, 26), date(2026, 9, 25)) == ["07/09 (seg) - Independência do Brasil"]


# ---------------------------------------------------------------------------
# Gerador: mesclagem
# ---------------------------------------------------------------------------
def test_mesclar_nao_duplica_e_preserva_edicao_do_rh():
    novas = gerador.gerar(2026, 2027)
    base = gerador.mesclar(None, novas)
    editado = base.copy()
    editado.loc[editado["data"] == "2026-06-04", "vale_para_empresa"] = "sim"
    de_novo = gerador.mesclar(editado, novas)
    assert len(de_novo) == len(base)
    assert de_novo.loc[de_novo["data"] == "2026-06-04", "vale_para_empresa"].iloc[0] == "sim"


def test_mesclar_troca_previsao_por_decreto_confirmado():
    novas = gerador.gerar(2026, 2026)
    obsoleta = pd.DataFrame([{
        "data": "2026-07-27", "nome": "Aniversário (previsão antiga)", "escopo": "municipal", "local": "Cianorte-PR",
        "vale_para_empresa": "", "situacao": "previsto", "fonte": "", "observacao": "",
    }])
    resultado = gerador.mesclar(obsoleta, novas)
    assert "2026-07-27" not in set(resultado["data"])
    assert "2026-07-26" in set(resultado["data"])


# ---------------------------------------------------------------------------
# Feriado na apuracao
# ---------------------------------------------------------------------------
def _batidas(dia: date, *horarios: str) -> list[datetime]:
    return [datetime.combine(dia, time(int(h[:2]), int(h[3:]))) for h in horarios]


FERIADO = date(2026, 9, 7)  # segunda-feira


def test_feriado_sem_batida_nao_e_falta():
    r = apurar_dia(PIS, FERIADO, [], JORNADA, feriado="Independência do Brasil")
    assert not r["falta"] and not r["precisa_validacao_rh"]
    assert r["status"] == "FERIADO" and r["feriado"] == "Independência do Brasil"


def test_mesmo_dia_sem_feriado_continua_sendo_falta():
    r = apurar_dia(PIS, FERIADO, [], JORNADA)
    assert r["falta"] and r["status"] == "FALTA_OU_AUSENCIA" and r["precisa_validacao_rh"]


def test_trabalho_em_feriado_vira_hora_extra_100():
    r = apurar_dia(PIS, FERIADO, _batidas(FERIADO, "08:00", "12:00", "13:00", "17:00"), JORNADA, feriado="Independência")
    assert r["status"] == "HE_100_FERIADO"
    assert r["he_100_min"] == 9 * 60 and r["he_50_min"] == 0 and r["atraso_min"] == 0
    assert not r["precisa_validacao_rh"]


def _apuracao_de_setembro(feriados) -> pd.DataFrame:
    batidas = {}
    for dia in pd.date_range("2026-09-01", "2026-09-30").date:
        if dia.weekday() < 5 and dia != FERIADO:
            batidas[(PIS, dia)] = _batidas(dia, "08:00", "12:00", "13:00", "17:00")
    batidas_por_dia = pd.DataFrame(
        [{"pis": pis, "data": dia, "batidas": b} for (pis, dia), b in batidas.items()]
    )
    return apurar_periodo(batidas_por_dia, {PIS: JORNADA}, date(2026, 9, 1), date(2026, 9, 30), feriados=feriados)


def test_feriado_nao_gera_falta_nem_tira_o_copr():
    com = consolidar_mes(_apuracao_de_setembro({FERIADO: "Independência do Brasil"}), None).iloc[0]
    assert com["dias_falta"] == 0 and com["dias_desconto_va"] == 0 and not com["perde_copr"]
    sem = consolidar_mes(_apuracao_de_setembro({}), None).iloc[0]   # o comportamento antigo
    assert sem["dias_falta"] == 1 and sem["perde_copr"]


# ---------------------------------------------------------------------------
# Cenario real de 09/2026 (so' roda se a pasta existir)
# ---------------------------------------------------------------------------
PASTA_REAL = Path("Dados e Docs Reais")


def _fila_real(feriados):
    from pipeline import carregar_cadastro_colaboradores, rodada_1_gerar_fila_de_excecoes

    cadastro = carregar_cadastro_colaboradores(str(PASTA_REAL / "cadastro_sintetico.xlsx"))
    return rodada_1_gerar_fila_de_excecoes(
        str(PASTA_REAL / "Cartão Ponto 09.2026.pdf"), cadastro, date(2026, 8, 26), date(2026, 9, 25),
        caminho_espelho=str(PASTA_REAL / "Espelho de Ponto 09.2026.pdf"), feriados=feriados,
    )


def test_cenario_real_setembro_2026_07_09_deixa_de_ser_falta():
    if not (PASTA_REAL / "Espelho de Ponto 09.2026.pdf").exists():
        return  # pasta real ausente (fora do git): nada a verificar
    _, fila_sem = _fila_real({})
    apuracao, fila_com = _fila_real(carregar_feriados())
    dia = pd.Timestamp(FERIADO)
    faltas_antes = fila_sem[(fila_sem["data"] == dia) & (fila_sem["status"] == "FALTA_OU_AUSENCIA")]
    faltas_depois = fila_com[(fila_com["data"] == dia) & (fila_com["status"] == "FALTA_OU_AUSENCIA")]
    assert len(faltas_antes) >= 10 and len(faltas_depois) == 0
    assert (apuracao.loc[pd.to_datetime(apuracao["data"]) == dia, "feriado"] == "Independência do Brasil").all()


if __name__ == "__main__":
    testes = [(n, f) for n, f in sorted(globals().items()) if n.startswith("test_") and callable(f)]
    for nome, funcao in testes:
        funcao()
        print(f"OK  {nome}")
    print(f"\n{len(testes)} teste(s) passaram.")

    if (PASTA_REAL / "Espelho de Ponto 09.2026.pdf").exists():
        _, fila_sem = _fila_real({})
        _, fila_com = _fila_real(carregar_feriados())
        dia = pd.Timestamp(FERIADO)
        n_antes = int(((fila_sem["data"] == dia) & (fila_sem["status"] == "FALTA_OU_AUSENCIA")).sum())
        n_depois = int(((fila_com["data"] == dia) & (fila_com["status"] == "FALTA_OU_AUSENCIA")).sum())
        print(f"\n07/09/2026 — colaboradores na fila como 'falta ou ausência': antes={n_antes} | depois={n_depois}")
        print(f"Linhas na fila de exceções: antes={len(fila_sem)} | depois={len(fila_com)}")
