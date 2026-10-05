"""
teste_falta_parcial.py
======================
Testes da deteccao de falta de meio periodo (dias com batidas faltando), do
cruzamento com o Espelho de Ponto e do cod. 8069 (horas de falta).

Rodar com pytest (`pytest teste_falta_parcial.py`) ou direto
(`python teste_falta_parcial.py`). Neste modo, se a pasta "Dados e Docs Reais"
existir (ela e' ignorada pelo git), roda tambem o cenario com os PDFs reais de
09/2026 e imprime o relatorio de conferencia contra os TOTAIS do Secullum.
"""

from __future__ import annotations

from datetime import date, datetime, time
from pathlib import Path

import pandas as pd

from business_rules import (
    Jornada, STATUS_FALTA_PARCIAL, apurar_dia, consolidar_mes, gerar_fila_validacao_rh,
)
from pipeline import (
    CRUZAMENTO_BATIDA_NAO_IMPORTADA, CRUZAMENTO_ERRO_RELOGIO, CRUZAMENTO_FALTA_PARCIAL,
    CRUZAMENTO_FALTA_REAL, _anexar_cruzamento_espelho, rodada_2_gerar_relacao_de_valores,
)

JORNADA = Jornada(
    entrada=time(8, 0), saida_almoco=time(12, 0), retorno_almoco=time(13, 0), saida=time(17, 0),
    dias_trabalho={0, 1, 2, 3, 4},
)  # carga de 8h (480 min)
DIA = date(2026, 9, 3)  # quinta-feira
PIS = "00000000099"


def _batidas(*horarios: str) -> list[datetime]:
    return [datetime.combine(DIA, time(int(h[:2]), int(h[3:]))) for h in horarios]


def _apurar(*horarios: str) -> dict:
    return apurar_dia(PIS, DIA, _batidas(*horarios), JORNADA)


# ---------------------------------------------------------------------------
# apurar_dia
# ---------------------------------------------------------------------------
def test_duas_batidas_so_da_tarde_sao_falta_parcial():
    r = _apurar("13:00", "17:29")  # faltou de manha (caso do William em 03/09)
    assert r["status"] == STATUS_FALTA_PARCIAL
    assert r["precisa_validacao_rh"]
    assert r["atraso_min"] == 0
    assert r["minutos_falta_parcial"] == 480 - 269


def test_duas_batidas_so_da_manha_sao_falta_parcial():
    r = _apurar("08:00", "12:03")  # faltou a tarde
    assert r["status"] == STATUS_FALTA_PARCIAL
    assert r["minutos_falta_parcial"] == 480 - 243


def test_tres_batidas_sinalizadas_com_minutos():
    r = _apurar("08:00", "12:00", "13:00")  # sem a saida
    assert r["status"] == STATUS_FALTA_PARCIAL
    assert r["minutos_falta_parcial"] == 240
    assert "08:00, 12:00, 13:00" in r["motivo_validacao"]


def test_uma_batida_vale_como_falta_de_dia_inteiro():
    r = _apurar("17:31")
    assert r["status"] == STATUS_FALTA_PARCIAL and r["precisa_validacao_rh"]
    assert r["minutos_falta_parcial"] == 480
    assert r["falta"]  # nenhum periodo completo: segue a regra de falta de dia inteiro
    resumo = consolidar_mes(pd.DataFrame([r]), None).iloc[0]
    assert resumo["dias_falta"] == 1 and resumo["horas_falta"] == 0 and resumo["dias_falta_parcial"] == 0


def test_turno_corrido_de_duas_batidas_nao_e_sinalizado():
    r = _apurar("08:00", "17:00")  # entrada e saida cobrem a jornada inteira
    assert r["status"] != STATUS_FALTA_PARCIAL
    assert not r["precisa_validacao_rh"]
    assert r["minutos_falta_parcial"] == 0


def test_quatro_batidas_normais_continuam_ok():
    r = _apurar("08:00", "12:00", "13:00", "17:00")
    assert r["status"] == "OK" and not r["precisa_validacao_rh"]


def test_mais_de_quatro_batidas_continuam_batida_incompleta():
    r = _apurar("08:00", "12:00", "12:30", "13:00", "17:00")
    assert r["status"] == "BATIDA_INCOMPLETA" and r["precisa_validacao_rh"]


# ---------------------------------------------------------------------------
# consolidar_mes
# ---------------------------------------------------------------------------
def _apuracao_com_falta_parcial() -> pd.DataFrame:
    return pd.DataFrame([_apurar("13:00", "17:29")])  # 211 min faltantes


def _decisao(classificacao: str) -> pd.DataFrame:
    return pd.DataFrame([{"pis": PIS, "data": "03/09/2026", "classificacao": classificacao}])


def test_falta_parcial_sem_decisao_conta_como_horas_e_tira_copr():
    r = consolidar_mes(_apuracao_com_falta_parcial(), None).iloc[0]
    assert r["horas_falta"] == round(211 / 60, 2)
    assert r["dias_falta"] == 0 and r["dias_falta_parcial"] == 1
    assert r["perde_copr"]
    assert r["dias_desconto_va"] == 0 and r["semanas_perde_dsr"] == 0  # VA/DSR seguem por dia inteiro


def test_falta_parcial_justificada_e_injustificada_somam_horas():
    for classificacao in ("FALTA_JUSTIFICADA", "FALTA_INJUSTIFICADA"):
        r = consolidar_mes(_apuracao_com_falta_parcial(), _decisao(classificacao)).iloc[0]
        assert r["horas_falta"] == round(211 / 60, 2), classificacao
        assert r["perde_copr"], classificacao
        assert r["dias_desconto_va"] == 0, classificacao


def test_erro_de_relogio_abona_a_falta_parcial():
    r = consolidar_mes(_apuracao_com_falta_parcial(), _decisao("ERRO_RELOGIO_ABONAR")).iloc[0]
    assert r["horas_falta"] == 0 and r["dias_falta_parcial"] == 0
    assert not r["perde_copr"]


def test_atraso_tambem_entra_nas_horas_de_falta():
    apuracao = pd.DataFrame([_apurar("08:00", "12:00", "13:00", "16:00")])  # saiu 1h antes
    r = consolidar_mes(apuracao, None).iloc[0]
    assert r["atraso_min"] == 60 and r["horas_falta"] == 1.0
    assert not r["perde_copr"]  # atraso nao e' falta (regra ja' existente)


# ---------------------------------------------------------------------------
# cruzamento com o Espelho
# ---------------------------------------------------------------------------
def _fila_parcial() -> pd.DataFrame:
    return gerar_fila_validacao_rh(_apuracao_com_falta_parcial())


def _espelho(*brutas: str, dia: date = DIA) -> pd.DataFrame:
    return pd.DataFrame([{"pis": PIS, "data": dia, "batidas_brutas": list(brutas)}])


def test_espelho_com_batida_a_mais_indica_erro_de_relogio():
    fila = _anexar_cruzamento_espelho(_fila_parcial(), _espelho("07:58", "13:00", "17:29"))
    linha = fila.iloc[0]
    assert linha["cruzamento_espelho"] == CRUZAMENTO_BATIDA_NAO_IMPORTADA
    assert "07:58, 13:00, 17:29" in linha["motivo_validacao"]


def test_espelho_sem_batida_a_mais_indica_falta_parcial():
    fila = _anexar_cruzamento_espelho(_fila_parcial(), _espelho("13:00", "17:29"))
    assert fila.iloc[0]["cruzamento_espelho"] == CRUZAMENTO_FALTA_PARCIAL


def test_batida_duplicada_no_espelho_nao_conta_como_batida_a_mais():
    fila = _anexar_cruzamento_espelho(_fila_parcial(), _espelho("13:00", "17:29", "17:30"))
    assert fila.iloc[0]["cruzamento_espelho"] == CRUZAMENTO_FALTA_PARCIAL


def test_dia_sem_linha_no_espelho_e_falta_parcial():
    fila = _anexar_cruzamento_espelho(_fila_parcial(), _espelho("08:00", dia=date(2026, 9, 4)))
    assert fila.iloc[0]["cruzamento_espelho"] == CRUZAMENTO_FALTA_PARCIAL


def test_dia_sem_batida_segue_a_regra_antiga():
    sem_batida = gerar_fila_validacao_rh(pd.DataFrame([_apurar()]))
    com = _anexar_cruzamento_espelho(sem_batida, _espelho("08:00"))
    sem = _anexar_cruzamento_espelho(sem_batida, _espelho())
    assert com.iloc[0]["cruzamento_espelho"] == CRUZAMENTO_ERRO_RELOGIO
    assert sem.iloc[0]["cruzamento_espelho"] == CRUZAMENTO_FALTA_REAL


# ---------------------------------------------------------------------------
# Relacao de valores: 8069
# ---------------------------------------------------------------------------
def _cadastro() -> pd.DataFrame:
    return pd.DataFrame([{
        "pis": PIS, "matricula": "99", "nome": "WILLIAN", "codigo_empresa": 1, "razao_social": "EMPRESA",
        "entrada": "08:00", "saida_almoco": "12:00", "retorno_almoco": "13:00", "saida": "17:00",
        "dias_trabalho": "0,1,2,3,4",
    }])


def test_codigo_8069_deixa_de_ser_zero():
    relacao = rodada_2_gerar_relacao_de_valores(
        _apuracao_com_falta_parcial(), _cadastro(),
        pd.DataFrame({"pis": [PIS], "valor_consignado": [0.0]}),
        pd.DataFrame({"pis": [PIS], "saldo_banco_horas": [0.0]}),
        None, "09/2026",
    )
    linha = relacao.iloc[0]
    assert linha["8069"] == round(211 / 60, 2)
    assert linha["8792"] == 0
    assert linha["_dias_falta_parcial"] == 1 and linha["_horas_falta_parcial"] == round(211 / 60, 2)
    assert not any(c.startswith("_sec_") for c in relacao.columns)  # sem Cartao PDF, sem bloco Secullum


def test_conferencia_secullum_so_aparece_com_totais_e_nao_altera_valores():
    apuracao = _apuracao_com_falta_parcial()
    apuracao.attrs["totais_secullum"] = pd.DataFrame(
        [{"pis": PIS, "faltas_dias": 8.0, "faltas_horas": 1085, "extra_50": 39, "extra_100": 0}]
    )
    args = (_cadastro(), pd.DataFrame({"pis": [PIS], "valor_consignado": [0.0]}),
            pd.DataFrame({"pis": [PIS], "saldo_banco_horas": [0.0]}), None, "09/2026")
    com = rodada_2_gerar_relacao_de_valores(apuracao, *args).iloc[0]
    sem = rodada_2_gerar_relacao_de_valores(_apuracao_com_falta_parcial(), *args).iloc[0]
    assert com["_sec_horas_falta"] == round(1085 / 60, 2)
    assert com["_dif_horas_falta"] == round(com["8069"] - 1085 / 60, 2)
    for codigo in ("0150", "0200", "8069", "8792", "8794", "0981", "0999"):
        assert com[codigo] == sem[codigo], codigo


# ---------------------------------------------------------------------------
# Cenario com os arquivos reais (so' roda se a pasta existir)
# ---------------------------------------------------------------------------
PASTA_REAL = Path("Dados e Docs Reais")


def _cenario_real(imprimir: bool = False):
    from pipeline import carregar_cadastro_colaboradores, carregar_consignados, rodada_1_gerar_fila_de_excecoes
    from banco_horas import consolidar_banco_horas
    from validacao import canonicalizar_pis

    cadastro = carregar_cadastro_colaboradores(str(PASTA_REAL / "cadastro_sintetico.xlsx"))
    cartao = str(PASTA_REAL / "Cartão Ponto 09.2026.pdf")
    espelho = str(PASTA_REAL / "Espelho de Ponto 09.2026.pdf")
    apuracao, fila = rodada_1_gerar_fila_de_excecoes(
        cartao, cadastro, date(2026, 8, 26), date(2026, 9, 25), caminho_espelho=espelho
    )
    consignados = carregar_consignados(str(PASTA_REAL / "Consignado Modificado.xlsx"), cadastro=cadastro)
    banco = consolidar_banco_horas(cadastro["pis"].tolist(), None, None)
    relacao = rodada_2_gerar_relacao_de_valores(apuracao, cadastro, consignados, banco, None, "09/2026")
    return cadastro, apuracao, fila, relacao, canonicalizar_pis("99")


def test_cenario_real_william_dias_parciais_sinalizados():
    if not (PASTA_REAL / "Espelho de Ponto 09.2026.pdf").exists():
        return  # pasta real ausente (fora do git): nada a verificar
    _, apuracao, fila, relacao, pis_william = _cenario_real()
    fila = fila.assign(data=pd.to_datetime(fila["data"]))
    for dia in ("2026-09-01", "2026-09-02", "2026-09-03", "2026-09-04", "2026-09-16", "2026-09-25"):
        linha = fila[(fila["pis"] == pis_william) & (fila["data"] == dia)]
        assert len(linha) == 1 and linha.iloc[0]["status"] == STATUS_FALTA_PARCIAL, dia
    assert relacao.loc[relacao["codigo_folha"].astype(str) == "99", "8069"].iloc[0] > 0


if __name__ == "__main__":
    testes = [(n, f) for n, f in sorted(globals().items()) if n.startswith("test_") and callable(f)]
    for nome, funcao in testes:
        funcao()
        print(f"OK  {nome}")
    print(f"\n{len(testes)} teste(s) passaram.")

    if (PASTA_REAL / "Espelho de Ponto 09.2026.pdf").exists():
        pd.set_option("display.width", 250)
        pd.set_option("display.max_columns", 40)
        pd.set_option("display.max_rows", 200)
        cadastro, apuracao, fila, relacao, pis_william = _cenario_real()
        fila = fila.assign(data=pd.to_datetime(fila["data"]))
        print("\n=== Matricula 99 (WILLIAN) — dias de interesse ===")
        dias = pd.to_datetime(["2026-09-01", "2026-09-02", "2026-09-03", "2026-09-04", "2026-09-16", "2026-09-25"])
        mostrar = fila[(fila["pis"] == pis_william) & (fila["data"].isin(dias))]
        print(mostrar[["data", "status", "n_batidas", "minutos_falta_parcial", "cruzamento_espelho"]].to_string(index=False))
        for _, linha in mostrar.iterrows():
            print(f"  {linha['data']:%d/%m}: {linha['motivo_validacao']}")
        william = relacao[relacao["codigo_folha"].astype(str) == "99"].iloc[0]
        print(f"\nWILLIAN -> 8069={william['8069']}  8792={william['8792']}  dias parciais={william['_dias_falta_parcial']}")
        print("\n=== Nosso x Secullum (todos) ===")
        cols = ["nome_colaborador", "8792", "_sec_dias_falta", "_dif_dias_falta", "8069", "_sec_horas_falta",
                "_dif_horas_falta", "0150", "_sec_he50", "_dif_he50", "0200", "_sec_he100", "_dif_he100"]
        print(relacao[cols].to_string(index=False))
