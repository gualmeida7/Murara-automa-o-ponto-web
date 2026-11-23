"""
Regra do Cooper: perde o valor quando as faltas do mes (dia inteiro + parcial, justificadas ou nao)
passam de LIMITE_FALTAS_COOPER_MIN (2:00). Atraso nao conta; `cooper_fixo` e' isento; quem nao tem
Cooper nao e' afetado. Dados totalmente ficticios.
"""
from __future__ import annotations

from datetime import date, datetime, time
from pathlib import Path

import pandas as pd

from business_rules import LIMITE_FALTAS_COOPER_MIN, Jornada, apurar_dia, consolidar_mes
from pipeline import rodada_2_gerar_relacao_de_valores
from planilha_thales import exportar_planilha_thales

JORNADA = Jornada(
    entrada=time(8, 0), saida_almoco=time(12, 0), retorno_almoco=time(13, 0), saida=time(17, 0),
    dias_trabalho={0, 1, 2, 3, 4},
)
PIS = "00000000001"
DIA = date(2026, 9, 3)  # quinta-feira


def _parcial(minutos: int, dia: date = DIA) -> dict:
    """Dia com falta parcial de `minutos` (3 batidas, valor de minutos ajustado)."""
    r = apurar_dia(PIS, dia, [datetime.combine(dia, time(h, m)) for h, m in ((8, 0), (12, 0), (13, 0))], JORNADA)
    r["minutos_falta_parcial"] = minutos
    return r


def _cadastro(cooper: float | None = 200.0, fixo: bool = False) -> pd.DataFrame:
    return pd.DataFrame([{
        "pis": PIS, "matricula": "1", "nome": "COLABORADOR", "codigo_empresa": 1, "razao_social": "EMPRESA",
        "entrada": "08:00", "saida_almoco": "12:00", "retorno_almoco": "13:00", "saida": "17:00",
        "dias_trabalho": "0,1,2,3,4", "cooper": cooper, "cooper_fixo": fixo,
    }])


def _relacao(apuracao: pd.DataFrame, cadastro: pd.DataFrame, decisoes=None) -> pd.Series:
    return rodada_2_gerar_relacao_de_valores(
        apuracao, cadastro, None, pd.DataFrame({"pis": [PIS], "saldo_banco_horas": [0.0]}), decisoes, "09/2026",
    ).iloc[0]


def test_limite_e_uma_constante_de_duas_horas():
    assert LIMITE_FALTAS_COOPER_MIN == 120


def test_exatamente_duas_horas_mantem_o_cooper():
    r = _relacao(pd.DataFrame([_parcial(120)]), _cadastro())
    assert r["_faltas_cooper_min"] == 120 and not r["_cooper_zerado"]


def test_duas_horas_e_um_minuto_perde_o_cooper():
    r = _relacao(pd.DataFrame([_parcial(121)]), _cadastro())
    assert r["_faltas_cooper_min"] == 121 and r["_cooper_zerado"]


def test_soma_de_varias_faltas_parciais_no_mes():
    apuracao = pd.DataFrame([_parcial(60, date(2026, 9, 3)), _parcial(61, date(2026, 9, 4))])
    assert _relacao(apuracao, _cadastro())["_cooper_zerado"]


def test_falta_justificada_tambem_conta():
    decisao = pd.DataFrame([{"pis": PIS, "data": "03/09/2026", "classificacao": "FALTA_JUSTIFICADA"}])
    r = _relacao(pd.DataFrame([_parcial(121)]), _cadastro(), decisao)
    assert r["_cooper_zerado"]


def test_falta_de_dia_inteiro_conta_a_carga_da_jornada():
    dia = apurar_dia(PIS, DIA, [], JORNADA)
    assert dia["minutos_falta_dia"] == 480
    assert consolidar_mes(pd.DataFrame([dia]), None).iloc[0]["minutos_faltas_cooper"] == 480
    assert _relacao(pd.DataFrame([dia]), _cadastro())["_cooper_zerado"]


def test_atraso_nao_conta():
    dia = _parcial(0)
    dia.update(status="OK", atraso_min=300, minutos_falta_parcial=0)
    r = _relacao(pd.DataFrame([dia]), _cadastro())
    assert r["_faltas_cooper_min"] == 0 and not r["_cooper_zerado"]


def test_erro_de_relogio_abonado_nao_conta():
    decisao = pd.DataFrame([{"pis": PIS, "data": "03/09/2026", "classificacao": "ERRO_RELOGIO_ABONAR"}])
    assert not _relacao(pd.DataFrame([_parcial(300)]), _cadastro(), decisao)["_cooper_zerado"]


def test_cooper_fixo_e_isento():
    r = _relacao(pd.DataFrame([_parcial(480)]), _cadastro(fixo=True))
    assert r["_faltas_cooper_min"] == 480 and not r["_cooper_zerado"]


def test_linha_sem_cooper_nao_e_afetada():
    for valor in (None, 0.0):
        assert not _relacao(pd.DataFrame([_parcial(480)]), _cadastro(cooper=valor))["_cooper_zerado"]


def test_planilha_thales_zera_o_cooper_e_anota(tmp_path: Path):
    apuracao = pd.DataFrame([_parcial(121)])
    cadastro = _cadastro()
    relacao = rodada_2_gerar_relacao_de_valores(
        apuracao, cadastro, None, pd.DataFrame({"pis": [PIS], "saldo_banco_horas": [0.0]}), None, "09/2026",
    )
    caminho = exportar_planilha_thales(relacao, apuracao, None, cadastro, "09/2026",
                                       date(2026, 8, 26), date(2026, 9, 25), str(tmp_path))
    from openpyxl import load_workbook
    ws = load_workbook(caminho).active
    linha = next(r for r in ws.iter_rows(values_only=True) if r[1] is None and r[0] == 1)
    assert linha[9] == 0 and "Cooper zerado" in linha[10]
