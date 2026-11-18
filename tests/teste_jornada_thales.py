"""Jornada de sexta, admissao, dias abonados para todos e a Planilha do Thales (dados f ficticios)."""
from __future__ import annotations

import tempfile
from datetime import date, datetime, time
from pathlib import Path

import pandas as pd

from banco_horas import consolidar_banco_horas
from business_rules import Jornada, apurar_dia, apurar_periodo
from pipeline import (
    carregar_cadastro_colaboradores, ler_dias_abonados, rodada_2_gerar_relacao_de_valores,
)
from planilha_thales import exportar_planilha_thales
from validacao import ArquivoInvalidoError

SEXTA, SEGUNDA = date(2026, 10, 9), date(2026, 10, 5)


def _jornada(**extra) -> Jornada:
    return Jornada(time(7, 30), time(12, 0), time(13, 0), time(17, 30), {0, 1, 2, 3, 4}, **extra)


def _batidas(dia: date, *horarios: str) -> list[datetime]:
    return [datetime.combine(dia, time.fromisoformat(h)) for h in horarios]


def test_sexta_usa_a_saida_da_sexta_e_os_outros_dias_nao_mudam():
    j = _jornada(saida_sexta=time(16, 30))
    sexta = apurar_dia("1", SEXTA, _batidas(SEXTA, "07:30", "12:00", "13:00", "16:30"), j)
    assert sexta["status"] == "OK" and sexta["atraso_min"] == 0 and sexta["he_50_min"] == 0
    # mesma saida numa segunda: 1h de atraso (a jornada de segunda vai ate' 17:30)
    segunda = apurar_dia("1", SEGUNDA, _batidas(SEGUNDA, "07:30", "12:00", "13:00", "16:30"), j)
    assert segunda["atraso_min"] == 60
    # sem `saida_sexta` a sexta continua como antes (compatibilidade com cadastro antigo)
    assert apurar_dia("1", SEXTA, _batidas(SEXTA, "07:30", "12:00", "13:00", "16:30"), _jornada())["atraso_min"] == 60


def test_antes_da_admissao_nao_e_falta():
    j = _jornada(admissao=date(2026, 10, 7))
    assert apurar_dia("1", SEGUNDA, [], j)["falta"] is False
    assert apurar_dia("1", date(2026, 10, 7), [], j)["falta"] is True


def test_dia_abonado_para_todos_nao_gera_falta_nem_fila_mas_dia_completo_e_apurado():
    j = _jornada()
    parcial = apurar_dia("1", SEGUNDA, _batidas(SEGUNDA, "07:29"), j, abono_coletivo="Falha do relógio")
    assert parcial["status"] == "ABONO_COLETIVO" and parcial["falta"] is False
    assert parcial["precisa_validacao_rh"] is False and parcial["atraso_min"] == 0
    vazio = apurar_dia("1", SEGUNDA, [], j, abono_coletivo="Falha do relógio")
    assert vazio["falta"] is False and vazio["precisa_validacao_rh"] is False
    completo = apurar_dia("1", SEGUNDA, _batidas(SEGUNDA, "07:30", "12:00", "13:00", "17:50"), j,
                          abono_coletivo="Falha do relógio")
    assert completo["he_50_min"] == 20    # 4 batidas: apuracao normal


def test_ler_dias_abonados():
    ini, fim = date(2026, 9, 26), date(2026, 10, 25)
    assert set(ler_dias_abonados("14/10/2026, 15/10/2026", ini, fim)) == {date(2026, 10, 14), date(2026, 10, 15)}
    assert ler_dias_abonados("", ini, fim) == {}
    for ruim in ("14-10-2026", "01/01/2026"):
        try:
            ler_dias_abonados(ruim, ini, fim)
        except ArquivoInvalidoError:
            continue
        raise AssertionError(f"deveria recusar {ruim!r}")


def test_planilha_do_thales_no_modelo_da_cliente():
    cadastro = pd.DataFrame([
        dict(pis="1", matricula="1", nome="ANA DA SILVA", codigo_empresa="20", razao_social="EMPRESA A LTDA",
             entrada="07:30", saida_almoco="12:00", retorno_almoco="13:00", saida="17:30", dias_trabalho="0,1,2,3,4",
             centro_custo="07 e 25", valor_va=960, cooper=200, desconto_folha=224, observacao="*venda 10X",
             admissao="01/01/2020", nota_rodape="atestado em 28/05, detenção desde 24/06"),
        dict(pis="2", matricula="2", nome="BETO NOVO", codigo_empresa="20", razao_social="EMPRESA A LTDA",
             entrada="07:30", saida_almoco="12:00", retorno_almoco="13:00", saida="17:30", dias_trabalho="0,1,2,3,4",
             valor_va=960, cooper=0, admissao="14/10/2026"),
    ])
    pasta = Path(tempfile.mkdtemp())
    cadastro.to_excel(pasta / "cad.xlsx", index=False)
    cad = carregar_cadastro_colaboradores(str(pasta / "cad.xlsx"))
    p1, p2 = cad["pis"].tolist()   # PIS canonico (11 digitos)\n    # Ana: 1 falta inteira injustificada (05/10) + HE de 50 minutos (06/10); Beto: admitido em 14/10, normal depois
    dias = pd.date_range("2026-10-05", "2026-10-16").date
    bat = {}
    for d in dias:
        if d.weekday() >= 5 or d in (date(2026, 10, 12),):
            continue
        saida = "16:30" if d.weekday() == 4 else "17:30"
        if d >= date(2026, 10, 14):
            bat[(p2, d)] = _batidas(d, "07:30", "12:00", "13:00", saida)
        if d != date(2026, 10, 5):
            bat[(p1, d)] = _batidas(d, "07:30", "12:00", "13:00", "18:20" if d == date(2026, 10, 6) else saida)
    batidas_por_dia = pd.DataFrame([{"pis": k[0], "data": k[1], "batidas": v} for k, v in bat.items()])
    jornadas = {p1: Jornada(time(7, 30), time(12), time(13), time(17, 30), {0, 1, 2, 3, 4}, saida_sexta=time(16, 30)),
                p2: Jornada(time(7, 30), time(12), time(13), time(17, 30), {0, 1, 2, 3, 4}, saida_sexta=time(16, 30),
                             admissao=date(2026, 10, 14))}
    ap = apurar_periodo(batidas_por_dia, jornadas, date(2026, 10, 5), date(2026, 10, 16),
                        feriados={date(2026, 10, 12): "Feriado"})
    rel = rodada_2_gerar_relacao_de_valores(ap, cad, None, consolidar_banco_horas(cad["pis"].tolist(), None, None),
                                            None, "10/2026")
    caminho = exportar_planilha_thales(rel, ap, None, cad, "10/2026", date(2026, 9, 26), date(2026, 10, 25), str(pasta))

    from openpyxl import load_workbook
    ws = load_workbook(caminho).active
    assert ws["A1"].value == "Fechamento 10/2026 " and ws["A2"].value == "Período 26/09/2026 a 25/10/2026"
    assert ws["B4"].value == "EMPRESA A LTDA:"
    assert [c.value for c in ws[5]][1:] == ["CC", "Colaborador", "Falta C/ At.", "Falta S/ At.", "H/E 50%", "H/E 100%",
                                            "Desc. em folha ", "Alelo Cartão", "Cooper", "Observações"]
    ana = [c.value for c in ws[6]]
    assert ana[:3] == [1, "07 e 25", "Ana da Silva"]
    assert ana[3] is None and ana[4] == "05/10/2026" and ana[5] == "00:50"      # falta sem atestado + HE 50%
    assert ana[7] == 224 and ana[8] == 960 - 32 and ana[9] == 200 and ana[10] == "*venda 10X"   # 1/30 de VA por falta
    beto = [c.value for c in ws[7]]
    assert beto[2] == "Beto Novo" and beto[4] is None
    assert beto[8] == 576 and beto[10] == "*Experiencia (Admitido em 14/10/2026)"    # 18/30 do VA (14 a 31/10)
    assert ws["I8"].value == "=SUM(I6:I7)" and ws["J8"].value == "=SUM(J6:J7)"
    assert ws["B9"].value == "Ana: atestado em 28/05, detenção desde 24/06" and ws["B10"].value is None
