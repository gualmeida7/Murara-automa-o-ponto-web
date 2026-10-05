"""
business_rules.py
==================
Cruza as batidas de ponto (ja agrupadas por dia pelo afd_parser) com a
jornada contratual de cada colaborador e aplica as regras de negocio da
empresa:

  - Calculo de horas extra 50% / 100%, atraso e falta.
  - Deteccao de inconsistencia (batida faltando / dia sem nenhuma batida)
    para validacao humana (RH so olha as excecoes, nao o arquivo inteiro).
  - Premio COPR: perde o premio integral se houver QUALQUER falta/ausencia
    no mes, mesmo justificada por atestado.
  - Vale Alimentacao: desconto proporcional de dias por falta SEM
    justificativa.
  - DSR: perda proporcional do DSR da semana quando ha falta injustificada
    na semana.

Nada aqui decide sozinho se uma falta e' justificada ou nao - isso exige
julgamento humano (atestado medico, aviso previo, etc.). O papel deste
modulo e' calcular tudo que PODE ser calculado automaticamente e isolar,
em uma tabela separada, exatamente as linhas que precisam de decisao humana.
Essa tabela e' o que alimenta o painel de aprovacao (dashboard_aprovacao.html).
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from itertools import combinations
from typing import Mapping

import pandas as pd

TOLERANCIA_DIARIA_MIN = 10       # Art. 58 §1º CLT: limite maximo diario
TOLERANCIA_POR_BATIDA_MIN = 5    # Art. 58 §1º CLT: limite por marcacao individual

# Dia com 1, 2 ou 3 batidas que NAO cobre a jornada inteira: possivel falta de
# meio periodo (ex.: faltou a tarde). Sempre vai para a fila do RH. Se nenhum
# periodo ficou completo, vale como falta de dia inteiro (falta=True).
STATUS_FALTA_PARCIAL = "FALTA_PARCIAL_CANDIDATA"
STATUS_FALTAS_A_CLASSIFICAR = ("FALTA_OU_AUSENCIA", STATUS_FALTA_PARCIAL)


@dataclass
class Jornada:
    """Jornada contratual de um colaborador."""
    entrada: time
    saida_almoco: time
    retorno_almoco: time
    saida: time
    dias_trabalho: set[int]  # 0=segunda ... 6=domingo (datetime.weekday())
    tolerancia_min: int = TOLERANCIA_DIARIA_MIN
    tolerancia_por_batida_min: int = TOLERANCIA_POR_BATIDA_MIN

    def carga_horaria_min(self) -> int:
        manha = datetime.combine(date.min, self.saida_almoco) - datetime.combine(date.min, self.entrada)
        tarde = datetime.combine(date.min, self.saida) - datetime.combine(date.min, self.retorno_almoco)
        return int((manha + tarde).total_seconds() // 60)


def _minutos_entre(t1: datetime, t2: datetime) -> int:
    return int((t2 - t1).total_seconds() // 60)


from datetime import datetime, time, date

def _ajustar_batida_pela_tolerancia(
    horario_real: datetime, 
    horario_programado: time, 
    dia: date, 
    tolerancia_min: int
) -> datetime:
    """
    Aplica a regra de tolerância de ponto segundo a CLT (Art. 58 § 1º / Súmula 366 TST):
    - Se a variação for de até `tolerancia_min` (ex: 5 min), considera o horário programado.
    - Se ultrapassar a tolerância, cobra/paga a TOTALIDADE do tempo real (não apenas o excedente).
    """
    programado_dt = datetime.combine(dia, horario_programado)
    desvio_min = (horario_real - programado_dt).total_seconds() / 60

    # Dentro do limite de tolerância: ajusta para o horário contratual/programado
    if abs(desvio_min) <= tolerancia_min:
        return programado_dt

    # Ultrapassou a tolerância: considera a marcação real integral
    return horario_real



def _hhmm(minutos: int) -> str:
    return f"{minutos // 60}h{minutos % 60:02d}"


def _slots_da_jornada(jornada: Jornada, dia: date) -> list[datetime]:
    """Os 4 horarios programados do dia: entrada, saida p/ almoco, retorno, saida."""
    return [
        datetime.combine(dia, t)
        for t in (jornada.entrada, jornada.saida_almoco, jornada.retorno_almoco, jornada.saida)
    ]


def _atribuir_batidas_aos_slots(batidas: list[datetime], slots: list[datetime]) -> tuple[int, ...]:
    """
    Descobre a qual marcacao programada (0=entrada ... 3=saida) cada batida
    corresponde, preservando a ordem e minimizando a distancia total ate' o
    horario programado. Assim "13:00 e 17:29" viram (retorno, saida) e nao
    (entrada, saida). Precisa de 1 a 4 batidas.
    """
    ordenadas = sorted(batidas)
    return min(
        combinations(range(len(slots)), len(ordenadas)),
        key=lambda idx: sum(abs((b - slots[i]).total_seconds()) for b, i in zip(ordenadas, idx)),
    )


def _minutos_nao_cobertos(batidas: list[datetime], idx_slots: tuple[int, ...], carga_min: int) -> int:
    """
    Estimativa de minutos da jornada sem cobertura de batidas: so' conta como
    trabalhado o periodo (manha/tarde) em que AMBAS as batidas existem. E'
    uma estimativa para o RH confirmar - o Secullum e' a referencia final.
    """
    por_slot = dict(zip(idx_slots, sorted(batidas)))
    coberto = 0
    for ini, fim in ((0, 1), (2, 3)):
        if ini in por_slot and fim in por_slot:
            coberto += _minutos_entre(por_slot[ini], por_slot[fim])
    return max(0, min(carga_min, carga_min - coberto))


def apurar_dia(
    pis: str, dia: date, batidas: list[datetime], jornada: Jornada, feriado: str | None = None,
) -> dict:
    """
    Apura um unico dia de um colaborador. Retorna um dict com o resultado
    do calculo automatico E um campo `precisa_validacao_rh` quando o dia
    nao pode ser fechado sem intervencao humana.

    Tolerancia da CLT (Art. 58 §1º): variacoes de ate' 5 minutos por
    marcacao, respeitado o limite de 10 minutos no total do dia, NAO geram
    hora extra nem atraso. As duas camadas sao aplicadas aqui: cada
    marcacao e' "encostada" no horario programado quando esta dentro dos 5
    min (via `_ajustar_batida_pela_tolerancia`), e o saldo do dia so' vira
    HE/atraso se ainda assim ultrapassar os 10 min agregados
    (`jornada.tolerancia_min`).

    `feriado` e' o nome do feriado do dia (ou None). Feriado e' dia nao util,
    como o domingo: sem batida nao e' falta; com batida, o tempo trabalhado
    vira hora extra 100%.
    """
    dow = dia.weekday()
    dia_deveria_trabalhar = dow in jornada.dias_trabalho and feriado is None
    n_batidas = len(batidas)

    base = {
        "pis": pis,
        "data": dia,
        "n_batidas": n_batidas,
        "he_50_min": 0,
        "he_100_min": 0,
        "atraso_min": 0,
        "minutos_falta_parcial": 0,
        "feriado": feriado or "",
        "falta": False,
        "status": "OK",
        "precisa_validacao_rh": False,
        "motivo_validacao": "",
    }

    # Domingo/feriado/dia nao util com batidas => tudo vira hora extra 100%
    # (a tolerancia por marcacao nao se aplica aqui: nao ha horario
    # programado de referencia num dia em que o colaborador nao deveria
    # trabalhar, entao qualquer minuto batido conta.)
    if not dia_deveria_trabalhar:
        if n_batidas >= 2:
            trabalhado_min = _minutos_entre(min(batidas), max(batidas))
            base["he_100_min"] = trabalhado_min
            base["status"] = "HE_100_FERIADO" if feriado else "HE_100_DIA_NAO_UTIL"
        elif feriado:
            base["status"] = "FERIADO"
        return base

    # Dia normal de trabalho, sem nenhuma batida => possivel falta
    if n_batidas == 0:
        base["falta"] = True
        base["status"] = "FALTA_OU_AUSENCIA"
        base["precisa_validacao_rh"] = True
        base["motivo_validacao"] = "Nenhuma batida no dia - confirmar se e' falta real ou erro de relogio"
        return base

    carga_min = jornada.carga_horaria_min()

    if n_batidas > 4:
        base["status"] = "BATIDA_INCOMPLETA"
        base["precisa_validacao_rh"] = True
        base["motivo_validacao"] = f"{n_batidas} batida(s) no dia - esperado 2 ou 4"
        return base

    # A jornada do cadastro sempre tem intervalo de almoco, entao o esperado sao 4
    # batidas. Com 1 a 3 batidas, so' 2 batidas em (entrada, saida) sao aceitas
    # como turno corrido; qualquer outra combinacao e' possivel falta parcial.
    if n_batidas < 4:
        idx_slots = _atribuir_batidas_aos_slots(batidas, _slots_da_jornada(jornada, dia))
        turno_corrido = n_batidas == 2 and idx_slots == (0, 3)
        if not turno_corrido:
            faltante_min = _minutos_nao_cobertos(batidas, idx_slots, carga_min)
            horarios = ", ".join(b.strftime("%H:%M") for b in sorted(batidas))
            base["status"] = STATUS_FALTA_PARCIAL
            base["minutos_falta_parcial"] = faltante_min
            base["precisa_validacao_rh"] = True
            if faltante_min >= carga_min:
                # Nenhum periodo (manha/tarde) ficou completo: na pratica e' o dia
                # inteiro ausente, como o Secullum conta (FALTAS em dias). Entra
                # como falta de dia inteiro, igual ao dia sem nenhuma batida.
                base["falta"] = True
                base["motivo_validacao"] = (
                    f"{n_batidas} batida(s) no dia ({horarios}) - nenhum periodo completo; "
                    "possível falta do dia inteiro"
                )
            else:
                base["motivo_validacao"] = (
                    f"{n_batidas} batida(s) no dia ({horarios}) - esperado 4; "
                    f"possível falta parcial de ~{_hhmm(faltante_min)}"
                )
            return base

    tol_batida = jornada.tolerancia_por_batida_min

    if n_batidas == 4:
        entrada, saida_almoco, retorno_almoco, saida = batidas
        intervalo_min = _minutos_entre(saida_almoco, retorno_almoco)

        entrada_aj = _ajustar_batida_pela_tolerancia(entrada, jornada.entrada, dia, tol_batida)
        saida_almoco_aj = _ajustar_batida_pela_tolerancia(saida_almoco, jornada.saida_almoco, dia, tol_batida)
        retorno_almoco_aj = _ajustar_batida_pela_tolerancia(retorno_almoco, jornada.retorno_almoco, dia, tol_batida)
        saida_aj = _ajustar_batida_pela_tolerancia(saida, jornada.saida, dia, tol_batida)

        trabalhado_min = (
            _minutos_entre(entrada_aj, saida_almoco_aj) + _minutos_entre(retorno_almoco_aj, saida_aj)
        )
        # A mesma tolerancia por marcacao da CLT (Art. 58 §1º) usada acima
        # pra "encostar" entrada/saida no horario programado tambem vale
        # aqui: um intervalo de 56-59 min e' so' a SAIDA_ALMOCO ou o
        # RETORNO_ALMOCO batendo 1-4 min fora do horario, o que a empresa
        # ja' tolera por marcacao - nao e' um intervalo curto de verdade.
        # So' sinaliza pra validacao humana o que sobra fora dessa folga.
        if intervalo_min < 60 - tol_batida:
            base["status"] = "INTERVALO_CURTO"
            base["precisa_validacao_rh"] = True
            base["motivo_validacao"] = f"Intervalo de almoco de apenas {intervalo_min} min"
    else:  # 2 batidas: jornada sem intervalo registrado (ex.: turno corrido)
        entrada, saida = batidas
        entrada_aj = _ajustar_batida_pela_tolerancia(entrada, jornada.entrada, dia, tol_batida)
        saida_aj = _ajustar_batida_pela_tolerancia(saida, jornada.saida, dia, tol_batida)
        trabalhado_min = _minutos_entre(entrada_aj, saida_aj)

    saldo_min = trabalhado_min - carga_min

    if saldo_min > jornada.tolerancia_min:
        base["he_50_min"] = saldo_min
    elif saldo_min < -jornada.tolerancia_min:
        base["atraso_min"] = abs(saldo_min)
        base["status"] = "ATRASO"

    return base


def apurar_periodo(
    batidas_por_dia: pd.DataFrame,
    jornadas: dict[str, Jornada],
    data_inicio: date,
    data_fim: date,
    feriados: Mapping[date, str] | None = None,
) -> pd.DataFrame:
    """
    Apura todos os colaboradores/dias do periodo de fechamento (ex.: 26 a 25,
    ou 26 a 03, conforme a politica da empresa), inclusive os dias em que o
    colaborador nao tem NENHUMA batida (que o parser nem gera linha) -
    esses sao os candidatos mais fortes a falta e precisam entrar na apuracao.

    `feriados` ({data: nome}, ver feriados.py) marca os dias que nao sao de
    trabalho para ninguem: neles a ausencia nao vira falta.
    """
    feriados = feriados or {}
    linhas = []
    dias_periodo = pd.date_range(data_inicio, data_fim, freq="D").date

    batidas_idx = batidas_por_dia.set_index(["pis", "data"])["batidas"].to_dict() if not batidas_por_dia.empty else {}

    for pis, jornada in jornadas.items():
        for dia in dias_periodo:
            batidas = batidas_idx.get((pis, dia), [])
            linhas.append(apurar_dia(pis, dia, batidas, jornada, feriados.get(dia)))

    return pd.DataFrame(linhas)


# ---------------------------------------------------------------------------
# Regras de negocio mensais (dependem da apuracao diaria + decisoes do RH)
# ---------------------------------------------------------------------------

def consolidar_mes(
    apuracao_diaria: pd.DataFrame,
    decisoes_rh: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """
    Junta a apuracao automatica com as decisoes humanas sobre as excecoes
    (vindas do painel de aprovacao) e consolida por colaborador (pis):

      - horas_extra_50 / horas_extra_100 (em horas, 2 casas)
      - dias_falta (dias INTEIROS de falta; faltas parciais ficam fora)
      - dias_falta_justificada / dias_falta_injustificada (dias inteiros)
      - dias_falta_parcial / horas_falta_parcial (dias com falta de meio periodo)
      - atraso_horas e horas_falta (para o cod. 8069 = atraso + falta parcial)
      - perde_copr (bool) - QUALQUER falta/ausencia, inteira ou parcial,
        justificada ou nao
      - dias_desconto_va - apenas faltas inteiras SEM justificativa
      - semanas_perde_dsr - semanas com falta inteira injustificada

    `decisoes_rh` e' um DataFrame opcional com colunas
    [pis, data, classificacao] onde classificacao in
    {"FALTA_JUSTIFICADA", "FALTA_INJUSTIFICADA", "ERRO_RELOGIO_ABONAR"}.
    Linhas de excecao sem decisao do RH sao tratadas, por seguranca, como
    FALTA_INJUSTIFICADA ate' serem revisadas (nunca fecham sozinhas a favor
    do colaborador nem da empresa sem rastro).

    Falta parcial (status FALTA_PARCIAL_CANDIDATA): "justificada" ou
    "injustificada" soma os minutos faltantes em horas_falta (nao vira dia
    de falta) e tira o COPR; "erro de relogio - abonar" nao conta nada. VA e
    DSR continuam por dia inteiro e nao sao afetados por falta parcial.
    """
    df = apuracao_diaria.copy()
    df["data"] = pd.to_datetime(df["data"])
    if "minutos_falta_parcial" not in df.columns:
        df["minutos_falta_parcial"] = 0

    if decisoes_rh is not None and not decisoes_rh.empty:
        decisoes = decisoes_rh.copy()
        decisoes["pis"] = decisoes["pis"].astype(str).str.strip()
        decisoes["data"] = pd.to_datetime(decisoes["data"], dayfirst=True)
        df = df.merge(decisoes[["pis", "data", "classificacao"]], on=["pis", "data"], how="left")
    else:
        df["classificacao"] = pd.NA

    def _classificacao_efetiva(row):
        if row["falta"] or row["status"] in STATUS_FALTAS_A_CLASSIFICAR:
            if pd.notna(row["classificacao"]):
                return row["classificacao"]
            return "FALTA_INJUSTIFICADA"  # default conservador ate' validacao
        return pd.NA

    df["classificacao_efetiva"] = df.apply(_classificacao_efetiva, axis=1)
    contam = df["classificacao_efetiva"].isin(["FALTA_JUSTIFICADA", "FALTA_INJUSTIFICADA"])
    # parcial = falta de meio periodo (horas). Se nenhum periodo ficou completo,
    # apurar_dia ja' marcou falta=True e o dia segue a regra de falta de dia inteiro.
    parcial = (df["status"] == STATUS_FALTA_PARCIAL) & ~df["falta"]
    df["e_falta"] = contam & ~parcial
    df["e_falta_justificada"] = df["e_falta"] & (df["classificacao_efetiva"] == "FALTA_JUSTIFICADA")
    df["e_falta_parcial"] = contam & parcial
    df["min_falta_parcial_efetiva"] = df["minutos_falta_parcial"].where(df["e_falta_parcial"], 0)
    df["semana"] = df["data"].dt.isocalendar().week

    resumo = df.groupby("pis").agg(
        horas_extra_50_min=("he_50_min", "sum"),
        horas_extra_100_min=("he_100_min", "sum"),
        atraso_min=("atraso_min", "sum"),
        falta_parcial_min=("min_falta_parcial_efetiva", "sum"),
        dias_falta_parcial=("e_falta_parcial", "sum"),
        dias_falta=("e_falta", "sum"),
        dias_falta_justificada=("e_falta_justificada", "sum"),
    ).reset_index()

    injustificadas = (
        df[(df["classificacao_efetiva"] == "FALTA_INJUSTIFICADA") & ~parcial]
        .groupby("pis")
        .agg(dias_falta_injustificada=("data", "count"), semanas_com_falta_injust=("semana", "nunique"))
        .reset_index()
    )
    resumo = resumo.merge(injustificadas, on="pis", how="left")
    resumo[["dias_falta_injustificada", "semanas_com_falta_injust"]] = resumo[
        ["dias_falta_injustificada", "semanas_com_falta_injust"]
    ].fillna(0).astype(int)

    resumo["horas_extra_50"] = (resumo["horas_extra_50_min"] / 60).round(2)
    resumo["horas_extra_100"] = (resumo["horas_extra_100_min"] / 60).round(2)
    resumo["atraso_horas"] = (resumo["atraso_min"] / 60).round(2)
    resumo["horas_falta_parcial"] = (resumo["falta_parcial_min"] / 60).round(2)
    # cod. 8069: horas de atraso (alem da tolerancia) + horas de falta parcial
    resumo["horas_falta"] = ((resumo["atraso_min"] + resumo["falta_parcial_min"]) / 60).round(2)
    for coluna in ("dias_falta", "dias_falta_justificada", "dias_falta_parcial"):
        resumo[coluna] = resumo[coluna].astype(int)

    # Regra COPR: qualquer falta (inteira ou parcial, justificada ou nao) => perde o premio
    resumo["perde_copr"] = (resumo["dias_falta"] > 0) | (resumo["dias_falta_parcial"] > 0)

    # Regra VA: apenas faltas inteiras SEM justificativa geram desconto proporcional
    resumo["dias_desconto_va"] = resumo["dias_falta_injustificada"]

    # DSR: perde o DSR da semana em que houve falta injustificada.
    # Aqui reportamos o numero de semanas afetadas; a conversao para R$/dia
    # de DSR e' feita na hora de montar a Relacao de Valores (cod. 8794),
    # pois depende do valor do dia de cada colaborador (fora do escopo do
    # ponto em si).
    resumo["semanas_perde_dsr"] = resumo["semanas_com_falta_injust"]

    return resumo[[
        "pis", "horas_extra_50", "horas_extra_100", "atraso_min", "atraso_horas",
        "horas_falta", "horas_falta_parcial", "dias_falta_parcial",
        "dias_falta", "dias_falta_justificada", "dias_falta_injustificada", "dias_desconto_va",
        "perde_copr", "semanas_perde_dsr",
    ]]


def gerar_fila_validacao_rh(apuracao_diaria: pd.DataFrame) -> pd.DataFrame:
    """
    Retorna somente as linhas que precisam de decisao humana (as excecoes).
    A coluna `data` fica como data de verdade (nao string) de proposito -
    quem exporta pra CSV (`pipeline.exportar_fila_validacao_rh`) e' quem
    decide a ordenacao e o formato de exibicao (dd/mm/aaaa para o RH).
    """
    fila = apuracao_diaria[apuracao_diaria["precisa_validacao_rh"]].copy()
    fila["data"] = pd.to_datetime(fila["data"])
    colunas = ["pis", "data", "status", "motivo_validacao", "n_batidas", "minutos_falta_parcial"]
    colunas = [c for c in colunas if c in fila.columns]
    return fila[colunas].sort_values(["data", "pis"]).reset_index(drop=True)
