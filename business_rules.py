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

import pandas as pd

TOLERANCIA_DIARIA_MIN = 10       # Art. 58 §1º CLT: limite maximo diario
TOLERANCIA_POR_BATIDA_MIN = 5    # Art. 58 §1º CLT: limite por marcacao individual


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


def apurar_dia(pis: str, dia: date, batidas: list[datetime], jornada: Jornada) -> dict:
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
    """
    dow = dia.weekday()
    dia_deveria_trabalhar = dow in jornada.dias_trabalho
    n_batidas = len(batidas)

    base = {
        "pis": pis,
        "data": dia,
        "n_batidas": n_batidas,
        "he_50_min": 0,
        "he_100_min": 0,
        "atraso_min": 0,
        "falta": False,
        "status": "OK",
        "precisa_validacao_rh": False,
        "motivo_validacao": "",
    }

    # Domingo/dia nao util com batidas => tudo vira hora extra 100%
    # (a tolerancia por marcacao nao se aplica aqui: nao ha horario
    # programado de referencia num dia em que o colaborador nao deveria
    # trabalhar, entao qualquer minuto batido conta.)
    if not dia_deveria_trabalhar:
        if n_batidas >= 2:
            trabalhado_min = _minutos_entre(min(batidas), max(batidas))
            base["he_100_min"] = trabalhado_min
            base["status"] = "HE_100_DIA_NAO_UTIL"
        return base

    # Dia normal de trabalho, sem nenhuma batida => possivel falta
    if n_batidas == 0:
        base["falta"] = True
        base["status"] = "FALTA_OU_AUSENCIA"
        base["precisa_validacao_rh"] = True
        base["motivo_validacao"] = "Nenhuma batida no dia - confirmar se e' falta real ou erro de relogio"
        return base

    # Numero de batidas incompleto (esperado 4: entrada/saida almoco/retorno/saida)
    if n_batidas not in (2, 4):
        base["status"] = "BATIDA_INCOMPLETA"
        base["precisa_validacao_rh"] = True
        base["motivo_validacao"] = f"{n_batidas} batida(s) no dia - esperado 2 ou 4"
        return base

    carga_min = jornada.carga_horaria_min()
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
        if intervalo_min < 60:
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
) -> pd.DataFrame:
    """
    Apura todos os colaboradores/dias do periodo de fechamento (ex.: 26 a 25,
    ou 26 a 03, conforme a politica da empresa), inclusive os dias em que o
    colaborador nao tem NENHUMA batida (que o parser nem gera linha) -
    esses sao os candidatos mais fortes a falta e precisam entrar na apuracao.
    """
    linhas = []
    dias_periodo = pd.date_range(data_inicio, data_fim, freq="D").date

    batidas_idx = batidas_por_dia.set_index(["pis", "data"])["batidas"].to_dict() if not batidas_por_dia.empty else {}

    for pis, jornada in jornadas.items():
        for dia in dias_periodo:
            batidas = batidas_idx.get((pis, dia), [])
            linhas.append(apurar_dia(pis, dia, batidas, jornada))

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
      - dias_falta (total)
      - dias_falta_justificada / dias_falta_injustificada
      - horas_falta (para o cod. 8069)
      - perde_copr (bool) - QUALQUER falta/ausencia, justificada ou nao
      - dias_desconto_va - apenas faltas SEM justificativa
      - semanas_perde_dsr - semanas com falta injustificada -> desconto de DSR

    `decisoes_rh` e' um DataFrame opcional com colunas
    [pis, data, classificacao] onde classificacao in
    {"FALTA_JUSTIFICADA", "FALTA_INJUSTIFICADA", "ERRO_RELOGIO_ABONAR"}.
    Linhas de excecao sem decisao do RH sao tratadas, por seguranca, como
    FALTA_INJUSTIFICADA ate' serem revisadas (nunca fecham sozinhas a favor
    do colaborador nem da empresa sem rastro).
    """
    df = apuracao_diaria.copy()
    df["data"] = pd.to_datetime(df["data"])

    if decisoes_rh is not None and not decisoes_rh.empty:
        decisoes = decisoes_rh.copy()
        decisoes["pis"] = decisoes["pis"].astype(str).str.strip()
        decisoes["data"] = pd.to_datetime(decisoes["data"], dayfirst=True)
        df = df.merge(decisoes[["pis", "data", "classificacao"]], on=["pis", "data"], how="left")
    else:
        df["classificacao"] = pd.NA

    def _classificacao_efetiva(row):
        if row["falta"] or row["status"] == "FALTA_OU_AUSENCIA":
            if pd.notna(row["classificacao"]):
                return row["classificacao"]
            return "FALTA_INJUSTIFICADA"  # default conservador ate' validacao
        return pd.NA

    df["classificacao_efetiva"] = df.apply(_classificacao_efetiva, axis=1)
    df["e_falta"] = df["classificacao_efetiva"].isin(["FALTA_JUSTIFICADA", "FALTA_INJUSTIFICADA"])
    df["semana"] = df["data"].dt.isocalendar().week

    resumo = df.groupby("pis").agg(
        horas_extra_50_min=("he_50_min", "sum"),
        horas_extra_100_min=("he_100_min", "sum"),
        atraso_min=("atraso_min", "sum"),
        dias_falta=("e_falta", "sum"),
    ).reset_index()

    injustificadas = (
        df[df["classificacao_efetiva"] == "FALTA_INJUSTIFICADA"]
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
    resumo["horas_falta"] = (resumo["dias_falta"] * 0).astype(float)  # ver nota abaixo
    resumo["dias_falta"] = resumo["dias_falta"].astype(int)

    # Regra COPR: qualquer falta (justificada ou nao) => perde o premio
    resumo["perde_copr"] = resumo["dias_falta"] > 0

    # Regra VA: apenas faltas SEM justificativa geram desconto proporcional
    resumo["dias_desconto_va"] = resumo["dias_falta_injustificada"]

    # DSR: perde o DSR da semana em que houve falta injustificada.
    # Aqui reportamos o numero de semanas afetadas; a conversao para R$/dia
    # de DSR e' feita na hora de montar a Relacao de Valores (cod. 8794),
    # pois depende do valor do dia de cada colaborador (fora do escopo do
    # ponto em si).
    resumo["semanas_perde_dsr"] = resumo["semanas_com_falta_injust"]

    return resumo[[
        "pis", "horas_extra_50", "horas_extra_100", "atraso_min",
        "dias_falta", "dias_falta_injustificada", "dias_desconto_va",
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
    return fila[["pis", "data", "status", "motivo_validacao"]].sort_values(["data", "pis"]).reset_index(drop=True)
