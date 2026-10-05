"""
feriados.py
===========
Calendario de feriados usado na apuracao do ponto. Um dia que e' feriado
deixa de ser dia de trabalho: nao gera falta (nem perda de COPR/VA/DSR) e,
se houver batida, o tempo trabalhado vira hora extra 100% - a mesma regra
que ja' valia para domingo.

Os feriados ficam em `feriados.csv` (gerado por `gerar_feriados.py`), que e'
lido SEM internet. Colunas:

    data               AAAA-MM-DD (dd/mm/aaaa tambem e' aceito, caso o Excel regrave o arquivo)
    nome               nome do feriado
    escopo             nacional | estadual | municipal | facultativo
    local              BR | PR | Cianorte-PR (informativo)
    vale_para_empresa  sim | nao | (vazio = usa o padrao do escopo, abaixo)
    situacao           lei | confirmado | previsto   (previsto = confirmar o decreto do ano)
    fonte, observacao  informativos

Padrao quando `vale_para_empresa` esta vazio: nacional e municipal valem;
estadual (19/12 - so' servidores publicos) e facultativo (Carnaval, Corpus
Christi, vesperas) NAO valem, a menos que a convencao coletiva ou a empresa
os conceda - nesse caso marque "sim" na linha, ou use `incluir_facultativos=True`
para todos os pontos facultativos de uma vez. Uma linha com "sim"/"nao"
explicito sempre prevalece sobre o padrao.
"""

from __future__ import annotations

import unicodedata
from datetime import date
from pathlib import Path
from typing import Mapping

import pandas as pd

from leitura_arquivos import ler_arquivo_generico
from validacao import ArquivoInvalidoError, verificar_colunas

CAMINHO_PADRAO = Path(__file__).with_name("feriados.csv")
NOME_AMIGAVEL = "Calendário de Feriados"

COLUNAS_OBRIGATORIAS = ["data", "nome", "escopo"]
ESCOPOS = ("nacional", "estadual", "municipal", "facultativo")
# Padrao de cada escopo quando a coluna `vale_para_empresa` esta vazia.
VALE_POR_PADRAO = {"nacional": True, "municipal": True, "estadual": False, "facultativo": False}

_SIM = {"sim", "s", "1", "true", "verdadeiro", "x"}
_NAO = {"nao", "n", "0", "false", "falso"}
_DIAS_SEMANA = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"]


def converter_datas(serie: pd.Series) -> pd.Series:
    """
    AAAA-MM-DD (padrao do arquivo) ou dd/mm/aaaa (se o Excel regravou). Cada formato
    e' lido de forma ESTRITA: `dayfirst=True` num parser flexivel inverte dia e mes das
    datas ISO ('2026-05-01' viraria 5 de janeiro). Valores invalidos viram NaT.
    """
    texto = serie.astype(str).str.strip().str.slice(0, 10)  # tolera '2026-09-07 00:00:00' vindo do Excel
    iso = pd.to_datetime(texto, format="%Y-%m-%d", errors="coerce")
    return iso.fillna(pd.to_datetime(texto, format="%d/%m/%Y", errors="coerce"))


def _sem_acento_minusculo(valor) -> str:
    texto = unicodedata.normalize("NFKD", str(valor))
    return "".join(c for c in texto if not unicodedata.combining(c)).strip().lower()


def _ler_vale(valor, linha_arquivo: int, caminho: Path) -> bool | None:
    """True/False quando preenchido; None quando vazio (vale o padrao do escopo)."""
    if pd.isna(valor) or str(valor).strip() == "":
        return None
    texto = _sem_acento_minusculo(valor)
    if texto in _SIM:
        return True
    if texto in _NAO:
        return False
    raise ArquivoInvalidoError(
        f"{NOME_AMIGAVEL} ('{caminho}'), linha {linha_arquivo}: o valor '{valor}' da coluna "
        "'vale_para_empresa' não é válido. Use 'sim', 'nao' ou deixe vazio."
    )


def carregar_calendario(caminho: str | Path | None = None, incluir_facultativos: bool = False) -> pd.DataFrame:
    """
    Le e valida o calendario. Devolve um DataFrame com `data` (datetime64),
    `nome`, `escopo`, `situacao` e a coluna `vale` (bool) ja' resolvida:
    valor explicito da linha > padrao do escopo (facultativo segue
    `incluir_facultativos`). Levanta `ArquivoInvalidoError` com a linha do problema.
    """
    caminho = Path(caminho) if caminho else CAMINHO_PADRAO
    df = ler_arquivo_generico(str(caminho), NOME_AMIGAVEL)
    verificar_colunas(df, COLUNAS_OBRIGATORIAS, NOME_AMIGAVEL, str(caminho))

    datas = converter_datas(df["data"])
    for posicao in datas[datas.isna()].index:
        raise ArquivoInvalidoError(
            f"{NOME_AMIGAVEL} ('{caminho}'), linha {posicao + 2}: a data '{df.at[posicao, 'data']}' "
            "não é válida. Use AAAA-MM-DD (ou dd/mm/aaaa)."
        )

    escopos = df["escopo"].map(_sem_acento_minusculo)
    for posicao in escopos[~escopos.isin(ESCOPOS)].index:
        raise ArquivoInvalidoError(
            f"{NOME_AMIGAVEL} ('{caminho}'), linha {posicao + 2}: o escopo '{df.at[posicao, 'escopo']}' "
            f"não é válido. Use um destes: {', '.join(ESCOPOS)}."
        )

    explicito = (
        df["vale_para_empresa"] if "vale_para_empresa" in df.columns else pd.Series([None] * len(df), index=df.index)
    )
    vale = []
    for posicao, valor in explicito.items():
        decisao = _ler_vale(valor, posicao + 2, caminho)
        if decisao is None:
            decisao = incluir_facultativos if escopos[posicao] == "facultativo" else VALE_POR_PADRAO[escopos[posicao]]
        vale.append(decisao)

    return pd.DataFrame({
        "data": datas,
        "nome": df["nome"].astype(str).str.strip(),
        "escopo": escopos,
        "situacao": df["situacao"].map(_sem_acento_minusculo) if "situacao" in df.columns else "",
        "vale": vale,
    })


def carregar_feriados(caminho: str | Path | None = None, incluir_facultativos: bool = False) -> dict[date, str]:
    """{data: nome} dos dias que contam como feriado para a empresa (so' linhas com `vale`)."""
    calendario = carregar_calendario(caminho, incluir_facultativos)
    feriados: dict[date, list[str]] = {}
    for data, nome in calendario.loc[calendario["vale"], ["data", "nome"]].itertuples(index=False):
        nomes = feriados.setdefault(data.date(), [])
        if nome not in nomes:
            nomes.append(nome)
    return {data: " / ".join(nomes) for data, nomes in sorted(feriados.items())}


def avisos_de_cobertura(
    data_inicio: date, data_fim: date, caminho: str | Path | None = None, incluir_facultativos: bool = False,
) -> list[str]:
    """
    Avisos para mostrar a quem fecha a folha: anos do periodo que o calendario
    nao cobre (feriado nenhum seria considerado) e feriados contados so' por
    PREVISAO (data fixa, sem o decreto do ano confirmado).
    """
    calendario = carregar_calendario(caminho, incluir_facultativos)
    avisos = []
    anos_no_arquivo = set(calendario["data"].dt.year)
    for ano in range(data_inicio.year, data_fim.year + 1):
        if ano not in anos_no_arquivo:
            avisos.append(
                f"O calendário de feriados não tem o ano {ano}: nenhum feriado de {ano} será considerado. "
                "Rode `python gerar_feriados.py` para incluir o ano."
            )

    no_periodo = calendario[
        calendario["vale"]
        & (calendario["situacao"] == "previsto")
        & (calendario["data"].dt.date >= data_inicio)
        & (calendario["data"].dt.date <= data_fim)
    ]
    for data, nome in no_periodo[["data", "nome"]].itertuples(index=False):
        avisos.append(
            f"Feriado considerado por previsão (confirme o decreto do ano): {data:%d/%m/%Y} - {nome}."
        )
    return avisos


def descrever_feriados(feriados: Mapping[date, str], data_inicio: date, data_fim: date) -> list[str]:
    """Linhas legiveis ('07/09 (seg) - Independência do Brasil') dos feriados que caem no periodo."""
    return [
        f"{dia:%d/%m} ({_DIAS_SEMANA[dia.weekday()]}) - {nome}"
        for dia, nome in sorted(feriados.items())
        if data_inicio <= dia <= data_fim
    ]
