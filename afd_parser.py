"""
afd_parser.py
=============
Le o arquivo AFD (Arquivo de Fonte de Dados) exportado pelo Secullum RH e
devolve um DataFrame com uma linha por batida de ponto:

    pis | data_hora | tipo_registro

O layout do AFD segue a Portaria MTP 671/2021 (que substituiu a Portaria
1510/2011). Os REP-P/REP-C do Secullum normalmente seguem esse padrao, mas
o layout EXATO (posicoes de cada campo) pode variar por versao do
equipamento/sistema. Antes de rodar em producao:

  1. Abra o .txt exportado pelo Secullum em um editor de texto.
  2. Confira a posicao inicial do campo TIPO DE REGISTRO (deve ser '3' nas
     linhas de marcacao de ponto) e ajuste as constantes abaixo se
     necessario (elas ficam centralizadas aqui de proposito).
  3. Rode `python afd_parser.py caminho/para/arquivo.txt --validar` para ver
     uma amostra parseada antes de usar em lote.

Este modulo eh tolerante a linhas mal formadas: registros que nao batem com
o layout esperado sao colocados em `erros_leitura` (lista retornada junto
com o DataFrame) em vez de derrubar o processamento inteiro - importante
porque um AFD de producao pode ter centenas de milhares de linhas e um
unico caractere fora do lugar nao pode travar o fechamento da folha.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# CONFIGURACAO DE LAYOUT - AJUSTE AQUI SE O SEU AFD DIVERGIR
# ---------------------------------------------------------------------------
# Posicoes 0-indexed, seguindo o padrao mais comum de REP-P (Portaria 671/2021)
# para registros tipo 3 (marcacao de ponto efetuada por empregado):
#   NSR (9) + TIPO (1) + PIS (12) + DATA/HORA (12, AAAAMMDDHHMM)a
LAYOUT_TIPO3 = {
    "nsr": (0, 9),
    "tipo_registro": (9, 10),
    "pis": (10, 22),
    "data_hora": (22, 34),  # AAAAMMDDHHMM
}
TIPO_MARCACAO_PONTO = "3"


@dataclass
class ResultadoParseAFD:
    batidas: pd.DataFrame
    erros_leitura: list = field(default_factory=list)
    total_linhas: int = 0
    total_batidas: int = 0


def _parse_linha(linha: str, numero_linha: int) -> dict | None:
    """Tenta interpretar uma linha do AFD como um registro tipo 3 (batida)."""
    if len(linha) < LAYOUT_TIPO3["data_hora"][1]:
        return None

    tipo = linha[slice(*LAYOUT_TIPO3["tipo_registro"])]
    if tipo != TIPO_MARCACAO_PONTO:
        return None  # cabecalho, trailer, ajuste manual, etc. - ignorado aqui

    pis = linha[slice(*LAYOUT_TIPO3["pis"])].strip()
    data_hora_str = linha[slice(*LAYOUT_TIPO3["data_hora"])].strip()

    try:
        data_hora = datetime.strptime(data_hora_str, "%Y%m%d%H%M")
    except ValueError:
        return {"__erro__": f"linha {numero_linha}: data/hora invalida ('{data_hora_str}')"}

    if not pis:
        return {"__erro__": f"linha {numero_linha}: PIS vazio"}

    return {"pis": pis, "data_hora": data_hora, "tipo_registro": tipo}


def parse_afd(caminho_arquivo: str | Path, encoding: str = "latin-1") -> ResultadoParseAFD:
    """Le um arquivo AFD e retorna todas as batidas de ponto (tipo 3)."""
    caminho_arquivo = Path(caminho_arquivo)
    registros = []
    erros = []
    total_linhas = 0

    with open(caminho_arquivo, "r", encoding=encoding) as f:
        for i, linha in enumerate(f, start=1):
            total_linhas += 1
            linha = linha.rstrip("\n\r")
            if not linha:
                continue
            resultado = _parse_linha(linha, i)
            if resultado is None:
                continue
            if "__erro__" in resultado:
                erros.append(resultado["__erro__"])
                continue
            registros.append(resultado)

    df = pd.DataFrame(registros, columns=["pis", "data_hora", "tipo_registro"])
    if not df.empty:
        df = df.sort_values(["pis", "data_hora"]).reset_index(drop=True)

    return ResultadoParseAFD(
        batidas=df,
        erros_leitura=erros,
        total_linhas=total_linhas,
        total_batidas=len(df),
    )


def agrupar_batidas_por_dia(df_batidas: pd.DataFrame) -> pd.DataFrame:
    """
    Transforma a lista de batidas (uma por linha) em uma linha por
    (pis, data), com as batidas do dia ordenadas em colunas
    b1..bN (entrada, saida almoco, retorno almoco, saida, extras...).

    Isso alimenta o business_rules.py, que decide o que cada par de
    batidas significa (jornada normal, hora extra, etc.).
    """
    if df_batidas.empty:
        return pd.DataFrame(columns=["pis", "data", "batidas"])

    df = df_batidas.copy()
    df["data"] = df["data_hora"].dt.date

    agrupado = (
        df.groupby(["pis", "data"])["data_hora"]
        .apply(lambda s: sorted(s.tolist()))
        .reset_index()
        .rename(columns={"data_hora": "batidas"})
    )
    return agrupado


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Parser de arquivo AFD (Secullum)")
    ap.add_argument("arquivo", help="Caminho do arquivo .txt AFD")
    ap.add_argument("--validar", action="store_true", help="Mostra amostra e estatisticas")
    args = ap.parse_args()

    resultado = parse_afd(args.arquivo)
    print(f"Linhas lidas: {resultado.total_linhas}")
    print(f"Batidas de ponto (tipo 3) encontradas: {resultado.total_batidas}")
    print(f"Linhas com erro: {len(resultado.erros_leitura)}")
    if resultado.erros_leitura:
        print("Primeiros erros:")
        for e in resultado.erros_leitura[:10]:
            print(f"  - {e}")
    if args.validar and not resultado.batidas.empty:
        print("\nAmostra de batidas parseadas:")
        print(resultado.batidas.head(10).to_string(index=False))
