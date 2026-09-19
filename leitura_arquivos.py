"""
leitura_arquivos.py
====================
Leitura universal e resiliente de arquivos de entrada (Excel ou CSV),
usada por todos os modulos que carregam dados (`pipeline.py`,
`banco_horas.py`). A ideia central: a pessoa do RH pode exportar um CSV
do jeito que o programa de origem mandar - separador virgula, ponto e
virgula ou tab, em UTF-8, UTF-8 com BOM ou Latin-1. Em vez de cada modulo
adivinhar isso sozinho (e falhar de um jeito diferente cada vez), toda
leitura passa por `ler_arquivo_generico()`, que tenta as combinacoes mais
comuns e so' desiste com uma mensagem amigavel, em portugues, se nenhuma
funcionar.
"""

from __future__ import annotations

import csv
from pathlib import Path

import pandas as pd

from validacao import ArquivoInvalidoError

_ENCODINGS_TENTATIVAS = ["utf-8-sig", "utf-8", "latin-1"]
_DELIMITADORES_TENTATIVAS = [",", ";", "\t"]


def _detectar_delimitador(caminho: Path, encoding: str) -> str | None:
    """Tenta usar o csv.Sniffer do Python; se nao conseguir, devolve None -
    quem chamou entao tenta os delimitadores comuns na forca bruta."""
    try:
        with open(caminho, "r", encoding=encoding, errors="strict") as f:
            amostra = f.read(4096)
        return csv.Sniffer().sniff(amostra, delimiters=",;\t").delimiter
    except Exception:
        return None


def _ler_csv_resiliente(caminho: Path) -> pd.DataFrame:
    """
    Tenta ler um CSV cruzando codificacoes (utf-8-sig, utf-8, latin-1) com
    delimitadores (virgula, ponto-e-virgula, tab). Usa o Sniffer do Python
    como primeiro palpite e cai para forca bruta se ele nao acertar.
    Considera "sucesso" a primeira combinacao que produz mais de uma
    coluna (uma unica coluna geralmente significa que o delimitador
    escolhido esta errado, mesmo que a leitura nao tenha dado erro).
    """
    ultimo_erro: Exception | None = None
    melhor_tentativa: pd.DataFrame | None = None

    for encoding in _ENCODINGS_TENTATIVAS:
        delimitador_detectado = _detectar_delimitador(caminho, encoding)
        candidatos = [delimitador_detectado] if delimitador_detectado else []
        candidatos += [d for d in _DELIMITADORES_TENTATIVAS if d not in candidatos]

        for delimitador in candidatos:
            try:
                df = pd.read_csv(caminho, encoding=encoding, sep=delimitador, engine="python")
            except Exception as e:
                ultimo_erro = e
                continue

            if df.shape[1] > 1:
                return df  # varias colunas reconhecidas -> delimitador certo
            if melhor_tentativa is None:
                melhor_tentativa = df  # guarda como fallback (arquivo pode ter 1 coluna mesmo)

    if melhor_tentativa is not None:
        return melhor_tentativa

    raise ArquivoInvalidoError(
        f"Não consegui ler o arquivo CSV '{caminho}' com nenhuma combinação comum de "
        "codificação/separador (testei UTF-8, UTF-8 com BOM e Latin-1; vírgula, ponto "
        "e vírgula e tabulação). O arquivo pode estar corrompido ou num formato incomum.\n"
        f"Detalhe técnico: {ultimo_erro}"
    )


def ler_arquivo_generico(caminho: str, nome_amigavel_arquivo: str, sheet_name: str | int = 0) -> pd.DataFrame:
    """
    Le um arquivo de entrada (.xlsx, .xls ou .csv) e devolve um DataFrame
    com o cabecalho ja' normalizado (nomes de coluna em minusculo, sem
    espaco nas pontas). Levanta `ArquivoInvalidoError`, com mensagem clara
    em portugues, se o arquivo nao existir, tiver extensao nao suportada,
    estiver corrompido/aberto em outro programa, ou vier vazio.

    `sheet_name` so' se aplica a arquivos Excel (.xlsx/.xls); e' ignorado
    para CSV.
    """
    caminho_path = Path(caminho)
    if not caminho_path.exists():
        raise ArquivoInvalidoError(
            f"O arquivo de {nome_amigavel_arquivo} não foi encontrado em '{caminho}'. "
            "Verifique se ele não foi movido, renomeado ou apagado."
        )

    sufixo = caminho_path.suffix.lower()
    try:
        if sufixo in (".xlsx", ".xls"):
            df = pd.read_excel(caminho_path, sheet_name=sheet_name)
        elif sufixo == ".csv":
            df = _ler_csv_resiliente(caminho_path)
        else:
            raise ArquivoInvalidoError(
                f"O arquivo de {nome_amigavel_arquivo} ('{caminho}') tem uma extensão não "
                f"suportada ('{sufixo or 'sem extensão'}'). Use .xlsx, .xls ou .csv."
            )
    except ArquivoInvalidoError:
        raise
    except Exception as e:
        raise ArquivoInvalidoError(
            f"Não consegui abrir o arquivo de {nome_amigavel_arquivo} ('{caminho}'). Ele pode "
            "estar corrompido, aberto em outro programa (feche o Excel, por exemplo) ou não "
            f"ser realmente um arquivo {sufixo or 'suportado'}.\nDetalhe técnico: {e}"
        ) from e

    if df.shape[1] == 0:
        raise ArquivoInvalidoError(
            f"O arquivo de {nome_amigavel_arquivo} ('{caminho}') foi aberto, mas não tem "
            "nenhuma coluna reconhecível."
        )
    if df.empty:
        raise ArquivoInvalidoError(
            f"O arquivo de {nome_amigavel_arquivo} ('{caminho}') foi aberto, mas está vazio "
            "(sem nenhuma linha de dado)."
        )

    df.columns = [str(c).strip().lower() for c in df.columns]
    return df
