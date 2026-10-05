"""
banco_horas.py
===============
Trata as DUAS excecoes de banco de horas citadas no processo:

  - "Sidney": banco de horas fica registrado no proprio Secullum. O
    Secullum tem um relatorio de banco de horas exportavel (Excel/CSV);
    aqui so' normalizamos esse export para o mesmo formato usado no
    restante do pipeline.

  - "Adriano": banco de horas controlado em planilha Excel paralela,
    semanal, com saldo acumulado do mes anterior + do mes atual.

O ponto critico pedido no briefing e' "sem corromper os dados dos demais":
por isso a funcao final faz um LEFT MERGE pelo `pis`/matricula contra a
base de todos os colaboradores - quem nao aparece em nenhuma das duas
fontes simplesmente fica com saldo 0 / NaN, sem afetar o calculo de mais
ninguem. Cada fonte tambem e' validada isoladamente antes do merge, para
que um erro na planilha do Adriano nunca derrube o processamento do
banco de horas do Sidney (ou vice-versa).
"""

from __future__ import annotations

import pandas as pd

from validacao import (
    verificar_colunas, normalizar_identificador_colaborador, converter_horas_decimais, ArquivoInvalidoError,
)
from leitura_arquivos import ler_arquivo_generico


def _coluna_em_horas(serie: pd.Series, nome_arquivo: str, caminho: str, coluna: str) -> pd.Series:
    """Converte a coluna inteira para horas decimais; um valor invalido aponta a linha da planilha."""
    return pd.Series(
        [converter_horas_decimais(v, f"{nome_arquivo} ('{caminho}'), linha {i + 2}", coluna) for i, v in serie.items()],
        index=serie.index, dtype=float,
    )


def ler_banco_horas_secullum(caminho_export: str, cadastro: pd.DataFrame | None = None) -> pd.DataFrame:
    """
    Le o relatorio de banco de horas exportado do proprio Secullum
    (cobre o caso do "Sidney"). Espera colunas Matricula/PIS e Saldo
    (em horas, podendo ser negativo). O saldo pode vir em decimal (virgula
    ou ponto), em "hh:mm" (inclusive "-05:10") ou em celula de hora do Excel
    - tudo vira horas decimais com 2 casas. Ajuste os nomes de coluna abaixo
    para bater com o export real do seu Secullum.

    Passe `cadastro` sempre que possível: o Secullum costuma identificar
    o colaborador pela matrícula interna, não pelo PIS — veja
    `validacao.normalizar_identificador_colaborador`.
    """
    df = ler_arquivo_generico(caminho_export, "Banco de Horas do Secullum")

    coluna_saldo = next((c for c in df.columns if "saldo" in c), None)
    coluna_pis = next((c for c in df.columns if "pis" in c or "matricula" in c), None)
    if coluna_saldo is None or coluna_pis is None:
        raise ArquivoInvalidoError(
            f"O Banco de Horas do Secullum ('{caminho_export}') está sem uma coluna de "
            f"PIS/matrícula e/ou de saldo.\nColunas encontradas: {list(df.columns)}."
        )
    out = df[[coluna_pis, coluna_saldo]].rename(columns={coluna_pis: "pis", coluna_saldo: "saldo_banco_horas"})
    out["pis"] = out["pis"].astype(str).str.strip()
    out["saldo_banco_horas"] = _coluna_em_horas(
        out["saldo_banco_horas"], "Banco de Horas do Secullum", caminho_export, coluna_saldo
    )
    if cadastro is not None:
        out = normalizar_identificador_colaborador(out, "pis", cadastro, "Banco de Horas do Secullum", caminho_export)
    out["fonte"] = "secullum_sidney"
    return out


def ler_banco_horas_adriano(
    caminho_planilha: str, pis_adriano: str, sheet_name: str | int = 0, cadastro: pd.DataFrame | None = None
) -> pd.DataFrame:
    """
    Le a planilha semanal paralela do Adriano e calcula o saldo do periodo.

    Formato esperado da planilha (ajuste os nomes de coluna conforme a
    planilha real):
        semana | horas_extras | horas_debito | saldo_anterior

    `pis_adriano` aceita tanto o PIS quanto a matrícula dele — se
    `cadastro` for informado, o valor é resolvido para o PIS canônico
    (mesma lógica usada para os outros arquivos auxiliares).

    Retorna uma unica linha (pis=pis_adriano, saldo_banco_horas=<acumulado>)
    pronta para entrar no mesmo merge do restante da equipe.
    """
    if not str(pis_adriano or "").strip():
        raise ArquivoInvalidoError(
            "O PIS/matrícula do Adriano não foi informado - sem ele não dá pra saber a "
            "quem atribuir o saldo de banco de horas dessa planilha."
        )

    df = ler_arquivo_generico(caminho_planilha, "Planilha semanal do Adriano", sheet_name=sheet_name)
    verificar_colunas(df, ["horas_extras", "horas_debito"], "Planilha semanal do Adriano", caminho_planilha)

    nome = "Planilha semanal do Adriano"
    saldo_anterior = 0.0
    if "saldo_anterior" in df.columns and not df["saldo_anterior"].dropna().empty:
        # saldo acumulado do mes anterior fica, por convencao, na 1a linha
        primeira = df["saldo_anterior"].dropna().index[0]
        saldo_anterior = converter_horas_decimais(
            df.at[primeira, "saldo_anterior"], f"{nome} ('{caminho_planilha}'), linha {primeira + 2}", "saldo_anterior"
        )

    horas_extras = _coluna_em_horas(df["horas_extras"], nome, caminho_planilha, "horas_extras")
    horas_debito = _coluna_em_horas(df["horas_debito"], nome, caminho_planilha, "horas_debito")
    saldo_total = round(saldo_anterior + float((horas_extras - horas_debito).sum()), 2)

    resultado = pd.DataFrame([{
        "pis": str(pis_adriano).strip(),
        "saldo_banco_horas": saldo_total,
        "fonte": "planilha_adriano",
    }])
    if cadastro is not None:
        resultado = normalizar_identificador_colaborador(
            resultado, "pis", cadastro, "Planilha semanal do Adriano (PIS/matrícula informado)", caminho_planilha
        )
    return resultado


def consolidar_banco_horas(
    todos_colaboradores_pis: list[str],
    df_secullum: pd.DataFrame | None,
    df_adriano: pd.DataFrame | None,
) -> pd.DataFrame:
    """
    Junta as duas fontes de excecao numa unica tabela cod. 0999, com uma
    linha por colaborador (0 para quem nao tem banco de horas registrado
    em nenhuma das duas fontes).
    """
    partes = [p for p in (df_secullum, df_adriano) if p is not None and not p.empty]
    base = pd.DataFrame({"pis": [str(p).strip() for p in todos_colaboradores_pis]})

    if not partes:
        base["saldo_banco_horas"] = 0.0
        return base

    empilhado = pd.concat(partes, ignore_index=True)

    duplicados = empilhado["pis"].duplicated(keep=False)
    if duplicados.any():
        raise ArquivoInvalidoError(
            "O mesmo PIS/matrícula aparece em mais de uma fonte de banco de horas "
            "(Secullum e planilha do Adriano) - verifique conflito para: "
            f"{empilhado.loc[duplicados, 'pis'].unique().tolist()}"
        )

    resultado = base.merge(empilhado[["pis", "saldo_banco_horas"]], on="pis", how="left")
    resultado["saldo_banco_horas"] = resultado["saldo_banco_horas"].fillna(0.0)
    return resultado
