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

from validacao import verificar_colunas, ArquivoInvalidoError
from leitura_arquivos import ler_arquivo_generico


def ler_banco_horas_secullum(caminho_export: str) -> pd.DataFrame:
    """
    Le o relatorio de banco de horas exportado do proprio Secullum
    (cobre o caso do "Sidney"). Espera colunas Matricula/PIS e Saldo
    (em horas, podendo ser negativo). Ajuste os nomes de coluna abaixo
    para bater com o export real do seu Secullum.
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
    out["fonte"] = "secullum_sidney"
    return out


def ler_banco_horas_adriano(caminho_planilha: str, pis_adriano: str, sheet_name: str | int = 0) -> pd.DataFrame:
    """
    Le a planilha semanal paralela do Adriano e calcula o saldo do periodo.

    Formato esperado da planilha (ajuste os nomes de coluna conforme a
    planilha real):
        semana | horas_extras | horas_debito | saldo_anterior

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

    saldo_anterior = 0.0
    if "saldo_anterior" in df.columns and not df["saldo_anterior"].dropna().empty:
        # saldo acumulado do mes anterior fica, por convencao, na 1a linha
        saldo_anterior = float(df["saldo_anterior"].dropna().iloc[0])

    saldo_periodo = (df["horas_extras"].fillna(0) - df["horas_debito"].fillna(0)).sum()
    saldo_total = round(saldo_anterior + float(saldo_periodo), 2)

    return pd.DataFrame([{
        "pis": str(pis_adriano).strip(),
        "saldo_banco_horas": saldo_total,
        "fonte": "planilha_adriano",
    }])


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
