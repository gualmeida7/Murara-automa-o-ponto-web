"""
gerar_feriados.py
=================
Gera (ou estende) o `feriados.csv` lido por `feriados.py`, SEM internet:
feriados nacionais e pontos facultativos saem de regra de calculo (datas
fixas + Pascoa), e os de Cianorte/PR de uma tabela com o que ja' foi
confirmado em decreto, mais a data fixa como PREVISAO para os anos seguintes.

Uso:
    python gerar_feriados.py                 # 2025 ate 2199 (limite da BrasilAPI)
    python gerar_feriados.py 2027 2030       # so' esses anos
    python gerar_feriados.py --verificar     # tambem confere os nacionais com a BrasilAPI (usa internet)

O arquivo existente e' MESCLADO, nunca sobrescrito: uma linha que ja' esta no
arquivo (mesma data, escopo e nome) e' mantida como esta - inclusive o "sim"/"nao"
que o RH tenha escrito em `vale_para_empresa`. Linhas de feriado municipal
"previsto" de um ano que passou a ter decreto confirmado nesta tabela sao trocadas
pelas confirmadas. Para recomeçar do zero, apague o arquivo e rode de novo.

Todo ano, quando sair o decreto do ano seguinte (o TJPR publica o calendario de
feriados locais por volta de novembro; a prefeitura pode transferir o aniversario
por decreto), acrescente o ano em MUNICIPAIS_CONFIRMADOS e rode o script.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
import urllib.error
import urllib.request
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

from feriados import converter_datas

CAMINHO_CSV = Path(__file__).with_name("feriados.csv")
ANO_INICIAL_PADRAO = 2025
ANO_FINAL_PADRAO = 2199  # a BrasilAPI (usada na verificacao) so' responde de 1900 a 2199
COLUNAS = ["data", "nome", "escopo", "local", "vale_para_empresa", "situacao", "fonte", "observacao"]

LOCAL_BR, LOCAL_PR, LOCAL_CIANORTE = "BR", "PR", "Cianorte-PR"

# ---------------------------------------------------------------------------
# Cianorte: o que ja' foi confirmado em decreto, por ano (mes, dia, nome, fonte, observacao).
# Fontes: Decretos Judiciarios do TJPR que consolidam os feriados locais das
# comarcas (645/2024 e 350/2025 para 2025; 621/2025 para 2026).
# ---------------------------------------------------------------------------
MUNICIPAIS_CONFIRMADOS: dict[int, list[tuple[int, int, str, str, str]]] = {
    2025: [
        (5, 13, "Dia da Padroeira do Município", "TJPR Decreto 645/2024", ""),
        (7, 26, "Aniversário do Município", "TJPR Decretos 645/2024 e 350/2025",
         "Em 2025 a comemoração foi transferida para 28/07 (Folha de Cianorte) mas o TJPR ainda lista o dia 26"),
        (7, 28, "Aniversário do Município (transferido do dia 26/07)", "TJPR Decreto 350/2025 e Folha de Cianorte",
         "Transferência válida só em 2025 (LC Municipal 180/2022)"),
    ],
    2026: [
        (5, 13, "Dia da Padroeira do Município", "TJPR Decreto 621/2025", ""),
        (7, 26, "Emancipação Política do Município", "TJPR Decreto 621/2025",
         "Caiu num domingo - confirmar se a prefeitura transferiu por decreto"),
    ],
}
# Data fixa usada como PREVISAO nos anos sem decreto confirmado.
MUNICIPAIS_PREVISTOS = [
    (5, 13, "Dia da Padroeira do Município"),
    (7, 26, "Aniversário / Emancipação Política do Município"),
]
OBS_PREVISAO_MUNICIPAL = "Previsão pela data fixa - confirmar o decreto do ano (pode haver transferência)"


# ---------------------------------------------------------------------------
# Regras de calculo
# ---------------------------------------------------------------------------
def pascoa(ano: int) -> date:
    """Domingo de Pascoa (calendario gregoriano - algoritmo de Meeus/Jones/Butcher)."""
    a, b, c = ano % 19, ano // 100, ano % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = (h + l - 7 * m + 114) % 31 + 1
    return date(ano, mes, dia)


def _linha(data: date, nome: str, escopo: str, local: str, situacao: str, fonte: str, observacao: str = "") -> dict:
    return {
        "data": data.isoformat(), "nome": nome, "escopo": escopo, "local": local,
        "vale_para_empresa": "", "situacao": situacao, "fonte": fonte, "observacao": observacao,
    }


def feriados_nacionais(ano: int) -> list[dict]:
    p = pascoa(ano)
    fixos = [
        (date(ano, 1, 1), "Confraternização Universal", "Lei federal"),
        (p - timedelta(days=2), "Sexta-feira Santa", "Lei federal (Páscoa calculada)"),
        (date(ano, 4, 21), "Tiradentes", "Lei federal"),
        (date(ano, 5, 1), "Dia do Trabalho", "Lei federal"),
        (date(ano, 9, 7), "Independência do Brasil", "Lei federal"),
        (date(ano, 10, 12), "Nossa Senhora Aparecida", "Lei federal"),
        (date(ano, 11, 2), "Finados", "Lei federal"),
        (date(ano, 11, 15), "Proclamação da República", "Lei federal"),
        (date(ano, 12, 25), "Natal", "Lei federal"),
    ]
    if ano >= 2024:  # Lei 14.759/2023
        fixos.append((date(ano, 11, 20), "Dia Nacional de Zumbi e da Consciência Negra", "Lei 14.759/2023"))
    return [_linha(d, n, "nacional", LOCAL_BR, "lei", f) for d, n, f in fixos]


def pontos_facultativos(ano: int) -> list[dict]:
    p = pascoa(ano)
    obs = "Ponto facultativo - só vale se a empresa ou a convenção coletiva conceder"
    itens = [
        (p - timedelta(days=48), "Segunda-feira de Carnaval"),
        (p - timedelta(days=47), "Terça-feira de Carnaval"),
        (p + timedelta(days=60), "Corpus Christi"),
        (date(ano, 12, 24), "Véspera de Natal"),
        (date(ano, 12, 31), "Véspera de Ano Novo"),
    ]
    return [_linha(d, n, "facultativo", LOCAL_BR, "previsto", "Regra de cálculo (Páscoa/data fixa)", obs) for d, n in itens]


def feriado_estadual_pr(ano: int) -> list[dict]:
    return [_linha(
        date(ano, 12, 19), "Emancipação Política do Paraná", "estadual", LOCAL_PR, "previsto",
        "Bem Paraná (decisão da Justiça do Trabalho)",
        "Segundo a Justiça do Trabalho não é feriado civil (só servidores) - confirmar na convenção coletiva",
    )]


def feriados_cianorte(ano: int) -> list[dict]:
    if ano in MUNICIPAIS_CONFIRMADOS:
        return [
            _linha(date(ano, mes, dia), nome, "municipal", LOCAL_CIANORTE, "confirmado", fonte, obs)
            for mes, dia, nome, fonte, obs in MUNICIPAIS_CONFIRMADOS[ano]
        ]
    return [
        _linha(date(ano, mes, dia), nome, "municipal", LOCAL_CIANORTE, "previsto",
               "Padrão dos anos confirmados (TJPR 2025 e 2026)", OBS_PREVISAO_MUNICIPAL)
        for mes, dia, nome in MUNICIPAIS_PREVISTOS
    ]


def gerar(ano_inicial: int, ano_final: int) -> pd.DataFrame:
    linhas: list[dict] = []
    for ano in range(ano_inicial, ano_final + 1):
        linhas += feriados_nacionais(ano) + pontos_facultativos(ano) + feriado_estadual_pr(ano) + feriados_cianorte(ano)
    return pd.DataFrame(linhas, columns=COLUNAS)


# ---------------------------------------------------------------------------
# Mesclagem com o arquivo existente
# ---------------------------------------------------------------------------
def _chave(df: pd.DataFrame) -> pd.Series:
    return df["data"].astype(str) + "|" + df["escopo"].astype(str) + "|" + df["nome"].astype(str)


def _ordenar(df: pd.DataFrame) -> pd.DataFrame:
    return df.sort_values(["data", "escopo", "nome"]).reset_index(drop=True)


def mesclar(existente: pd.DataFrame | None, novas: pd.DataFrame) -> pd.DataFrame:
    """Mantem as linhas ja' existentes (inclusive edicoes do RH) e acrescenta so' as que faltam."""
    if existente is None or existente.empty:
        return _ordenar(novas)

    existente = existente.reindex(columns=COLUNAS).fillna("")
    datas = converter_datas(existente["data"])  # aceita ISO e dd/mm/aaaa sem inverter dia e mes
    if datas.isna().any():
        raise ValueError(f"{CAMINHO_CSV.name} tem data inválida na linha {datas.index[datas.isna()][0] + 2}")
    existente = existente.assign(data=datas.dt.strftime("%Y-%m-%d"))

    # Previsao de um ano que agora tem decreto confirmado nesta tabela: sai a previsao, entra a confirmada.
    previsao_obsoleta = (
        (existente["escopo"] == "municipal")
        & (existente["situacao"] == "previsto")
        & datas.dt.year.isin(set(MUNICIPAIS_CONFIRMADOS))
    )
    existente = existente[~previsao_obsoleta]

    faltantes = novas[~_chave(novas).isin(set(_chave(existente)))]
    return _ordenar(pd.concat([existente, faltantes], ignore_index=True))


def gravar(df: pd.DataFrame, caminho: Path = CAMINHO_CSV) -> None:
    # ';' + utf-8-sig: abre direto no Excel em portugues; o leitor do projeto detecta os dois.
    df[COLUNAS].to_csv(caminho, sep=";", index=False, encoding="utf-8-sig")


# ---------------------------------------------------------------------------
# Verificacao opcional contra a BrasilAPI (usa internet)
# ---------------------------------------------------------------------------
def _buscar_brasilapi(ano: int) -> list[dict]:
    # A BrasilAPI (Cloudflare) recusa o User-Agent padrao do Python (erro 1010): identifica-se.
    requisicao = urllib.request.Request(
        f"https://brasilapi.com.br/api/feriados/v1/{ano}",
        headers={"User-Agent": "fechamento-folha-feriados/1.0 (uso interno)", "Accept": "application/json"},
    )
    for tentativa in range(3):
        try:
            with urllib.request.urlopen(requisicao, timeout=20) as resposta:
                return json.load(resposta)
        except urllib.error.HTTPError as e:
            if e.code != 429 or tentativa == 2:
                raise
            time.sleep(2 * (tentativa + 1))
    return []


def verificar_com_brasilapi(ano_inicial: int, ano_final: int) -> list[str]:
    """
    Confere, ano a ano, as datas nacionais + Carnaval/Corpus Christi/Pascoa
    geradas aqui contra as que a BrasilAPI devolve. Devolve as divergencias.
    """
    divergencias = []
    for ano in range(ano_inicial, ano_final + 1):
        api = {item["date"] for item in _buscar_brasilapi(ano)}
        nossas = {l["data"] for l in feriados_nacionais(ano)}
        p = pascoa(ano)
        nossas |= {(p - timedelta(days=48)).isoformat(), (p - timedelta(days=47)).isoformat(),
                   (p + timedelta(days=60)).isoformat(), p.isoformat()}
        if api != nossas:
            divergencias.append(f"{ano}: só na API {sorted(api - nossas)} | só aqui {sorted(nossas - api)}")
        time.sleep(0.05)
    return divergencias


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Gera/estende o feriados.csv (nacional, facultativos, PR e Cianorte).")
    ap.add_argument("ano_inicial", nargs="?", type=int, default=ANO_INICIAL_PADRAO)
    ap.add_argument("ano_final", nargs="?", type=int, default=ANO_FINAL_PADRAO)
    ap.add_argument("--verificar", action="store_true", help="confere os nacionais com a BrasilAPI (usa internet)")
    args = ap.parse_args(argv)
    if args.ano_inicial > args.ano_final:
        ap.error("ano_inicial maior que ano_final")

    if args.verificar:
        print(f"Verificando {args.ano_inicial}-{args.ano_final} na BrasilAPI...")
        divergencias = verificar_com_brasilapi(args.ano_inicial, min(args.ano_final, 2199))
        for d in divergencias:
            print("  DIVERGE:", d)
        if divergencias:
            print("Nada foi gravado. Corrija a regra ou investigue a divergência.")
            return 1
        print("  OK: todas as datas batem com a BrasilAPI.")

    existente = pd.read_csv(CAMINHO_CSV, sep=";", encoding="utf-8-sig", dtype=str) if CAMINHO_CSV.exists() else None
    resultado = mesclar(existente, gerar(args.ano_inicial, args.ano_final))
    gravar(resultado)
    print(f"{CAMINHO_CSV.name}: {len(resultado)} linha(s), anos {args.ano_inicial}-{args.ano_final} garantidos.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
