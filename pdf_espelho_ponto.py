"""
pdf_espelho_ponto.py
====================
Le o relatorio "ESPELHO DE PONTO ELETRONICO" exportado em PDF pelo Secullum RH.

Diferente do Cartao Ponto (resultado consolidado), o Espelho e' a trilha de
auditoria: cada dia traz as MARCACOES REGISTRADAS NO PONTO ELETRONICO (as
batidas brutas do relogio, bloco da esquerda), a jornada realizada (o que o
Secullum decidiu considerar, bloco da direita) e, quando houve, o tratamento
automatico aplicado ("Batida duplicada", "Alocar batidas" ...).

Este modulo extrai so' o que o cruzamento precisa: por colaborador e por dia,
as batidas brutas e o tratamento aplicado. O bloco da esquerda e' separado do
da direita pela POSICAO das palavras (o texto corrido perde as celulas vazias).

O Espelho identifica o colaborador pelo NOME e, quando o Secullum tem, pelo
PIS (a coluna fica vazia para quem nao tem PIS cadastrado). A conversao para
o PIS canonico do cadastro e' feita por `mapear_para_pis`.
"""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from cartao_ponto_pdf import _abrir_pdf, _agrupar_em_linhas, _sem_acento
from validacao import ArquivoInvalidoError, canonicalizar_pis

REGEX_PERIODO = re.compile(r"(\d{2}/\d{2}/\d{4})\s+at\S*\s+(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
REGEX_DIA_MES = re.compile(r"^\d{2}/\d{2}$")
REGEX_HORA_SIMPLES = re.compile(r"^\d{1,2}:\d{2}$")
REGEX_PIS = re.compile(r"^\d{11,12}$")

# Marcacoes brutas ficam a esquerda; a "jornada realizada" comeca em ~x=239 no layout atual.
# Recalculado por pagina a partir do titulo "JORNADAREALIZADA" quando existir.
LIMITE_X_BRUTAS_PADRAO = 230
# Colunas do final da linha (DURACAO, HORARIO, OCOR., MOTIVO) comecam depois das 6 colunas de jornada.
TITULO_MOTIVO = "MOTIVO"


@dataclass
class ResultadoEspelhoPonto:
    colaboradores: pd.DataFrame          # nome | pis | cpf
    dias: pd.DataFrame                   # nome | pis | data | batidas_brutas (list[str]) | motivo
    avisos: list[str] = field(default_factory=list)


def _resolver_data(dia_mes: str, inicio: date, fim: date) -> date | None:
    """O Espelho imprime 'dd/mm' sem ano; escolhe o ano que cai dentro do periodo do relatorio."""
    dia, mes = (int(x) for x in dia_mes.split("/"))
    for ano in sorted({inicio.year, fim.year}):
        try:
            candidata = date(ano, mes, dia)
        except ValueError:
            continue
        if inicio <= candidata <= fim:
            return candidata
    return None


def _ler_periodo(texto_pagina: str) -> tuple[date, date] | None:
    casamento = REGEX_PERIODO.search(texto_pagina)
    if not casamento:
        return None
    return tuple(datetime.strptime(g, "%d/%m/%Y").date() for g in casamento.groups())  # type: ignore[return-value]


def _extrair_colaborador(linhas: list[list[dict]]) -> dict | None:
    """Linha de rotulos 'NOME: NºPIS/PASEP: CPF: ...' e a linha logo abaixo com os valores."""
    for indice, linha in enumerate(linhas[:-1]):
        if _sem_acento(linha[0]["text"]).upper() != "NOME:":
            continue
        rotulos = {_sem_acento(p["text"]).upper(): p["x0"] for p in linha}
        x_pis = next((x for r, x in rotulos.items() if "PIS" in r), None)
        x_cpf = rotulos.get("CPF:")
        # ate' o proximo rotulo a direita do CPF (ADMISSAO) - limita a coluna do CPF
        x_fim_cpf = next((x for r, x in rotulos.items() if r.startswith("ADMISS")), float("inf"))
        valores = linhas[indice + 1]
        x_inicio_pis = x_pis if x_pis is not None else float("inf")
        nome = " ".join(p["text"] for p in valores if p["x0"] < min(x_inicio_pis, x_cpf or float("inf")) - 2)
        pis = " ".join(
            p["text"] for p in valores if x_pis is not None and x_pis - 2 <= p["x0"] < (x_cpf or float("inf")) - 2
        )
        cpf = " ".join(p["text"] for p in valores if x_cpf is not None and x_cpf - 2 <= p["x0"] < x_fim_cpf - 2)
        pis_limpo = re.sub(r"\D", "", pis)
        return {
            "nome": nome.strip(),
            "pis": canonicalizar_pis(pis_limpo) if REGEX_PIS.match(pis_limpo) else "",
            "cpf": cpf.strip(),
        }
    return None


def parse_espelho_pdf(caminho_arquivo: str | Path) -> ResultadoEspelhoPonto:
    caminho_arquivo = Path(caminho_arquivo)
    if not caminho_arquivo.exists():
        raise ArquivoInvalidoError(f"O Espelho de Ponto em PDF não foi encontrado em '{caminho_arquivo}'.")

    colaboradores: dict[str, dict] = {}
    dias: list[dict] = []
    avisos: list[str] = []

    with _abrir_pdf(caminho_arquivo) as pdf:
        for pagina in pdf.pages:
            palavras = pagina.extract_words()
            linhas = _agrupar_em_linhas(palavras)
            periodo = _ler_periodo(pagina.extract_text() or "")
            colaborador = _extrair_colaborador(linhas)
            if colaborador is None or periodo is None:
                continue

            titulo_realizada = next(
                (p for p in palavras if _sem_acento(p["text"]).upper().startswith("JORNADAREALIZADA")), None
            )
            # a coluna "ENTRADA" da jornada realizada fica sob o titulo; as brutas terminam antes dela.
            primeira_entrada = sorted(
                (p for p in palavras if _sem_acento(p["text"]).upper() == "ENTRADA"), key=lambda p: p["x0"]
            )
            limite_x = (
                titulo_realizada["x0"] - 2 if titulo_realizada
                else (primeira_entrada[3]["x0"] - 2 if len(primeira_entrada) > 3 else LIMITE_X_BRUTAS_PADRAO)
            )
            motivo_x = next((p["x0"] for p in palavras if p["text"].upper() == TITULO_MOTIVO), None)

            chave = colaborador["pis"] or colaborador["nome"]
            colaboradores[chave] = colaborador

            for linha in linhas:
                if not REGEX_DIA_MES.match(linha[0]["text"]):
                    continue
                data_dia = _resolver_data(linha[0]["text"], *periodo)
                if data_dia is None:
                    avisos.append(f"página {pagina.page_number}: data '{linha[0]['text']}' fora do período - ignorada")
                    continue
                brutas = [
                    p["text"] for p in linha[1:]
                    if p["x0"] < limite_x and REGEX_HORA_SIMPLES.match(p["text"])
                ]
                motivo = " ".join(p["text"] for p in linha if motivo_x is not None and p["x0"] >= motivo_x - 2)
                dias.append({
                    "nome": colaborador["nome"],
                    "pis": colaborador["pis"],
                    "data": data_dia,
                    "batidas_brutas": brutas,
                    "motivo": motivo or None,
                })

    if not dias:
        raise ArquivoInvalidoError(
            f"O PDF '{caminho_arquivo}' foi aberto, mas não encontrei nenhuma linha de Espelho de Ponto "
            "(colunas 'DIA / MARCAÇÕES REGISTRADAS NO PONTO ELETRÔNICO'). "
            "Confira se é o relatório 'Espelho de Ponto Eletrônico' do Secullum."
        )

    return ResultadoEspelhoPonto(
        colaboradores=pd.DataFrame(list(colaboradores.values())),
        dias=pd.DataFrame(dias),
        avisos=avisos,
    )


def _normalizar_nome(nome: str) -> str:
    return re.sub(r"\s+", "", _sem_acento(str(nome))).upper()


def mapear_para_pis(resultado: ResultadoEspelhoPonto, cadastro: pd.DataFrame) -> pd.DataFrame:
    """
    Devolve `resultado.dias` com 'pis' trocado pelo PIS canonico do cadastro.
    Tenta o PIS impresso no Espelho primeiro e, quando ele vem vazio, o NOME
    (comparado sem acento/espacos). Quem nao casa vira aviso e e' descartado -
    mesma politica de `leitores_ponto.ler_arquivo_ponto`.
    """
    pis_cadastro = set(cadastro["pis"].astype(str).str.strip())
    nome_para_pis = {
        _normalizar_nome(n): str(p).strip() for n, p in zip(cadastro["nome"], cadastro["pis"])
    }

    def _resolver(linha) -> str | None:
        if linha["pis"] in pis_cadastro:
            return linha["pis"]
        return nome_para_pis.get(_normalizar_nome(linha["nome"]))

    dias = resultado.dias.copy()
    dias["pis"] = dias.apply(_resolver, axis=1)
    for nome in sorted(dias.loc[dias["pis"].isna(), "nome"].unique()):
        resultado.avisos.append(f"'{nome}' do Espelho de Ponto não bate com o cadastro (PIS/nome) - ignorado")
    return dias.dropna(subset=["pis"]).reset_index(drop=True)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Parser do Espelho de Ponto (PDF) do Secullum RH")
    ap.add_argument("arquivo", help="Caminho do PDF")
    args = ap.parse_args()
    r = parse_espelho_pdf(args.arquivo)
    print(f"Colaboradores: {len(r.colaboradores)} | Dias lidos: {len(r.dias)}")
    print(r.colaboradores.to_string(index=False))
    for aviso in r.avisos:
        print(f"  [aviso] {aviso}")
