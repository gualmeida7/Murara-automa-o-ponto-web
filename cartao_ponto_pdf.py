"""
cartao_ponto_pdf.py
===================
Le o relatorio "CARTAO PONTO" exportado em PDF pelo Secullum RH e devolve,
para cada colaborador do arquivo:

  - os dados cadastrais impressos no cabecalho (nome, Nº folha, CPF, CNPJ,
    empresa, departamento, funcao, admissao);
  - uma linha por dia com as marcacoes (ENT.1, SAI.1 ... SAI.3) e os
    totalizadores do proprio Secullum (NORMAIS, FALTAS, EX50%, EX100%,
    AJUSTE, DSR.DEB, ...);
  - a linha TOTAIS do periodo.

Por que ler pela POSICAO das palavras e nao pelo texto corrido:
o texto extraido do PDF perde as celulas vazias. "31/08/2026 - Seg 07:30
12:10 13:10 04:40 04:20" nao diz se 04:20 e' FALTAS H ou EX50%. Por isso
cada valor e' atribuido a coluna cujo titulo (na linha do cabecalho
"DATA ENTRADA1SAIDA1 ...") esta imediatamente acima dele. As colunas sao
recalculadas a cada pagina, entao um relatorio com colunas em outra ordem
ou com colunas a mais continua sendo lido corretamente.

O Cartao Ponto NAO traz PIS - o identificador do colaborador e' o
"Nº FOLHA" (matricula). A conversao para PIS e' feita por
`leitores_ponto.ler_arquivo_ponto()`, com o Cadastro de Colaboradores.

Uso rapido para conferir um PDF novo:

    python cartao_ponto_pdf.py "Cartao Ponto 09.2026.pdf"
"""

from __future__ import annotations

import argparse
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from validacao import ArquivoInvalidoError

# ---------------------------------------------------------------------------
# CONFIGURACAO DE LAYOUT
# ---------------------------------------------------------------------------
# Colunas de marcacao: o Secullum imprime "ENTRADA1SAIDA1" como UM titulo
# so' cobrindo duas celulas; a metade direita e' a SAIDA.
REGEX_TITULO_PAR_MARCACAO = re.compile(r"^ENTRADA(\d)SAIDA\1$")

# Titulo impresso (sem acento, maiusculo) -> nome da coluna no DataFrame.
# Titulos fora desta lista continuam sendo lidos, com o nome normalizado.
TITULOS_TOTALIZADORES = {
    "NORMAIS": "normais",
    "FALTASD": "faltas_dias",
    "FALTASH": "faltas_horas",
    "EX50%": "extra_50",
    "EX100%": "extra_100",
    "AJUSTE": "ajuste",
    "DSR.DEB": "dsr_debito",
    "NOT.TOT.": "noturno_total",
    "JUSTPA.": "justificativa_parcial",
}
COLUNAS_EM_DIAS = {"faltas_dias"}

# Rotulo do cabecalho do colaborador (sem acento) -> campo de ColaboradorCartaoPonto.
ROTULOS_CABECALHO = {
    "EMPRESA:": "empresa",
    "CNPJ:": "cnpj",
    "INSCRICAO:": "inscricao",
    "DEPARTAMENTO:": "departamento",
    "FUNCAO:": "funcao",
    "NOME:": "nome",
    "ADMISSAO:": "admissao",
    "CPF:": "cpf",
    "C.T.P.S.:": "ctps",
    "NOFOLHA:": "matricula",
}
# O valor de cada rotulo fica na linha logo abaixo dele (~10pt no layout atual).
DISTANCIA_MIN_VALOR_ROTULO = 4
DISTANCIA_MAX_VALOR_ROTULO = 16
# A partir daqui, a direita do cabecalho, fica o quadro "HORARIO DE TRABALHO".
LIMITE_X_CABECALHO = 480

TOLERANCIA_ALINHAMENTO_COLUNA = 6   # pt que um valor pode "vazar" a esquerda do titulo
TOLERANCIA_MESMA_LINHA = 3          # pt de diferenca de 'top' ainda na mesma linha

REGEX_DATA_LINHA = re.compile(r"^\d{2}/\d{2}/\d{4}$")
REGEX_HORA = re.compile(r"^(-?\d{1,3}):(\d{2})([*¨^]?)$")
REGEX_NUMERO = re.compile(r"^-?\d+(,\d+)?$")
MARCA_BATIDA_MANUAL = "*"


@dataclass(frozen=True)
class ColaboradorCartaoPonto:
    matricula: str
    nome: str
    cpf: str = ""
    cnpj: str = ""
    empresa: str = ""
    departamento: str = ""
    funcao: str = ""
    admissao: str = ""
    inscricao: str = ""
    ctps: str = ""


@dataclass
class ResultadoCartaoPonto:
    colaboradores: pd.DataFrame
    dias: pd.DataFrame
    totais: pd.DataFrame
    avisos: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class _Coluna:
    nome: str
    x0: float


# ---------------------------------------------------------------------------
# Conversao de valores
# ---------------------------------------------------------------------------
def _sem_acento(texto: str) -> str:
    """'SAÍDA' -> 'SAIDA', 'NºFOLHA:' -> 'NoFOLHA:' (NFKD troca 'º' por 'o')."""
    decomposto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in decomposto if not unicodedata.combining(c))


def hhmm_para_minutos(valor: str) -> int | None:
    """'158:33' -> 9513, '-09:00' -> -540. Devolve None se nao for hora."""
    casamento = REGEX_HORA.match(valor)
    if not casamento:
        return None
    horas, minutos = int(casamento.group(1)), int(casamento.group(2))
    sinal = -1 if casamento.group(1).startswith("-") else 1
    return sinal * (abs(horas) * 60 + minutos)


def _numero_br(valor: str) -> float | None:
    """'3,00' -> 3.0, '1' -> 1.0."""
    if not REGEX_NUMERO.match(valor):
        return None
    return float(valor.replace(",", "."))


# ---------------------------------------------------------------------------
# Geometria da pagina
# ---------------------------------------------------------------------------
def _agrupar_em_linhas(palavras: list[dict]) -> list[list[dict]]:
    """Agrupa as palavras extraidas por linha visual, ordenadas da esquerda p/ direita."""
    linhas: list[list[dict]] = []
    for palavra in sorted(palavras, key=lambda p: (p["top"], p["x0"])):
        if linhas and abs(linhas[-1][0]["top"] - palavra["top"]) <= TOLERANCIA_MESMA_LINHA:
            linhas[-1].append(palavra)
        else:
            linhas.append([palavra])
    return [sorted(linha, key=lambda p: p["x0"]) for linha in linhas]


def _localizar_cabecalho_tabela(linhas: list[list[dict]]) -> int | None:
    for indice, linha in enumerate(linhas):
        if linha[0]["text"] == "DATA" and len(linha) > 3:
            return indice
    return None


def _montar_colunas(linha_cabecalho: list[dict]) -> list[_Coluna]:
    colunas = []
    for palavra in linha_cabecalho[1:]:  # [0] e' "DATA"
        titulo = _sem_acento(palavra["text"]).upper()
        par = REGEX_TITULO_PAR_MARCACAO.match(titulo)
        if par:
            numero = par.group(1)
            meio = (palavra["x0"] + palavra["x1"]) / 2
            colunas.append(_Coluna(f"ent{numero}", palavra["x0"]))
            colunas.append(_Coluna(f"sai{numero}", meio))
        else:
            nome = TITULOS_TOTALIZADORES.get(titulo, re.sub(r"\W+", "_", titulo.lower()).strip("_"))
            colunas.append(_Coluna(nome, palavra["x0"]))
    return sorted(colunas, key=lambda c: c.x0)


def _coluna_da_palavra(palavra: dict, colunas: list[_Coluna]) -> _Coluna | None:
    """A coluna e' o titulo mais a direita que ainda comeca antes da palavra."""
    candidatas = [c for c in colunas if c.x0 <= palavra["x0"] + TOLERANCIA_ALINHAMENTO_COLUNA]
    return candidatas[-1] if candidatas else None


# ---------------------------------------------------------------------------
# Cabecalho do colaborador
# ---------------------------------------------------------------------------
def _extrair_colaborador(palavras: list[dict]) -> ColaboradorCartaoPonto | None:
    rotulos = [
        (ROTULOS_CABECALHO[_sem_acento(p["text"]).upper()], p)
        for p in palavras
        if _sem_acento(p["text"]).upper() in ROTULOS_CABECALHO and p["x0"] < LIMITE_X_CABECALHO
    ]
    if not rotulos:
        return None

    campos = {}
    for campo, rotulo in rotulos:
        limite_direito = min(
            [outro["x0"] for _, outro in rotulos
             if abs(outro["top"] - rotulo["top"]) <= TOLERANCIA_MESMA_LINHA and outro["x0"] > rotulo["x0"]]
            + [LIMITE_X_CABECALHO]
        )
        valor = [
            p for p in palavras
            if DISTANCIA_MIN_VALOR_ROTULO < p["top"] - rotulo["top"] < DISTANCIA_MAX_VALOR_ROTULO
            and rotulo["x0"] - TOLERANCIA_ALINHAMENTO_COLUNA <= p["x0"] < limite_direito
        ]
        campos[campo] = " ".join(p["text"] for p in sorted(valor, key=lambda p: p["x0"]))

    if not campos.get("matricula") and not campos.get("nome"):
        return None
    return ColaboradorCartaoPonto(**{"matricula": "", "nome": "", **campos})


# ---------------------------------------------------------------------------
# Linhas da tabela diaria
# ---------------------------------------------------------------------------
def _ler_celulas(linha: list[dict], colunas: list[_Coluna]) -> dict[str, list[str]]:
    celulas: dict[str, list[str]] = {}
    for palavra in linha:
        coluna = _coluna_da_palavra(palavra, colunas)
        if coluna is not None:
            celulas.setdefault(coluna.nome, []).append(palavra["text"])
    return celulas


def _converter_totalizador(nome_coluna: str, texto: str):
    if nome_coluna in COLUNAS_EM_DIAS:
        return _numero_br(texto)
    minutos = hhmm_para_minutos(texto)
    return minutos if minutos is not None else _numero_br(texto)


def _montar_registro_dia(
    data_dia: date, celulas: dict[str, list[str]], colunas_marcacao: list[str]
) -> dict:
    """
    Celulas de marcacao podem trazer hora ('07:30'), hora manual ('07:30*')
    ou uma ocorrencia em texto ('ATESTAD', 'FALTA'). Ocorrencias sao
    guardadas a parte - nao sao batidas.
    """
    registro: dict = {"data": data_dia, "batidas_manuais": 0}
    ocorrencias = []
    for nome in colunas_marcacao:
        texto = " ".join(celulas.get(nome, []))
        casamento = REGEX_HORA.match(texto)
        if casamento:
            registro[nome] = f"{int(casamento.group(1)):02d}:{casamento.group(2)}"
            registro["batidas_manuais"] += casamento.group(3) == MARCA_BATIDA_MANUAL
        else:
            registro[nome] = None
            if texto:
                ocorrencias.append(texto)
    registro["ocorrencia"] = ", ".join(dict.fromkeys(ocorrencias)) or None

    for nome, textos in celulas.items():
        if nome not in colunas_marcacao:
            registro[nome] = _converter_totalizador(nome, " ".join(textos))
    return registro


def _ler_totais(celulas: dict[str, list[str]]) -> dict:
    return {nome: _converter_totalizador(nome, " ".join(textos)) for nome, textos in celulas.items()}


# ---------------------------------------------------------------------------
# Orquestracao
# ---------------------------------------------------------------------------
def _abrir_pdf(caminho: Path):
    try:
        import pdfplumber
    except ImportError as e:
        raise ArquivoInvalidoError(
            "Para ler o Cartão Ponto em PDF é preciso instalar o pacote 'pdfplumber' "
            "(pip install pdfplumber)."
        ) from e
    try:
        return pdfplumber.open(caminho)
    except Exception as e:
        raise ArquivoInvalidoError(
            f"Não consegui abrir o PDF '{caminho}'. Ele pode estar corrompido ou protegido "
            f"por senha.\nDetalhe técnico: {e}"
        ) from e


def parse_cartao_ponto_pdf(caminho_arquivo: str | Path) -> ResultadoCartaoPonto:
    """
    Le todas as paginas do PDF. Paginas sem tabela (ex.: a pagina de
    assinatura que o Secullum imprime apos cada colaborador) sao ignoradas;
    uma tabela sem cabecalho de colaborador na mesma pagina e' atribuida
    ao ultimo colaborador lido (relatorio que quebrou pagina no meio).
    """
    caminho_arquivo = Path(caminho_arquivo)
    if not caminho_arquivo.exists():
        raise ArquivoInvalidoError(f"O Cartão Ponto em PDF não foi encontrado em '{caminho_arquivo}'.")

    colaboradores: dict[str, ColaboradorCartaoPonto] = {}
    dias: list[dict] = []
    totais: list[dict] = []
    avisos: list[str] = []
    colaborador_atual: ColaboradorCartaoPonto | None = None

    with _abrir_pdf(caminho_arquivo) as pdf:
        for pagina in pdf.pages:
            palavras = pagina.extract_words()
            linhas = _agrupar_em_linhas(palavras)
            indice_cabecalho = _localizar_cabecalho_tabela(linhas)
            if indice_cabecalho is None:
                continue

            topo_tabela = linhas[indice_cabecalho][0]["top"]
            colaborador_atual = (
                _extrair_colaborador([p for p in palavras if p["top"] < topo_tabela]) or colaborador_atual
            )
            if colaborador_atual is None:
                avisos.append(f"página {pagina.page_number}: tabela sem colaborador identificado - ignorada")
                continue
            if not colaborador_atual.matricula:
                avisos.append(
                    f"página {pagina.page_number}: {colaborador_atual.nome} sem Nº FOLHA - "
                    "não será possível cruzar com o cadastro"
                )
            colaboradores[colaborador_atual.matricula or colaborador_atual.nome] = colaborador_atual

            colunas = _montar_colunas(linhas[indice_cabecalho])
            colunas_marcacao = [c.nome for c in colunas if c.nome.startswith(("ent", "sai"))]
            identificacao = {"matricula": colaborador_atual.matricula, "nome": colaborador_atual.nome}

            for linha in linhas[indice_cabecalho + 1:]:
                primeira = linha[0]["text"]
                celulas = _ler_celulas(linha, colunas)
                if primeira == "TOTAIS":
                    totais.append({**identificacao, **_ler_totais(celulas)})
                elif REGEX_DATA_LINHA.match(primeira):
                    data_dia = datetime.strptime(primeira, "%d/%m/%Y").date()
                    dias.append({**identificacao, **_montar_registro_dia(data_dia, celulas, colunas_marcacao)})

    if not dias:
        raise ArquivoInvalidoError(
            f"O PDF '{caminho_arquivo}' foi aberto, mas não encontrei nenhuma tabela de "
            "Cartão Ponto do Secullum (linha de títulos 'DATA ENTRADA1 SAÍDA1 ...'). "
            "Confira se é o relatório 'Cartão Ponto' e não outro relatório do Secullum."
        )

    return ResultadoCartaoPonto(
        colaboradores=pd.DataFrame([vars(c) for c in colaboradores.values()]),
        dias=pd.DataFrame(dias),
        totais=pd.DataFrame(totais),
        avisos=avisos,
    )


def dias_para_batidas(df_dias: pd.DataFrame) -> pd.DataFrame:
    """
    Converte as marcacoes diarias no mesmo formato de `afd_parser.parse_afd`
    (uma linha por batida: pis | data_hora | tipo_registro), para o resto
    do pipeline nao precisar saber de onde o ponto veio. Aqui 'pis' ainda
    contem a MATRICULA - a troca pelo PIS acontece em leitores_ponto.py.
    """
    colunas_marcacao = [c for c in df_dias.columns if re.fullmatch(r"(ent|sai)\d", c)]
    longo = df_dias.melt(id_vars=["matricula", "data"], value_vars=colunas_marcacao, value_name="hora").dropna(
        subset=["hora"]
    )
    batidas = pd.DataFrame({
        "pis": longo["matricula"].astype(str),
        "data_hora": pd.to_datetime(longo["data"].astype(str) + " " + longo["hora"], format="%Y-%m-%d %H:%M"),
        "tipo_registro": "PDF",
    })
    return batidas.sort_values(["pis", "data_hora"]).reset_index(drop=True)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Parser do Cartão Ponto (PDF) do Secullum RH")
    ap.add_argument("arquivo", help="Caminho do PDF")
    args = ap.parse_args()

    resultado = parse_cartao_ponto_pdf(args.arquivo)
    print(f"Colaboradores: {len(resultado.colaboradores)}")
    print(f"Dias lidos: {len(resultado.dias)}")
    for aviso in resultado.avisos:
        print(f"  [aviso] {aviso}")
    print("\nTotais por colaborador (horas em minutos):")
    print(resultado.totais.to_string(index=False))
