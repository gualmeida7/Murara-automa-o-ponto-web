"""
leitores_ponto.py
=================
Ponto unico de entrada para ler o ponto, qualquer que seja o formato
exportado pelo Secullum:

    .txt / .afd  -> AFD (Portaria 671/2021)          -> afd_parser.py
    .pdf         -> relatorio "Cartao Ponto" (PDF)   -> cartao_ponto_pdf.py

Cada formato e' uma estrategia (`LeitorPonto`) e todas devolvem o mesmo
`ResultadoLeituraPonto`, cuja tabela `batidas` tem exatamente o formato
que o resto do pipeline ja' usa (pis | data_hora | tipo_registro). Para
suportar um formato novo basta criar outra classe com `extensoes` e
`ler()` e registra-la em `LEITORES_DISPONIVEIS` - nada em business_rules.py
ou pipeline.py precisa mudar.

O AFD identifica o colaborador pelo PIS; o Cartao Ponto em PDF, pela
matricula (Nº FOLHA). Quando o Cadastro de Colaboradores e' informado,
os identificadores sao convertidos para o PIS canonico do cadastro (PIS
primeiro, matricula depois - a mesma regra dos consignados e banco de horas).
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from afd_parser import parse_afd
from cartao_ponto_pdf import dias_para_batidas, parse_cartao_ponto_pdf
from pdf_espelho_ponto import mapear_para_pis, parse_espelho_pdf
from validacao import ArquivoInvalidoError

NOME_AMIGAVEL_ARQUIVO_PONTO = "Arquivo de Ponto"


@dataclass
class ResultadoLeituraPonto:
    batidas: pd.DataFrame
    erros_leitura: list[str] = field(default_factory=list)
    # So' preenchidos por formatos que trazem a apuracao do proprio Secullum
    # (Cartao Ponto em PDF). Servem para conferencia contra o calculo do sistema.
    espelho_diario: pd.DataFrame | None = None
    totais_secullum: pd.DataFrame | None = None
    colaboradores: pd.DataFrame | None = None


class LeitorPonto(ABC):
    extensoes: tuple[str, ...] = ()

    def aceita(self, caminho: Path) -> bool:
        return caminho.suffix.lower() in self.extensoes

    @abstractmethod
    def ler(self, caminho: Path) -> ResultadoLeituraPonto:
        """Le o arquivo; a coluna 'pis' de `batidas` traz o identificador como veio no arquivo."""


class LeitorAFDTexto(LeitorPonto):
    extensoes = (".txt", ".afd")

    def ler(self, caminho: Path) -> ResultadoLeituraPonto:
        resultado = parse_afd(caminho)
        return ResultadoLeituraPonto(batidas=resultado.batidas, erros_leitura=resultado.erros_leitura)


class LeitorCartaoPontoPDF(LeitorPonto):
    extensoes = (".pdf",)

    def ler(self, caminho: Path) -> ResultadoLeituraPonto:
        resultado = parse_cartao_ponto_pdf(caminho)
        return ResultadoLeituraPonto(
            batidas=dias_para_batidas(resultado.dias),
            erros_leitura=resultado.avisos,
            espelho_diario=resultado.dias,
            totais_secullum=resultado.totais,
            colaboradores=resultado.colaboradores,
        )


LEITORES_DISPONIVEIS: tuple[LeitorPonto, ...] = (LeitorAFDTexto(), LeitorCartaoPontoPDF())
EXTENSOES_ACEITAS = tuple(ext for leitor in LEITORES_DISPONIVEIS for ext in leitor.extensoes)


def escolher_leitor(caminho: Path) -> LeitorPonto:
    for leitor in LEITORES_DISPONIVEIS:
        if leitor.aceita(caminho):
            return leitor
    raise ArquivoInvalidoError(
        f"O {NOME_AMIGAVEL_ARQUIVO_PONTO} ('{caminho}') tem uma extensão não suportada "
        f"('{caminho.suffix or 'sem extensão'}'). Use {', '.join(EXTENSOES_ACEITAS)}."
    )


TOLERANCIA_BATIDA_DUPLICADA_MIN = 2


def ler_espelho_como_ponto(caminho: str | Path, cadastro: pd.DataFrame) -> ResultadoLeituraPonto:
    """
    Usa o Espelho de Ponto (PDF) como UNICA fonte de batidas: as marcacoes brutas
    do relogio viram `batidas` no mesmo formato do AFD, ja' com o PIS do cadastro.
    Marcacoes a menos de 2 min uma da outra sao a mesma batida duplicada pelo relogio.
    """
    resultado = parse_espelho_pdf(caminho)
    dias = mapear_para_pis(resultado, cadastro)
    linhas = [
        {"pis": str(pis), "data_hora": pd.Timestamp(f"{data} {hora}")}
        for pis, data, brutas in zip(dias["pis"], dias["data"], dias["batidas_brutas"])
        for hora in brutas
    ]
    batidas = pd.DataFrame(linhas, columns=["pis", "data_hora"])
    if not batidas.empty:
        batidas = batidas.sort_values(["pis", "data_hora"]).reset_index(drop=True)
        intervalo = batidas.groupby("pis")["data_hora"].diff()
        batidas = batidas[~(intervalo < pd.Timedelta(minutes=TOLERANCIA_BATIDA_DUPLICADA_MIN))]
    batidas = batidas.assign(tipo_registro="ESPELHO").reset_index(drop=True)
    return ResultadoLeituraPonto(batidas=batidas, erros_leitura=resultado.avisos)


def _identificadores_do_arquivo(resultado: ResultadoLeituraPonto) -> set[str]:
    """Inclui quem aparece no espelho sem nenhuma batida (ex.: afastado o mes todo)."""
    ids = set(resultado.batidas["pis"].astype(str).str.strip())
    if resultado.colaboradores is not None and not resultado.colaboradores.empty:
        ids |= set(resultado.colaboradores["matricula"].astype(str).str.strip())
    return ids - {""}


def _mapa_identificador_para_pis(identificadores: set[str], cadastro: pd.DataFrame) -> dict[str, str]:
    """
    Mesma regra de `validacao.normalizar_identificador_colaborador` (PIS
    primeiro, matricula depois), mas sem interromper o processo: um
    colaborador do ponto que nao esta no cadastro (ex.: desligado) fica de
    fora da apuracao, como ja' acontecia com o AFD, e vira um aviso.
    """
    pis_cadastro = cadastro["pis"].astype(str).str.strip()
    matricula_para_pis = dict(zip(cadastro["matricula"].astype(str).str.strip(), pis_cadastro))
    pis_validos = set(pis_cadastro)
    mapa = {}
    for identificador in identificadores:
        if identificador in pis_validos:
            mapa[identificador] = identificador
        elif identificador in matricula_para_pis:
            mapa[identificador] = matricula_para_pis[identificador]
    return mapa


def _aplicar_pis(df: pd.DataFrame | None, coluna_origem: str, mapa: dict) -> pd.DataFrame | None:
    if df is None or df.empty:
        return df
    return df.assign(pis=df[coluna_origem].astype(str).str.strip().map(mapa))


def ler_arquivo_ponto(caminho: str | Path, cadastro: pd.DataFrame | None = None) -> ResultadoLeituraPonto:
    """
    Le o arquivo de ponto escolhendo a estrategia pela extensao. Com
    `cadastro`, troca matricula/PIS do arquivo pelo PIS canonico do cadastro
    em `batidas` e, quando existirem, em `espelho_diario` e `totais_secullum`.
    Batidas de quem nao esta no cadastro sao descartadas com um aviso em
    `erros_leitura`.
    """
    caminho = Path(caminho)
    resultado = escolher_leitor(caminho).ler(caminho)
    if cadastro is None:
        return resultado

    identificadores = _identificadores_do_arquivo(resultado)
    mapa = _mapa_identificador_para_pis(identificadores, cadastro)
    for desconhecido in sorted(identificadores - mapa.keys()):
        resultado.erros_leitura.append(
            f"identificador '{desconhecido}' do ponto não bate com PIS nem matrícula do cadastro - ignorado"
        )

    batidas = resultado.batidas.assign(pis=resultado.batidas["pis"].astype(str).str.strip().map(mapa))
    resultado.batidas = batidas.dropna(subset=["pis"]).reset_index(drop=True)
    resultado.espelho_diario = _aplicar_pis(resultado.espelho_diario, "matricula", mapa)
    resultado.totais_secullum = _aplicar_pis(resultado.totais_secullum, "matricula", mapa)
    return resultado
