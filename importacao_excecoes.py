"""
importacao_excecoes.py
======================
Importa a planilha de revisão de exceções (gerada por `exportar_fila_excecoes_xlsx`) depois de preenchida
por quem decide (Thales), devolvendo as decisões prontas para aplicar na fila da etapa 4.

Por que existe: sem isso o RH precisaria digitar, linha a linha, decisões que já estão numa planilha.

Regras (nenhuma regra de cálculo muda; a decisão continua sendo humana):
- A linha da planilha é casada com a fila por Colaborador + Empresa + Data (a planilha não tem PIS).
  Nome e empresa são normalizados (maiúsculas, sem acento, espaços colapsados); a data aceita
  dd/mm/aaaa em texto e datetime do Excel.
- Só entram decisões idênticas às opções válidas da coluna Decisão (ignorando maiúsculas/minúsculas e espaços
  sobrando). Decisão fora da lista é recusada e relatada, nunca "adivinhada".
- Decisão em branco = linha intocada (continua como falta injustificada por padrão, o critério conservador).
- Linhas repetidas com a mesma chave: se forem idênticas aplica-se uma vez; se divergirem, nenhuma é aplicada
  (não dá para saber qual vale) e o conflito é relatado.
- Chave que casa com mais de uma linha da fila (ex.: homônimos na mesma empresa) é recusada como ambígua.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from leitura_arquivos import ler_arquivo_generico
from validacao import ArquivoInvalidoError

NOME_ABA = "Exceções"
COLUNAS_OBRIGATORIAS = ("colaborador", "empresa", "data", "decisão")
COLUNA_OBSERVACAO = "observação"


@dataclass
class ResultadoImportacao:
    """Resultado da leitura da planilha. `decisoes` é indexado pelo índice da fila (nunca por posição)."""

    decisoes: pd.DataFrame = field(
        default_factory=lambda: pd.DataFrame(columns=["classificacao", "observacao"])
    )
    aplicadas: int = 0
    em_branco: int = 0
    sobrescritas: int = 0
    nao_encontradas: list[dict] = field(default_factory=list)
    invalidas: list[dict] = field(default_factory=list)
    duplicadas: list[dict] = field(default_factory=list)
    ambiguas: list[dict] = field(default_factory=list)

    @property
    def total_problemas(self) -> int:
        return len(self.nao_encontradas) + len(self.invalidas) + len(self.duplicadas) + len(self.ambiguas)


def _normalizar_texto(valor) -> str:
    """Maiúsculas, sem acento e com espaços colapsados — para casar nomes digitados de formas diferentes."""
    if valor is None or pd.isna(valor):
        return ""
    sem_acento = "".join(c for c in unicodedata.normalize("NFD", str(valor)) if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", sem_acento).strip().upper()


def _normalizar_decisao(valor) -> str:
    """Compara decisões ignorando caixa e espaços extras (mas respeitando acentos e o travessão)."""
    if valor is None or pd.isna(valor):
        return ""
    return re.sub(r"\s+", " ", str(valor).replace("\xa0", " ")).strip().casefold()


def _normalizar_data(valor) -> pd.Timestamp | None:
    """Aceita datetime do Excel e texto dd/mm/aaaa (também dd-mm-aaaa e aaaa-mm-dd). None se não for data."""
    if valor is None or pd.isna(valor):
        return None
    if isinstance(valor, (pd.Timestamp,)) or hasattr(valor, "year"):
        return pd.Timestamp(valor).normalize()
    texto = str(valor).strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}.*", texto):
        data = pd.to_datetime(texto, errors="coerce")
    else:
        data = pd.to_datetime(texto, format="%d/%m/%Y", errors="coerce")
        if pd.isna(data):
            data = pd.to_datetime(texto.replace("-", "/"), format="%d/%m/%Y", errors="coerce")
    return None if pd.isna(data) else data.normalize()


def _texto_observacao(valor) -> str:
    return "" if valor is None or pd.isna(valor) else str(valor).strip()


def _ler_aba_excecoes(arquivo: str | Path) -> pd.DataFrame:
    caminho = Path(arquivo)
    nome_amigavel = "revisão de exceções preenchida"
    if caminho.suffix.lower() != ".xlsx":
        raise ArquivoInvalidoError(
            f"A planilha preenchida ('{caminho.name}') precisa ser um arquivo .xlsx exportado pelo sistema."
        )
    try:
        abas = pd.ExcelFile(caminho).sheet_names
    except FileNotFoundError:
        abas = None  # deixa ler_arquivo_generico dar a mensagem padrão de "não encontrado"
    except Exception:
        abas = None  # arquivo corrompido: a mensagem amigável vem de ler_arquivo_generico
    if abas is not None:
        aba = next((a for a in abas if _normalizar_texto(a) == _normalizar_texto(NOME_ABA)), None)
        if aba is None:
            raise ArquivoInvalidoError(
                f"A planilha enviada não tem a aba \"{NOME_ABA}\" (abas encontradas: {', '.join(abas) or 'nenhuma'}). "
                "Envie o arquivo exportado pelo botão de download da etapa 4, preenchido na coluna Decisão."
            )
    else:
        aba = NOME_ABA
    df = ler_arquivo_generico(str(caminho), nome_amigavel, sheet_name=aba)
    faltando = [c for c in COLUNAS_OBRIGATORIAS if c not in df.columns]
    if faltando:
        raise ArquivoInvalidoError(
            f"A aba \"{NOME_ABA}\" não tem a(s) coluna(s) esperada(s): {', '.join(c.capitalize() for c in faltando)}. "
            "Não altere nem apague os cabeçalhos da planilha exportada."
        )
    return df


def importar_decisoes_excecoes_xlsx(
    arquivo: str | Path, fila: pd.DataFrame, opcoes_decisao: list[str]
) -> ResultadoImportacao:
    """
    Lê a aba "Exceções" da planilha preenchida e casa cada linha com `fila` (colunas nome_colaborador,
    razao_social, data, classificacao). `opcoes_decisao` são os rótulos válidos da coluna Decisão
    (o "" é ignorado). Não altera `fila`: devolve `ResultadoImportacao` com as decisões indexadas pelo
    índice da fila e o resumo (aplicadas, em branco, sobrescritas, não encontradas, inválidas, duplicadas).
    Levanta `ArquivoInvalidoError` (em português) se a aba ou as colunas esperadas estiverem ausentes.
    """
    df = _ler_aba_excecoes(arquivo)
    validas = {_normalizar_decisao(o): o for o in opcoes_decisao if o}

    # Índice de busca da fila: (colaborador, empresa, data) -> índices da fila
    chaves_fila: dict[tuple, list] = {}
    datas_fila = pd.to_datetime(fila["data"], errors="coerce")
    for idx, nome, empresa, data in zip(fila.index, fila["nome_colaborador"], fila["razao_social"], datas_fila):
        if pd.isna(data):
            continue
        chaves_fila.setdefault((_normalizar_texto(nome), _normalizar_texto(empresa), data.normalize()), []).append(idx)

    resultado = ResultadoImportacao()
    candidatas: dict[tuple, list[dict]] = {}  # chave -> linhas válidas (para detectar repetidas)

    for _, linha in df.iterrows():
        colaborador = "" if pd.isna(linha["colaborador"]) else str(linha["colaborador"]).strip()
        empresa = "" if pd.isna(linha["empresa"]) else str(linha["empresa"]).strip()
        if not colaborador and not empresa and pd.isna(linha["data"]) and pd.isna(linha["decisão"]):
            continue  # linha totalmente vazia no fim da planilha
        data = _normalizar_data(linha["data"])
        info = {
            "colaborador": colaborador,
            "empresa": empresa,
            "data": data.strftime("%d/%m/%Y") if data is not None else str(linha["data"]),
        }

        decisao_bruta = _normalizar_decisao(linha["decisão"])
        if not decisao_bruta:
            resultado.em_branco += 1
            continue
        if decisao_bruta not in validas:
            resultado.invalidas.append({**info, "motivo": f"decisão \"{str(linha['decisão']).strip()}\" não é uma opção válida"})
            continue
        if data is None:
            resultado.invalidas.append({**info, "motivo": "data ilegível"})
            continue

        chave = (_normalizar_texto(colaborador), _normalizar_texto(empresa), data)
        candidatas.setdefault(chave, []).append({
            **info,
            "classificacao": validas[decisao_bruta],
            "observacao": _texto_observacao(linha[COLUNA_OBSERVACAO]) if COLUNA_OBSERVACAO in df.columns else "",
        })

    linhas_aplicar: dict = {}
    for chave, itens in candidatas.items():
        indices = chaves_fila.get(chave, [])
        if not indices:
            resultado.nao_encontradas.append({**itens[0], "motivo": "não está na fila de exceções"})
            continue
        if len(indices) > 1:
            resultado.ambiguas.append({**itens[0], "motivo": "mais de uma linha da fila com o mesmo colaborador, empresa e data"})
            continue
        distintos = {(i["classificacao"], i["observacao"]) for i in itens}
        if len(itens) > 1:
            conflito = len(distintos) > 1
            resultado.duplicadas.append({
                **itens[0],
                "motivo": (f"{len(itens)} linhas repetidas com decisões diferentes — nenhuma foi aplicada"
                           if conflito else f"{len(itens)} linhas repetidas e idênticas — aplicada uma vez"),
            })
            if conflito:
                continue
        linhas_aplicar[indices[0]] = itens[0]

    if linhas_aplicar:
        resultado.decisoes = pd.DataFrame(
            {"classificacao": [v["classificacao"] for v in linhas_aplicar.values()],
             "observacao": [v["observacao"] for v in linhas_aplicar.values()]},
            index=pd.Index(list(linhas_aplicar), name=fila.index.name),
        )
        atuais = fila.loc[resultado.decisoes.index, "classificacao"].fillna("").astype(str)
        resultado.sobrescritas = int(((atuais != "") & (atuais != resultado.decisoes["classificacao"])).sum())
    resultado.aplicadas = len(resultado.decisoes)
    return resultado
