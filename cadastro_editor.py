"""
cadastro_editor.py
==================
Cadastro de Colaboradores editavel dentro do sistema. Concentra, fora da tela, tudo o que a tela
`app_web.py` precisa para tratar a planilha da cliente como proposta e nao como verdade absoluta:

* `ler_cadastro_para_edicao` - leitor TOLERANTE: reconhece os cabecalhos reais da cliente (P/Q/R), escolhe
  o vale (Alelo/Ticket) pela empresa, descarta linhas-lixo AVISANDO, normaliza tipos e, so' na importacao,
  le as cores de preenchimento (verde forte = valor fixo; verde claro = desconto manual do vale);
* `validar_cadastro` - validacao por linha (erro bloqueia o fechamento; aviso nao);
* `cadastro_para_pipeline` - garante os tipos que o restante do codigo espera;
* `cadastro_para_xlsx_bytes` / JSON - planilha limpa para baixar e formato de gravacao no SQLite;
* `comparar_cadastros` / `aplicar_aceites` - diferencas entre o cadastro salvo e uma planilha reenviada.

Depois da importacao as colunas explicitas (`va_fixo`, `cooper_fixo`, `va_desconto_manual`) sao a verdade:
a cor nunca mais e' consultada. O arquivo original nunca e' alterado.
Dados de folha sao sensiveis: mensagens citam matricula, nunca nome/PIS/valor de pessoa real em logs/testes.
"""

from __future__ import annotations

import hashlib
import io
import json
import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime, time
from pathlib import Path

import numpy as np
import pandas as pd

from leitura_arquivos import ler_arquivo_generico
from validacao import (
    ArquivoInvalidoError, canonicalizar_pis, converter_valor_monetario, ler_sim_nao, sem_acento_minusculo,
)

# ---------------------------------------------------------------------------------------------- colunas
SECOES = {
    "Identificação": ["matricula", "nome", "pis", "codigo_empresa", "razao_social", "cnpj", "centro_custo"],
    "Jornada": ["entrada", "saida_almoco", "retorno_almoco", "saida", "saida_sexta", "dias_trabalho", "admissao"],
    "Benefícios": ["beneficio", "valor_va", "va_fixo", "va_desconto_manual", "cooper", "cooper_fixo", "desconto_folha"],
    "Opcionais": ["banco_horas", "registra_ponto", "valor_dia_dsr", "observacao", "nota_rodape"],
}
COLUNAS_CANONICAS = [c for cols in SECOES.values() for c in cols]
COLUNAS_HORA = ["entrada", "saida_almoco", "retorno_almoco", "saida", "saida_sexta"]
COLUNAS_DINHEIRO = ["valor_va", "cooper", "desconto_folha", "valor_dia_dsr"]
COLUNAS_SIM_NAO = {"va_fixo": False, "cooper_fixo": False, "va_desconto_manual": False,
                   "banco_horas": False, "registra_ponto": True}
COLUNAS_TEXTO = ["matricula", "nome", "pis", "codigo_empresa", "razao_social", "cnpj", "centro_custo",
                 "dias_trabalho", "beneficio", "observacao", "nota_rodape"]
COLUNAS_OBRIGATORIAS_JORNADA = ["entrada", "saida_almoco", "retorno_almoco", "saida"]

BENEFICIO_ALELO = "Alelo Cartão"
BENEFICIO_TICKET = "Ticket"
BENEFICIOS_PADRAO = [BENEFICIO_ALELO, BENEFICIO_TICKET]
# Combinado com a cliente: coluna "Alelo Alimentacao" vale para a empresa 001, "Ticket Alimentacao" para a 002.
EMPRESA_ALELO, EMPRESA_TICKET = "001", "002"

_COL_VALE_ALELO, _COL_VALE_TICKET = "_vale_alelo", "_vale_ticket"
_ALIAS = {
    "cooper card assiduidade": "cooper", "cooper card": "cooper",
    "alelo alimentacao": _COL_VALE_ALELO, "alelo": _COL_VALE_ALELO,
    "ticket alimentacao": _COL_VALE_TICKET, "ticket": _COL_VALE_TICKET,
    "centro de custo": "centro_custo", "data admissao": "admissao", "data de admissao": "admissao",
}

_COR_VERDE_FORTE = "00B050"
SITUACAO_OK, SITUACAO_AVISO, SITUACAO_ERRO = "OK", "Aviso", "Erro"


@dataclass(frozen=True)
class Problema:
    """Uma pendencia do cadastro. `severidade`: 'erro' (bloqueia o fechamento) ou 'aviso'."""
    matricula: str
    campo: str
    severidade: str
    mensagem: str


@dataclass
class ResultadoCarga:
    df: pd.DataFrame
    avisos: list[str] = field(default_factory=list)
    colunas_ausentes: list[str] = field(default_factory=list)   # colunas canonicas que a planilha nao trazia
    linhas_planilha: dict[str, int] = field(default_factory=dict)   # matricula -> linha do Excel (1 = cabecalho)


@dataclass(frozen=True)
class Diferenca:
    """Diferenca entre o cadastro salvo e a planilha reenviada. `tipo`: 'novo', 'removido' ou 'alterado'."""
    matricula: str
    tipo: str
    campo: str          # '' para novo/removido (a linha inteira)
    valor_salvo: str
    valor_novo: str

    @property
    def chave(self) -> tuple[str, str]:
        return (self.matricula, self.campo)


# ---------------------------------------------------------------------------------------------- normalizacao
def _vazio(v) -> bool:
    if v is None or v is pd.NaT:
        return True
    if isinstance(v, (float, np.floating)) and np.isnan(v):
        return True
    return isinstance(v, str) and v.strip() == ""


def _texto(v) -> str:
    """Texto como digitado; float inteiro (14.0) vira '14'. Vazio vira ''."""
    if _vazio(v):
        return ""
    if isinstance(v, (float, np.floating)) and float(v).is_integer():
        return str(int(v))
    if isinstance(v, (bool, np.bool_)):
        return str(v)
    t = str(v).strip()
    return re.sub(r"^(\d+)\.0+$", r"\1", t)


def _hora(v) -> str:
    """Hora como 'HH:MM'; texto que nao seja hora fica como digitado (a validacao aponta)."""
    if _vazio(v):
        return ""
    if isinstance(v, (datetime, pd.Timestamp)):
        return f"{v.hour:02d}:{v.minute:02d}"
    if isinstance(v, time):
        return f"{v.hour:02d}:{v.minute:02d}"
    t = str(v).strip()
    m = re.fullmatch(r"(\d{1,2}):(\d{2})(?::\d{2})?", t)
    if m and int(m.group(1)) < 24 and int(m.group(2)) < 60:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return t


def _data(v) -> date | None:
    if _vazio(v):
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.date()
    if isinstance(v, date):
        return v
    t = str(v).strip()
    for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(t[:10] if formato == "%Y-%m-%d" else t, formato).date()
        except ValueError:
            continue
    return None


def _dinheiro(v):
    """float; vazio vira NaN; texto que nao seja valor em R$ fica como digitado (a validacao aponta)."""
    if _vazio(v):
        return np.nan
    try:
        r = converter_valor_monetario(v, "cadastro", "valor")
    except ArquivoInvalidoError:
        return str(v).strip()
    return np.nan if r is None else float(r)


def _sim_nao(serie: pd.Series, padrao: bool) -> pd.Series:
    def um(v):
        try:
            r = ler_sim_nao(v, "cadastro", "coluna", padrao)
        except ArquivoInvalidoError:
            return padrao
        return padrao if r is None else bool(r)

    return pd.Series([um(v) for v in serie], index=serie.index, dtype=bool)


def _codigo_empresa_chave(codigo: str) -> str:
    return codigo.zfill(3) if codigo.isdigit() else codigo


def normalizar_para_edicao(df: pd.DataFrame) -> pd.DataFrame:
    """Devolve o cadastro com todas as colunas canonicas, na ordem das secoes e com tipos estaveis."""
    out = pd.DataFrame(index=range(len(df)))
    origem = df.reset_index(drop=True)
    for col in COLUNAS_CANONICAS:
        serie = origem[col] if col in origem.columns else pd.Series([None] * len(origem), dtype=object)
        if col in COLUNAS_TEXTO:
            out[col] = [_texto(v) for v in serie]
        elif col in COLUNAS_HORA:
            out[col] = [_hora(v) for v in serie]
        elif col == "admissao":
            out[col] = pd.Series([_data(v) for v in serie], dtype=object)
        elif col in COLUNAS_DINHEIRO:
            valores = [_dinheiro(v) for v in serie]
            out[col] = pd.Series(valores, dtype=float if all(not isinstance(v, str) for v in valores) else object)
        else:
            out[col] = _sim_nao(serie, COLUNAS_SIM_NAO[col]) if col in origem.columns else COLUNAS_SIM_NAO[col]
    for col in COLUNAS_SIM_NAO:
        out[col] = out[col].astype(bool)
    return out


# ---------------------------------------------------------------------------------------------- carregador
def _chave_cabecalho(nome) -> str:
    chave = re.sub(r"\s+", " ", sem_acento_minusculo(nome))
    if chave in _ALIAS:
        return _ALIAS[chave]
    return chave.replace(" ", "_")


def normalizar_cabecalhos(colunas) -> list[str]:
    """Nomes canonicos: sem acento/maiuscula/espaco nas pontas; P/Q/R da cliente mapeados por alias."""
    return [_chave_cabecalho(c) for c in colunas]


def _faixas(linhas: list[int]) -> str:
    faixas, inicio, anterior = [], None, None
    for n in sorted(linhas):
        if inicio is None:
            inicio = anterior = n
        elif n == anterior + 1:
            anterior = n
        else:
            faixas.append((inicio, anterior))
            inicio = anterior = n
    if inicio is not None:
        faixas.append((inicio, anterior))
    return ", ".join(f"{a} a {b}" if a != b else str(a) for a, b in faixas)


def _ler_cores(caminho: str) -> dict[int, dict[str, bool]]:
    """
    So' na importacao: cor de preenchimento por linha do Excel -> flags. Verde forte (FF00B050) na coluna do
    Cooper = `cooper_fixo`; na coluna do vale = `va_fixo`. Verde claro (tema 9, tint ~0,8) em qualquer celula
    da linha = `va_desconto_manual`.
    """
    from openpyxl import load_workbook

    ws = load_workbook(caminho, data_only=True).worksheets[0]
    cab = normalizar_cabecalhos([c.value for c in ws[1]])
    cores: dict[int, dict[str, bool]] = {}
    for linha in ws.iter_rows(min_row=2):
        marcas = {"cooper_fixo": False, "va_fixo": False, "va_desconto_manual": False}
        for celula in linha:
            if celula.fill is None or celula.fill.fill_type != "solid":
                continue
            cor = celula.fill.fgColor
            nome_col = cab[celula.column - 1] if celula.column - 1 < len(cab) else ""
            if cor.type == "rgb" and isinstance(cor.rgb, str) and cor.rgb.upper().endswith(_COR_VERDE_FORTE):
                if nome_col == "cooper":
                    marcas["cooper_fixo"] = True
                elif nome_col in ("valor_va", _COL_VALE_ALELO, _COL_VALE_TICKET):
                    marcas["va_fixo"] = True
            elif cor.type == "theme" and cor.theme == 9 and abs((cor.tint or 0) - 0.8) < 0.05:
                marcas["va_desconto_manual"] = True
        if any(marcas.values()):
            cores[linha[0].row] = marcas
    return cores


def _ha_preenchimento(caminho: str) -> bool:
    """True se a planilha tem ao menos uma celula com cor de preenchimento (qualquer cor)."""
    from openpyxl import load_workbook

    ws = load_workbook(caminho, data_only=True).worksheets[0]
    return any(c.fill is not None and c.fill.fill_type == "solid" for linha in ws.iter_rows() for c in linha)


_DATA_BARRAS = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{2}|\d{4})")


def _interpretar_admissoes_em_texto(bruto: pd.DataFrame, avisos: list[str]) -> None:
    """
    Admissao digitada como TEXTO com barras ('3/13/26') pode ser dd/mm/aa ou mm/dd/aa (planilhas salvas por
    ferramentas em ingles). A ordem e' descoberta pela propria coluna: um primeiro numero acima de 12 so' pode ser
    dia; um segundo acima de 12 so' pode ser dia no formato mm/dd. Datas com ano de 2 digitos e de 4 digitos vem de
    digitacoes diferentes e sao avaliadas em grupos separados. Sem nenhuma evidencia vale dd/mm (padrao do
    sistema). Datas que continuarem ambiguas por evidencias conflitantes ficam em branco, com aviso: nunca se
    adivinha uma admissao, que muda o proporcional do vale e as faltas ignoradas.
    """
    if "admissao" not in bruto.columns:
        return
    grupos: dict[int, dict[int, re.Match]] = {}
    for i, v in bruto["admissao"].items():
        m = _DATA_BARRAS.fullmatch(v.strip()) if isinstance(v, str) else None
        if m:
            grupos.setdefault(len(m.group(3)), {})[i] = m
    convertidas = 0
    em_branco: list[str] = []
    for partes in grupos.values():
        ha_dia_primeiro = any(int(m.group(1)) > 12 for m in partes.values())
        ha_mes_primeiro = any(int(m.group(2)) > 12 for m in partes.values())
        mes_primeiro = ha_mes_primeiro and not ha_dia_primeiro
        misto = ha_dia_primeiro and ha_mes_primeiro
        for i, m in partes.items():
            a, b, ano = int(m.group(1)), int(m.group(2)), int(m.group(3))
            ano += 2000 if ano < 100 else 0
            if misto:      # so' as que a propria data desambigua
                dia, mes = (a, b) if a > 12 and b <= 12 else (b, a) if b > 12 and a <= 12 else (None, None)
            else:
                dia, mes = (b, a) if mes_primeiro else (a, b)
            try:
                bruto.at[i, "admissao"] = date(ano, mes, dia) if dia else None
            except (TypeError, ValueError):
                bruto.at[i, "admissao"], dia = None, None
            if dia is None:
                em_branco.append(_texto(bruto.at[i, "matricula"]) or f"da linha {bruto.at[i, '_linha_excel']}")
            elif mes_primeiro:
                convertidas += 1
    if convertidas:
        avisos.append(
            f"As datas de admissão em texto com ano de 2 dígitos estão no formato mês/dia/ano (ex.: '3/13/26'). Foram "
            f"lidas assim ({convertidas} datas). Confira as admissões na tela antes de fechar."
        )
    if em_branco:
        avisos.append(
            "Há datas de admissão em texto que misturam dia/mês e mês/dia e não puderam ser distinguidas; ficaram em "
            f"branco: matrículas {', '.join(em_branco)}. Preencha-as na tela."
        )


def _escolher_vale(df: pd.DataFrame) -> pd.DataFrame:
    """Une as colunas Alelo/Ticket em `valor_va` + `beneficio`, conforme a empresa de cada linha."""
    df = df.copy()
    for col in ("valor_va", "beneficio", _COL_VALE_ALELO, _COL_VALE_TICKET):
        if col not in df.columns:
            df[col] = None
    valores, rotulos = [], []
    for _, linha in df.iterrows():
        cod = _codigo_empresa_chave(_texto(linha.get("codigo_empresa")))
        rotulo_empresa = {EMPRESA_ALELO: BENEFICIO_ALELO, EMPRESA_TICKET: BENEFICIO_TICKET}.get(cod)
        pares = [(_COL_VALE_ALELO, BENEFICIO_ALELO), (_COL_VALE_TICKET, BENEFICIO_TICKET)]
        if cod == EMPRESA_TICKET:
            pares.reverse()
        valor, rotulo = linha["valor_va"], _texto(linha["beneficio"]) or rotulo_empresa
        if _vazio(valor):
            valor = None
            for coluna, nome_beneficio in pares:
                if not _vazio(linha[coluna]) and (rotulo_empresa is None or nome_beneficio == rotulo_empresa):
                    valor, rotulo = linha[coluna], _texto(linha["beneficio"]) or nome_beneficio
                    break
            else:   # empresa fora do combinado: usa o primeiro vale preenchido
                for coluna, nome_beneficio in pares:
                    if not _vazio(linha[coluna]):
                        valor, rotulo = linha[coluna], _texto(linha["beneficio"]) or nome_beneficio
                        break
        valores.append(valor)
        rotulos.append(rotulo)
    df["valor_va"], df["beneficio"] = valores, rotulos
    return df.drop(columns=[_COL_VALE_ALELO, _COL_VALE_TICKET])


def ler_cadastro_para_edicao(caminho: str) -> ResultadoCarga:
    """
    Le a planilha do cadastro sem abortar por problemas de preenchimento (isso e' papel de `validar_cadastro`).
    Levanta `ArquivoInvalidoError` so' se o arquivo for ilegivel ou nao tiver nem matricula nem PIS.
    """
    bruto = ler_arquivo_generico(caminho, "Cadastro de Colaboradores", dtype=object)
    avisos: list[str] = []
    bruto.columns = normalizar_cabecalhos(bruto.columns)
    bruto = bruto.loc[:, ~pd.Index(bruto.columns).duplicated()]
    if "matricula" not in bruto.columns and "pis" not in bruto.columns:
        raise ArquivoInvalidoError(
            f"O Cadastro de Colaboradores ('{caminho}') não tem a coluna 'matricula' nem 'pis'. "
            f"Colunas encontradas: {list(bruto.columns)}."
        )
    bruto["_linha_excel"] = bruto.index + 2

    outras = [c for c in bruto.columns if c not in COLUNAS_CANONICAS and c not in (_COL_VALE_ALELO, _COL_VALE_TICKET, "_linha_excel")
              and not c.startswith("unnamed")]
    if outras:
        avisos.append(f"Colunas não reconhecidas, ignoradas: {', '.join(outras)}.")

    for col in ("matricula", "pis"):
        if col not in bruto.columns:
            bruto[col] = None
    sem_id = bruto["matricula"].map(_vazio) & bruto["pis"].map(_vazio)
    campos_dados = [c for c in ("codigo_empresa", "razao_social", *COLUNAS_OBRIGATORIAS_JORNADA, "dias_trabalho") if c in bruto.columns]
    sem_dados = bruto["pis"].map(_vazio) & (
        bruto[campos_dados].apply(lambda col: col.map(_vazio)).all(axis=1) if campos_dados else True
    )
    lixo = sem_id | sem_dados
    if lixo.any():
        linhas = [int(n) for n in bruto.loc[lixo, "_linha_excel"]]
        avisos.append(
            f"{len(linhas)} linhas ignoradas: sem matrícula/PIS ou sem nenhum dado de colaborador "
            f"(legendas, anotações, linhas em branco), linhas {_faixas(linhas)} da planilha."
        )
    bruto = bruto.loc[~lixo].reset_index(drop=True)

    for col in ("banco_horas", "registra_ponto", "valor_dia_dsr"):
        if col not in bruto.columns:
            avisos.append(
                f"A planilha não tem a coluna '{col}'; foi criada com o valor padrão "
                f"({'sim' if COLUNAS_SIM_NAO.get(col, False) else 'não' if col in COLUNAS_SIM_NAO else 'vazio'})."
            )
    ausentes = [c for c in ("razao_social", "codigo_empresa", *COLUNAS_OBRIGATORIAS_JORNADA, "dias_trabalho", "nome")
                if c not in bruto.columns]
    if ausentes:
        avisos.append(f"Colunas obrigatórias ausentes na planilha (criadas vazias para preenchimento): {', '.join(ausentes)}.")

    for col in COLUNAS_SIM_NAO:
        if col in bruto.columns:
            for i, v in bruto[col].items():
                try:
                    ler_sim_nao(v, "cadastro", col, None)
                except ArquivoInvalidoError:
                    avisos.append(f"Matrícula {_texto(bruto.at[i, 'matricula']) or '(vazia)'}: valor '{v}' em '{col}' "
                                  "não é sim/não; considerado o valor padrão.")
    colunas_ausentes = [c for c in COLUNAS_CANONICAS if c not in bruto.columns and c not in ("beneficio", "valor_va", "cooper")]
    _interpretar_admissoes_em_texto(bruto, avisos)
    bruto = _escolher_vale(bruto)
    df = normalizar_para_edicao(bruto)

    if str(caminho).lower().endswith(".xlsx"):
        try:
            cores = _ler_cores(caminho)
            if not cores and not _ha_preenchimento(caminho):
                avisos.append(
                    "A planilha não traz nenhuma cor de preenchimento. Se a cliente marcava em verde os valores fixos "
                    "do vale/Cooper e o desconto manual do vale, essa informação não veio no arquivo: 'va_fixo', "
                    "'cooper_fixo' e 'va_desconto_manual' ficaram como 'não'. Confira essas colunas na tela "
                    "(planilha reenviada sem cores pode ter sido salva por outra ferramenta)."
                )
            for i, linha_excel in enumerate(bruto["_linha_excel"]):
                for flag, valor in cores.get(int(linha_excel), {}).items():
                    if valor:
                        df.at[i, flag] = True
        except Exception as e:
            avisos.append(
                f"Não foi possível ler as cores de preenchimento da planilha ({e}); 'va_fixo', 'cooper_fixo' e "
                "'va_desconto_manual' ficaram como 'não'. Marque-os na tela, se for o caso."
            )

    for col in COLUNAS_DINHEIRO:
        invalidos = df[col].map(lambda v: isinstance(v, str))
        for i in df.index[invalidos]:
            avisos.append(f"Matrícula {df.at[i, 'matricula'] or '(vazia)'}: valor '{df.at[i, col]}' em '{col}' não é um valor em R$; "
                          "deixado em branco para correção.")
            df.at[i, col] = np.nan
        df[col] = pd.to_numeric(df[col], errors="coerce")
    # Admissao que o leitor nao conseguiu interpretar: avisa (a celula fica vazia).
    for i, valor in enumerate(bruto["admissao"] if "admissao" in bruto.columns else []):
        if not _vazio(valor) and df.at[i, "admissao"] is None:
            avisos.append(f"Matrícula {df.at[i, 'matricula'] or '(vazia)'}: admissão '{valor}' não está em dd/mm/aaaa; deixada em branco.")
    linhas_planilha = {m: int(n) for m, n in zip(df["matricula"], bruto["_linha_excel"]) if m}
    return ResultadoCarga(df=df, avisos=avisos, colunas_ausentes=colunas_ausentes, linhas_planilha=linhas_planilha)


# ---------------------------------------------------------------------------------------------- validacao
def tem_erros(problemas: list[Problema]) -> bool:
    return any(p.severidade == "erro" for p in problemas)


def _cnpj_valido(cnpj: str) -> bool:
    d = re.sub(r"\D", "", cnpj)
    if len(d) != 14 or len(set(d)) == 1:
        return False

    def dv(base: str) -> str:
        pesos = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2] if len(base) == 12 else [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        resto = sum(int(a) * b for a, b in zip(base, pesos)) % 11
        return str(0 if resto < 2 else 11 - resto)

    return d[12] == dv(d[:12]) and d[13] == dv(d[:13])


def _formatar_cnpj(d: str) -> str:
    return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"


def _hora_valida(v: str) -> bool:
    return bool(re.fullmatch(r"([01]?\d|2[0-3]):[0-5]\d", v))


def validar_cadastro(df: pd.DataFrame, data_fim: date | None = None,
                     linhas_planilha: dict[str, int] | None = None) -> list[Problema]:
    """
    Todas as pendencias do cadastro, erros primeiro. `data_fim` (fim do periodo) habilita o aviso de admissao.
    `linhas_planilha` (matricula -> linha do Excel, de `ResultadoCarga`) faz a mensagem citar tambem a linha da
    planilha enviada; colaboradores acrescentados na tela nao tem linha e ficam so' com a matricula.
    """
    d = df.reset_index(drop=True)
    problemas: list[Problema] = []

    def rotulo(i: int) -> str:
        mat = _texto(d.at[i, "matricula"]) if "matricula" in d.columns else ""
        return mat or "(sem matrícula)"

    def add(i: int, campo: str, sev: str, msg: str) -> None:
        linha = (linhas_planilha or {}).get(rotulo(i))
        onde = f" (linha {linha} da planilha)" if linha else ""
        problemas.append(Problema(rotulo(i), campo, sev, f"Matrícula {rotulo(i)}{onde}: {msg}"))

    def col(nome: str):
        return d[nome] if nome in d.columns else pd.Series([None] * len(d), dtype=object)

    pis_canon = [canonicalizar_pis(v) if not _vazio(v) else "" for v in col("pis")]
    contagem_mat = Counter(_texto(v) for v in col("matricula") if _texto(v))
    contagem_pis = Counter(p for p in pis_canon if p)

    cnpjs_validos_por_empresa: dict[str, set[str]] = {}
    for i in range(len(d)):
        cnpj = re.sub(r"\D", "", _texto(col("cnpj").iloc[i]))
        if cnpj and _cnpj_valido(cnpj):
            cnpjs_validos_por_empresa.setdefault(_texto(col("codigo_empresa").iloc[i]), set()).add(cnpj)

    padroes: dict[str, list[tuple]] = {}
    for i in range(len(d)):
        padrao = tuple(_hora(col(c).iloc[i]) for c in ("entrada", "saida_almoco", "retorno_almoco"))
        if all(_hora_valida(p) for p in padrao):
            padroes.setdefault(_texto(col("codigo_empresa").iloc[i]), []).append(padrao)
    padrao_maioria: dict[str, tuple] = {}
    for emp, lista in padroes.items():
        padrao, qtd = Counter(lista).most_common(1)[0]
        if len(lista) >= 3 and qtd > len(lista) / 2:
            padrao_maioria[emp] = padrao

    for i in range(len(d)):
        mat, pis_i = _texto(col("matricula").iloc[i]), pis_canon[i]
        if not mat:
            add(i, "matricula", "erro", "matrícula não preenchida.")
        elif contagem_mat[mat] > 1:
            add(i, "matricula", "erro", f"matrícula repetida ({contagem_mat[mat]} linhas).")
        if not pis_i:
            add(i, "pis", "erro", "PIS não preenchido.")
        elif contagem_pis[pis_i] > 1:
            add(i, "pis", "erro", f"PIS repetido em {contagem_pis[pis_i]} linhas (cada colaborador precisa de um PIS próprio).")
        elif mat and mat.isdigit() and pis_i == canonicalizar_pis(mat):
            add(i, "pis", "aviso",
                "o PIS é igual à matrícula com zeros à esquerda. Isso só funciona com o Cartão Ponto (PDF), que identifica "
                "por matrícula; com o arquivo AFD (PIS real) as batidas não serão encontradas.")

        for campo in ("razao_social", "codigo_empresa"):
            if not _texto(col(campo).iloc[i]):
                add(i, campo, "erro", f"'{campo}' não preenchido.")
        for campo in COLUNAS_OBRIGATORIAS_JORNADA:
            h = _hora(col(campo).iloc[i])
            if not h:
                add(i, campo, "erro", f"'{campo}' não preenchido.")
            elif not _hora_valida(h):
                add(i, campo, "erro", f"'{campo}' com hora inválida ('{h}'). Use hh:mm, por exemplo 08:00.")
        sexta = _hora(col("saida_sexta").iloc[i])
        dias = _texto(col("dias_trabalho").iloc[i]).replace(" ", "")
        if sexta and not _hora_valida(sexta):
            add(i, "saida_sexta", "erro", f"'saida_sexta' com hora inválida ('{sexta}').")
        if not dias:
            add(i, "dias_trabalho", "erro", "'dias_trabalho' não preenchido (use 0,1,2,3,4; 0 = segunda).")
        elif not re.fullmatch(r"[0-6](,[0-6])*", dias):
            add(i, "dias_trabalho", "erro", f"'dias_trabalho' fora do formato ('{dias}'). Use números de 0 a 6 separados por vírgula, "
                                            "por exemplo 0,1,2,3,4 (0 = segunda).")
        elif not sexta and "4" in dias.split(","):
            add(i, "saida_sexta", "aviso", "'saida_sexta' não preenchida; a sexta usará a mesma saída dos outros dias.")

        for campo in COLUNAS_DINHEIRO:
            v = col(campo).iloc[i]
            if isinstance(v, str) and v.strip():
                add(i, campo, "erro", f"'{campo}' não é um valor em R$ válido ('{v}').")
            elif not _vazio(v) and not isinstance(v, str) and float(v) < 0:
                add(i, campo, "erro", f"'{campo}' negativo.")

        adm_bruta = col("admissao").iloc[i]
        adm = _data(adm_bruta)
        if not _vazio(adm_bruta) and adm is None:
            add(i, "admissao", "erro", f"'admissao' inválida ('{adm_bruta}'). Use dd/mm/aaaa.")
        elif adm is not None and data_fim is not None and adm > data_fim:
            add(i, "admissao", "aviso", f"admissão em {adm:%d/%m/%Y}, depois do fim do período ({data_fim:%d/%m/%Y}); "
                                        "o colaborador não terá dias apurados neste fechamento.")

        cnpj = _texto(col("cnpj").iloc[i])
        if cnpj:
            digitos = re.sub(r"\D", "", cnpj)
            if not _cnpj_valido(digitos):
                msg = f"CNPJ '{cnpj}' com dígitos verificadores inválidos."
                sugestoes = sorted(
                    c for c in cnpjs_validos_por_empresa.get(_texto(col("codigo_empresa").iloc[i]), set())
                    if c[:8] == digitos[:8]
                )
                if sugestoes:
                    msg += f" Outro CNPJ da mesma empresa, com a mesma raiz, é válido: {_formatar_cnpj(sugestoes[0])} (possível correção)."
                add(i, "cnpj", "aviso", msg)

        emp = _texto(col("codigo_empresa").iloc[i])
        padrao = tuple(_hora(col(c).iloc[i]) for c in ("entrada", "saida_almoco", "retorno_almoco"))
        if emp in padrao_maioria and all(_hora_valida(p) for p in padrao) and padrao != padrao_maioria[emp]:
            add(i, "entrada", "aviso", "entrada/almoço diferentes da maioria da empresa "
                                       f"({' / '.join(padrao_maioria[emp])}). Confirme se está certo.")

    if "banco_horas" not in d.columns:
        problemas.append(Problema("(todas)", "banco_horas", "aviso",
                                  "A coluna 'banco_horas' não existe: ninguém será tratado como colaborador em banco de horas."))
    problemas.sort(key=lambda p: p.severidade != "erro")
    return problemas


def situacao_por_matricula(problemas: list[Problema]) -> dict[str, str]:
    """Matricula -> 'Erro' | 'Aviso' (matriculas sem problema nao aparecem; use .get(m, 'OK'))."""
    situacao: dict[str, str] = {}
    for p in problemas:
        if p.severidade == "erro":
            situacao[p.matricula] = SITUACAO_ERRO
        else:
            situacao.setdefault(p.matricula, SITUACAO_AVISO)
    return situacao


def com_situacao(df: pd.DataFrame, problemas: list[Problema]) -> pd.DataFrame:
    """Copia do cadastro com a coluna somente-leitura 'situacao' (OK / Aviso / Erro) na frente."""
    situacao = situacao_por_matricula(problemas)
    out = df.copy()
    out.insert(0, "situacao", [situacao.get(_texto(m) or "(sem matrícula)", SITUACAO_OK) for m in out["matricula"]])
    return out


# ---------------------------------------------------------------------------------------------- pipeline / saida
def cadastro_para_pipeline(df: pd.DataFrame) -> pd.DataFrame:
    """DataFrame com os tipos que o pipeline espera: PIS canonico, booleanos, floats e datas."""
    out = normalizar_para_edicao(df.drop(columns=["situacao"], errors="ignore"))
    out["pis"] = out["pis"].map(canonicalizar_pis)
    for col in COLUNAS_DINHEIRO:
        out[col] = pd.to_numeric(out[col], errors="coerce").astype(float)
    for col in COLUNAS_SIM_NAO:
        out[col] = out[col].astype(bool)
    return out


def cadastro_para_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """Planilha limpa, cabecalhos canonicos, so' os colaboradores (sem legendas). Pode ser reenviada ao sistema."""
    from openpyxl import Workbook

    d = normalizar_para_edicao(df.drop(columns=["situacao"], errors="ignore"))
    wb = Workbook()
    ws = wb.active
    ws.title = "Cadastro"
    ws.append(COLUNAS_CANONICAS)
    for _, linha in d.iterrows():
        valores = []
        for col in COLUNAS_CANONICAS:
            v = linha[col]
            if col in COLUNAS_SIM_NAO:
                v = "sim" if v else "não"
            elif col in COLUNAS_DINHEIRO:
                v = None if _vazio(v) else float(v)
            elif col == "admissao":
                v = v if isinstance(v, date) else None
            else:
                v = v if v != "" else None
            valores.append(v)
        ws.append(valores)
    coluna_adm = COLUNAS_CANONICAS.index("admissao") + 1
    for linha in ws.iter_rows(min_row=2, min_col=coluna_adm, max_col=coluna_adm):
        for celula in linha:
            celula.number_format = "DD/MM/YYYY"
    for col in ("centro_custo", "matricula", "pis", "codigo_empresa"):   # texto: preserva zeros a esquerda
        c = COLUNAS_CANONICAS.index(col) + 1
        for linha in ws.iter_rows(min_row=2, min_col=c, max_col=c):
            for celula in linha:
                celula.number_format = "@"
    ws.freeze_panes = "C2"
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def cadastro_para_json(df: pd.DataFrame) -> str:
    """Serializacao estavel (ordenada por matricula) usada na gravacao e no hash do conteudo."""
    d = normalizar_para_edicao(df.drop(columns=["situacao"], errors="ignore"))
    registros = []
    for _, linha in d.iterrows():
        reg = {}
        for col in COLUNAS_CANONICAS:
            v = linha[col]
            if isinstance(v, date):
                v = v.isoformat()
            elif isinstance(v, (np.bool_, bool)):
                v = bool(v)
            elif isinstance(v, (float, np.floating)):
                v = None if np.isnan(v) else float(v)
            reg[col] = v
        registros.append(reg)
    registros.sort(key=lambda r: (r["matricula"], r["pis"]))
    return json.dumps(registros, ensure_ascii=False, sort_keys=True)


def cadastro_de_json(texto: str) -> pd.DataFrame:
    return normalizar_para_edicao(pd.DataFrame(json.loads(texto), columns=COLUNAS_CANONICAS))


def hash_do_cadastro(df: pd.DataFrame) -> str:
    return hashlib.sha256(cadastro_para_json(df).encode("utf-8")).hexdigest()


# ---------------------------------------------------------------------------------------------- diferencas
def _canon(v) -> str:
    if _vazio(v):
        return ""
    if isinstance(v, (bool, np.bool_)):
        return "sim" if v else "não"
    if isinstance(v, date):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, (float, np.floating)):
        return f"{float(v):.2f}"
    return str(v)


def comparar_cadastros(salvo: pd.DataFrame, novo: pd.DataFrame) -> list[Diferenca]:
    """Por matricula: linhas novas, removidas e campos alterados do `novo` (planilha) em relacao ao `salvo` (sistema)."""
    a = normalizar_para_edicao(salvo.drop(columns=["situacao"], errors="ignore")).set_index("matricula", drop=False)
    b = normalizar_para_edicao(novo.drop(columns=["situacao"], errors="ignore")).set_index("matricula", drop=False)
    a, b = a[~a.index.duplicated()], b[~b.index.duplicated()]
    difs: list[Diferenca] = []
    for mat in b.index.difference(a.index):
        difs.append(Diferenca(mat, "novo", "", "", ""))
    for mat in a.index.difference(b.index):
        difs.append(Diferenca(mat, "removido", "", "", ""))
    for mat in a.index.intersection(b.index):
        for col in COLUNAS_CANONICAS:
            if col != "matricula" and _canon(a.at[mat, col]) != _canon(b.at[mat, col]):
                difs.append(Diferenca(mat, "alterado", col, _canon(a.at[mat, col]), _canon(b.at[mat, col])))
    return sorted(difs, key=lambda x: (x.tipo != "alterado", x.matricula, x.campo))


def aplicar_aceites(salvo: pd.DataFrame, novo: pd.DataFrame, aceitas: set[tuple[str, str]]) -> pd.DataFrame:
    """Aplica ao cadastro salvo apenas as diferencas aceitas (`Diferenca.chave`); o resto fica como esta no sistema."""
    a = normalizar_para_edicao(salvo.drop(columns=["situacao"], errors="ignore"))
    b = normalizar_para_edicao(novo.drop(columns=["situacao"], errors="ignore")).set_index("matricula", drop=False)
    for dif in comparar_cadastros(a, b.reset_index(drop=True)):
        if dif.chave not in aceitas:
            continue
        if dif.tipo == "novo":
            a = pd.concat([a, b.loc[[dif.matricula]].reset_index(drop=True)], ignore_index=True)
        elif dif.tipo == "removido":
            a = a[a["matricula"] != dif.matricula].reset_index(drop=True)
        else:
            a.loc[a["matricula"] == dif.matricula, dif.campo] = b.at[dif.matricula, dif.campo]
    return normalizar_para_edicao(a)


def reconciliar_com_salvo(enviado: pd.DataFrame, salvo: pd.DataFrame | None) -> tuple[pd.DataFrame, list[Diferenca]]:
    """
    Decisao padrao do reenvio: o cadastro editado no sistema e' a verdade; a planilha so' PROPOE mudancas.
    Devolve (cadastro de trabalho, diferencas a mostrar). Sem cadastro salvo, a planilha vira o cadastro.
    """
    if salvo is None or salvo.empty:
        return normalizar_para_edicao(enviado), []
    return normalizar_para_edicao(salvo), comparar_cadastros(salvo, enviado)


def situacao_por_linha(editado: pd.DataFrame, problemas: list[Problema]) -> pd.Series:
    """Situacao (OK/Aviso/Erro) de cada linha do cadastro editado, pelo indice da linha (coluna da tela)."""
    por_matricula = situacao_por_matricula(problemas)
    return pd.Series(
        [por_matricula.get(_texto(m) or "(sem matrícula)", SITUACAO_OK) for m in editado["matricula"]],
        index=editado.index, dtype=object,
    )
