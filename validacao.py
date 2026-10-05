"""
validacao.py
============
Validacao defensiva compartilhada pelos modulos que carregam arquivos de
entrada (cadastro, consignados, banco de horas). O objetivo e' nunca deixar
o Pandas estourar um `KeyError` cru no meio do processamento - em vez
disso, a ausencia de uma coluna obrigatoria e' detectada logo no
carregamento e vira uma mensagem clara, em portugues, dizendo exatamente
qual arquivo e qual coluna esta faltando.
"""

from __future__ import annotations

import re
import unicodedata
from datetime import datetime, time, timedelta

import numpy as np
import pandas as pd


class ArquivoInvalidoError(ValueError):
    """
    Erro amigavel levantado quando um arquivo de entrada esta ilegivel ou
    sem uma coluna obrigatoria. E' subclasse de ValueError de proposito -
    qualquer codigo que ja trata ValueError (como a interface grafica em
    app_gui.py) continua funcionando sem precisar de nenhum ajuste.
    """


TAMANHO_PIS = 11  # PIS/NIS/PASEP tem 11 digitos (a formatacao com pontos/traco e so' visual)


def canonicalizar_pis(valor) -> str:
    """
    Reduz qualquer formatacao de PIS (com pontos/traco, ou com zeros de
    preenchimento) a uma string de 11 digitos comparavel.

    Por que isso existe: o campo PIS do AFD (Portaria 671/2021) tem 12
    caracteres de largura fixa, preenchido com zero a esquerda -
    "016151387504" la' e' o mesmo PIS "16151387504" (11 digitos) do
    Cadastro de Colaboradores ou do Espelho de Ponto. Sem essa normalizacao,
    o AFD de producao nunca bate com ninguem no cadastro - toda a apuracao
    do ponto ficaria vazia silenciosamente (nenhum erro visivel, so'
    zero batida encontrada pra todo mundo).

    A conta e' "tira os zeros a esquerda, depois recoloca ate' completar
    11 digitos": isso remove exatamente o preenchimento de largura do AFD
    sem arriscar apagar um zero que faca parte de verdade do PIS (ex.: um
    PIS real "00123456789" sobrevive intacto, porque volta a ter 11
    digitos no final). Caracteres nao numericos (pontos, traco, espaco)
    sao descartados antes de mais nada, pra aceitar tambem PIS digitado
    com a mascara "123.45678.90-1".
    """
    digitos = re.sub(r"\D", "", str(valor))
    digitos = digitos.lstrip("0")
    return digitos.rjust(TAMANHO_PIS, "0")


def verificar_colunas(
    df: pd.DataFrame,
    colunas_obrigatorias: list[str],
    nome_amigavel_arquivo: str,
    caminho: str,
) -> None:
    """
    Confere se `df` tem todas as `colunas_obrigatorias`. Se faltar
    qualquer uma, interrompe o processo com uma mensagem clara em vez de
    deixar o erro estourar mais adiante, sem contexto, como um KeyError.
    """
    faltando = [c for c in colunas_obrigatorias if c not in df.columns]
    if faltando:
        raise ArquivoInvalidoError(
            f"O arquivo de {nome_amigavel_arquivo} ('{caminho}') está sem a(s) "
            f"coluna(s) obrigatória(s): {', '.join(faltando)}.\n"
            f"Colunas encontradas no arquivo: {list(df.columns)}.\n"
            "Verifique se este é o arquivo certo, ou se o nome das colunas "
            "bate com o esperado (a comparação diferencia maiúsculas de "
            "minúsculas e não ignora espaços extras)."
        )


def validar_pis_existem(
    pis_no_arquivo: list[str] | pd.Series,
    pis_validos: list[str] | pd.Series,
    nome_amigavel_arquivo: str,
    caminho: str,
) -> None:
    """
    Checagem pre-voo: confere que todo PIS/matrícula citado num arquivo
    opcional (banco de horas do Sidney, do Adriano, consignados) existe
    de fato no Cadastro de Colaboradores, ANTES de qualquer calculo
    pesado comecar. Um PIS digitado errado nesses arquivos hoje passaria
    batido (a linha simplesmente nunca casaria com ninguem no merge) -
    isso avisa a usuaria na hora, apontando exatamente qual PIS e' o
    problema, em vez de a folha de alguem sair sem o banco de horas dela
    silenciosamente.
    """
    encontrados = {str(p).strip() for p in pis_no_arquivo if str(p).strip()}
    validos = {str(p).strip() for p in pis_validos}
    desconhecidos = sorted(encontrados - validos)
    if desconhecidos:
        raise ArquivoInvalidoError(
            f"O arquivo de {nome_amigavel_arquivo} ('{caminho}') cita o(s) PIS/matrícula "
            f"{', '.join(desconhecidos)}, que não consta(m) no Cadastro de Colaboradores.\n"
            "Verifique se o PIS foi digitado corretamente nesse arquivo, ou se falta "
            "cadastrar esse colaborador."
        )


def normalizar_identificador_colaborador(
    df: pd.DataFrame,
    coluna_id: str,
    cadastro: pd.DataFrame,
    nome_amigavel_arquivo: str,
    caminho: str,
) -> pd.DataFrame:
    """
    Muitos arquivos auxiliares (consignados, banco de horas) não trazem o
    PIS — trazem a MATRÍCULA interna da empresa (um número curto, tipo
    "96" ou "97"), que é um identificador diferente do PIS (o número de
    11 dígitos usado como chave em todo o resto do pipeline). Isso é
    normal e legítimo: nem todo sistema de origem expõe o PIS.

    Sem este passo, um valor de matrícula cai na coluna que o detector
    genérico rotula como "pis" (ele procura qualquer coluna com "pis" OU
    "matricula" no nome) e nunca bate com o PIS de ninguém no cadastro —
    o sistema então acusa "PIS não encontrado" para gente que, na
    verdade, está cadastrada certinha, só que identificada por matrícula
    naquele arquivo específico.

    Esta função aceita as duas coisas: tenta casar cada valor primeiro
    como PIS e, se não achar, como matrícula; grava de volta sempre o
    PIS canônico do cadastro, para o resto do pipeline continuar
    trabalhando só com PIS. Só levanta erro se um valor não bater com
    PIS *nem* com matrícula de ninguém.
    """
    cadastro_pis = cadastro["pis"].astype(str).str.strip()
    mapa_matricula_para_pis = dict(zip(cadastro["matricula"].astype(str).str.strip(), cadastro_pis))
    pis_validos = set(cadastro_pis)

    def resolver(valor):
        v = str(valor).strip()
        if v in pis_validos:
            return v
        if v in mapa_matricula_para_pis:
            return mapa_matricula_para_pis[v]
        return None

    resolvidos = df[coluna_id].map(resolver)
    nao_resolvidos = sorted(df.loc[resolvidos.isna(), coluna_id].astype(str).str.strip().unique().tolist())
    if nao_resolvidos:
        raise ArquivoInvalidoError(
            f"O arquivo de {nome_amigavel_arquivo} ('{caminho}') cita o(s) identificador(es) "
            f"{', '.join(nao_resolvidos)}, que não bate(m) nem com o PIS nem com a matrícula "
            "de nenhum colaborador no Cadastro de Colaboradores.\n"
            "Verifique se o valor foi digitado corretamente nesse arquivo, ou se falta "
            "cadastrar esse colaborador."
        )

    df = df.copy()
    df[coluna_id] = resolvidos
    return df


# ---------------------------------------------------------------------------
# Conversores compartilhados (sim/nao, horas, valores em R$)
# ---------------------------------------------------------------------------
_SIM = {"sim", "s", "1", "true", "verdadeiro", "x"}
_NAO = {"nao", "n", "0", "false", "falso"}


def sem_acento_minusculo(valor) -> str:
    texto = unicodedata.normalize("NFKD", str(valor))
    return "".join(c for c in texto if not unicodedata.combining(c)).strip().lower()


def _texto_da_celula(valor) -> str:
    # 1.0 (Excel com celulas vazias na coluna vira float) tem que valer "1"
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor)


def ler_sim_nao(valor, onde: str, coluna: str, padrao: bool | None = None) -> bool | None:
    """
    Le uma celula sim/nao (sim/s/1/true/x e nao/n/0/false, sem diferenciar
    maiusculas nem acento). Vazio devolve `padrao`. Qualquer outra coisa levanta
    `ArquivoInvalidoError`; `onde` e' o trecho "arquivo ('caminho'), linha N" da mensagem.
    """
    if valor is None or pd.isna(valor) or str(valor).strip() == "":
        return padrao
    texto = sem_acento_minusculo(_texto_da_celula(valor))
    if texto in _SIM:
        return True
    if texto in _NAO:
        return False
    raise ArquivoInvalidoError(
        f"{onde}: o valor '{valor}' da coluna '{coluna}' não é válido. Use 'sim' ou 'não' (ou deixe vazio)."
    )


def _horas_de_valor(valor) -> float:
    """Converte UM valor em horas decimais (sem arredondar). Levanta ValueError se nao entender."""
    if isinstance(valor, (bool, np.bool_)):
        raise ValueError("valor lógico não é horas")
    if isinstance(valor, (pd.Timedelta, timedelta, np.timedelta64)):
        return pd.Timedelta(valor).total_seconds() / 3600
    if isinstance(valor, datetime):  # celula [h]:mm do Excel > 24h vem como data a partir de 30/12/1899
        return (valor - datetime(1899, 12, 30)).total_seconds() / 3600
    if isinstance(valor, time):
        return valor.hour + valor.minute / 60 + valor.second / 3600
    if isinstance(valor, (int, float, np.integer, np.floating)):
        return float(valor)

    texto = str(valor).strip().replace(" ", "")
    sinal = -1 if texto.startswith("-") else 1
    sem_sinal = texto[1:] if texto[:1] in ("+", "-") else texto
    if ":" in sem_sinal:
        partes = sem_sinal.split(":")
        if len(partes) not in (2, 3) or not all(p.isdigit() for p in partes) or int(partes[1]) > 59:
            raise ValueError("hh:mm inválido")
        horas, minutos = int(partes[0]), int(partes[1])
        segundos = int(partes[2]) if len(partes) == 3 else 0
        return sinal * (horas + minutos / 60 + segundos / 3600)
    if "," in sem_sinal:  # "1.234,5" (milhar com ponto) ou "5,5"
        sem_sinal = sem_sinal.replace(".", "").replace(",", ".")
    if not re.fullmatch(r"\d+(\.\d+)?", sem_sinal):
        raise ValueError("número inválido")
    return sinal * float(sem_sinal)


def converter_horas_decimais(valor, onde: str, coluna: str) -> float:
    """
    Horas decimais (2 casas, como as demais colunas de horas) a partir de numero
    (virgula ou ponto), "hh:mm" (inclusive negativo, "-05:10") ou dos
    datetime.time / timedelta que o pandas devolve para celulas de hora do Excel.
    Vazio vale 0. Valor invalido levanta `ArquivoInvalidoError` ("onde" traz arquivo e linha).
    """
    if valor is None or (not isinstance(valor, (time, timedelta, np.timedelta64)) and pd.isna(valor))             or str(valor).strip() == "":
        return 0.0
    try:
        return round(_horas_de_valor(valor), 2)
    except (ValueError, OverflowError):
        raise ArquivoInvalidoError(
            f"{onde}: o valor '{valor}' da coluna '{coluna}' não é um número de horas válido. "
            "Use um número (ex.: 5,5), horas e minutos (ex.: 05:30 ou -05:10) ou uma célula de hora do Excel."
        ) from None


def converter_valor_monetario(valor, onde: str, coluna: str) -> float | None:
    """Numero ou texto brasileiro ("62,50", "R$ 62,50", "1.234,56") -> float. Vazio -> None."""
    if valor is None or pd.isna(valor) or str(valor).strip() == "":
        return None
    if isinstance(valor, (int, float, np.integer, np.floating)) and not isinstance(valor, (bool, np.bool_)):
        return float(valor)
    texto = re.sub(r"^R\$\s*", "", str(valor).strip(), flags=re.IGNORECASE).replace(" ", "")
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    if not re.fullmatch(r"\d+(\.\d+)?", texto):
        raise ArquivoInvalidoError(
            f"{onde}: o valor '{valor}' da coluna '{coluna}' não é um valor em R$ válido. Use por exemplo 62,50."
        )
    return float(texto)
