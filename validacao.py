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

from typing import Iterable
import pandas as pd


class ArquivoInvalidoError(ValueError):
    """
    Erro amigavel levantado quando um arquivo de entrada esta ilegivel ou
    sem uma coluna obrigatoria. E' subclasse de ValueError de proposito -
    qualquer codigo que ja trata ValueError (como a interface grafica em
    app_gui.py) continua funcionando sem precisar de nenhum ajuste.
    """


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


def _normalizar_id(valor: object) -> str:
    """
    Função auxiliar para normalizar PIS/Matrícula:
    - Remove decimais flutuantes do Pandas (ex: '96.0' -> '96')
    - Remove espaços em branco
    - Remove zeros à esquerda (ex: '0096' -> '96')
    """
    if pd.isna(valor):
        return ""
    
    texto = str(valor).strip()
    
    # Caso venha como float convertido para string (ex: '96.0')
    if texto.endswith(".0"):
        texto = texto[:-2]
        
    # Remove zeros à esquerda mantendo pelo menos um '0' se o valor for exatamente zero
    texto_limpo = texto.lstrip("0")
    return texto_limpo if texto_limpo else ("0" if texto else "")


def validar_pis_existem(
    pis_no_arquivo: Iterable[object] | pd.Series,
    pis_validos: Iterable[object] | pd.Series,
    nome_amigavel_arquivo: str,
    caminho: str,
) -> None:
    """
    Checagem pre-voo: confere que todo PIS/matrícula citado num arquivo
    opcional (banco de horas do Sidney, do Adriano, consignados) existe
    de fato no Cadastro de Colaboradores, ANTES de qualquer calculo
    pesado comecar.
    """
    encontrados_map: dict[str, str] = {}
    for p in pis_no_arquivo:
        orig = str(p).strip()
        norm = _normalizar_id(p)
        if norm and orig.lower() != "nan":
            encontrados_map[norm] = orig

    # Constrói o conjunto de identificadores válidos (PIS e Matrículas)
    validos_set: set[str] = set()
    for p in pis_validos:
        orig = str(p).strip()
        norm = _normalizar_id(p)
        if norm and orig.lower() != "nan":
            validos_set.add(orig)
            validos_set.add(norm)

    # Identifica quais itens do ficheiro opcional não existem no cadastro
    desconhecidos = [
        orig for norm, orig in encontrados_map.items()
        if norm not in validos_set and orig not in validos_set
    ]
    desconhecidos = sorted(list(set(desconhecidos)))

    if desconhecidos:
        raise ArquivoInvalidoError(
            f"O arquivo de {nome_amigavel_arquivo} ('{caminho}') cita o(s) PIS/matrícula "
            f"{', '.join(desconhecidos)}, que não consta(m) no Cadastro de Colaboradores.\n"
            "Verifique se o PIS/matrícula foi digitado corretamente nesse arquivo, ou se falta "
            "cadastrar esse colaborador no Cadastro de Colaboradores."
        )