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
