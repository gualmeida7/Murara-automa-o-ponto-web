"""
pipeline.py
===========
Orquestra o fechamento de folha ponta a ponta:

    Ponto (AFD .txt ou Cartao Ponto .pdf) + cadastro de colaboradores + consignados (Excel)
    + banco de horas (Secullum / planilha Adriano) + decisoes do RH
    (vindas do painel de aprovacao)
        -> "Relacao de Valores" (um arquivo por empresa, no layout
           exigido pela contabilidade)
        -> fila de excecoes para validacao do RH (antes do fechamento)

Como rodar (depois de ajustar os caminhos dos arquivos reais no bloco
`if __name__ == "__main__"` no final do arquivo):

    python pipeline.py

Duas rodadas por fechamento:
  1. Rodada de EXCECOES: gera `fila_validacao_rh.csv`, que e' carregada no
     dashboard_aprovacao.html para o RH/gestor classificar cada dia
     estranho (falta real, erro de relogio, atestado, etc.).
  2. Rodada FINAL: le as decisoes exportadas do dashboard
     (`decisoes_rh.csv`) e gera a Relacao de Valores definitiva.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, time
from pathlib import Path
from typing import Mapping

import numpy as np
import pandas as pd

from afd_parser import agrupar_batidas_por_dia
from leitores_ponto import ler_arquivo_ponto
from pdf_espelho_ponto import mapear_para_pis, parse_espelho_pdf
from business_rules import (
    Jornada, apurar_periodo, consolidar_mes, gerar_fila_validacao_rh, STATUS_FALTA_PARCIAL,
)
from banco_horas import consolidar_banco_horas, ler_banco_horas_secullum, ler_banco_horas_adriano
from validacao import (
    verificar_colunas,
    validar_pis_existem,
    normalizar_identificador_colaborador,
    canonicalizar_pis,
    converter_valor_monetario,
    ler_sim_nao,
    ArquivoInvalidoError,
)
from leitura_arquivos import ler_arquivo_generico
from feriados import avisos_de_cobertura, carregar_feriados

# ---------------------------------------------------------------------------
# Codigos de evento exigidos pela contabilidade (Secao 4 do briefing)
# ---------------------------------------------------------------------------
COD_HE_50 = "0150"
COD_HE_100 = "0200"
COD_HORAS_FALTA = "8069"
COD_DIAS_FALTA = "8792"
COD_DESCONTO_DSR = "8794"
COD_ADIANTAMENTO_CONSIGNADO = "0981"
COD_OUTROS_BANCO_HORAS = "0999"

# O escritorio desconta o consignado direto do extrato (confirmado pela cliente): o 0981 fica
# vazio e o extrato e' opcional. True = volta a preencher o 0981 com o extrato enviado.
PREENCHER_CONSIGNADO = False

COLUNAS_RELACAO_VALORES = [
    "codigo_empresa", "razao_social", "competencia",
    "codigo_folha", "nome_colaborador",
    COD_HE_50, COD_HE_100, COD_HORAS_FALTA, COD_DIAS_FALTA,
    COD_DESCONTO_DSR, COD_ADIANTAMENTO_CONSIGNADO, COD_OUTROS_BANCO_HORAS,
]

# Colunas que o Cadastro de Colaboradores precisa ter - validadas logo no
# carregamento, para o erro apontar exatamente o arquivo/coluna em vez de
# estourar um KeyError sem contexto no meio do calculo.
COLUNAS_OBRIGATORIAS_CADASTRO = [
    "pis", "matricula", "nome", "codigo_empresa", "razao_social",
    "entrada", "saida_almoco", "retorno_almoco", "saida", "dias_trabalho",
]

# Colunas OPCIONAIS do cadastro (cadastro antigo sem elas segue como sempre):
#   banco_horas     (padrao nao) - colaborador no banco de horas;
#   registra_ponto  (padrao sim) - colaborador usa o relogio de ponto;
#   valor_dia_dsr   (sem padrao) - valor em R$ de um dia de DSR, so' usado com unidade_dsr="valor".
COL_BANCO_HORAS = "banco_horas"
COL_REGISTRA_PONTO = "registra_ponto"
COL_VALOR_DIA_DSR = "valor_dia_dsr"

# --- PREMISSAS A CONFIRMAR COM A CLIENTE (ver README, "Banco de horas") -------
# Quem esta no banco de horas tem as horas do relogio levadas ao banco, nao a folha.
# True = tambem as horas de domingo/feriado (0200) vao para o banco, nao so' a HE 50%.
BANCO_ABSORVE_HE_100 = True
# True = atraso e falta parcial (8069) sao debitados do banco, em vez de descontados na folha.
BANCO_ABSORVE_ATRASO_E_FALTA_PARCIAL = True

# Como o codigo 8794 e' lancado: "dias" (quantidade de DSR perdidos) ou "valor" (R$).
# A unidade NAO esta confirmada com a contabilidade - por isso e' uma configuracao explicita.
UNIDADES_DSR = ("dias", "valor")
UNIDADE_DSR_PADRAO = "dias"

# Terceiro estado do "Premio COPR" para quem nao registra ponto: nao da' para dizer Perde/Nao Perde.
COPR_CONFERIR = "Conferir"


def carregar_cadastro_colaboradores(caminho: str) -> pd.DataFrame:
    """
    Cadastro mestre. Colunas esperadas (ajuste os nomes se seu cadastro
    real usar outros):
        pis, matricula, nome, codigo_empresa, razao_social,
        entrada, saida_almoco, retorno_almoco, saida, dias_trabalho
        (dias_trabalho = string "0,1,2,3,4" para seg-sex, 0=segunda)

    Levanta `ArquivoInvalidoError` (com o caminho e a(s) coluna(s) que
    faltam) se o arquivo nao tiver todas as colunas obrigatorias.

    Colunas opcionais: `banco_horas` (sim/nao, padrao nao), `registra_ponto`
    (sim/nao, padrao sim) e `valor_dia_dsr` (R$). `banco_horas` e `registra_ponto`
    saem SEMPRE como booleanos, mesmo que o arquivo nao tenha a coluna;
    `valor_dia_dsr` so' existe no resultado se o arquivo a tiver.
    """
    df = ler_arquivo_generico(caminho, "Cadastro de Colaboradores")
    verificar_colunas(df, COLUNAS_OBRIGATORIAS_CADASTRO, "Cadastro de Colaboradores", caminho)
    # canonicalizar_pis (e nao so' strip): o AFD do Secullum traz o PIS com
    # zero de preenchimento a esquerda (12 digitos) - sem reduzir os dois
    # lados (cadastro E afd_parser) ao mesmo formato de 11 digitos, nenhuma
    # batida do AFD real bateria com ninguem aqui.
    df["pis"] = df["pis"].map(canonicalizar_pis)
    df[COL_BANCO_HORAS] = _ler_coluna_sim_nao(df, COL_BANCO_HORAS, False, caminho)
    df[COL_REGISTRA_PONTO] = _ler_coluna_sim_nao(df, COL_REGISTRA_PONTO, True, caminho)
    if COL_VALOR_DIA_DSR in df.columns:
        df[COL_VALOR_DIA_DSR] = pd.Series(
            [
                converter_valor_monetario(v, f"Cadastro de Colaboradores ('{caminho}'), linha {i + 2}", COL_VALOR_DIA_DSR)
                for i, v in df[COL_VALOR_DIA_DSR].items()
            ],
            index=df.index, dtype=float,
        )
    return df


def _ler_coluna_sim_nao(df: pd.DataFrame, coluna: str, padrao: bool, caminho: str) -> pd.Series:
    """Coluna sim/nao do cadastro como booleanos; coluna ausente ou celula vazia = `padrao`."""
    if coluna not in df.columns:
        return pd.Series(padrao, index=df.index, dtype=bool)
    return pd.Series(
        [
            ler_sim_nao(v, f"Cadastro de Colaboradores ('{caminho}'), linha {i + 2}", coluna, padrao)
            for i, v in df[coluna].items()
        ],
        index=df.index, dtype=bool,
    )


def _flag_cadastro(df_cadastro: pd.DataFrame, coluna: str, padrao: bool) -> pd.Series:
    """Booleano por colaborador; cadastro montado sem a coluna (ou sem o valor) usa o padrao."""
    if coluna not in df_cadastro.columns:
        return pd.Series(padrao, index=df_cadastro.index, dtype=bool)
    return df_cadastro[coluna].fillna(padrao).astype(bool)


def montar_jornadas(df_cadastro: pd.DataFrame) -> dict[str, Jornada]:
    jornadas = {}
    for _, row in df_cadastro.iterrows():
        try:
            jornadas[row["pis"]] = Jornada(
                entrada=pd.to_datetime(row["entrada"]).time(),
                saida_almoco=pd.to_datetime(row["saida_almoco"]).time(),
                retorno_almoco=pd.to_datetime(row["retorno_almoco"]).time(),
                saida=pd.to_datetime(row["saida"]).time(),
                dias_trabalho=set(int(d) for d in str(row["dias_trabalho"]).split(",")),
            )
        except Exception as e:
            raise ArquivoInvalidoError(
                f"Não consegui interpretar a jornada do colaborador com PIS '{row.get('pis')}' "
                f"no Cadastro de Colaboradores. Confira se 'entrada', 'saida_almoco', "
                f"'retorno_almoco' e 'saida' estão em formato de hora (ex.: 08:00) e se "
                f"'dias_trabalho' está no formato \"0,1,2,3,4\" (0=segunda). Detalhe: {e}"
            ) from e
    return jornadas


def _detectar_coluna_valor_consignado(colunas: list[str]) -> str | None:
    """
    Acha a coluna do valor da parcela mensal do consignado. Não basta
    procurar "valor" ou "desconto" na primeira coluna que bater: o
    extrato padrão de averbação de consignados (eSocial/Secullum) traz
    colunas como "competenciaInicioDesconto" e "competenciaFimDesconto"
    (datas, não valores) ANTES de "valorParcela" na planilha - "desconto"
    aparece no nome dessas colunas de data também. Por isso a busca
    primeiro restringe aos candidatos que não são claramente uma
    competência/data, prioriza "valor" (que cobre "valorparcela",
    "valoremprestimo", "valorliberado" - nessa ordem, e "valorparcela" é
    a parcela mensal, a correta para descontar na folha) e só cai para
    "desconto" como últimO recurso, para não quebrar extratos simples
    que usam um nome de coluna como "valor_desconto".
    """
    candidatos = [c for c in colunas if "competencia" not in c and "data" not in c]
    return (
        next((c for c in candidatos if "valor" in c), None)
        or next((c for c in candidatos if "desconto" in c), None)
    )


def carregar_consignados(caminho: str | list[str], cadastro: pd.DataFrame | None = None) -> pd.DataFrame:
    """
    Extrato de consignados. Espera colunas matricula/pis + valor da
    parcela mensal (ex.: "valor_desconto", ou "valorParcela" no layout
    padrão de averbação eSocial/Secullum - ver `_detectar_coluna_valor_consignado`).

    Aceita um único caminho ou uma LISTA de caminhos: cada empresa
    costuma emitir seu próprio extrato de consignados (CNPJs diferentes),
    então o fechamento muitas vezes precisa somar mais de um arquivo
    antes da checagem pré-voo e do cruzamento com a Relação de Valores.

    Um mesmo colaborador pode ter mais de um empréstimo ativo no mês
    (várias linhas) - o valor de todas é somado por colaborador antes de
    retornar.

    Passe `cadastro` (o DataFrame já carregado de
    `carregar_cadastro_colaboradores`) sempre que possível: muitos
    extratos de consignados identificam o colaborador pela MATRÍCULA
    interna, não pelo PIS, e sem o cadastro em mãos não dá para saber
    qual dos dois está na coluna de identificação — veja
    `validacao.normalizar_identificador_colaborador` para o porquê.
    """
    caminhos = [caminho] if isinstance(caminho, str) else list(caminho)
    partes = []
    for caminho_arquivo in caminhos:
        df = ler_arquivo_generico(caminho_arquivo, "Extrato de Consignados")
        col_id = next((c for c in df.columns if "pis" in c or "matricula" in c), None)
        col_valor = _detectar_coluna_valor_consignado(list(df.columns))
        if col_id is None or col_valor is None:
            raise ArquivoInvalidoError(
                f"O Extrato de Consignados ('{caminho_arquivo}') está sem uma coluna de identificação "
                f"(pis/matrícula) e/ou de valor da parcela (valor/desconto).\n"
                f"Colunas encontradas: {list(df.columns)}."
            )
        parte = df[[col_id, col_valor]].rename(columns={col_id: "pis", col_valor: "valor_consignado"})
        parte["pis"] = parte["pis"].astype(str).str.strip()
        if cadastro is not None:
            parte = normalizar_identificador_colaborador(parte, "pis", cadastro, "Extrato de Consignados", caminho_arquivo)
        partes.append(parte)

    out = pd.concat(partes, ignore_index=True)
    return out.groupby("pis", as_index=False)["valor_consignado"].sum()


def executar_checagens_preflight(
    df_cadastro: pd.DataFrame,
    df_consignados: pd.DataFrame | None = None,
    df_banco_secullum: pd.DataFrame | None = None,
    df_banco_adriano: pd.DataFrame | None = None,
    caminho_consignados: str = "",
    caminho_banco_secullum: str = "",
    caminho_banco_adriano: str = "",
) -> None:
    """
    Roda ANTES do processamento pesado (leitura/apuracao do AFD): confere
    que todo PIS/matrícula citado nos arquivos opcionais existe de fato no
    Cadastro de Colaboradores. Sem essa checagem, um PIS digitado errado
    no banco de horas do Adriano, por exemplo, simplesmente nao casaria
    com ninguem no merge la' na frente - a folha de alguem sairia sem o
    banco de horas dela, silenciosamente, sem nenhum aviso.
    """
    pis_validos = df_cadastro["pis"]

    if df_consignados is not None and not df_consignados.empty:
        validar_pis_existem(df_consignados["pis"], pis_validos, "Extrato de Consignados", caminho_consignados)

    if df_banco_secullum is not None and not df_banco_secullum.empty:
        validar_pis_existem(
            df_banco_secullum["pis"], pis_validos, "Banco de Horas do Secullum", caminho_banco_secullum
        )

    if df_banco_adriano is not None and not df_banco_adriano.empty:
        validar_pis_existem(
            df_banco_adriano["pis"], pis_validos, "Planilha semanal do Adriano", caminho_banco_adriano
        )

    _validar_banco_so_para_quem_esta_no_banco(
        df_cadastro, df_banco_secullum, "Banco de Horas do Secullum", caminho_banco_secullum
    )
    _validar_banco_so_para_quem_esta_no_banco(
        df_cadastro, df_banco_adriano, "Planilha semanal do Adriano", caminho_banco_adriano
    )


def _validar_banco_so_para_quem_esta_no_banco(
    df_cadastro: pd.DataFrame, df_banco: pd.DataFrame | None, nome_arquivo: str, caminho: str,
) -> None:
    """
    Saldo de banco de horas para quem NAO esta marcado `banco_horas = sim` no cadastro:
    as horas extras sairiam na folha (0150/0200) e tambem no banco (0999), pagas duas
    vezes. Linha com saldo 0 nao paga nada duas vezes (export que lista todo mundo), entao passa.
    """
    if df_banco is None or df_banco.empty:
        return
    no_banco = set(df_cadastro.loc[_flag_cadastro(df_cadastro, COL_BANCO_HORAS, False), "pis"])
    com_saldo = df_banco.loc[df_banco["saldo_banco_horas"].fillna(0) != 0, "pis"]
    fora = sorted({str(p).strip() for p in com_saldo} - no_banco)
    if fora:
        nomes = df_cadastro.drop_duplicates("pis").set_index("pis")["nome"]
        quem = ", ".join(f"{nomes.get(p, p)} (PIS {p})" for p in fora)
        raise ArquivoInvalidoError(
            f"O arquivo de {nome_arquivo} ('{caminho}') traz saldo de banco de horas para quem não está "
            f"marcado como banco de horas no Cadastro de Colaboradores: {quem}.\n"
            "Com isso as horas extras dessa pessoa seriam pagas na folha (0150/0200) E também contadas "
            "no banco de horas (0999).\n"
            "Se ela realmente está no banco de horas, preencha 'sim' na coluna 'banco_horas' do cadastro; "
            "senão, tire o saldo dela desse arquivo."
        )


def avisos_banco_horas(
    df_cadastro: pd.DataFrame,
    df_banco_secullum: pd.DataFrame | None = None,
    df_banco_adriano: pd.DataFrame | None = None,
) -> list[str]:
    """
    AVISOS (nao erros): colaboradores marcados `banco_horas = sim` sem nenhuma linha de saldo
    nos arquivos enviados - o 0999 deles sairia 0 (e as horas do relogio nao vao para a folha).
    """
    partes = [d["pis"] for d in (df_banco_secullum, df_banco_adriano) if d is not None and not d.empty]
    com_saldo = {str(p).strip() for parte in partes for p in parte}
    no_banco = df_cadastro[_flag_cadastro(df_cadastro, COL_BANCO_HORAS, False)]
    return [
        f"{nome} está marcado(a) como banco de horas, mas nenhum arquivo de banco de horas traz o saldo "
        "dessa pessoa: o Banco de Horas (0999) dela sairá 0."
        for pis, nome in zip(no_banco["pis"], no_banco["nome"])
        if str(pis).strip() not in com_saldo
    ]


def rodada_1_gerar_fila_de_excecoes(
    caminho_ponto: str,
    df_cadastro: pd.DataFrame,
    data_inicio: date,
    data_fim: date,
    caminho_espelho: str | None = None,
    feriados: Mapping[date, str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Retorna (apuracao_diaria_completa, fila_para_validacao_rh).
    `caminho_ponto` pode ser o AFD (.txt) ou o Cartao Ponto do Secullum
    (.pdf) - ver leitores_ponto.py.

    `feriados` ({data: nome}) sao os dias que nao contam como falta e cujo
    trabalho vale hora extra 100%. Sem ele, le o calendario padrao
    (`feriados.csv`, ver feriados.py) e imprime os avisos de cobertura.

    Com `caminho_espelho` (Espelho de Ponto em PDF), cada dia sem batida na
    fila e' cruzado com as batidas brutas do relogio - ver
    `_anexar_cruzamento_espelho`.
    """
    if feriados is None:
        feriados = carregar_feriados()
        for aviso in avisos_de_cobertura(data_inicio, data_fim):
            print(f"[aviso] {aviso}")
    resultado_ponto = ler_arquivo_ponto(caminho_ponto, df_cadastro)
    if resultado_ponto.erros_leitura:
        print(f"[aviso] {len(resultado_ponto.erros_leitura)} aviso(s) na leitura do ponto - ver log.")

    batidas_por_dia = agrupar_batidas_por_dia(resultado_ponto.batidas)
    # Quem nao registra ponto (registra_ponto = nao) fica fora da apuracao: sem dia, sem excecao, sem falta.
    jornadas = montar_jornadas(df_cadastro[_flag_cadastro(df_cadastro, COL_REGISTRA_PONTO, True)])
    apuracao = apurar_periodo(batidas_por_dia, jornadas, data_inicio, data_fim, feriados=feriados)
    # Cadastrado que registra ponto mas nao tem NENHUMA batida no arquivo (ex.: desligado que
    # continua no cadastro): sem isto o mes inteiro viraria falta calada. Viaja em attrs, como
    # os totais do Secullum, para o app avisar o RH antes de fechar.
    com_batida = set(batidas_por_dia["pis"]) if not batidas_por_dia.empty else set()
    apuracao.attrs["sem_ponto"] = [pis for pis in jornadas if pis not in com_batida]
    # Totais impressos pelo proprio Secullum (so' no Cartao Ponto em PDF): viajam
    # junto da apuracao para a conferencia da rodada 2, sem mudar o retorno.
    if resultado_ponto.totais_secullum is not None:
        apuracao.attrs["totais_secullum"] = resultado_ponto.totais_secullum
    fila = gerar_fila_validacao_rh(apuracao)
    fila = _anexar_contexto_secullum(fila, resultado_ponto.espelho_diario)
    if caminho_espelho:
        resultado_espelho = parse_espelho_pdf(caminho_espelho)
        batidas_brutas = mapear_para_pis(resultado_espelho, df_cadastro)
        if resultado_espelho.avisos:
            print(f"[aviso] {len(resultado_espelho.avisos)} aviso(s) na leitura do Espelho de Ponto - ver log.")
        fila = _anexar_cruzamento_espelho(fila, batidas_brutas)
    return apuracao, fila


CRUZAMENTO_ERRO_RELOGIO = "Provável erro de relógio — verificar"
CRUZAMENTO_FALTA_REAL = "Provável falta real — confirmar"
CRUZAMENTO_BATIDA_NAO_IMPORTADA = "Provável erro de relógio / batida não importada — verificar"
CRUZAMENTO_FALTA_PARCIAL = "Provável falta parcial — confirmar"

# Batidas brutas do Espelho a menos de 2 min uma da outra sao a mesma marcacao
# duplicada pelo relogio ("Batida duplicada"), nao uma batida a mais.
TOLERANCIA_BATIDA_DUPLICADA_MIN = 2


def _minutos_do_dia(hhmm: str) -> int:
    horas, minutos = hhmm.split(":")
    return int(horas) * 60 + int(minutos)


def _batidas_brutas_distintas(brutas: list[str]) -> list[str]:
    distintas: list[str] = []
    for hora in sorted(brutas, key=_minutos_do_dia):
        if not distintas or _minutos_do_dia(hora) - _minutos_do_dia(distintas[-1]) > TOLERANCIA_BATIDA_DUPLICADA_MIN:
            distintas.append(hora)
    return distintas


def _anexar_cruzamento_espelho(fila: pd.DataFrame, dias_espelho: pd.DataFrame) -> pd.DataFrame:
    """
    Usa o Espelho de Ponto (batidas brutas do relogio) para ajudar o RH nos dias
    em que o Cartao Ponto esta incompleto:
      - dia SEM batida (FALTA_OU_AUSENCIA): se o Espelho mostra alguma batida,
        "Provavel erro de relogio - verificar"; senao "Provavel falta real - confirmar";
      - dia com batidas faltando (FALTA_PARCIAL_CANDIDATA): se o Espelho tem MAIS
        batidas distintas que o Cartao, "Provavel erro de relogio / batida nao
        importada - verificar" (e lista as batidas brutas no motivo); senao
        "Provavel falta parcial - confirmar".
    Adiciona a coluna `cruzamento_espelho` (NaN nas demais linhas) e acrescenta
    o resultado ao `motivo_validacao`. Nao decide nada: a operadora continua
    classificando cada linha na tela de aprovacao.
    """
    if fila.empty:
        return fila.assign(cruzamento_espelho=pd.Series(dtype=str))

    brutas_por_dia = {
        (pis, pd.Timestamp(data)): _batidas_brutas_distintas(list(brutas))
        for pis, data, brutas in zip(dias_espelho["pis"], dias_espelho["data"], dias_espelho["batidas_brutas"])
    }

    def _classificar(linha):
        brutas = brutas_por_dia.get((linha["pis"], pd.Timestamp(linha["data"])), [])
        if linha["status"] == "FALTA_OU_AUSENCIA":
            return CRUZAMENTO_ERRO_RELOGIO if brutas else CRUZAMENTO_FALTA_REAL
        if linha["status"] == STATUS_FALTA_PARCIAL:
            return CRUZAMENTO_BATIDA_NAO_IMPORTADA if len(brutas) > linha["n_batidas"] else CRUZAMENTO_FALTA_PARCIAL
        return pd.NA

    def _detalhe_espelho(linha):
        brutas = brutas_por_dia.get((linha["pis"], pd.Timestamp(linha["data"])), [])
        texto = f" — Espelho de Ponto: {linha['cruzamento_espelho']}"
        if brutas and linha["cruzamento_espelho"] == CRUZAMENTO_BATIDA_NAO_IMPORTADA:
            texto += f" (batidas no Espelho: {', '.join(brutas)})"
        return texto

    fila = fila.copy()
    fila["cruzamento_espelho"] = fila.apply(_classificar, axis=1)
    tem = fila["cruzamento_espelho"].notna()
    if tem.any():
        fila.loc[tem, "motivo_validacao"] = fila.loc[tem, "motivo_validacao"] + fila.loc[tem].apply(_detalhe_espelho, axis=1)
    return fila


def _anexar_contexto_secullum(fila: pd.DataFrame, espelho_diario: pd.DataFrame | None) -> pd.DataFrame:
    """
    Quando o ponto veio do Cartao Ponto em PDF (`cartao_ponto_pdf.py`), cada
    dia ja' traz a ocorrencia que o PROPRIO Secullum anotou (ex.: 'ATESTAD',
    'FALTA') - informacao que o AFD (.txt) simplesmente nao tem. Isso NAO
    decide nada automaticamente (o RH continua tendo que confirmar cada
    linha, exatamente como hoje) - so' poupa a pessoa de abrir o PDF de
    novo pra descobrir que aquele dia sem batida sinalizado como "confirmar
    se e' falta real ou erro de relogio" ja' tinha um atestado lancado.

    Com AFD a coluna `ocorrencia_secullum` fica vazia (NaN) - nenhum
    comportamento muda pra quem so' usa o AFD.
    """
    if fila.empty:
        return fila.assign(ocorrencia_secullum=pd.Series(dtype=str))
    if espelho_diario is None or espelho_diario.empty or "ocorrencia" not in espelho_diario.columns:
        return fila.assign(ocorrencia_secullum=pd.NA)

    notas = (
        espelho_diario[["pis", "data", "ocorrencia"]]
        .dropna(subset=["ocorrencia"])
        .rename(columns={"ocorrencia": "ocorrencia_secullum"})
    )
    notas["data"] = pd.to_datetime(notas["data"])
    fila = fila.merge(notas, on=["pis", "data"], how="left")

    tem_nota = fila["ocorrencia_secullum"].notna()
    fila.loc[tem_nota, "motivo_validacao"] = (
        fila.loc[tem_nota, "motivo_validacao"]
        + " — Secullum já registrou: '" + fila.loc[tem_nota, "ocorrencia_secullum"] + "'"
    )
    return fila


def exportar_fila_validacao_rh(fila: pd.DataFrame, df_cadastro: pd.DataFrame, caminho_csv: str) -> pd.DataFrame:
    """
    Prepara e grava a fila de excecoes num CSV pronto para o RH abrir
    direto no Excel:
      - cruza com o cadastro pra trazer nome_colaborador e razao_social
        (o RH nao precisa abrir outro arquivo pra saber quem e' quem);
      - ordena por nome do colaborador e, dentro do mesmo colaborador,
        cronologicamente pela data (a ordenacao cronologica acontece
        ANTES de converter a data pra string, senao "05/12/2026" ficaria
        antes de "20/01/2026" numa ordenacao de texto);
      - formata a data como dd/mm/aaaa (padrao brasileiro);
      - grava em 'utf-8-sig', que e' o utf-8 com o BOM que faz o Excel
        abrir acentos corretamente sem pedir importacao de texto.
    Retorna o DataFrame final tambem, para quem quiser inspecionar sem
    reabrir o CSV.
    """
    colunas_saida = ["pis", "nome_colaborador", "razao_social", "data", "status", "motivo_validacao"]
    if "ocorrencia_secullum" in fila.columns:
        colunas_saida.append("ocorrencia_secullum")

    if fila.empty:
        fila_final = fila.assign(nome_colaborador=pd.Series(dtype=str), razao_social=pd.Series(dtype=str))
        fila_final = fila_final.reindex(columns=colunas_saida)
        fila_final.to_csv(caminho_csv, index=False, encoding="utf-8-sig")
        return fila_final

    info_colaborador = df_cadastro[["pis", "nome", "razao_social"]].rename(columns={"nome": "nome_colaborador"})
    fila_com_nome = fila.merge(info_colaborador, on="pis", how="left")

    fila_com_nome["data"] = pd.to_datetime(fila_com_nome["data"])
    fila_ordenada = fila_com_nome.sort_values(["nome_colaborador", "data"]).reset_index(drop=True)
    fila_ordenada["data"] = fila_ordenada["data"].dt.strftime("%d/%m/%Y")

    fila_final = fila_ordenada[colunas_saida]
    fila_final.to_csv(caminho_csv, index=False, encoding="utf-8-sig")
    return fila_final


def rodada_2_gerar_relacao_de_valores(
    apuracao_diaria: pd.DataFrame,
    df_cadastro: pd.DataFrame,
    df_consignados: pd.DataFrame | None,
    df_banco_horas: pd.DataFrame,
    decisoes_rh: pd.DataFrame | None,
    competencia: str,
    valor_dia_dsr_por_pis: dict[str, float] | None = None,
    unidade_dsr: str = UNIDADE_DSR_PADRAO,
) -> pd.DataFrame:
    """
    Gera a tabela final no layout exigido pela contabilidade, uma linha por colaborador.

    Banco de horas: quem tem `banco_horas = sim` no cadastro tem as horas do relogio
    levadas ao banco, nao a folha - 0150 = 0 e, conforme as constantes `BANCO_ABSORVE_*`,
    0200 e 8069 tambem. O 0999 continua vindo de `df_banco_horas`. Falta de dia inteiro
    (8792, COPR, VA, DSR) nao muda. Quem tem `registra_ponto = nao` nao tem apuracao: zero
    nos codigos do relogio e "Conferir" no premio COPR.

    8794 (`unidade_dsr`): "dias" = DSR perdidos (semanas com falta inteira injustificada);
    "valor" = semanas x valor do dia de DSR (coluna `valor_dia_dsr` do cadastro; o dict
    `valor_dia_dsr_por_pis` sobrepoe a coluna). Em "valor", faltar o valor de quem perdeu
    DSR levanta `ArquivoInvalidoError`.
    """
    if unidade_dsr not in UNIDADES_DSR:
        raise ValueError(f"unidade_dsr deve ser um destes: {', '.join(UNIDADES_DSR)} (recebido: {unidade_dsr!r})")
    resumo_mes = consolidar_mes(apuracao_diaria, decisoes_rh)

    df = df_cadastro.merge(resumo_mes, on="pis", how="left")
    if PREENCHER_CONSIGNADO and df_consignados is not None:
        df = df.merge(df_consignados, on="pis", how="left")
    else:
        df["valor_consignado"] = float("nan")  # 0981 vazio: o escritorio desconta pelo extrato
    df = df.merge(df_banco_horas, on="pis", how="left")

    for col in ["horas_extra_50", "horas_extra_100", "horas_falta", "atraso_horas", "horas_falta_parcial",
                "dias_falta", "dias_falta_parcial", "dias_falta_justificada", "dias_falta_injustificada",
                "dias_desconto_va", "semanas_perde_dsr", "valor_consignado", "saldo_banco_horas"]:
        if col in df.columns and not (col == "valor_consignado" and not PREENCHER_CONSIGNADO):
            df[col] = df[col].fillna(0)

    df["valor_desconto_dsr"], valor_dia_dsr = _calcular_8794(df, unidade_dsr, valor_dia_dsr_por_pis)

    # Banco de horas / sem relogio. O df continua com os valores DO RELOGIO (a conferencia
    # com o Secullum, la' embaixo, compara com eles); so' a saida roteia para o banco.
    no_banco = _flag_cadastro(df, COL_BANCO_HORAS, False)
    registra_ponto = _flag_cadastro(df, COL_REGISTRA_PONTO, True)
    he50_banco = df["horas_extra_50"].where(no_banco, 0.0)
    he100_banco = df["horas_extra_100"].where(no_banco & BANCO_ABSORVE_HE_100, 0.0)
    atraso_falta_banco = df["horas_falta"].where(no_banco & BANCO_ABSORVE_ATRASO_E_FALTA_PARCIAL, 0.0)
    perde_copr = df["perde_copr"].fillna(False).astype(object).where(registra_ponto, COPR_CONFERIR)

    saida = pd.DataFrame({
        "codigo_empresa": df["codigo_empresa"],
        "razao_social": df["razao_social"],
        "competencia": competencia,
        "codigo_folha": df["matricula"],
        "nome_colaborador": df["nome"],
        COD_HE_50: df["horas_extra_50"] - he50_banco,
        COD_HE_100: df["horas_extra_100"] - he100_banco,
        COD_HORAS_FALTA: df["horas_falta"] - atraso_falta_banco,
        COD_DIAS_FALTA: df["dias_falta"],
        COD_DESCONTO_DSR: df["valor_desconto_dsr"],
        COD_ADIANTAMENTO_CONSIGNADO: df["valor_consignado"],
        COD_OUTROS_BANCO_HORAS: df["saldo_banco_horas"],
        # colunas auxiliares - uteis para conferencia interna do RH, NAO
        # fazem parte do layout formal enviado a contabilidade. Ficam com
        # prefixo "_" de proposito, para o exportador saber separa-las
        # automaticamente na aba "Conferência RH" e nunca vazarem para a
        # aba "Contabilidade":
        "_perde_premio_copr": perde_copr,
        "_banco_horas": no_banco.map({True: "Sim", False: "Não"}),
        "_he50_banco": he50_banco,
        "_he100_banco": he100_banco,
        "_atraso_falta_parcial_banco": atraso_falta_banco,
        "_dias_falta_total": df["dias_falta"],
        "_dias_falta_justificada": df["dias_falta_justificada"],
        "_dias_falta_injustificada": df["dias_falta_injustificada"],
        "_dias_falta_parcial": df["dias_falta_parcial"],
        "_horas_falta_parcial": df["horas_falta_parcial"],
        "_atraso_horas": df["atraso_horas"],
        "_dias_desconto_va": df["dias_desconto_va"],
        "_semanas_perde_dsr": df["semanas_perde_dsr"],
    })
    if valor_dia_dsr is not None:
        saida["_valor_dia_dsr"] = valor_dia_dsr

    totais_secullum = apuracao_diaria.attrs.get("totais_secullum")
    if totais_secullum is not None:
        saida = saida.join(_conferencia_com_secullum(df, totais_secullum))
    return saida


def _calcular_8794(
    df: pd.DataFrame, unidade_dsr: str, valor_dia_dsr_por_pis: Mapping[str, float] | None,
) -> tuple[pd.Series, pd.Series | None]:
    """
    (valor do 8794, valor do dia de DSR - so' em "valor"). Em "dias" nao precisa de dado
    de salario. Em "valor", o dia vem do dict (prioridade) ou da coluna `valor_dia_dsr` do
    cadastro, e quem perdeu DSR sem valor derruba o fechamento: nunca vira 0 em silencio.
    """
    semanas = df["semanas_perde_dsr"]
    if unidade_dsr == "dias":
        return semanas.astype(int), None

    valor_dia = df[COL_VALOR_DIA_DSR].astype(float) if COL_VALOR_DIA_DSR in df.columns else pd.Series(np.nan, index=df.index)
    if valor_dia_dsr_por_pis:
        valor_dia = df["pis"].map(valor_dia_dsr_por_pis).astype(float).combine_first(valor_dia)
    sem_valor = (semanas > 0) & valor_dia.isna()
    if sem_valor.any():
        quem = ", ".join(f"{n} (matrícula {m})" for n, m in zip(df.loc[sem_valor, "nome"], df.loc[sem_valor, "matricula"]))
        raise ArquivoInvalidoError(
            "O desconto de DSR (8794) foi pedido em valor (R$), mas falta o valor do dia de DSR de quem "
            f"perdeu DSR no período: {quem}.\n"
            f"Preencha a coluna '{COL_VALOR_DIA_DSR}' no Cadastro de Colaboradores, ou lance o 8794 em dias."
        )
    return (semanas * valor_dia.fillna(0.0)).round(2), valor_dia


def _conferencia_com_secullum(df: pd.DataFrame, totais_secullum: pd.DataFrame) -> pd.DataFrame:
    """
    Colunas informativas "Nosso x Secullum" (so' quando o ponto veio do Cartao
    Ponto em PDF, que traz a linha TOTAIS do proprio Secullum). NAO altera
    nenhum valor da folha: serve para o RH ver onde o calculo do sistema se
    afasta do Secullum. Tudo em dias ou horas decimais; o Secullum imprime
    horas como hh:mm, que o parser ja' converte para minutos.
    """
    sec = totais_secullum.dropna(subset=["pis"]).drop_duplicates("pis").set_index("pis")

    def _do_secullum(coluna: str, em_horas: bool) -> pd.Series:
        if coluna not in sec.columns:
            return pd.Series(float("nan"), index=df.index)
        valores = df["pis"].map(sec[coluna]).astype(float)
        return (valores / 60).round(2) if em_horas else valores

    nosso = {
        "dias_falta": df["dias_falta"],
        "horas_falta": df["horas_falta"],
        "he50": df["horas_extra_50"],
        "he100": df["horas_extra_100"],
    }
    do_sec = {
        "dias_falta": _do_secullum("faltas_dias", em_horas=False),
        "horas_falta": _do_secullum("faltas_horas", em_horas=True),
        "he50": _do_secullum("extra_50", em_horas=True),
        "he100": _do_secullum("extra_100", em_horas=True),
    }
    colunas = {}
    for chave in nosso:
        colunas[f"_sec_{chave}"] = do_sec[chave]
    for chave in nosso:
        colunas[f"_dif_{chave}"] = (nosso[chave] - do_sec[chave]).round(2)
    return pd.DataFrame(colunas, index=df.index)


def exportar_por_empresa(df_relacao: pd.DataFrame, pasta_saida: str) -> list[Path]:
    """
    Gera um .xlsx por empresa com DUAS abas:
      - "Contabilidade": estritamente as colunas do layout exigido pelo
        escritorio contabil (nenhuma coluna auxiliar, para nao arriscar
        rejeicao no sistema deles).
      - "Conferência RH": as mesmas colunas + as colunas auxiliares
        (prefixo "_" em `df_relacao`) para auditoria interna, com
        destaque visual de quem perde premio/tem faltas.
    """
    pasta = Path(pasta_saida)
    pasta.mkdir(parents=True, exist_ok=True)
    arquivos = []
    for (codigo, razao), grupo in df_relacao.groupby(["codigo_empresa", "razao_social"]):
        nome_arquivo = f"relacao_de_valores_{codigo}_{razao.replace(' ', '_')}.xlsx"
        caminho = pasta / nome_arquivo
        _escrever_workbook_formatado(grupo, caminho)
        arquivos.append(caminho)
    return arquivos


# ---------------------------------------------------------------------------
# Exportação Excel com openpyxl - layout, formatação e formatação condicional
# ---------------------------------------------------------------------------

_COR_CABECALHO = "1B4D3E"       # verde corporativo
_COR_TEXTO_CABECALHO = "FFFFFF"
_COR_BORDA = "D9D9D9"           # cinza claro
_COR_ALERTA_FUNDO = "FBE4E1"    # vermelho bem suave
_COR_ALERTA_TEXTO = "8C2E22"
LIMITE_FALTAS_ALTO = 3          # a partir de quantas faltas o RH ve' o destaque de "elevado"

# Nomes de exibicao das colunas do layout contabil (na ordem em que aparecem)
_COLUNAS_CONTABIL = [
    "codigo_empresa", "razao_social", "competencia", "codigo_folha", "nome_colaborador",
    COD_HE_50, COD_HE_100, COD_HORAS_FALTA, COD_DIAS_FALTA,
    COD_DESCONTO_DSR, COD_ADIANTAMENTO_CONSIGNADO, COD_OUTROS_BANCO_HORAS,
]
_CABECALHOS_CONTABIL = {
    "codigo_empresa": "Código Empresa", "razao_social": "Razão Social", "competencia": "Competência",
    "codigo_folha": "Código Folha", "nome_colaborador": "Nome do Colaborador",
    COD_HE_50: "HE 50% (0150)", COD_HE_100: "HE 100% (0200)",
    COD_HORAS_FALTA: "Horas Falta (8069)", COD_DIAS_FALTA: "Dias Falta (8792)",
    COD_DESCONTO_DSR: "Desconto DSR (8794)", COD_ADIANTAMENTO_CONSIGNADO: "Consignado (0981)",
    COD_OUTROS_BANCO_HORAS: "Banco de Horas (0999)",
}
_COLUNAS_AUX = [
    "_perde_premio_copr", "_dias_falta_total", "_dias_falta_justificada", "_dias_falta_injustificada",
    "_dias_falta_parcial", "_horas_falta_parcial", "_atraso_horas", "_dias_desconto_va", "_semanas_perde_dsr",
    "_banco_horas", "_he50_banco", "_he100_banco", "_atraso_falta_parcial_banco",
]
# So' existe quando o 8794 foi lancado em valor (R$).
_COLUNAS_AUX_DSR_VALOR = ["_valor_dia_dsr"]
# Bloco "Nosso x Secullum": so' existe quando o ponto veio do Cartao Ponto em PDF.
_COLUNAS_CONFERENCIA_SECULLUM = [
    "_sec_dias_falta", "_dif_dias_falta", "_sec_horas_falta", "_dif_horas_falta",
    "_sec_he50", "_dif_he50", "_sec_he100", "_dif_he100",
]
_CABECALHOS_AUX = {
    "_perde_premio_copr": "Prêmio COPR", "_dias_falta_total": "Dias de Falta (total)",
    "_dias_falta_justificada": "Dias Falta Justificada", "_dias_falta_injustificada": "Dias Falta Injustificada",
    "_dias_falta_parcial": "Dias Falta Parcial", "_horas_falta_parcial": "Horas Falta Parcial",
    "_atraso_horas": "Horas de Atraso",
    "_banco_horas": "Banco de horas", "_he50_banco": "HE 50% enviada ao banco",
    "_he100_banco": "HE 100% enviada ao banco", "_atraso_falta_parcial_banco": "Atraso/falta parcial enviados ao banco",
    "_valor_dia_dsr": "Valor do dia de DSR",
    "_dias_desconto_va": "Dias Desconto VA", "_semanas_perde_dsr": "Semanas c/ Perda de DSR",
    "_sec_dias_falta": "Dias Falta (Secullum)", "_dif_dias_falta": "Dif. Dias Falta (nosso - Secullum)",
    "_sec_horas_falta": "Horas Falta (Secullum)", "_dif_horas_falta": "Dif. Horas Falta",
    "_sec_he50": "HE 50% (Secullum)", "_dif_he50": "Dif. HE 50%",
    "_sec_he100": "HE 100% (Secullum)", "_dif_he100": "Dif. HE 100%",
}

# Colunas monetarias/numericas que levam formato "0,00" (2 casas decimais explicitas)
_COLUNAS_DECIMAIS = {
    COD_HE_50, COD_HE_100, COD_HORAS_FALTA, COD_DESCONTO_DSR, COD_ADIANTAMENTO_CONSIGNADO, COD_OUTROS_BANCO_HORAS,
    "_horas_falta_parcial", "_atraso_horas", "_sec_horas_falta", "_dif_horas_falta",
    "_he50_banco", "_he100_banco", "_atraso_falta_parcial_banco", "_valor_dia_dsr",
    "_sec_he50", "_dif_he50", "_sec_he100", "_dif_he100",
}
# Colunas de quantidade inteira (ainda assim formatadas com precisao explicita)
_COLUNAS_INTEIRAS = {
    COD_DIAS_FALTA, "_dias_falta_total", "_dias_falta_justificada", "_dias_falta_injustificada",
    "_dias_falta_parcial", "_dias_desconto_va", "_semanas_perde_dsr", "_sec_dias_falta", "_dif_dias_falta",
}
# Colunas centralizadas (codigos, matricula, competencia)
_COLUNAS_CENTRALIZADAS = {"codigo_empresa", "competencia", "codigo_folha", "_perde_premio_copr", "_banco_horas"}


def _perde_copr(valor) -> bool:
    """So' bool de verdade e' perda: "Conferir" (sem ponto) e' texto e nao deve ser destacado."""
    return isinstance(valor, (bool, np.bool_)) and bool(valor)


def _escrever_workbook_formatado(grupo: pd.DataFrame, caminho: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    fonte_padrao = "Arial"
    fill_cabecalho = PatternFill("solid", fgColor=_COR_CABECALHO)
    fonte_cabecalho = Font(name=fonte_padrao, bold=True, color=_COR_TEXTO_CABECALHO)
    borda_fina = Border(*[Side(style="thin", color=_COR_BORDA)] * 4)
    fill_alerta = PatternFill("solid", fgColor=_COR_ALERTA_FUNDO)
    fonte_alerta = Font(name=fonte_padrao, color=_COR_ALERTA_TEXTO)

    def alinhamento_para(nome_coluna: str) -> Alignment:
        if nome_coluna in _COLUNAS_CENTRALIZADAS:
            return Alignment(horizontal="center", vertical="center")
        if nome_coluna in _COLUNAS_DECIMAIS or nome_coluna in _COLUNAS_INTEIRAS:
            return Alignment(horizontal="right", vertical="center")
        return Alignment(horizontal="left", vertical="center")

    def formato_numero_para(nome_coluna: str) -> str | None:
        if nome_coluna in _COLUNAS_DECIMAIS:
            return "0.00"
        if nome_coluna in _COLUNAS_INTEIRAS:
            return "0"
        return None

    def escrever_aba(ws, colunas: list[str], cabecalhos: dict, com_alerta: bool):
        ws.append([cabecalhos.get(c, c) for c in colunas])
        for cel in ws[1]:
            cel.fill = fill_cabecalho
            cel.font = fonte_cabecalho
            cel.alignment = Alignment(horizontal="center", vertical="center")
            cel.border = borda_fina
        ws.freeze_panes = "A2"

        for _, linha in grupo.iterrows():
            valores = []
            for c in colunas:
                v = linha[c]
                if c == "_perde_premio_copr":
                    v = COPR_CONFERIR if v == COPR_CONFERIR else ("Perde" if bool(v) else "Não Perde")
                elif pd.isna(v):
                    v = None  # celula em branco (ex.: colaborador sem linha de TOTAIS no Secullum)
                valores.append(v)
            ws.append(valores)

        n_linhas = len(grupo)
        for i in range(2, n_linhas + 2):
            linha_df = grupo.iloc[i - 2]
            destacar = com_alerta and (
                _perde_copr(linha_df.get("_perde_premio_copr", False))
                or int(linha_df.get("_dias_falta_total", 0) or 0) >= LIMITE_FALTAS_ALTO
            )
            for j, c in enumerate(colunas, start=1):
                cel = ws.cell(row=i, column=j)
                cel.font = fonte_alerta if destacar else Font(name=fonte_padrao)
                cel.alignment = alinhamento_para(c)
                cel.border = borda_fina
                fmt = formato_numero_para(c)
                if fmt:
                    cel.number_format = fmt
                if destacar:
                    cel.fill = fill_alerta

        for j, c in enumerate(colunas, start=1):
            letra = get_column_letter(j)
            maior = max(
                [len(str(cabecalhos.get(c, c)))]
                + [len(str(cabecalhos.get(c, c)) if pd.isna(linha[c]) else str(linha[c])) for _, linha in grupo.iterrows()]
            )
            ws.column_dimensions[letra].width = min(max(maior + 3, 10), 42)

        ws.auto_filter.ref = f"A1:{get_column_letter(len(colunas))}{n_linhas + 1}"
        _configurar_impressao(ws, len(colunas), n_linhas + 1)

    wb = Workbook()
    ws_contabil = wb.active
    ws_contabil.title = "Contabilidade"
    escrever_aba(ws_contabil, _COLUNAS_CONTABIL, _CABECALHOS_CONTABIL, com_alerta=False)

    ws_rh = wb.create_sheet("Conferência RH")
    colunas_rh = (
        _COLUNAS_CONTABIL + _COLUNAS_AUX
        + [c for c in _COLUNAS_AUX_DSR_VALOR if c in grupo.columns]
        + [c for c in _COLUNAS_CONFERENCIA_SECULLUM if c in grupo.columns]
    )
    cabecalhos_rh = {**_CABECALHOS_CONTABIL, **_CABECALHOS_AUX}
    escrever_aba(ws_rh, colunas_rh, cabecalhos_rh, com_alerta=True)

    wb.save(caminho)


# ---------------------------------------------------------------------------
# Exportação da fila de exceções + decisões da revisão (etapa 4 do app)
# ---------------------------------------------------------------------------

_COR_PENDENTE_FUNDO = "FFF2CC"   # ambar suave: sem decisao ainda
_COR_PENDENTE_TEXTO = "7F6000"
_COR_RELOGIO_FUNDO = "DDEBF7"    # azul suave: provavel erro de relogio
_COR_RELOGIO_TEXTO = "1F3864"
_COR_ZEBRA = "F7F7F7"
_FONTE_PADRAO = "Arial"
_CARACTERES_ILEGAIS_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_CARACTERES_INVALIDOS_ABA = re.compile(r"[\[\]:*?/\\]")

# (cabecalho, coluna da fila, largura minima, largura maxima, alinhamento, formato)
_COLUNAS_FILA_XLSX = [
    ("Colaborador", "nome_colaborador", 24, 40, "left", None),
    ("Empresa", "razao_social", 24, 40, "left", None),
    ("Data", "data", 12, 12, "center", "DD/MM/YYYY"),
    ("Situação", "situacao", 22, 36, "left", None),
    ("Motivo", "motivo_validacao", 45, 70, "left", None),
    ("Decisão", "classificacao", 24, 32, "left", None),
    ("Status", None, 13, 13, "center", None),
    ("Observação", "observacao", 40, 70, "left", None),
]


def nome_aba_seguro(nome: str, existentes=()) -> str:
    """Nome de aba valido no Excel: sem []:*?/\\, ate 31 caracteres, sem repetir (sem diferenciar maiusculas)."""
    base = _CARACTERES_INVALIDOS_ABA.sub(" ", str(nome)).strip().strip("'") or "Planilha"
    base = base[:31]
    usados = {e.lower() for e in existentes}
    candidato, n = base, 1
    while candidato.lower() in usados:
        n += 1
        sufixo = f" ({n})"
        candidato = base[:31 - len(sufixo)] + sufixo
    return candidato


def _texto_seguro(valor) -> str | None:
    """Texto sem caracteres de controle (o openpyxl recusa) e sem virar formula quando comeca com '='."""
    if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
        return None
    return _CARACTERES_ILEGAIS_XML.sub("", str(valor))


def _configurar_impressao(ws, ultima_coluna: int, ultima_linha: int, titulo: str | None = None) -> None:
    """Paisagem, 1 pagina de largura, cabecalho repetido, margens e rodape 'Pagina X de Y'."""
    from openpyxl.utils import get_column_letter

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"
    ws.print_area = f"A1:{get_column_letter(max(ultima_coluna, 1))}{max(ultima_linha, 1)}"
    ws.page_margins.left = ws.page_margins.right = 0.5
    ws.page_margins.top = 0.75
    ws.page_margins.bottom = 0.75
    ws.oddFooter.center.text = "Página &P de &N"
    ws.oddFooter.center.size = 9
    if titulo:
        ws.oddHeader.left.text = titulo
        ws.oddHeader.left.size = 9


def _escrever_tabela_formatada(ws, df: pd.DataFrame, colunas: list[tuple], regras: list[tuple] | None = None,
                               zebra: bool = True, congelar_colunas: int = 1, titulo_impressao: str | None = None) -> None:
    """
    Helper unico de estilo para tabelas de revisao. `colunas` = lista de
    (cabecalho, chave_no_df | None, largura_min, largura_max, alinhamento, number_format);
    chave None deixa a coluna em branco para as `regras`/o chamador preencherem depois.
    `regras` = [(formula, intervalo_de_colunas "A:H", fundo, texto)] em formatacao
    condicional (mais leve que preencher celula a celula); a formula e' escrita para
    a primeira linha de dados e usa referencias relativas.
    """
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    borda = Border(*[Side(style="thin", color=_COR_BORDA)] * 4)
    n_linhas, n_colunas = len(df), len(colunas)

    for j, (cabecalho, *_resto) in enumerate(colunas, start=1):
        cel = ws.cell(row=1, column=j, value=cabecalho)
        cel.font = Font(name=_FONTE_PADRAO, bold=True, color=_COR_TEXTO_CABECALHO)
        cel.fill = PatternFill("solid", fgColor=_COR_CABECALHO)
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cel.border = borda
    ws.row_dimensions[1].height = 30

    maiores = [len(c[0]) for c in colunas]
    for i, (_, linha) in enumerate(df.iterrows(), start=2):
        for j, (_cab, chave, _mn, _mx, alinh, formato) in enumerate(colunas, start=1):
            valor = linha[chave] if chave is not None and chave in df.columns else None
            if isinstance(valor, pd.Timestamp):
                valor = None if pd.isna(valor) else valor.to_pydatetime().date()
            elif isinstance(valor, str) or valor is None or (not hasattr(valor, "item") and pd.isna(valor)):
                valor = _texto_seguro(valor)
            elif hasattr(valor, "item"):  # numpy -> python (numero de verdade, nunca texto)
                valor = None if pd.isna(valor) else valor.item()
            cel = ws.cell(row=i, column=j, value=valor)
            if isinstance(valor, str):
                cel.data_type = "s"  # texto que comeca com "=" nao pode virar formula
                maiores[j - 1] = max(maiores[j - 1], max((len(p) for p in valor.split("\n")), default=0))
            cel.font = Font(name=_FONTE_PADRAO)
            cel.alignment = Alignment(horizontal=alinh, vertical="center" if alinh == "center" else "top", wrap_text=alinh == "left")
            cel.border = borda
            if formato:
                cel.number_format = formato

    for j, (_cab, _chave, minimo, maximo, *_r) in enumerate(colunas, start=1):
        ws.column_dimensions[get_column_letter(j)].width = min(max(maiores[j - 1] + 3, minimo), maximo)

    ultima = get_column_letter(n_colunas)
    ws.freeze_panes = f"{get_column_letter(congelar_colunas + 1)}2"
    ws.auto_filter.ref = f"A1:{ultima}{max(n_linhas + 1, 1)}"
    ws.sheet_view.showGridLines = False

    if n_linhas:
        for formula, intervalo, fundo, texto in (regras or []):
            ini, fim = intervalo.split(":")
            ws.conditional_formatting.add(
                f"{ini}2:{fim}{n_linhas + 1}",
                FormulaRule(formula=[formula], fill=PatternFill("solid", bgColor=fundo, fgColor=fundo), font=Font(color=texto)),
            )
        if zebra:  # por ultimo = menor prioridade; as cores semanticas ganham
            ws.conditional_formatting.add(
                f"A2:{ultima}{n_linhas + 1}",
                FormulaRule(formula=["MOD(ROW(),2)=0"], fill=PatternFill("solid", bgColor=_COR_ZEBRA, fgColor=_COR_ZEBRA)),
            )
    _configurar_impressao(ws, n_colunas, n_linhas + 1, titulo_impressao)


def exportar_fila_excecoes_xlsx(fila: pd.DataFrame, competencia: str, opcoes_decisao: list[str],
                                periodo: str = "", filtros_descricao: str = "",
                                total_geral: int | None = None) -> bytes:
    """
    Planilha de auditoria da revisao de excecoes (so' apresentacao; nada do
    calculo muda): aba "Resumo" (periodo, totais, classificadas x pendentes,
    contagem por decisao e por empresa) e aba "Exceções" (uma linha por dia,
    ordenada por empresa > colaborador > data, com Status Pendente/Classificada).
    `opcoes_decisao` = rotulos validos da coluna Decisao, na ordem de exibicao.
    Para exportar so' um recorte, passe a fila ja filtrada em `fila`, a descricao dos
    filtros em `filtros_descricao` e o tamanho da fila completa em `total_geral`
    (o Resumo mostra "Exibindo N de M"; todas as contagens sao do recorte).
    """
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    colunas_fila = [c for c in ("nome_colaborador", "razao_social", "data", "situacao", "motivo_validacao",
                                "classificacao", "observacao") if c in fila.columns]
    df = fila.reindex(columns=colunas_fila).copy()
    for coluna in ("nome_colaborador", "razao_social", "situacao", "motivo_validacao", "classificacao", "observacao"):
        if coluna in df.columns:
            df[coluna] = df[coluna].fillna("").astype(str)
    if "data" in df.columns:
        df["data"] = pd.to_datetime(df["data"], errors="coerce")
    ordem = [c for c in ("razao_social", "nome_colaborador", "data") if c in df.columns]
    df = df.sort_values(ordem, kind="stable").reset_index(drop=True) if ordem else df

    classificada = df["classificacao"].ne("") if "classificacao" in df.columns else pd.Series(False, index=df.index)
    df["_status"] = np.where(classificada, "Classificada", "Pendente")

    colunas = [(c[0], "_status" if c[0] == "Status" else c[1], *c[2:]) for c in _COLUNAS_FILA_XLSX]
    colunas = [c for c in colunas if c[1] is None or c[1] in df.columns]
    letras = {c[0]: chr(ord("A") + k) for k, c in enumerate(colunas)}
    sit, dec, sts = letras.get("Situação"), letras.get("Decisão"), letras.get("Status")
    ultima = chr(ord("A") + len(colunas) - 1)
    regras = []
    if sit:
        regras.append((f'LEFT(${sit}2,23)="Provável erro de relógio"', f"A:{ultima}", _COR_RELOGIO_FUNDO, _COR_RELOGIO_TEXTO))
        regras.append((f'LEFT(${sit}2,14)="Provável falta"', f"A:{ultima}", _COR_ALERTA_FUNDO, _COR_ALERTA_TEXTO))
    if sts:
        regras.append((f'${sts}2="Pendente"', f"{dec or sts}:{sts}", _COR_PENDENTE_FUNDO, _COR_PENDENTE_TEXTO))

    wb = Workbook()
    ws_resumo = wb.active
    ws_resumo.title = "Resumo"
    ws = wb.create_sheet(nome_aba_seguro("Exceções", [ws_resumo.title]))
    _escrever_tabela_formatada(ws, df, colunas, regras, titulo_impressao=f"Revisão de exceções — {competencia}")

    # --- Resumo ---------------------------------------------------------------
    borda = Border(*[Side(style="thin", color=_COR_BORDA)] * 4)
    fonte = Font(name=_FONTE_PADRAO)
    fonte_neg = Font(name=_FONTE_PADRAO, bold=True)

    def titulo(linha: int, texto: str, largura: int = 2):
        for col in range(1, largura + 1):
            cel = ws_resumo.cell(row=linha, column=col, value=texto if col == 1 else None)
            cel.font = Font(name=_FONTE_PADRAO, bold=True, color=_COR_TEXTO_CABECALHO)
            cel.fill = PatternFill("solid", fgColor=_COR_CABECALHO)
            cel.alignment = Alignment(vertical="center")

    def par(linha: int, rotulo: str, valor, col_valor: int = 2):
        a = ws_resumo.cell(row=linha, column=1, value=_texto_seguro(rotulo))
        b = ws_resumo.cell(row=linha, column=col_valor, value=valor if not isinstance(valor, str) else _texto_seguro(valor))
        for cel in (a, b):
            cel.font, cel.border = fonte, borda
        b.alignment = Alignment(horizontal="right")
        if isinstance(valor, str):
            b.data_type = "s"

    total, n_classif = len(df), int(classificada.sum())
    titulo(1, f"Revisão de exceções — competência {competencia}", 3)
    ws_resumo.row_dimensions[1].height = 26
    ws_resumo["A2"] = f"Período: {periodo}" if periodo else None
    ws_resumo["A2"].font = fonte
    if filtros_descricao or total_geral is not None:
        partes = [f"Filtros: {filtros_descricao or 'Nenhum'}"]
        if total_geral is not None:
            partes.append(f"Exibindo {len(df)} de {total_geral} exceções")
        ws_resumo["A3"] = " · ".join(partes)
        ws_resumo["A3"].font = fonte
    empresas = sorted(e for e in df["razao_social"].unique() if e) if "razao_social" in df.columns else []
    linha = 4
    titulo(linha, "Totais", 3)
    for rotulo, valor in (("Empresas", ", ".join(empresas) if len(empresas) <= 3 else f"{len(empresas)} empresas"),
                          ("Total de exceções", total), ("Classificadas", n_classif), ("Pendentes", total - n_classif)):
        linha += 1
        par(linha, rotulo, valor)
    linha += 2
    titulo(linha, "Por decisão", 3)
    contagem = df["classificacao"].value_counts() if "classificacao" in df.columns else pd.Series(dtype=int)
    for opcao in [o for o in opcoes_decisao if o] + [o for o in contagem.index if o and o not in opcoes_decisao]:
        linha += 1
        par(linha, opcao, int(contagem.get(opcao, 0)))
    linha += 1
    par(linha, "Sem decisão (entram como falta injustificada)", total - n_classif)
    if empresas:
        linha += 2
        titulo(linha, "Por empresa", 3)
        for k, rotulo in enumerate(("Total", "Pendentes"), start=2):
            ws_resumo.cell(row=linha, column=k, value=rotulo)
        for empresa in empresas:
            linha += 1
            da_empresa = df["razao_social"] == empresa
            par(linha, empresa, int(da_empresa.sum()))
            cel = ws_resumo.cell(row=linha, column=3, value=int((da_empresa & ~classificada).sum()))
            cel.font, cel.border, cel.alignment = fonte, borda, Alignment(horizontal="right")
    ws_resumo.column_dimensions["A"].width = min(max([len(c.value) for c in ws_resumo["A"] if isinstance(c.value, str)] + [30]) + 3, 60)
    ws_resumo.column_dimensions["B"].width = 18
    ws_resumo.column_dimensions["C"].width = 14
    ws_resumo.sheet_view.showGridLines = False
    ws_resumo.page_setup.orientation = "portrait"
    ws_resumo.page_setup.fitToWidth, ws_resumo.page_setup.fitToHeight = 1, 0
    ws_resumo.sheet_properties.pageSetUpPr.fitToPage = True
    ws_resumo.oddFooter.center.text = "Página &P de &N"

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


if __name__ == "__main__":
    # --------------------------------------------------------------------
    # AJUSTE OS CAMINHOS ABAIXO PARA OS ARQUIVOS REAIS ANTES DE RODAR
    # --------------------------------------------------------------------
    CAMINHO_AFD = "dados/ponto.txt"
    CAMINHO_CADASTRO = "dados/cadastro_colaboradores.xlsx"
    CAMINHO_CONSIGNADOS = "dados/consignados.xlsx"
    CAMINHO_BANCO_HORAS_SECULLUM = "dados/banco_horas_secullum.xlsx"
    CAMINHO_PLANILHA_ADRIANO = "dados/banco_horas_adriano.xlsx"
    PIS_ADRIANO = "00000000001"
    DATA_INICIO = date(2026, 7, 26)
    DATA_FIM = date(2026, 8, 25)
    COMPETENCIA = "08/2026"

    cadastro = carregar_cadastro_colaboradores(CAMINHO_CADASTRO)

    # Le os arquivos leves primeiro (consignados, banco de horas) e roda a
    # checagem pre-voo ANTES do processamento pesado do AFD - se um PIS
    # estiver errado em algum desses arquivos, a usuaria fica sabendo em
    # segundos, sem esperar o AFD inteiro ser processado pra descobrir.
    consignados = carregar_consignados(CAMINHO_CONSIGNADOS, cadastro=cadastro)
    banco_secullum = ler_banco_horas_secullum(CAMINHO_BANCO_HORAS_SECULLUM, cadastro=cadastro) if Path(CAMINHO_BANCO_HORAS_SECULLUM).exists() else None
    banco_adriano = ler_banco_horas_adriano(CAMINHO_PLANILHA_ADRIANO, PIS_ADRIANO, cadastro=cadastro) if Path(CAMINHO_PLANILHA_ADRIANO).exists() else None

    executar_checagens_preflight(
        cadastro, consignados, banco_secullum, banco_adriano,
        caminho_consignados=CAMINHO_CONSIGNADOS,
        caminho_banco_secullum=CAMINHO_BANCO_HORAS_SECULLUM,
        caminho_banco_adriano=CAMINHO_PLANILHA_ADRIANO,
    )
    banco_horas = consolidar_banco_horas(cadastro["pis"].tolist(), banco_secullum, banco_adriano)

    # Rodada 1: gera as excecoes para o RH validar no dashboard (aqui sim
    # entra o processamento pesado - ler e apurar o AFD inteiro)
    apuracao, fila = rodada_1_gerar_fila_de_excecoes(CAMINHO_AFD, cadastro, DATA_INICIO, DATA_FIM)
    fila_exportada = exportar_fila_validacao_rh(fila, cadastro, "saida/fila_validacao_rh.csv")
    print(f"{len(fila_exportada)} excecao(oes) geradas para validacao do RH -> saida/fila_validacao_rh.csv")
    print("Abra o dashboard_aprovacao.html, carregue esse CSV, classifique as excecoes e exporte decisoes_rh.csv")

    # ---- (o RH classifica no dashboard e exporta decisoes_rh.csv aqui) ----

    # Rodada 2: com as decisoes do RH em maos, gera a Relacao de Valores final
    decisoes = pd.read_csv("saida/decisoes_rh.csv") if Path("saida/decisoes_rh.csv").exists() else None
    relacao_valores = rodada_2_gerar_relacao_de_valores(
        apuracao, cadastro, consignados, banco_horas, decisoes, COMPETENCIA
    )
    arquivos = exportar_por_empresa(relacao_valores, "saida")
    for a in arquivos:
        print(f"Gerado: {a}")
