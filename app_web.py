"""
app_web.py
==========
Versão 100% web (Streamlit) do Sistema de Fechamento de Folha, numa
única tela interativa. Substitui o app_gui.py (Tkinter) e os arquivos
soltos fila_validacao_rh.csv / dashboard_aprovacao.html / decisoes_rh.csv
por um fluxo único, em memória, com estado guardado em st.session_state.

Este arquivo é só a CASCA de interface: toda a lógica de cálculo continua
em afd_parser.py, business_rules.py, banco_horas.py e pipeline.py, sem
nenhuma duplicação — os mesmos módulos que o app_gui.py (Tkinter) usa.

Como rodar (veja também o guia completo na conversa):
    pip install -r requirements_web.txt
    streamlit run app_web.py
"""

from __future__ import annotations

import hashlib
import io
import shutil
import tempfile
import zipfile
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from pipeline import (
    carregar_cadastro_colaboradores,
    carregar_consignados,
    avisos_banco_horas,
    executar_checagens_preflight,
    rodada_1_gerar_fila_de_excecoes,
    rodada_2_gerar_relacao_de_valores,
    exportar_por_empresa,
    COLUNAS_OBRIGATORIAS_CADASTRO,
    UNIDADE_DSR_PADRAO,
    UNIDADES_DSR,
)
import db_sessao
from feriados import avisos_de_cobertura, carregar_feriados, descrever_feriados
from banco_horas import consolidar_banco_horas, ler_banco_horas_secullum, ler_banco_horas_adriano
from cartao_ponto_pdf import _sem_acento
from leitura_arquivos import ler_arquivo_generico
from leitores_ponto import EXTENSOES_ACEITAS as EXTENSOES_PONTO
from validacao import ArquivoInvalidoError

# ---------------------------------------------------------------------------
# Config da página
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Fechamento de Folha — Maçaneiro e Gonzaga / Cianorte Tubos",
    page_icon=":material/receipt_long:",
    layout="wide",
)

db_sessao.inicializar_banco()

STATUS_LEGIVEL = {
    "FALTA_OU_AUSENCIA": "Falta ou ausência",
    "BATIDA_INCOMPLETA": "Batida incompleta",
    "FALTA_PARCIAL_CANDIDATA": "Falta parcial (batidas faltando)",
    "INTERVALO_CURTO": "Intervalo curto",
}
OPCOES_CLASSIFICACAO = [
    "",
    "Falta injustificada",
    "Falta justificada",
    "Erro de relógio — abonar",
    "Atestado médico",
]
LABEL_PARA_CODIGO = {
    "Falta injustificada": "FALTA_INJUSTIFICADA",
    "Falta justificada": "FALTA_JUSTIFICADA",
    "Erro de relógio — abonar": "ERRO_RELOGIO_ABONAR",
    # Atestado médico = falta justificada para o cálculo (business_rules.py não muda):
    # não gera desconto de VA/DSR, mas continua tirando o COPR, como toda falta justificada.
    "Atestado médico": "FALTA_JUSTIFICADA",
}

# ---------------------------------------------------------------------------
# Sessão: pasta temporária única por sessão do navegador
# ---------------------------------------------------------------------------
if "pasta_temp" not in st.session_state:
    st.session_state.pasta_temp = tempfile.mkdtemp(prefix="folha_web_")

PASTA_TEMP = Path(st.session_state.pasta_temp)


def salvar_upload_em_disco(arquivo_subido) -> str:
    """
    Streamlit entrega o arquivo enviado como bytes em memória. As funções
    de carregamento (carregar_cadastro_colaboradores, etc.) já são
    robustas e testadas para ler de um CAMINHO em disco — em vez de
    duplicar essa lógica para ler direto de bytes, gravamos o arquivo
    numa pasta temporária exclusiva desta sessão do navegador e
    reaproveitamos as mesmas funções sem nenhuma alteração. Do ponto de
    vista de quem usa o sistema, isso é invisível: nenhum arquivo aparece
    pra baixar, procurar ou gerenciar manualmente.
    """
    destino = PASTA_TEMP / arquivo_subido.name
    with open(destino, "wb") as f:
        f.write(arquivo_subido.getbuffer())
    return str(destino)


# ---------------------------------------------------------------------------
# Classificação automática dos arquivos soltos na área única de upload
# ---------------------------------------------------------------------------
def _classificar_pdf_ponto(caminho: str) -> str:
    """
    Lê o cabeçalho da 1ª página: "ESPELHO DE PONTO ELETRÔNICO" ou "CARTÃO
    PONTO". Se o texto não puder ser lido, cai para o nome do arquivo.
    """
    cabecalho = ""
    try:
        import pdfplumber

        with pdfplumber.open(caminho) as pdf:
            cabecalho = (pdf.pages[0].extract_text() or "")[:500]
    except Exception:
        pass
    cabecalho = _sem_acento(cabecalho).upper()
    if "ESPELHO DE PONTO" in cabecalho:
        return "espelho_ponto"
    if "CARTAO PONTO" in cabecalho:
        return "cartao_ponto"

    nome = _sem_acento(Path(caminho).name).lower()
    if "espelho" in nome:
        return "espelho_ponto"
    if "cartao" in nome:
        return "cartao_ponto"
    return "desconhecido"


def classificar_arquivo(caminho: str) -> str:
    """
    Devolve um dos rótulos: 'cartao_ponto', 'espelho_ponto', 'afd',
    'cadastro', 'consignados', 'banco_secullum', 'banco_adriano' ou
    'desconhecido'. Para planilhas usa a mesma lógica de detecção de coluna
    já usada dentro dos carregadores (validacao.py / banco_horas.py), então
    a classificação nunca diverge do que o carregador real vai aceitar.
    """
    sufixo = Path(caminho).suffix.lower()
    if sufixo in (".txt", ".afd"):
        return "afd"
    if sufixo == ".pdf":
        return _classificar_pdf_ponto(caminho)
    try:
        df = ler_arquivo_generico(caminho, "arquivo enviado")
    except ArquivoInvalidoError:
        return "desconhecido"

    colunas = set(df.columns)
    if set(COLUNAS_OBRIGATORIAS_CADASTRO).issubset(colunas):
        return "cadastro"
    if {"horas_extras", "horas_debito"}.issubset(colunas):
        return "banco_adriano"
    tem_pis = any("pis" in c or "matricula" in c for c in colunas)
    tem_saldo = any("saldo" in c for c in colunas)
    if tem_pis and tem_saldo:
        return "banco_secullum"
    tem_valor = any("valor" in c or "desconto" in c for c in colunas)
    if tem_pis and tem_valor:
        return "consignados"
    return "desconhecido"


RÓTULO_AMIGÁVEL = {
    "cartao_ponto": "Cartão Ponto (PDF Secullum)",
    "espelho_ponto": "Espelho de Ponto (PDF Secullum)",
    "afd": "Arquivo AFD (.txt — exportação direta do Secullum)",
    "cadastro": "Cadastro de Colaboradores",
    "consignados": "Extrato de Consignados",
    "banco_secullum": "Banco de Horas — Secullum",
    "banco_adriano": "Banco de Horas — Planilha do Adriano",
    "desconhecido": "Não identificado (escolha manualmente)",
}


PREFIXOS_OCORRENCIA_JUSTIFICAM_FALTA = ("ATESTAD",)  # atestado medico - sempre falta justificada


def _sugerir_classificacao(ocorrencia_secullum) -> str:
    """
    Pré-seleciona "Falta justificada" quando o PRÓPRIO Secullum já
    registrou um atestado médico para aquele dia (ver
    `pipeline._anexar_contexto_secullum` - só existe quando o ponto veio
    do Cartão Ponto em PDF). A operadora continua vendo e podendo mudar a
    decisão no dropdown normalmente - isto só poupa o clique mais óbvio,
    nunca fecha nada sem passar pela tela de aprovação.
    """
    if isinstance(ocorrencia_secullum, str) and ocorrencia_secullum.upper().startswith(PREFIXOS_OCORRENCIA_JUSTIFICAM_FALTA):
        return "Falta justificada"
    return ""


def preparar_fila_para_tabela(fila: pd.DataFrame, cadastro: pd.DataFrame) -> pd.DataFrame:
    """
    Mesma transformação de pipeline.exportar_fila_validacao_rh (cruza com
    o cadastro, ordena por nome e depois por data, formata dd/mm/aaaa) —
    só que devolve o DataFrame pronto para a tabela interativa em vez de
    gravar um CSV em disco. Também já adiciona as colunas editáveis
    'classificacao' e 'observacao' que a operadora preenche na tela.
    """
    if fila.empty:
        return fila.assign(nome_colaborador=pd.Series(dtype=str), razao_social=pd.Series(dtype=str))

    info = cadastro[["pis", "nome", "razao_social"]].rename(columns={"nome": "nome_colaborador"})
    fila_com_nome = fila.merge(info, on="pis", how="left")
    fila_com_nome["data"] = pd.to_datetime(fila_com_nome["data"])
    fila_ordenada = fila_com_nome.sort_values(["nome_colaborador", "data"]).reset_index(drop=True)
    fila_ordenada["data_exibicao"] = fila_ordenada["data"].dt.strftime("%d/%m/%Y")
    fila_ordenada["situacao"] = fila_ordenada["status"].map(STATUS_LEGIVEL).fillna(fila_ordenada["status"])
    if "cruzamento_espelho" in fila_ordenada.columns:
        # Com o Espelho de Ponto, "Falta ou ausência" vira "Provável erro de relógio" ou "Provável falta real".
        fila_ordenada["situacao"] = fila_ordenada["cruzamento_espelho"].fillna(fila_ordenada["situacao"])
    fila_ordenada["classificacao"] = (
        fila_ordenada["ocorrencia_secullum"].map(_sugerir_classificacao)
        if "ocorrencia_secullum" in fila_ordenada.columns else ""
    )
    fila_ordenada["observacao"] = ""
    return fila_ordenada



# ---------------------------------------------------------------------------
# Apoio de interface (só apresentação — nenhuma regra de negócio aqui)
# ---------------------------------------------------------------------------
LARGURA_LEITURA = 900  # px — largura máxima das seções de texto; só a tabela de exceções usa a página toda
ETAPAS = ["Arquivos", "Período", "Processamento", "Revisão", "Download"]
TIPOS_ARQUIVO_UNICO = (
    "cadastro", "cartao_ponto", "espelho_ponto", "afd", "banco_secullum", "banco_adriano",
)
CHAVE_EDITOR = "editor_excecoes"
CHAVES_FILTRO = ("filtro_empresa", "filtro_colaborador", "filtro_situacao")


@contextmanager
def secao_leitura(rotulo: str, aberta: bool):
    """Expander de etapa com o conteúdo limitado à largura de leitura."""
    with st.expander(rotulo, expanded=aberta):
        with st.container(width=LARGURA_LEITURA):
            yield


def passo_atual() -> int:
    """Etapa ativa (1 a 5), derivada só do que já existe no session_state."""
    if st.session_state.get("arquivos_finais"):
        return 5
    if "fila_tabela" in st.session_state:
        return 4
    if st.session_state.get("preflight_ok"):
        return 3
    if st.session_state.get("arquivos_confirmados"):
        return 2
    return 1


def agendar_mensagem(nivel: str, texto: str) -> None:
    """Guarda uma mensagem para aparecer no topo depois do st.rerun() de uma transição de etapa."""
    st.session_state.setdefault("mensagens_transicao", []).append((nivel, texto))


def resumo_periodo() -> str:
    inicio = st.session_state.get("data_inicio_input")
    fim = st.session_state.get("data_fim_input")
    competencia = st.session_state.get("competencia_input", "")
    if inicio and fim:
        return f"{inicio:%d/%m/%Y} a {fim:%d/%m/%Y} · competência {competencia}"
    return f"competência {competencia}"


def contar_decisoes(fila: pd.DataFrame) -> tuple[int, int]:
    """(classificadas, total) da fila de exceções."""
    return int((fila["classificacao"] != "").sum()), len(fila)


def filtrar_fila(fila: pd.DataFrame, empresas, colaboradores, situacoes) -> pd.DataFrame:
    """Filtro SÓ de exibição: devolve um recorte que mantém o índice original da fila."""
    mascara = pd.Series(True, index=fila.index)
    if empresas:
        mascara &= fila["razao_social"].isin(empresas)
    if colaboradores:
        mascara &= fila["nome_colaborador"].isin(colaboradores)
    if situacoes:
        mascara &= fila["situacao"].isin(situacoes)
    return fila[mascara]


def limpar_estado_editor() -> None:
    """Descarta as edições pendentes do data_editor (sem isso elas seriam reaplicadas por cima de novos dados)."""
    st.session_state.pop(CHAVE_EDITOR, None)
    chave_ativa = st.session_state.pop("_chave_editor_ativa", None)
    if chave_ativa:
        st.session_state.pop(chave_ativa, None)


def aplicar_decisao_em_lote() -> None:
    """on_click do botão 'Aplicar': grava a mesma decisão nas linhas atualmente filtradas, por índice."""
    fila = st.session_state.fila_tabela
    filtradas = filtrar_fila(
        fila,
        st.session_state.get("filtro_empresa"),
        st.session_state.get("filtro_colaborador"),
        st.session_state.get("filtro_situacao"),
    )
    fila.loc[filtradas.index, "classificacao"] = st.session_state.decisao_em_lote
    limpar_estado_editor()


def estilo_situacao(valor: str) -> str:
    """Cor de fundo translúcida (legível nos modos claro e escuro) para as situações do cruzamento com o Espelho."""
    if valor.startswith("Provável erro de relógio"):
        return "background-color: rgba(59, 130, 246, 0.22)"
    if valor.startswith("Provável falta"):
        return "background-color: rgba(239, 68, 68, 0.22)"
    return ""


def empresa_do_arquivo(nome_arquivo: str, cadastro: pd.DataFrame) -> str:
    """Razão social da empresa a que o arquivo gerado pertence (o nome do arquivo traz a razão com '_')."""
    for razao in cadastro["razao_social"].dropna().unique():
        if razao.replace(" ", "_") in nome_arquivo:
            return razao
    return "Empresa"


# ---------------------------------------------------------------------------
# Cabeçalho, progresso e áreas reservadas (preenchidas no fim do script, com o estado já atualizado)
# ---------------------------------------------------------------------------
st.title("Fechamento de Folha de Pagamento")
st.caption("Maçaneiro e Gonzaga LTDA · Cianorte Tubos LTDA")

passo = passo_atual()
with st.container(horizontal=True):
    for numero, nome_etapa in enumerate(ETAPAS, start=1):
        if numero < passo:
            st.badge(f"{numero}. {nome_etapa}", icon=":material/check:", color="green")
        elif numero == passo:
            st.badge(f"{numero}. {nome_etapa}", color="primary")
        else:
            st.badge(f"{numero}. {nome_etapa}", color="gray")

slot_mensagens = st.container()
slot_metricas = st.container()
slot_avisos = st.container()

st.sidebar.subheader("Fechamento")
sb_competencia = st.sidebar.container()
sb_rascunho = st.sidebar.container()
sb_acoes = st.sidebar.container()

with slot_mensagens:
    for nivel, texto in st.session_state.pop("mensagens_transicao", []):
        getattr(st, nivel)(texto)

# Avisos que não impedem o fechamento; são reunidos aqui e mostrados num único expander no fim do script.
avisos: list[tuple[str, str]] = []

with sb_competencia:
    st.caption("Competência")
    if not st.session_state.get("arquivos_confirmados"):
        st.write("Será definida após o envio dos arquivos.")

# ===========================================================================
# ETAPA 1 — Upload centralizado (arrastar e soltar)
# ===========================================================================
resumo_arquivos = ""
if passo > 1:
    qtd_arquivos = sum(len(v) for v in st.session_state.arquivos_confirmados.values())
    resumo_arquivos = f" — {qtd_arquivos} arquivo(s) confirmado(s)"

with secao_leitura(f"1. Arquivos{resumo_arquivos}", passo == 1):
    st.write(
        "Envie todos os arquivos do fechamento de uma vez: Cadastro de Colaboradores, "
        "Cartão Ponto e Espelho de Ponto (PDF) ou o AFD (.txt), Extrato de Consignados e, se houver, "
        "o Banco de Horas do Secullum e/ou a planilha semanal do Adriano. "
        "O sistema identifica o tipo de cada arquivo."
    )

    arquivos_subidos = st.file_uploader(
        "Arraste os arquivos aqui ou clique para selecionar",
        type=[ext.lstrip(".") for ext in EXTENSOES_PONTO] + ["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="uploader_geral",
    )

    # Cada tipo guarda uma LISTA de caminhos, não um único - várias empresas
    # costumam emitir extratos de consignados separados (um por CNPJ), por
    # exemplo, e nada deve ser descartado silenciosamente se a pessoa soltar
    # dois arquivos classificados com o mesmo tipo.
    tipos_confirmados: dict[str, list[str]] = {}
    if arquivos_subidos:
        linhas_classificacao = []
        for arq in arquivos_subidos:
            caminho = salvar_upload_em_disco(arq)
            tipo_detectado = classificar_arquivo(caminho)
            linhas_classificacao.append({"arquivo": arq.name, "caminho": caminho, "tipo": tipo_detectado})

        st.write("**Arquivos recebidos** — ajuste o tipo se algo foi identificado incorretamente.")
        opcoes_tipo = list(RÓTULO_AMIGÁVEL.keys())
        for i, item in enumerate(linhas_classificacao):
            with st.container(border=True):
                col_nome, col_status, col_tipo = st.columns([4, 2, 4], vertical_alignment="center")
                with col_nome:
                    st.write(item["arquivo"])
                    st.caption(f"Detectado: {RÓTULO_AMIGÁVEL[item['tipo']]}")
                with col_tipo:
                    escolha = st.selectbox(
                        f"Tipo — {item['arquivo']}",
                        options=opcoes_tipo,
                        index=opcoes_tipo.index(item["tipo"]),
                        format_func=lambda t: RÓTULO_AMIGÁVEL[t],
                        key=f"tipo_{item['arquivo']}_{i}",
                        label_visibility="collapsed",
                    )
                with col_status:
                    if escolha == "desconhecido":
                        st.badge("Requer revisão", icon=":material/warning:", color="orange")
                    elif escolha != item["tipo"]:
                        st.badge("Tipo ajustado", icon=":material/check:", color="blue")
                    else:
                        st.badge("Identificado", icon=":material/check:", color="green")
            tipos_confirmados.setdefault(escolha, []).append(item["caminho"])

    # Checklist dos tipos obrigatórios, sempre visível antes de confirmar
    tem_ponto_ck = "afd" in tipos_confirmados or "cartao_ponto" in tipos_confirmados
    requisitos = [
        ("Cadastro de Colaboradores", "cadastro" in tipos_confirmados),
        ("Ponto: Cartão Ponto (PDF) ou Arquivo AFD (.txt)", tem_ponto_ck),
        ("Extrato de Consignados", "consignados" in tipos_confirmados),
    ]
    st.write("**Arquivos obrigatórios**")
    for descricao, presente in requisitos:
        if presente:
            st.markdown(f":green[:material/check_circle:] {descricao}")
        else:
            st.markdown(f":orange[:material/pending:] {descricao} — pendente")

    if "espelho_ponto" in tipos_confirmados and "cartao_ponto" not in tipos_confirmados:
        avisos.append((
            "warning",
            "O Espelho de Ponto foi enviado sem o Cartão Ponto correspondente. A conferência cruzada ficará limitada.",
        ))

    if arquivos_subidos:
        # Tipos dos quais só faz sentido existir UM arquivo por fechamento -
        # mais de um classificado assim é quase certamente um erro de
        # classificação (ex.: dois cadastros enviados por engano), e seguir
        # silenciosamente usando só o primeiro escondería o problema.
        pis_adriano = ""
        if "banco_adriano" in tipos_confirmados:
            pis_adriano = st.text_input(
                "PIS/matrícula do Adriano (a planilha dele não traz essa informação)",
                key="pis_adriano_input",
            )

        if st.button("Confirmar arquivos e continuar", type="primary"):
            # Slot "ponto": AFD sozinho, Cartão Ponto sozinho, ou Cartão + Espelho (conferência
            # cruzada completa). O Espelho sem Cartão nem AFD não tem o que cruzar.
            tem_ponto = "afd" in tipos_confirmados or "cartao_ponto" in tipos_confirmados
            faltando = [t for t in ("cadastro", "consignados") if t not in tipos_confirmados]
            if not tem_ponto:
                faltando.insert(1, "cartao_ponto")
            duplicados = [t for t in TIPOS_ARQUIVO_UNICO if len(tipos_confirmados.get(t, [])) > 1]
            if faltando:
                st.error(
                    "Ainda faltam arquivos obrigatórios: "
                    + ", ".join(RÓTULO_AMIGÁVEL[t] for t in faltando)
                    + (" (ou o Arquivo AFD)" if "cartao_ponto" in faltando else "")
                    + ". Envie-os e classifique corretamente antes de continuar."
                )
            elif duplicados:
                st.error(
                    "Só é esperado um arquivo de cada um destes tipos, mas mais de um foi "
                    "classificado assim: " + ", ".join(RÓTULO_AMIGÁVEL[t] for t in duplicados) + ". "
                    "Confira a classificação de cada arquivo acima (se houver mais de um extrato de "
                    "Consignados, isso é esperado — um por empresa/CNPJ, por exemplo — e não gera este erro)."
                )
            elif "banco_adriano" in tipos_confirmados and not pis_adriano.strip():
                st.error("A planilha do Adriano foi enviada, mas o PIS/matrícula dele não foi informado.")
            else:
                st.session_state.arquivos_confirmados = tipos_confirmados
                st.session_state.pis_adriano = pis_adriano.strip()
                # limpa qualquer processamento anterior, já que os arquivos mudaram
                for chave in ("cadastro_df", "preflight_ok", "apuracao_diaria", "fila_tabela", "arquivos_finais",
                              "feriados_periodo", "avisos_feriados", "avisos_banco_horas"):
                    st.session_state.pop(chave, None)
                agendar_mensagem("success", "Arquivos confirmados. Defina o período na etapa 2.")
                st.rerun()

# ===========================================================================
# ETAPA 2 — Período + checagem pré-voo instantânea
# ===========================================================================
if st.session_state.get("arquivos_confirmados"):
    resumo_etapa2 = f" — {resumo_periodo()}" if passo > 2 else ""
    with secao_leitura(f"2. Período e validação{resumo_etapa2}", passo == 2):
        col_a, col_b = st.columns(2)
        with col_a:
            data_inicio = st.date_input("Início do período", key="data_inicio_input")
        with col_b:
            data_fim = st.date_input("Fim do período", key="data_fim_input")
        with sb_competencia:
            competencia = st.text_input("Competência (mm/aaaa)", value=data_fim.strftime("%m/%Y"), key="competencia_input")
        st.caption("A competência é informada na barra lateral.")

        if st.button("Validar dados", type="primary"):
            arquivos = st.session_state.arquivos_confirmados
            try:
                cadastro = carregar_cadastro_colaboradores(arquivos["cadastro"][0])
                consignados = carregar_consignados(arquivos["consignados"], cadastro=cadastro)
                banco_secullum = (
                    ler_banco_horas_secullum(arquivos["banco_secullum"][0], cadastro=cadastro)
                    if "banco_secullum" in arquivos else None
                )
                banco_adriano = (
                    ler_banco_horas_adriano(arquivos["banco_adriano"][0], st.session_state.pis_adriano, cadastro=cadastro)
                    if "banco_adriano" in arquivos
                    else None
                )
                executar_checagens_preflight(
                    cadastro, consignados, banco_secullum, banco_adriano,
                    caminho_consignados=", ".join(arquivos["consignados"]),
                    caminho_banco_secullum=arquivos.get("banco_secullum", [""])[0],
                    caminho_banco_adriano=arquivos.get("banco_adriano", [""])[0],
                )
            except ArquivoInvalidoError as e:
                st.session_state.preflight_ok = False
                st.error(f"**Não foi possível continuar — há um problema nos dados:**\n\n{e}")
            else:
                st.session_state.preflight_ok = True
                st.session_state.cadastro_df = cadastro
                st.session_state.consignados_df = consignados
                st.session_state.banco_horas_df = consolidar_banco_horas(
                    cadastro["pis"].tolist(), banco_secullum, banco_adriano
                )
                st.session_state.avisos_banco_horas = avisos_banco_horas(cadastro, banco_secullum, banco_adriano)
                agendar_mensagem("success", "Validação concluída: os dados são consistentes. Prossiga para o processamento.")
                st.rerun()

    # Avisos (nao impedem de continuar): ficam na tela depois do rerun, por isso saem do session_state.
    if st.session_state.get("preflight_ok"):
        for aviso in st.session_state.get("avisos_banco_horas", []):
            avisos.append(("warning", aviso))

# ===========================================================================
# ETAPA 3 — Processar o ponto (etapa pesada, só roda sob demanda)
# ===========================================================================
if st.session_state.get("preflight_ok"):
    resumo_etapa3 = ""
    if passo > 3:
        resumo_etapa3 = f" — {len(st.session_state.fila_tabela)} exceção(ões) para revisão"
    with secao_leitura(f"3. Processamento{resumo_etapa3}", passo == 3):
        competencia_atual = st.session_state.get("competencia_input")
        sessao_salva = db_sessao.obter_sessao(competencia_atual) if competencia_atual else None
        if sessao_salva and sessao_salva["status"] == "em_andamento":
            decisoes_salvas = db_sessao.carregar_decisoes(competencia_atual)
            n = int((decisoes_salvas["classificacao"] != "").sum())
            ultimo_save = datetime.fromisoformat(sessao_salva["atualizado_em"]).strftime("%d/%m/%Y às %H:%M")
            st.info(
                f"**Rascunho encontrado — fechamento {competencia_atual}**  \n"
                f"{n} de {sessao_salva['total_excecoes']} exceções já classificadas. "
                f"Último save: {ultimo_save}.  \n"
                f"Clique em **Processar** para recarregar e continuar de onde parou, "
                f"ou use **Recomeçar** na barra lateral para descartar."
            )

        incluir_facultativos = st.checkbox(
            "Contar pontos facultativos (Carnaval, Corpus Christi, vésperas de Natal e Ano Novo) como feriado",
            value=False,
            key="incluir_facultativos",
            help=(
                "Marque só se a empresa ou a convenção coletiva concede esses dias. Os feriados nacionais e os de "
                "Cianorte já são sempre considerados (calendário em feriados.csv). Clique em Processar de novo "
                "depois de mudar."
            ),
        )

        if st.button("Processar ponto e calcular fechamento", type="primary"):
            try:
                feriados = carregar_feriados(incluir_facultativos=incluir_facultativos)
                avisos_feriados = avisos_de_cobertura(data_inicio, data_fim, incluir_facultativos=incluir_facultativos)
            except ArquivoInvalidoError as e:
                st.error(f"**Não foi possível ler o calendário de feriados:** {e}")
                st.stop()
            st.session_state.feriados_periodo = descrever_feriados(feriados, data_inicio, data_fim)
            st.session_state.avisos_feriados = avisos_feriados

            with st.spinner("Lendo o ponto e apurando dia a dia. Isto pode levar alguns segundos..."):
                arquivos = st.session_state.arquivos_confirmados
                # AFD (.txt) tem prioridade quando enviado; senão usa o Cartão Ponto (PDF).
                caminho_ponto = (arquivos.get("afd") or arquivos["cartao_ponto"])[0]
                apuracao, fila = rodada_1_gerar_fila_de_excecoes(
                    caminho_ponto, st.session_state.cadastro_df, data_inicio, data_fim,
                    caminho_espelho=arquivos.get("espelho_ponto", [None])[0],
                    feriados=feriados,
                )
                st.session_state.apuracao_diaria = apuracao
                fila_tabela = preparar_fila_para_tabela(fila, st.session_state.cadastro_df)

                # Restaura decisões salvas (cruza por pis + data) e registra a sessão
                if not fila_tabela.empty:
                    salvas = db_sessao.carregar_decisoes(competencia)
                    if not salvas.empty:
                        salvas = salvas.rename(columns={"data": "data_exibicao"})
                        fila_tabela = fila_tabela.merge(
                            salvas.rename(columns={"classificacao": "_cls", "observacao": "_obs"}),
                            on=["pis", "data_exibicao"], how="left",
                        )
                        tem = fila_tabela["_cls"].notna()
                        fila_tabela.loc[tem, "classificacao"] = fila_tabela.loc[tem, "_cls"]
                        fila_tabela.loc[tem, "observacao"] = fila_tabela.loc[tem, "_obs"]
                        fila_tabela = fila_tabela.drop(columns=["_cls", "_obs"])
                        st.session_state.rascunho_restaurado = int(tem.sum())
                    db_sessao.criar_ou_atualizar_sessao(
                        competencia, len(fila_tabela),
                        "concluido" if (fila_tabela["classificacao"] != "").all() else "em_andamento",
                    )

                st.session_state.fila_tabela = fila_tabela
                st.session_state.competencia_final = competencia
                st.session_state.pop("editor_excecoes", None)
                limpar_estado_editor()
                for chave in CHAVES_FILTRO:
                    st.session_state.pop(chave, None)
            agendar_mensagem(
                "success",
                f"Processamento concluído. {len(st.session_state.fila_tabela)} exceção(ões) encontrada(s) para revisão.",
            )
            restauradas = st.session_state.pop("rascunho_restaurado", 0)
            if restauradas:
                agendar_mensagem(
                    "info",
                    f"Rascunho restaurado: {restauradas} decisão(ões) anterior(es) recuperada(s). "
                    "Pode continuar de onde parou.",
                )
            st.rerun()

    if "feriados_periodo" in st.session_state:
        if st.session_state.feriados_periodo:
            avisos.append((
                "info",
                "**Feriados considerados neste período** (ausência neles não é falta; trabalho neles vale "
                "hora extra 100%): " + " · ".join(st.session_state.feriados_periodo),
            ))
        else:
            avisos.append(("info", "Nenhum feriado considerado neste período."))
        for aviso in st.session_state.get("avisos_feriados", []):
            avisos.append(("warning", aviso))

# ===========================================================================
# ETAPA 4 — Tabela de aprovação embutida (sem CSV intermediário)
# ===========================================================================
if "fila_tabela" in st.session_state:
    resumo_etapa4 = ""
    if passo > 4:
        c_res, t_res = contar_decisoes(st.session_state.fila_tabela)
        resumo_etapa4 = f" — {c_res} de {t_res} exceções classificadas"
    with st.expander(f"4. Revisão de exceções{resumo_etapa4}", expanded=passo == 4):
        if st.session_state.fila_tabela.empty:
            st.success("Nenhuma exceção encontrada neste período. Prossiga para a geração do fechamento.")
        else:
            with st.container(width=LARGURA_LEITURA):
                st.write(
                    "Só os dias que o sistema não conseguiu decidir sozinho aparecem aqui. "
                    "Classifique cada linha no menu da coluna Decisão: o que ficar em branco entra como "
                    "**falta injustificada** por padrão (o mesmo critério conservador de sempre)."
                )

            fila_completa = st.session_state.fila_tabela

            # --- Filtros (só de exibição) ---------------------------------------
            opcoes_filtro = {
                "filtro_empresa": sorted(fila_completa["razao_social"].dropna().unique()),
                "filtro_colaborador": sorted(fila_completa["nome_colaborador"].dropna().unique()),
                "filtro_situacao": sorted(fila_completa["situacao"].dropna().unique()),
            }
            for chave, opcoes in opcoes_filtro.items():
                if chave in st.session_state:  # descarta seleções que não existem mais na fila atual
                    st.session_state[chave] = [v for v in st.session_state[chave] if v in opcoes]
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                filtro_empresa = st.multiselect(
                    "Empresa", opcoes_filtro["filtro_empresa"], key="filtro_empresa", placeholder="Todas",
                )
            with col_f2:
                filtro_colaborador = st.multiselect(
                    "Colaborador", opcoes_filtro["filtro_colaborador"], key="filtro_colaborador", placeholder="Todos",
                )
            with col_f3:
                filtro_situacao = st.multiselect(
                    "Situação", opcoes_filtro["filtro_situacao"], key="filtro_situacao", placeholder="Todas",
                )
            fila_filtrada = filtrar_fila(fila_completa, filtro_empresa, filtro_colaborador, filtro_situacao)
            filtrando = len(fila_filtrada) != len(fila_completa) or any(
                (filtro_empresa, filtro_colaborador, filtro_situacao)
            )

            with st.container(horizontal=True, vertical_alignment="bottom"):
                st.selectbox(
                    "Decisão para todas as linhas filtradas",
                    options=OPCOES_CLASSIFICACAO[1:],
                    key="decisao_em_lote",
                    width=300,
                )
                st.button(
                    f"Aplicar a {len(fila_filtrada)} linha(s)",
                    on_click=aplicar_decisao_em_lote,
                    disabled=fila_filtrada.empty,
                    help="Substitui a decisão de todas as linhas exibidas na tabela, inclusive as já classificadas.",
                )
            st.caption(
                f"Exibindo {len(fila_filtrada)} de {len(fila_completa)} exceções. "
                "Fundo azul: provável erro de relógio. Fundo vermelho: provável falta real."
            )

            # Cada combinação de filtros usa seu próprio data_editor: as edições pendentes do widget são
            # posicionais e, se o recorte mudasse com a mesma chave, cairiam em linhas erradas. As decisões já
            # feitas não se perdem porque são gravadas em st.session_state.fila_tabela a cada rerun (abaixo).
            if filtrando:
                assinatura = repr((sorted(filtro_empresa), sorted(filtro_colaborador), sorted(filtro_situacao)))
                chave_editor = f"{CHAVE_EDITOR}_{hashlib.md5(assinatura.encode()).hexdigest()[:8]}"
            else:
                chave_editor = CHAVE_EDITOR
            chave_anterior = st.session_state.get("_chave_editor_ativa")
            if chave_anterior and chave_anterior != chave_editor:
                st.session_state.pop(chave_anterior, None)
            st.session_state["_chave_editor_ativa"] = chave_editor

            colunas_exibidas = ["nome_colaborador", "razao_social", "data_exibicao", "situacao", "motivo_validacao", "classificacao", "observacao"]
            tabela_editada = st.data_editor(
                fila_filtrada[colunas_exibidas].style.map(estilo_situacao, subset=["situacao"]),
                key=chave_editor,
                hide_index=True,
                width="stretch",
                disabled=["nome_colaborador", "razao_social", "data_exibicao", "situacao", "motivo_validacao"],
                column_config={
                    "nome_colaborador": st.column_config.TextColumn("Colaborador"),
                    "razao_social": st.column_config.TextColumn("Empresa"),
                    "data_exibicao": st.column_config.TextColumn("Data"),
                    "situacao": st.column_config.TextColumn("Situação"),
                    "motivo_validacao": st.column_config.TextColumn("Motivo"),
                    "classificacao": st.column_config.SelectboxColumn("Decisão", options=OPCOES_CLASSIFICACAO, required=False),
                    "observacao": st.column_config.TextColumn("Observação (opcional)"),
                },
            )
            # guarda de volta no session_state por ÍNDICE (não por posição) —
            # se a operadora clicar num cabeçalho pra ordenar a tabela na tela,
            # uma atribuição posicional (.values) juntaria a decisão certa com
            # a linha errada. .loc por índice alinha corretamente mesmo assim.
            # O recorte filtrado mantém o índice original, então o mesmo vale para os filtros.
            st.session_state.fila_tabela.loc[tabela_editada.index, colunas_exibidas] = tabela_editada[colunas_exibidas]

            # Salva automaticamente no SQLite (só grava o que mudou desde o último save)
            competencia_salva = st.session_state.competencia_final
            fila_atual = st.session_state.fila_tabela
            decididas = fila_atual[fila_atual["classificacao"] != ""]
            houve_mudanca = db_sessao.salvar_decisoes_em_lote(
                competencia_salva,
                decididas.rename(columns={"data_exibicao": "data"})[["pis", "data", "classificacao", "observacao"]],
            )
            classificadas = len(decididas)
            total = len(fila_atual)
            status = "concluido" if classificadas == total else "em_andamento"
            sessao = db_sessao.obter_sessao(competencia_salva)
            if houve_mudanca or sessao is None or sessao["status"] != status or sessao["total_excecoes"] != total:
                db_sessao.criar_ou_atualizar_sessao(competencia_salva, total, status)
                sessao = db_sessao.obter_sessao(competencia_salva)

            pendentes = total - classificadas
            if pendentes:
                st.info(f"{pendentes} linha(s) sem decisão serão tratadas como falta injustificada por padrão.")
            elif sessao:
                st.success("Todas as exceções foram classificadas. O fechamento final já pode ser gerado.")

# ===========================================================================
# ETAPA 5 — Gerar e baixar o Excel final
# ===========================================================================
if "apuracao_diaria" in st.session_state:
    st.subheader("5. Download")
    with st.container(width=LARGURA_LEITURA):
        ROTULOS_UNIDADE_DSR = {
            "dias": "Em dias de DSR (padrão)",
            "valor": "Em valor (R$) — exige a coluna valor_dia_dsr no cadastro",
        }
        unidade_dsr = st.radio(
            "Como lançar o Desconto DSR (8794)?",
            options=list(UNIDADES_DSR),
            index=UNIDADES_DSR.index(UNIDADE_DSR_PADRAO),
            format_func=ROTULOS_UNIDADE_DSR.get,
            key="unidade_dsr",
            help=(
                "Em dias: o 8794 sai com a quantidade de DSR perdidos (uma por semana com falta de dia inteiro "
                "sem justificativa). Em valor: essa quantidade vezes o valor do dia de DSR de cada colaborador, "
                "lido da coluna valor_dia_dsr do cadastro. Confirme com a contabilidade qual unidade ela espera."
            ),
        )

        if st.button("Gerar fechamento final", type="primary" if not st.session_state.get("arquivos_finais") else "secondary"):
            with st.spinner("Consolidando os dados e montando as planilhas..."):
                fila = st.session_state.fila_tabela
                if not fila.empty:
                    decididas = fila[fila["classificacao"] != ""].copy()
                    decisoes = pd.DataFrame({
                        "pis": decididas["pis"],
                        "data": decididas["data_exibicao"],
                        "classificacao": decididas["classificacao"].map(LABEL_PARA_CODIGO),
                    })
                else:
                    decisoes = None

                try:
                    relacao = rodada_2_gerar_relacao_de_valores(
                        st.session_state.apuracao_diaria,
                        st.session_state.cadastro_df,
                        st.session_state.consignados_df,
                        st.session_state.banco_horas_df,
                        decisoes,
                        st.session_state.competencia_final,
                        unidade_dsr=unidade_dsr,
                    )
                except ArquivoInvalidoError as e:
                    st.error(f"**Não foi possível gerar o fechamento:**\n\n{e}")
                    st.stop()
                pasta_saida = PASTA_TEMP / "saida"
                if pasta_saida.exists():
                    shutil.rmtree(pasta_saida)
                arquivos_gerados = exportar_por_empresa(relacao, str(pasta_saida))

                arquivos_finais = {}
                for caminho in arquivos_gerados:
                    arquivos_finais[caminho.name] = caminho.read_bytes()
                st.session_state.arquivos_finais = arquivos_finais
            agendar_mensagem(
                "success",
                f"{len(st.session_state.arquivos_finais)} arquivo(s) gerado(s). Disponíveis para download na etapa 5.",
            )
            st.rerun()

        if st.session_state.get("arquivos_finais"):
            arquivos_finais = st.session_state.arquivos_finais
            varios = len(arquivos_finais) > 1

            if varios:
                buffer_zip = io.BytesIO()
                with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                    for nome, conteudo in arquivos_finais.items():
                        zf.writestr(nome, conteudo)
                st.download_button(
                    "Baixar todos os arquivos (.zip)",
                    data=buffer_zip.getvalue(),
                    file_name="relacao_de_valores_todas_empresas.zip",
                    mime="application/zip",
                    type="primary",
                    icon=":material/download:",
                )

            for nome, conteudo in arquivos_finais.items():
                with st.container(border=True):
                    col_info, col_baixar = st.columns([3, 1], vertical_alignment="center")
                    with col_info:
                        st.write(f"**{empresa_do_arquivo(nome, st.session_state.cadastro_df)}**")
                        st.caption(nome)
                    with col_baixar:
                        st.download_button(
                            "Baixar",
                            data=conteudo,
                            file_name=nome,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="secondary" if varios else "primary",
                            key=f"download_{nome}",
                        )

# ===========================================================================
# Áreas reservadas: métricas, avisos e barra lateral (já com o estado desta execução)
# ===========================================================================
with slot_metricas:
    if "cadastro_df" in st.session_state:
        cadastro_ref = st.session_state.cadastro_df
        fila_ref = st.session_state.get("fila_tabela")
        classificadas_ref, total_ref = contar_decisoes(fila_ref) if fila_ref is not None else (0, 0)
        colunas_metricas = st.columns(6)
        colunas_metricas[0].metric("Colaboradores", len(cadastro_ref), border=True)
        colunas_metricas[1].metric("Empresas", cadastro_ref["razao_social"].nunique(), border=True)
        colunas_metricas[2].metric("Exceções", total_ref if fila_ref is not None else "—", border=True)
        colunas_metricas[3].metric("Classificadas", classificadas_ref if fila_ref is not None else "—", border=True)
        colunas_metricas[4].metric("Pendentes", total_ref - classificadas_ref if fila_ref is not None else "—", border=True)
        colunas_metricas[5].metric(
            "Feriados no período",
            len(st.session_state.feriados_periodo) if "feriados_periodo" in st.session_state else "—",
            border=True,
        )
        if fila_ref is not None and total_ref:
            st.progress(
                classificadas_ref / total_ref,
                text=f"Exceções classificadas: {classificadas_ref} de {total_ref}",
            )

with slot_avisos:
    if avisos:
        with st.expander(f"Avisos ({len(avisos)})"):
            for nivel, texto in avisos:
                icone = ":material/warning:" if nivel == "warning" else ":material/info:"
                getattr(st, nivel)(texto, icon=icone)

with sb_rascunho:
    st.divider()
    st.caption("Rascunho")
    if "fila_tabela" in st.session_state and "competencia_final" in st.session_state:
        classificadas_sb, total_sb = contar_decisoes(st.session_state.fila_tabela)
        sessao_sb = db_sessao.obter_sessao(st.session_state.competencia_final)
        if sessao_sb:
            ultimo_save_sb = datetime.fromisoformat(sessao_sb["atualizado_em"]).strftime("%d/%m/%Y às %H:%M")
            st.write(f"Último save: {ultimo_save_sb}")
        else:
            st.write("Sem alterações salvas ainda.")
        if total_sb:
            st.progress(classificadas_sb / total_sb, text=f"{classificadas_sb} de {total_sb} classificadas")
        else:
            st.write("Nenhuma exceção a classificar.")
    else:
        st.write("Nenhum rascunho em andamento.")

with sb_acoes:
    if "fila_tabela" in st.session_state and "competencia_final" in st.session_state:
        st.divider()
        with st.popover("Recomeçar", icon=":material/restart_alt:", width="stretch"):
            st.write(
                f"Descartar todas as decisões do fechamento {st.session_state.competencia_final}? "
                "Esta ação não pode ser desfeita."
            )
            if st.button("Confirmar descarte", type="primary", key="confirmar_recomecar"):
                db_sessao.limpar_sessao(st.session_state.competencia_final)
                for chave in ("fila_tabela", "apuracao_diaria", "arquivos_finais", "editor_excecoes",
                              "feriados_periodo", "avisos_feriados"):
                    st.session_state.pop(chave, None)
                limpar_estado_editor()
                for chave in CHAVES_FILTRO:
                    st.session_state.pop(chave, None)
                st.rerun()

# ---------------------------------------------------------------------------
# Rodapé
# ---------------------------------------------------------------------------
st.divider()
st.caption(
    "As decisões de classificação são salvas automaticamente no banco local (SQLite) a cada alteração. "
    "Ao atualizar a página ou fechar a aba, os arquivos enviados e a validação precisam ser refeitos; "
    "depois, ao processar o mesmo mês, as decisões já salvas são recuperadas."
)
