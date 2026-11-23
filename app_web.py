"""
app_web.py
==========
Versão 100% web (Streamlit) do Sistema de Fechamento de Folha, numa
única tela interativa, sem arquivos intermediários soltos
(fila_validacao_rh.csv / decisoes_rh.csv), com um fluxo único, em memória, com estado guardado em st.session_state.

Este arquivo é só a CASCA de interface: toda a lógica de cálculo continua
em afd_parser.py, business_rules.py, banco_horas.py e pipeline.py, sem
nenhuma duplicação.

Como rodar (veja também o guia completo na conversa):
    pip install -r requirements.txt
    streamlit run app_web.py
"""

from __future__ import annotations

import hashlib
import io
import shutil
import tempfile
import zipfile
from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from pipeline import (
    carregar_cadastro_colaboradores,
    carregar_consignados,
    PREENCHER_CONSIGNADO,
    aviso_consignados_ignorados,
    avisos_banco_horas,
    executar_checagens_preflight,
    rodada_1_gerar_fila_de_excecoes,
    ler_dias_abonados,
    rodada_2_gerar_relacao_de_valores,
    exportar_por_empresa,
    exportar_fila_excecoes_xlsx,
    COLUNAS_OBRIGATORIAS_CADASTRO,
    UNIDADE_DSR_PADRAO,
    UNIDADES_DSR,
)
import db_sessao
from cadastro_editor import (
    BENEFICIOS_PADRAO, COLUNAS_CANONICAS, SECOES, aplicar_aceites, cadastro_de_json, cadastro_para_json,
    cadastro_para_pipeline, cadastro_para_xlsx_bytes, hash_do_cadastro, ler_cadastro_para_edicao,
    normalizar_cabecalhos, resumo_conferencia, normalizar_para_edicao, reconciliar_com_salvo, situacao_por_linha, tem_erros, validar_cadastro,
)
from datas_br import formatar_data_br, parse_data_br
from planilha_thales import exportar_planilha_thales
from feriados import avisos_de_cobertura, carregar_feriados, descrever_feriados
from banco_horas import consolidar_banco_horas, ler_banco_horas_secullum, ler_banco_horas_adriano
from cartao_ponto_pdf import _sem_acento
from leitura_arquivos import ler_arquivo_generico
from importacao_excecoes import importar_decisoes_excecoes_xlsx
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
    if "RESUMO BANCO" in cabecalho:
        return "banco_secullum"

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

    colunas = set(df.columns) | set(normalizar_cabecalhos(df.columns))   # P/Q/R e variantes de cabeçalho
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
    inicio = st.session_state.get("periodo_inicio")
    fim = st.session_state.get("periodo_fim")
    competencia = st.session_state.get("competencia_input", "")
    if inicio and fim:
        return f"{inicio:%d/%m/%Y} a {fim:%d/%m/%Y} · competência {competencia}"
    return f"competência {competencia}"


def contar_decisoes(fila: pd.DataFrame) -> tuple[int, int]:
    """(classificadas, total) da fila de exceções."""
    if fila.empty or "classificacao" not in fila.columns:   # sem exceções (ex.: todos os dias abonados)
        return 0, len(fila)
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


def aplicar_planilha_importada() -> None:
    """
    on_click do botão 'Aplicar decisões da planilha': lê a planilha preenchida, atualiza a fila por ÍNDICE,
    grava no SQLite, atualiza o status da sessão e descarta o estado do editor (para a tabela mostrar o importado).
    O resumo fica em st.session_state.resultado_importacao e é exibido logo abaixo do botão.
    """
    arquivo = st.session_state.get("upload_planilha_decisoes")
    if arquivo is None:
        return
    st.session_state.pop("resultado_importacao", None)
    st.session_state.pop("erro_importacao", None)
    try:
        resultado = importar_decisoes_excecoes_xlsx(
            salvar_upload_em_disco(arquivo), st.session_state.fila_tabela, OPCOES_CLASSIFICACAO,
        )
    except ArquivoInvalidoError as e:
        st.session_state.erro_importacao = str(e)
        return

    fila = st.session_state.fila_tabela
    if not resultado.decisoes.empty:
        fila.loc[resultado.decisoes.index, ["classificacao", "observacao"]] = resultado.decisoes[["classificacao", "observacao"]]
        competencia = st.session_state.competencia_final
        db_sessao.salvar_decisoes_em_lote(
            competencia,
            fila.loc[resultado.decisoes.index].rename(columns={"data_exibicao": "data"})[
                ["pis", "data", "classificacao", "observacao"]
            ],
        )
        classificadas = int((fila["classificacao"] != "").sum())
        db_sessao.criar_ou_atualizar_sessao(
            competencia, len(fila), "concluido" if classificadas == len(fila) else "em_andamento",
        )
    limpar_estado_editor()
    st.session_state.resultado_importacao = resultado


def estilo_situacao(valor: str) -> str:
    """Cor de fundo translúcida (legível nos modos claro e escuro) para as situações do cruzamento com o Espelho."""
    if valor.startswith("Provável erro de relógio"):
        return "background-color: rgba(59, 130, 246, 0.22)"
    if valor.startswith("Provável falta"):
        return "background-color: rgba(239, 68, 68, 0.22)"
    return ""


def empresa_do_arquivo(nome_arquivo: str, cadastro: pd.DataFrame) -> str:
    """Razão social da empresa a que o arquivo gerado pertence (o nome do arquivo traz a razão com '_')."""
    if nome_arquivo.startswith("planilha_thales"):
        return "Resumo do fechamento (Planilha do Thales)"
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

slot_avisos = st.container()   # logo abaixo das etapas, para o aviso ser visto antes de tudo
slot_mensagens = st.container()
slot_metricas = st.container()

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
    st.markdown(
        "**Como enviar os arquivos**\n\n"
        "1. **Cadastro de Colaboradores** (planilha Excel) — obrigatório.\n"
        "2. **Ponto** — escolha uma das opções abaixo:\n"
        "    - **Espelho de Ponto** (PDF do Secullum) sozinho; ou\n"
        "    - **Cartão Ponto** (PDF do Secullum) ou **Arquivo AFD** (.txt); ou\n"
        "    - **os dois juntos** (Cartão Ponto ou AFD, mais o Espelho de Ponto), o que permite "
        "uma conferência mais completa entre o que o relógio registrou e o que foi importado.\n"
        "3. **Opcionais** — Extrato de Consignados (um por empresa), Banco de Horas do Secullum "
        "e planilha semanal do Adriano, se houver.\n\n"
        "Arraste todos os arquivos de uma só vez. O sistema identifica o tipo de cada um e o "
        "mostra na lista abaixo; se algum for reconhecido de forma errada, corrija no campo ao lado."
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
    tem_ponto_ck = any(t in tipos_confirmados for t in ("afd", "cartao_ponto", "espelho_ponto"))
    requisitos = [
        ("Cadastro de Colaboradores", "cadastro" in tipos_confirmados),
        ("Ponto: Espelho de Ponto (PDF), Cartão Ponto (PDF) ou Arquivo AFD (.txt)", tem_ponto_ck),
    ]
    if PREENCHER_CONSIGNADO:
        requisitos.append(("Extrato de Consignados", "consignados" in tipos_confirmados))
    st.write("**Arquivos obrigatórios**")
    for descricao, presente in requisitos:
        if presente:
            st.markdown(f":green[:material/check_circle:] {descricao}")
        else:
            st.markdown(f":orange[:material/pending:] {descricao} — pendente")

    if "espelho_ponto" in tipos_confirmados and not ("cartao_ponto" in tipos_confirmados or "afd" in tipos_confirmados):
        avisos.append((
            "info",
            "Apenas o Espelho de Ponto foi enviado: ele será usado como fonte das batidas do relógio. "
            "Para a conferência cruzada, envie também o Cartão Ponto ou o Arquivo AFD.",
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
            # Slot "ponto": Espelho sozinho, AFD, Cartão Ponto, ou Cartão/AFD + Espelho (conferência cruzada).
            tem_ponto = any(t in tipos_confirmados for t in ("afd", "cartao_ponto", "espelho_ponto"))
            obrigatorios = ("cadastro", "consignados") if PREENCHER_CONSIGNADO else ("cadastro",)
            faltando = [t for t in obrigatorios if t not in tipos_confirmados]
            if not tem_ponto:
                faltando.insert(1, "espelho_ponto")
            duplicados = [t for t in TIPOS_ARQUIVO_UNICO if len(tipos_confirmados.get(t, [])) > 1]
            if faltando:
                st.error(
                    "Ainda faltam arquivos obrigatórios: "
                    + ", ".join(RÓTULO_AMIGÁVEL[t] for t in faltando)
                    + (" (ou o Cartão Ponto, ou o Arquivo AFD)" if "espelho_ponto" in faltando else "")
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
                st.session_state.pis_adriano = pis_adriano.strip()
                # limpa qualquer processamento anterior, já que os arquivos mudaram
                for chave in ("cadastro_df", "preflight_ok", "apuracao_diaria", "fila_tabela", "arquivos_finais",
                              "feriados_periodo", "avisos_feriados", "avisos_banco_horas", "aviso_consignados",
                              "cadastro_hash_validado"):
                    st.session_state.pop(chave, None)
                try:
                    carga = ler_cadastro_para_edicao(tipos_confirmados["cadastro"][0])
                except ArquivoInvalidoError as e:
                    st.error(f"**Não foi possível ler o Cadastro de Colaboradores:**\n\n{e}")
                    st.stop()
                st.session_state.arquivos_confirmados = tipos_confirmados
                salvo = db_sessao.carregar_cadastro_atual()
                trabalho, diferencas = reconciliar_com_salvo(
                    carga.df, cadastro_de_json(salvo["linhas_json"]) if salvo else None,
                )
                st.session_state.update(
                    cadastro_trabalho=trabalho, cadastro_editado=trabalho, cadastro_enviado=carga.df,
                    cadastro_diferencas=diferencas, cadastro_avisos_carga=carga.avisos,
                    cadastro_linhas_planilha=carga.linhas_planilha,
                    cadastro_formato_admissao=carga.formato_admissao, cadastro_sem_cores=carga.sem_cores,
                    cadastro_nome_arquivo=Path(tipos_confirmados["cadastro"][0]).name,
                    cadastro_versao_editor=st.session_state.get("cadastro_versao_editor", 0) + 1,
                )
                st.session_state.pop("cadastro_situacao_exibida", None)
                agendar_mensagem("success", "Arquivos confirmados. Revise o cadastro e defina o período na etapa 2.")
                st.rerun()

# ===========================================================================
# Cadastro de colaboradores - tela editavel (a planilha enviada e' so' a proposta; o cadastro do sistema vale)
# ===========================================================================
st.session_state.cadastro_bloqueado = False
if st.session_state.get("arquivos_confirmados") and "cadastro_trabalho" in st.session_state:
    base = st.session_state.cadastro_trabalho
    editado_ant = st.session_state.cadastro_editado
    versao_editor = st.session_state.cadastro_versao_editor
    with st.expander("Cadastro de colaboradores — revisar, corrigir e salvar", expanded=passo <= 2):
        st.caption(
            "A planilha enviada não é alterada. O que for corrigido aqui passa a valer para o fechamento "
            "e pode ser salvo no sistema ou baixado em planilha limpa."
        )
        sem_cores = st.session_state.get("cadastro_sem_cores", False)
        for aviso_carga in st.session_state.cadastro_avisos_carga:
            if sem_cores and "nenhuma cor de preenchimento" in aviso_carga:
                continue    # substituido pelo aviso curto do quadro de conferencia, logo abaixo
            st.warning(aviso_carga, icon=":material/warning:")
        quadro_conferencia = st.container()   # preenchido depois do editor, com o cadastro ja' editado

        # Diferencas entre a planilha reenviada e o cadastro ja' salvo: nada e' sobrescrito sem aceite.
        diferencas = st.session_state.cadastro_diferencas
        if diferencas:
            with st.container(border=True):
                st.markdown(
                    f"**A planilha enviada difere do cadastro salvo no sistema ({len(diferencas)} diferença(s)).** "
                    "Por padrão vale o cadastro do sistema; marque o que deseja aceitar da planilha."
                )
                rotulo_tipo = {"novo": "Colaborador novo na planilha", "removido": "Ausente na planilha", "alterado": "Campo alterado"}
                quadro = pd.DataFrame([
                    {"aceitar": False, "matricula": d.matricula, "diferença": rotulo_tipo[d.tipo], "campo": d.campo,
                     "no sistema": d.valor_salvo, "na planilha": d.valor_novo}
                    for d in diferencas
                ])
                escolha = st.data_editor(
                    quadro, hide_index=True, key=f"editor_diferencas_{versao_editor}",
                    disabled=["matricula", "diferença", "campo", "no sistema", "na planilha"],
                    column_config={"aceitar": st.column_config.CheckboxColumn("Aceitar")},
                )
                col_a, col_b = st.columns(2)
                if col_a.button("Aplicar os itens aceitos", disabled=not escolha["aceitar"].any()):
                    aceitas = {d.chave for d, ok in zip(diferencas, escolha["aceitar"]) if ok}
                    novo = aplicar_aceites(editado_ant, st.session_state.cadastro_enviado, aceitas)
                    st.session_state.update(
                        cadastro_trabalho=novo, cadastro_editado=novo, cadastro_versao_editor=versao_editor + 1,
                        cadastro_diferencas=[d for d in diferencas if d.chave not in aceitas],
                    )
                    st.session_state.pop("cadastro_situacao_exibida", None)
                    st.rerun()
                if col_b.button("Manter o cadastro do sistema (dispensar diferenças)"):
                    st.session_state.cadastro_diferencas = []
                    st.rerun()

        secao_escolhida = st.segmented_control(
            "Seção", ["Todas", *SECOES], default="Todas", key="cadastro_secao", label_visibility="collapsed",
        ) or "Todas"
        colunas_visiveis = (
            COLUNAS_CANONICAS if secao_escolhida == "Todas"
            else ["matricula", "nome", *[c for c in SECOES[secao_escolhida] if c not in ("matricula", "nome")]]
        )
        exibida = st.session_state.get("cadastro_situacao_exibida")
        tabela = base.copy()
        tabela.insert(0, "situacao", exibida.reindex(tabela.index).fillna("") if exibida is not None else "")
        opcoes_beneficio = sorted(set(BENEFICIOS_PADRAO) | {b for b in base["beneficio"] if b})
        horas = r"^([01]?\d|2[0-3]):[0-5]\d$"
        config = {
            "situacao": st.column_config.TextColumn("Situação", disabled=True, width="small", pinned=True),
            "matricula": st.column_config.TextColumn("Matrícula", pinned=True),
            "nome": st.column_config.TextColumn("Nome"),
            "pis": st.column_config.TextColumn("PIS"),
            "codigo_empresa": st.column_config.TextColumn("Cód. empresa"),
            "razao_social": st.column_config.TextColumn("Razão social"),
            "cnpj": st.column_config.TextColumn("CNPJ"),
            "centro_custo": st.column_config.TextColumn("Centro de custo"),
            **{c: st.column_config.TextColumn(r, help="hh:mm, ex.: 08:00", validate=horas, width="small")
               for c, r in (("entrada", "Entrada"), ("saida_almoco", "Saída almoço"), ("retorno_almoco", "Retorno almoço"),
                            ("saida", "Saída"), ("saida_sexta", "Saída sexta"))},
            "dias_trabalho": st.column_config.TextColumn("Dias de trabalho", help="0=segunda ... 6=domingo, ex.: 0,1,2,3,4"),
            "admissao": st.column_config.DateColumn("Admissão", format="DD/MM/YYYY"),
            "beneficio": st.column_config.SelectboxColumn("Vale", options=opcoes_beneficio),
            "valor_va": st.column_config.NumberColumn("Valor do vale", format="R$ %.2f", min_value=0.0),
            "va_fixo": st.column_config.CheckboxColumn("Vale fixo", help="Marque quem tem vale fixo (verde escuro na planilha): não sofre o desconto de 1/30 por falta."),
            "va_desconto_manual": st.column_config.CheckboxColumn(
                "Vale c/ desconto manual", help="Marque quem tem desconto manual no vale (verde claro na planilha)."),
            "cooper": st.column_config.NumberColumn("Cooper", format="R$ %.2f", min_value=0.0),
            "cooper_fixo": st.column_config.CheckboxColumn("Cooper fixo", help="Marque quem tem Cooper fixo (verde escuro na planilha)."),
            "desconto_folha": st.column_config.NumberColumn("Desconto em folha", format="R$ %.2f", min_value=0.0),
            "banco_horas": st.column_config.CheckboxColumn("Banco de horas"),
            "registra_ponto": st.column_config.CheckboxColumn("Registra ponto"),
            "valor_dia_dsr": st.column_config.NumberColumn("Valor do dia de DSR", format="R$ %.2f", min_value=0.0),
            "observacao": st.column_config.TextColumn("Observação"),
            "nota_rodape": st.column_config.TextColumn("Nota de rodapé"),
        }
        retorno = st.data_editor(
            tabela, num_rows="dynamic", hide_index=True, key=f"editor_cadastro_{versao_editor}",
            column_order=["situacao", *colunas_visiveis], column_config=config, disabled=["situacao"],
        )
        editado = normalizar_para_edicao(retorno.drop(columns=["situacao"]))
        st.session_state.cadastro_editado = editado
        with quadro_conferencia:
            resumo = resumo_conferencia(editado)
            st.markdown("**Conferência da leitura do cadastro**")
            linhas_resumo = [f"{resumo['total']} colaborador(es)."]
            for emp, tipos in resumo["vale_por_empresa"].items():
                linhas_resumo.append(f"Empresa {emp}: " + ", ".join(f"{q} com {t}" for t, q in tipos.items()) + ".")
            linhas_resumo.append(f"Sem vale: {resumo['sem_vale']}. Com Cooper: {resumo['cooper']}.")
            linhas_resumo.append(
                f"Marcados como vale fixo: {resumo['va_fixo']}; Cooper fixo: {resumo['cooper_fixo']}; "
                f"vale com desconto manual: {resumo['va_desconto_manual']}."
            )
            formato = st.session_state.get("cadastro_formato_admissao") or "dia/mês/ano"
            amostra = "; ".join(f"matrícula {m}: {d}" for m, d in resumo["amostra_admissoes"]) or "nenhuma admissão preenchida"
            linhas_resumo.append(f"Datas de admissão lidas como {formato} ({resumo['admissoes_lidas']} preenchidas). Exemplos: {amostra}.")
            st.markdown("\n".join(f"- {t}" for t in linhas_resumo))
            if sem_cores and not (resumo["va_fixo"] or resumo["cooper_fixo"] or resumo["va_desconto_manual"]):
                st.info("A planilha chegou sem cores; marque abaixo quem tem valor fixo ou desconto manual.",
                        icon=":material/info:")
        problemas = validar_cadastro(editado, st.session_state.get("periodo_fim"),
                                     st.session_state.get("cadastro_linhas_planilha"))
        st.session_state.cadastro_bloqueado = tem_erros(problemas)

        erros = [p for p in problemas if p.severidade == "erro"]
        avisos_cad = [p for p in problemas if p.severidade == "aviso"]
        if erros:
            st.error(f"**{len(erros)} erro(s) impedem o fechamento.** Corrija-os na tabela acima.", icon=":material/error:")
            st.dataframe(pd.DataFrame([{"Matrícula": p.matricula, "Campo": p.campo, "Problema": p.mensagem} for p in erros]),
                         hide_index=True)
        else:
            st.success("Nenhum erro bloqueante no cadastro.", icon=":material/check_circle:")
        if avisos_cad:
            with st.expander(f"{len(avisos_cad)} aviso(s) — não bloqueiam"):
                st.dataframe(pd.DataFrame([{"Matrícula": p.matricula, "Campo": p.campo, "Aviso": p.mensagem}
                                           for p in avisos_cad]), hide_index=True)

        col_salvar, col_baixar = st.columns(2)
        if col_salvar.button("Salvar cadastro", icon=":material/save:"):
            id_versao = db_sessao.salvar_cadastro(
                cadastro_para_json(editado), st.session_state.cadastro_nome_arquivo, hash_do_cadastro(editado),
            )
            if id_versao is None:
                st.error("Não foi possível salvar o cadastro no banco de dados.")
            else:
                st.success(f"Cadastro salvo (versão {id_versao}, {len(editado)} colaborador(es)).")
        col_baixar.download_button(
            "Baixar cadastro corrigido (.xlsx)", data=cadastro_para_xlsx_bytes(editado),
            file_name="cadastro_corrigido.xlsx", icon=":material/download:",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Coluna "Situacao": atualizada (uma vez) quando as edicoes mudam o resultado da validacao.
        nova_situacao = situacao_por_linha(retorno, problemas)
        # Trava: no máximo UM rerun seguido, para nunca entrar em laço se o estado do editor oscilar.
        if exibida is None or not nova_situacao.equals(exibida):
            st.session_state.cadastro_situacao_exibida = nova_situacao
            if not st.session_state.pop("_situacao_rerun", False):
                st.session_state._situacao_rerun = True
                st.rerun()
        else:
            st.session_state.pop("_situacao_rerun", None)

    # O cadastro do sistema e' a verdade: se mudou depois da validacao, a validacao deixa de valer.
    hash_validado = st.session_state.get("cadastro_hash_validado")
    if hash_validado and hash_validado != hash_do_cadastro(st.session_state.cadastro_editado):
        for chave in ("preflight_ok", "cadastro_df", "apuracao_diaria", "fila_tabela", "arquivos_finais", "cadastro_hash_validado"):
            st.session_state.pop(chave, None)
        limpar_estado_editor()
        agendar_mensagem("warning", "O cadastro foi alterado depois da validação. Valide os dados de novo (etapa 2).")
        st.rerun()

# ===========================================================================
# ETAPA 2 — Período + checagem pré-voo instantânea
# ===========================================================================
if st.session_state.get("arquivos_confirmados"):
    resumo_etapa2 = f" — {resumo_periodo()}" if passo > 2 else ""
    with secao_leitura(f"2. Período e validação{resumo_etapa2}", passo == 2):
        # Estado antigo (date ou string ISO) vira dd/mm/aaaa; padrão = hoje, como antes.
        for chave in ("data_inicio_input", "data_fim_input"):
            st.session_state[chave] = formatar_data_br(st.session_state.get(chave, date.today()))
        col_a, col_b = st.columns(2)
        with col_a:
            texto_inicio = st.text_input(
                "Início do período", key="data_inicio_input",
                placeholder="dd/mm/aaaa", max_chars=10,
            )
        with col_b:
            texto_fim = st.text_input(
                "Fim do período", key="data_fim_input",
                placeholder="dd/mm/aaaa", max_chars=10,
            )
        data_inicio = parse_data_br(texto_inicio)
        data_fim = parse_data_br(texto_fim)
        if data_inicio is None:
            st.error("Data de início inválida. Use o formato dd/mm/aaaa.")
        if data_fim is None:
            st.error("Data de fim inválida. Use o formato dd/mm/aaaa.")
        if data_inicio and data_fim and data_inicio > data_fim:
            st.error("A data de início não pode ser posterior à data de fim.")
        periodo_valido = bool(data_inicio and data_fim and data_inicio <= data_fim)
        st.session_state.periodo_inicio = data_inicio if periodo_valido else None
        st.session_state.periodo_fim = data_fim if periodo_valido else None
        with sb_competencia:
            competencia = st.text_input(
                "Competência (mm/aaaa)", value=data_fim.strftime("%m/%Y") if data_fim else "", key="competencia_input",
            )
        st.caption("A competência é informada na barra lateral.")

        if st.session_state.get("cadastro_bloqueado"):
            st.error("Há erros no cadastro de colaboradores. Corrija-os na tabela de cadastro antes de validar os dados.")
        if st.button("Validar dados", type="primary",
                     disabled=not periodo_valido or bool(st.session_state.get("cadastro_bloqueado"))):
            arquivos = st.session_state.arquivos_confirmados
            try:
                cadastro = cadastro_para_pipeline(st.session_state.cadastro_editado)
                consignados = (
                    carregar_consignados(arquivos["consignados"], cadastro=cadastro)
                    if PREENCHER_CONSIGNADO and "consignados" in arquivos else None
                )
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
                    caminho_consignados=", ".join(arquivos.get("consignados", [])),
                    caminho_banco_secullum=arquivos.get("banco_secullum", [""])[0],
                    caminho_banco_adriano=arquivos.get("banco_adriano", [""])[0],
                )
            except ArquivoInvalidoError as e:
                st.session_state.preflight_ok = False
                st.error(f"**Não foi possível continuar — há um problema nos dados:**\n\n{e}")
            else:
                st.session_state.preflight_ok = True
                st.session_state.cadastro_hash_validado = hash_do_cadastro(st.session_state.cadastro_editado)
                st.session_state.cadastro_df = cadastro
                st.session_state.consignados_df = consignados
                st.session_state.aviso_consignados = aviso_consignados_ignorados(consignados)
                st.session_state.banco_horas_df = consolidar_banco_horas(
                    cadastro["pis"].tolist(), banco_secullum, banco_adriano
                )
                st.session_state.avisos_banco_horas = avisos_banco_horas(cadastro, banco_secullum, banco_adriano)
                agendar_mensagem("success", "Validação concluída: os dados são consistentes. Prossiga para o processamento.")
                st.rerun()

    # Avisos (nao impedem de continuar): ficam na tela depois do rerun, por isso saem do session_state.
    if st.session_state.get("preflight_ok"):
        if st.session_state.get("aviso_consignados"):
            avisos.append(("warning", st.session_state.aviso_consignados))
        for aviso in st.session_state.get("avisos_banco_horas", []):
            avisos.append(("warning", aviso))

# ===========================================================================
# ETAPA 3 — Processar o ponto (etapa pesada, só roda sob demanda)
# ===========================================================================
if st.session_state.get("preflight_ok") and st.session_state.get("periodo_inicio"):
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

        texto_abonos = st.text_input(
            "Dias abonados para todos (opcional)",
            key="dias_abonados_input",
            placeholder="14/10/2026, 15/10/2026",
            help=(
                "Use quando o relógio falhou para todo mundo no dia: quem não fechou as 4 batidas não leva falta, "
                "atraso nem perde DSR/VA/COPR, e o dia não vai para a fila. Dia com as 4 batidas é apurado normal. "
                "Formato dd/mm/aaaa, separados por vírgula. Clique em Processar de novo depois de mudar."
            ),
        )

        if st.button("Processar ponto e calcular fechamento", type="primary"):
            try:
                dias_abonados = ler_dias_abonados(texto_abonos, data_inicio, data_fim)
            except ArquivoInvalidoError as e:
                st.error(f"**Dias abonados:** {e}")
                st.stop()
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
                # AFD (.txt) tem prioridade; depois o Cartão Ponto (PDF); sem nenhum dos dois, o Espelho é a fonte.
                lista_ponto = arquivos.get("afd") or arquivos.get("cartao_ponto")
                caminho_ponto = lista_ponto[0] if lista_ponto else None
                apuracao, fila = rodada_1_gerar_fila_de_excecoes(
                    caminho_ponto, st.session_state.cadastro_df, data_inicio, data_fim,
                    caminho_espelho=arquivos.get("espelho_ponto", [None])[0],
                    feriados=feriados,
                    dias_abonados=dias_abonados,
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

    apuracao_ref = st.session_state.get("apuracao_diaria")
    if apuracao_ref is not None and "cadastro_df" in st.session_state:
        cad = st.session_state.cadastro_df
        registra = cad["registra_ponto"].fillna(True).astype(bool) if "registra_ponto" in cad.columns else True
        sem_ponto = cad[cad["pis"].isin(apuracao_ref.attrs.get("sem_ponto", [])) & registra]
        if not sem_ponto.empty:
            nomes = {f"{r.nome} (PIS {r.pis})": r.pis for r in sem_ponto.itertuples()}
            with st.container(width=LARGURA_LEITURA, border=True):
                st.warning(
                    f"**{len(sem_ponto)} colaborador(es) estão no cadastro mas não têm nenhuma batida no ponto.** "
                    "Se nada for feito, o período inteiro deles entra como falta. Isso costuma ser desligado "
                    "que ficou no cadastro, ou quem não usa o relógio.",
                    icon=":material/warning:",
                )
                escolhidos = st.multiselect("Colaboradores sem ponto", list(nomes), key="sem_ponto_escolhidos")
                acao = st.radio(
                    "O que fazer com os selecionados?",
                    ["Desligado: tirar do fechamento", "Não registra ponto: manter no relatório com COPR \"Conferir\""],
                    key="sem_ponto_acao",
                )
                if st.button("Aplicar e reprocessar", disabled=not escolhidos):
                    pis_escolhidos = [nomes[n] for n in escolhidos]
                    if acao.startswith("Desligado"):
                        st.session_state.cadastro_df = cad[~cad["pis"].isin(pis_escolhidos)].reset_index(drop=True)
                    else:
                        novo = cad.copy()
                        if "registra_ponto" not in novo.columns:
                            novo["registra_ponto"] = True
                        novo["registra_ponto"] = novo["registra_ponto"].fillna(True).astype(bool)
                        novo.loc[novo["pis"].isin(pis_escolhidos), "registra_ponto"] = False
                        st.session_state.cadastro_df = novo
                    for chave in ("apuracao_diaria", "fila_tabela", "sem_ponto_escolhidos"):
                        st.session_state.pop(chave, None)
                    limpar_estado_editor()
                    agendar_mensagem(
                        "info",
                        f"{len(pis_escolhidos)} colaborador(es) ajustado(s) só nesta sessão (a planilha de cadastro "
                        "não foi alterada). Clique em **Processar** de novo.",
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

            # Esconde o "Download as CSV" da barra da tabela (CSV cru que o Excel abre tudo na coluna A):
            # a exportação suportada é a planilha .xlsx abaixo. Preenchido depois da tabela para já
            # levar a última edição.
            st.html(
                "<style>[data-testid='stElementToolbarButton']:has(button[aria-label='Download as CSV'])"
                "{display:none}</style>"
            )
            slot_exportar = st.container()

            # "data" (datetime) em vez de "data_exibicao" (texto dd/mm/aaaa): o texto ordenava por dia, não por mês
            colunas_exibidas = ["nome_colaborador", "razao_social", "data", "situacao", "motivo_validacao", "classificacao", "observacao"]
            tabela_editada = st.data_editor(
                fila_filtrada[colunas_exibidas].style.map(estilo_situacao, subset=["situacao"]),
                key=chave_editor,
                hide_index=True,
                width="stretch",
                disabled=["nome_colaborador", "razao_social", "data", "situacao", "motivo_validacao"],
                column_config={
                    "nome_colaborador": st.column_config.TextColumn("Colaborador"),
                    "razao_social": st.column_config.TextColumn("Empresa"),
                    "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
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

            with slot_exportar:
                # Recorte recalculado sobre a fila JÁ atualizada com as edições desta execução
                # (fila_filtrada, lá em cima, ainda não tem a última decisão).
                exportar_filtrada = filtrar_fila(fila_atual, filtro_empresa, filtro_colaborador, filtro_situacao)
                nome_base = f"revisao_excecoes_{str(st.session_state.competencia_final).replace('/', '-')}"
                mime_xlsx = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                descricao_filtros = "; ".join(
                    f"{rotulo}: {', '.join(valores)}"
                    for rotulo, valores in (
                        ("Empresa", filtro_empresa), ("Colaborador", filtro_colaborador), ("Situação", filtro_situacao),
                    ) if valores
                )
                with st.container(horizontal=True):
                    st.download_button(
                        f"Baixar planilha filtrada (.xlsx) — {len(exportar_filtrada)} linha(s)"
                        if filtrando else "Baixar planilha da revisão (.xlsx)",
                        data=exportar_fila_excecoes_xlsx(
                            exportar_filtrada, str(st.session_state.competencia_final), OPCOES_CLASSIFICACAO,
                            resumo_periodo(), descricao_filtros, len(fila_atual),
                        ),
                        file_name=f"{nome_base}{'_filtrado' if filtrando else ''}.xlsx",
                        mime=mime_xlsx,
                        icon=":material/download:",
                        disabled=exportar_filtrada.empty,
                        key="download_revisao_filtrada",
                        help="Uma informação por célula, com resumo por decisão e empresa.",
                    )
                    if filtrando:
                        st.download_button(
                            f"Baixar fila completa (.xlsx) — {len(fila_atual)} linha(s)",
                            data=exportar_fila_excecoes_xlsx(
                                fila_atual, str(st.session_state.competencia_final), OPCOES_CLASSIFICACAO,
                                resumo_periodo(),
                            ),
                            file_name=f"{nome_base}.xlsx",
                            mime=mime_xlsx,
                            icon=":material/download:",
                            type="tertiary",
                            key="download_revisao_completa",
                        )

                with st.expander("Importar planilha preenchida", icon=":material/upload_file:"):
                    st.caption(
                        "Fluxo: baixe a planilha acima → quem decide preenche a coluna Decisão (e, se quiser, a "
                        "Observação) → envie o arquivo aqui. Serve a planilha filtrada ou a fila completa. "
                        "Decisão em branco não altera a linha; a planilha prevalece sobre decisões já marcadas."
                    )
                    st.file_uploader(
                        "Importar planilha preenchida", type=["xlsx"], key="upload_planilha_decisoes",
                        label_visibility="collapsed",
                    )
                    st.button(
                        "Aplicar decisões da planilha",
                        on_click=aplicar_planilha_importada,
                        disabled=st.session_state.get("upload_planilha_decisoes") is None,
                        icon=":material/playlist_add_check:",
                    )
                    if st.session_state.get("erro_importacao"):
                        st.error(f"**Não foi possível importar a planilha:**\n\n{st.session_state.erro_importacao}")
                    resultado_imp = st.session_state.get("resultado_importacao")
                    if resultado_imp is not None:
                        st.success(
                            f"{resultado_imp.aplicadas} decisão(ões) aplicada(s) a partir da planilha"
                            + (f" ({resultado_imp.sobrescritas} sobrescreveram decisão já marcada)"
                               if resultado_imp.sobrescritas else "")
                            + f". {resultado_imp.em_branco} linha(s) em branco ficaram como estavam."
                        )
                        for titulo_p, itens in (
                            ("não encontrada(s) na fila", resultado_imp.nao_encontradas),
                            ("com decisão inválida", resultado_imp.invalidas),
                            ("repetida(s) na planilha", resultado_imp.duplicadas),
                            ("ambígua(s) na fila", resultado_imp.ambiguas),
                        ):
                            if itens:
                                detalhes = "\n".join(
                                    f"- {i['colaborador']} · {i['empresa']} · {i['data']}: {i['motivo']}" for i in itens
                                )
                                st.warning(f"**{len(itens)} linha(s) {titulo_p}:**\n\n{detalhes}")

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
                arquivos_gerados.append(exportar_planilha_thales(
                    relacao, st.session_state.apuracao_diaria, decisoes, st.session_state.cadastro_df,
                    st.session_state.competencia_final,
                    st.session_state.periodo_inicio, st.session_state.periodo_fim, str(pasta_saida),
                ))

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

def acao_sugerida(texto: str) -> str | None:
    """Frase de ação para um aviso conhecido (None se não houver orientação específica)."""
    if "banco de horas" in texto and "nenhum arquivo" in texto:
        return (
            "Na etapa 1, envie o export do Secullum ou a planilha do Adriano com o saldo dessa pessoa; "
            "ou, se ela não está no banco, desmarque `banco_horas` no cadastro."
        )
    if "sem o Cartão Ponto" in texto:
        return "Na etapa 1, envie também o Cartão Ponto do mesmo período."
    if "calendário de feriados não tem o ano" in texto:
        return "Rode `python gerar_feriados.py` para incluir o ano e reinicie o app."
    if "por previsão" in texto:
        return "Confira o decreto do ano e, se a data mudou, regere o calendário com `gerar_feriados.py`."
    return None


qtd_atencao_total = sum(1 for nivel, _ in avisos if nivel == "warning")

with slot_avisos:
    if avisos:
        # Atenção (warning) primeiro; sempre aberto e em destaque, não escondido num expander.
        avisos_ordenados = sorted(avisos, key=lambda a: a[0] != "warning")
        qtd_atencao = qtd_atencao_total
        with st.container(border=True):
            st.markdown(
                f"#### :material/warning: Atenção — {len(avisos)} aviso(s)"
                if qtd_atencao
                else f"#### :material/info: {len(avisos)} informação(ões)"
            )
            for nivel, texto in avisos_ordenados:
                icone = ":material/warning:" if nivel == "warning" else ":material/info:"
                getattr(st, nivel)(texto, icon=icone)
                acao = acao_sugerida(texto) if nivel == "warning" else None
                if acao:
                    st.caption(f":material/arrow_forward: **O que fazer:** {acao}")

with sb_rascunho:
    st.divider()
    if qtd_atencao_total:
        st.badge(f"{qtd_atencao_total} aviso(s) de atenção", icon=":material/warning:", color="orange")
        st.caption("Veja o quadro no topo da página.")
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
