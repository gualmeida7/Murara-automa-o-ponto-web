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

import io
import shutil
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from pipeline import (
    carregar_cadastro_colaboradores,
    carregar_consignados,
    executar_checagens_preflight,
    rodada_1_gerar_fila_de_excecoes,
    rodada_2_gerar_relacao_de_valores,
    exportar_por_empresa,
    COLUNAS_OBRIGATORIAS_CADASTRO,
)
from banco_horas import consolidar_banco_horas, ler_banco_horas_secullum, ler_banco_horas_adriano
from leitura_arquivos import ler_arquivo_generico
from validacao import ArquivoInvalidoError

# ---------------------------------------------------------------------------
# Config da página
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Fechamento de Folha — Maçaneiro e Gonzaga / Cianorte Tubos",
    page_icon="✅",
    layout="wide",
)

STATUS_LEGIVEL = {
    "FALTA_OU_AUSENCIA": "Falta ou ausência",
    "BATIDA_INCOMPLETA": "Batida incompleta",
    "INTERVALO_CURTO": "Intervalo curto",
}
OPCOES_CLASSIFICACAO = ["", "Falta injustificada", "Falta justificada", "Erro de relógio — abonar"]
LABEL_PARA_CODIGO = {
    "Falta injustificada": "FALTA_INJUSTIFICADA",
    "Falta justificada": "FALTA_JUSTIFICADA",
    "Erro de relógio — abonar": "ERRO_RELOGIO_ABONAR",
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
def classificar_arquivo(caminho: str) -> str:
    """
    Devolve um dos rótulos: 'cadastro', 'afd', 'consignados',
    'banco_secullum', 'banco_adriano' ou 'desconhecido'. Usa a mesma
    lógica de detecção de coluna já usada dentro dos carregadores
    (validacao.py / banco_horas.py), então a classificação nunca diverge
    do que o carregador real vai aceitar.
    """
    sufixo = Path(caminho).suffix.lower()
    if sufixo in (".txt", ".afd"):
        return "afd"
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
    "cadastro": "Cadastro de Colaboradores",
    "afd": "Arquivo AFD (Ponto)",
    "consignados": "Extrato de Consignados",
    "banco_secullum": "Banco de Horas — Secullum",
    "banco_adriano": "Banco de Horas — Planilha do Adriano",
    "desconhecido": "Não identificado (escolha manualmente)",
}


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
    fila_ordenada["classificacao"] = ""
    fila_ordenada["observacao"] = ""
    return fila_ordenada


# ---------------------------------------------------------------------------
# Cabeçalho
# ---------------------------------------------------------------------------
st.title("✅ Fechamento de Folha de Pagamento")
st.caption("Maçaneiro e Gonzaga LTDA · Cianorte Tubos LTDA — tudo numa tela só, sem arquivo intermediário para gerenciar.")

# ===========================================================================
# SEÇÃO 1 — Upload centralizado (arrastar e soltar)
# ===========================================================================
st.header("1. Enviar os arquivos")
st.write(
    "Arraste todos os arquivos do fechamento para a área abaixo de uma vez só — "
    "Cadastro de Colaboradores, Arquivo AFD (ponto), Extrato de Consignados e, se houver, "
    "o Banco de Horas do Secullum e/ou a planilha semanal do Adriano. "
    "O sistema identifica sozinho o que é cada arquivo."
)

arquivos_subidos = st.file_uploader(
    "Arraste os arquivos aqui ou clique para selecionar",
    type=["txt", "xlsx", "xls", "csv"],
    accept_multiple_files=True,
    key="uploader_geral",
)

if arquivos_subidos:
    linhas_classificacao = []
    for arq in arquivos_subidos:
        caminho = salvar_upload_em_disco(arq)
        tipo_detectado = classificar_arquivo(caminho)
        linhas_classificacao.append({"arquivo": arq.name, "caminho": caminho, "tipo": tipo_detectado})

    st.write("**Confira o que o sistema identificou** (ajuste no dropdown se algo saiu errado):")
    opcoes_tipo = list(RÓTULO_AMIGÁVEL.keys())
    tipos_confirmados = {}
    for i, item in enumerate(linhas_classificacao):
        col1, col2 = st.columns([2, 3])
        with col1:
            st.write(f"📄 {item['arquivo']}")
        with col2:
            escolha = st.selectbox(
                f"Tipo — {item['arquivo']}",
                options=opcoes_tipo,
                index=opcoes_tipo.index(item["tipo"]),
                format_func=lambda t: RÓTULO_AMIGÁVEL[t],
                key=f"tipo_{item['arquivo']}_{i}",
                label_visibility="collapsed",
            )
        tipos_confirmados[escolha] = item["caminho"]

    pis_adriano = ""
    if "banco_adriano" in tipos_confirmados:
        pis_adriano = st.text_input(
            "PIS/matrícula do Adriano (a planilha dele não traz essa informação)",
            key="pis_adriano_input",
        )

    if st.button("Confirmar arquivos e continuar", type="primary"):
        faltando = [t for t in ("cadastro", "afd", "consignados") if t not in tipos_confirmados]
        if faltando:
            st.error(
                "Ainda faltam arquivos obrigatórios: "
                + ", ".join(RÓTULO_AMIGÁVEL[t] for t in faltando)
                + ". Envie-os e classifique corretamente antes de continuar."
            )
        elif "banco_adriano" in tipos_confirmados and not pis_adriano.strip():
            st.error("Você enviou a planilha do Adriano, mas não informou o PIS/matrícula dele.")
        else:
            st.session_state.arquivos_confirmados = tipos_confirmados
            st.session_state.pis_adriano = pis_adriano.strip()
            # limpa qualquer processamento anterior, já que os arquivos mudaram
            for chave in ("cadastro_df", "preflight_ok", "apuracao_diaria", "fila_tabela", "arquivos_finais"):
                st.session_state.pop(chave, None)
            st.success("Arquivos confirmados. Continue na seção 2 abaixo.")

# ===========================================================================
# SEÇÃO 2 — Período + checagem pré-voo instantânea
# ===========================================================================
if st.session_state.get("arquivos_confirmados"):
    st.header("2. Período do fechamento e checagem pré-voo")

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        data_inicio = st.date_input("Início do período", key="data_inicio_input")
    with col_b:
        data_fim = st.date_input("Fim do período", key="data_fim_input")
    with col_c:
        competencia = st.text_input("Competência (mm/aaaa)", value=data_fim.strftime("%m/%Y"), key="competencia_input")

    if st.button("🔎 Validar dados (checagem pré-voo)", type="primary"):
        arquivos = st.session_state.arquivos_confirmados
        try:
            cadastro = carregar_cadastro_colaboradores(arquivos["cadastro"])
            consignados = carregar_consignados(arquivos["consignados"])
            banco_secullum = (
                ler_banco_horas_secullum(arquivos["banco_secullum"]) if "banco_secullum" in arquivos else None
            )
            banco_adriano = (
                ler_banco_horas_adriano(arquivos["banco_adriano"], st.session_state.pis_adriano)
                if "banco_adriano" in arquivos
                else None
            )
            executar_checagens_preflight(
                cadastro, consignados, banco_secullum, banco_adriano,
                caminho_consignados=arquivos["consignados"],
                caminho_banco_secullum=arquivos.get("banco_secullum", ""),
                caminho_banco_adriano=arquivos.get("banco_adriano", ""),
            )
        except ArquivoInvalidoError as e:
            st.session_state.preflight_ok = False
            st.error(f"**Não foi possível continuar — encontramos um problema nos dados:**\n\n{e}")
        else:
            st.session_state.preflight_ok = True
            st.session_state.cadastro_df = cadastro
            st.session_state.consignados_df = consignados
            st.session_state.banco_horas_df = consolidar_banco_horas(
                cadastro["pis"].tolist(), banco_secullum, banco_adriano
            )
            st.success("✅ Checagem pré-voo concluída — todos os dados batem. Pode continuar.")

# ===========================================================================
# SEÇÃO 3 — Processar o ponto (etapa pesada, só roda sob demanda)
# ===========================================================================
if st.session_state.get("preflight_ok"):
    st.header("3. Processar o ponto e calcular o fechamento")
    if st.button("⚙️ Processar Ponto e Calcular Fechamento", type="primary"):
        with st.spinner("Lendo o AFD e apurando ponto a ponto — pode levar alguns segundos..."):
            arquivos = st.session_state.arquivos_confirmados
            apuracao, fila = rodada_1_gerar_fila_de_excecoes(
                arquivos["afd"], st.session_state.cadastro_df, data_inicio, data_fim
            )
            st.session_state.apuracao_diaria = apuracao
            st.session_state.fila_tabela = preparar_fila_para_tabela(fila, st.session_state.cadastro_df)
            st.session_state.competencia_final = competencia
        st.success(f"Processamento concluído. {len(st.session_state.fila_tabela)} exceção(ões) encontrada(s) para revisão.")

# ===========================================================================
# SEÇÃO 4 — Tabela de aprovação embutida (sem CSV intermediário)
# ===========================================================================
if "fila_tabela" in st.session_state:
    st.header("4. Revisar exceções")

    if st.session_state.fila_tabela.empty:
        st.success("Nenhuma exceção encontrada neste período — pode ir direto para a geração do fechamento.")
    else:
        st.write(
            "Só os dias que o sistema não conseguiu decidir sozinho aparecem aqui. "
            "Classifique cada linha no dropdown — o que ficar em branco entra como "
            "**falta injustificada** por padrão (o mesmo critério conservador de sempre)."
        )
        colunas_exibidas = ["nome_colaborador", "razao_social", "data_exibicao", "situacao", "motivo_validacao", "classificacao", "observacao"]
        tabela_editada = st.data_editor(
            st.session_state.fila_tabela[colunas_exibidas],
            key="editor_excecoes",
            hide_index=True,
            use_container_width=True,
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
        st.session_state.fila_tabela.loc[tabela_editada.index, colunas_exibidas] = tabela_editada[colunas_exibidas]

        pendentes = (st.session_state.fila_tabela["classificacao"] == "").sum()
        if pendentes:
            st.info(f"{pendentes} linha(s) ainda sem decisão — serão tratadas como falta injustificada por padrão.")
        else:
            st.success("Todas as exceções já têm uma decisão.")

# ===========================================================================
# SEÇÃO 5 — Gerar e baixar o Excel final
# ===========================================================================
if "apuracao_diaria" in st.session_state:
    st.header("5. Gerar o fechamento final")

    if st.button("📊 Gerar Fechamento Final", type="primary"):
        with st.spinner("Consolidando tudo e montando as planilhas..."):
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

            relacao = rodada_2_gerar_relacao_de_valores(
                st.session_state.apuracao_diaria,
                st.session_state.cadastro_df,
                st.session_state.consignados_df,
                st.session_state.banco_horas_df,
                decisoes,
                st.session_state.competencia_final,
            )
            pasta_saida = PASTA_TEMP / "saida"
            if pasta_saida.exists():
                shutil.rmtree(pasta_saida)
            arquivos_gerados = exportar_por_empresa(relacao, str(pasta_saida))

            arquivos_finais = {}
            for caminho in arquivos_gerados:
                arquivos_finais[caminho.name] = caminho.read_bytes()
            st.session_state.arquivos_finais = arquivos_finais
        st.success(f"{len(st.session_state.arquivos_finais)} arquivo(s) gerado(s) — prontos para baixar abaixo.")

if st.session_state.get("arquivos_finais"):
    st.subheader("Baixar")
    arquivos_finais = st.session_state.arquivos_finais

    if len(arquivos_finais) > 1:
        buffer_zip = io.BytesIO()
        with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            for nome, conteudo in arquivos_finais.items():
                zf.writestr(nome, conteudo)
        st.download_button(
            "⬇️ Baixar todos os arquivos (.zip)",
            data=buffer_zip.getvalue(),
            file_name="relacao_de_valores_todas_empresas.zip",
            mime="application/zip",
            type="primary",
        )

    for nome, conteudo in arquivos_finais.items():
        st.download_button(
            f"⬇️ {nome}",
            data=conteudo,
            file_name=nome,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ---------------------------------------------------------------------------
# Rodapé
# ---------------------------------------------------------------------------
st.divider()
st.caption(
    "O progresso desta página fica salvo enquanto a aba do navegador permanecer aberta. "
    "Atualizar a página (F5) ou fechar a aba reinicia o fechamento do zero."
)
