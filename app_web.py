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
from datetime import datetime
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
import db_sessao
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
    page_icon="✅",
    layout="wide",
)

db_sessao.inicializar_banco()

STATUS_LEGIVEL = {
    "FALTA_OU_AUSENCIA": "Falta ou ausência",
    "BATIDA_INCOMPLETA": "Batida incompleta",
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
    "Cadastro de Colaboradores, Cartão Ponto e Espelho de Ponto (PDF) — ou o AFD .txt —, Extrato de Consignados e, se houver, "
    "o Banco de Horas do Secullum e/ou a planilha semanal do Adriano. "
    "O sistema identifica sozinho o que é cada arquivo."
)

arquivos_subidos = st.file_uploader(
    "Arraste os arquivos aqui ou clique para selecionar",
    type=[ext.lstrip(".") for ext in EXTENSOES_PONTO] + ["xlsx", "xls", "csv"],
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
    # Cada tipo guarda uma LISTA de caminhos, não um único - várias empresas
    # costumam emitir extratos de consignados separados (um por CNPJ), por
    # exemplo, e nada deve ser descartado silenciosamente se a pessoa soltar
    # dois arquivos classificados com o mesmo tipo.
    tipos_confirmados: dict[str, list[str]] = {}
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
        tipos_confirmados.setdefault(escolha, []).append(item["caminho"])

    # Tipos dos quais só faz sentido existir UM arquivo por fechamento -
    # mais de um classificado assim é quase certamente um erro de
    # classificação (ex.: dois cadastros enviados por engano), e seguir
    # silenciosamente usando só o primeiro escondería o problema.
    TIPOS_ARQUIVO_UNICO = (
        "cadastro", "cartao_ponto", "espelho_ponto", "afd", "banco_secullum", "banco_adriano",
    )

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
        if "espelho_ponto" in tipos_confirmados and "cartao_ponto" not in tipos_confirmados:
            st.warning(
                "O Espelho de Ponto foi enviado sem o Cartão Ponto correspondente. "
                "A conferência cruzada ficará limitada."
            )
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

    competencia_atual = st.session_state.get("competencia_input")
    sessao_salva = db_sessao.obter_sessao(competencia_atual) if competencia_atual else None
    if sessao_salva and sessao_salva["status"] == "em_andamento":
        decisoes_salvas = db_sessao.carregar_decisoes(competencia_atual)
        n = int((decisoes_salvas["classificacao"] != "").sum())
        ultimo_save = datetime.fromisoformat(sessao_salva["atualizado_em"]).strftime("%d/%m/%Y às %H:%M")
        st.info(
            f"📋 **Rascunho encontrado — Fechamento {competencia_atual}**  \n"
            f"{n} de {sessao_salva['total_excecoes']} exceções já classificadas. "
            f"Último save: {ultimo_save}.  \n"
            f"Clique em **Processar** para recarregar e continuar de onde parou, "
            f"ou use *Recomeçar* na seção 4 para descartar."
        )

    if st.button("⚙️ Processar Ponto e Calcular Fechamento", type="primary"):
        with st.spinner("Lendo o ponto e apurando dia a dia — pode levar alguns segundos..."):
            arquivos = st.session_state.arquivos_confirmados
            # AFD (.txt) tem prioridade quando enviado; senão usa o Cartão Ponto (PDF).
            caminho_ponto = (arquivos.get("afd") or arquivos["cartao_ponto"])[0]
            apuracao, fila = rodada_1_gerar_fila_de_excecoes(
                caminho_ponto, st.session_state.cadastro_df, data_inicio, data_fim,
                caminho_espelho=arquivos.get("espelho_ponto", [None])[0],
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
        st.success(f"Processamento concluído. {len(st.session_state.fila_tabela)} exceção(ões) encontrada(s) para revisão.")
        restauradas = st.session_state.pop("rascunho_restaurado", 0)
        if restauradas:
            st.info(f"📋 Rascunho restaurado — {restauradas} decisão(ões) anterior(es) recuperada(s). Pode continuar de onde parou.")

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
            st.info(f"{pendentes} linha(s) ainda sem decisão — serão tratadas como falta injustificada por padrão.")
        if sessao:
            if status == "concluido":
                st.success("✅ Todas as exceções classificadas. Pronto para gerar o fechamento final.")
            else:
                st.caption(
                    f"✓ Rascunho salvo automaticamente — "
                    f"{classificadas} de {total} exceções classificadas "
                    f"(Último save: {sessao['atualizado_em'][-8:-3]})"
                )

    with st.expander("⚠️ Opções"):
        if st.button("🗑️ Recomeçar — descartar todas as decisões deste mês"):
            db_sessao.limpar_sessao(st.session_state.competencia_final)
            for chave in ("fila_tabela", "apuracao_diaria", "arquivos_finais", "editor_excecoes"):
                st.session_state.pop(chave, None)
            st.rerun()

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
