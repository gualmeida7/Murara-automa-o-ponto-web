"""
app_gui.py
==========
Interface grafica (tkinter) para a usuaria do administrativo rodar o
fechamento de folha sem terminal, sem linha de comando e sem editor de
codigo. Este arquivo fica na MESMA pasta que afd_parser.py, business_rules.py,
banco_horas.py e pipeline.py - ele so' importa essas funcoes e desenha os
botoes por cima.

Duas coisas alem do pedido original, com justificativa:

1. Foi adicionado um botao "Selecionar Cadastro de Colaboradores" -
   sem ele o programa nao sabe a jornada nem a empresa de cada
   funcionario, entao nao ha como calcular nada. E' obrigatorio.
2. Banco de horas ficou como DOIS botoes opcionais, um para o export do
   Secullum (cobre o "Sidney") e outro para a planilha semanal do
   "Adriano" (que exige o PIS/matricula dele, pedido num campo de texto
   ao lado do botao). Isso segue exatamente a excecao descrita no
   processo original - os dois sao a mesma "categoria" (banco de horas),
   so' de fontes diferentes.

O botao "Processar Folha de Pagamento" roda o fechamento inteiro num
unico clique: gera a Relacao de Valores por empresa E a lista de
excecoes que ainda precisam de revisao humana (dias sem batida, batida
incompleta, etc.) - essas ficam com uma classificacao conservadora
("falta injustificada") ate' serem revisadas no painel
`dashboard_aprovacao.html`, que continua sendo o lugar certo para essa
revisao antes do envio definitivo a' contabilidade.
"""

import os
import sys
import threading
import traceback
from datetime import date, timedelta
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

from afd_parser import parse_afd
from pipeline import (
    carregar_cadastro_colaboradores,
    carregar_consignados,
    rodada_1_gerar_fila_de_excecoes,
    rodada_2_gerar_relacao_de_valores,
    exportar_por_empresa,
    exportar_fila_validacao_rh,
    executar_checagens_preflight,
)
from banco_horas import consolidar_banco_horas, ler_banco_horas_secullum, ler_banco_horas_adriano
from validacao import ArquivoInvalidoError


def periodo_padrao(hoje: date | None = None) -> tuple[date, date]:
    """Fechamento padrao: dia 26 do mes anterior ate' dia 25 do mes atual."""
    hoje = hoje or date.today()
    fim = hoje.replace(day=25) if hoje.day >= 26 else (hoje.replace(day=1) - timedelta(days=1)).replace(day=25)
    inicio = (fim.replace(day=1) - timedelta(days=1)).replace(day=26)
    return inicio, fim


class AplicativoFolha(tk.Tk):
    _ETAPAS = [
        "Lendo cadastro de colaboradores...",
        "Lendo consignados e banco de horas...",
        "Validando dados (checagem pré-voo)...",
        "Lendo e apurando o ponto (AFD) - pode demorar um pouco...",
        "Aplicando regras e gerando a Relação de Valores...",
        "Salvando arquivos Excel...",
    ]

    def __init__(self):
        super().__init__()
        self.title("Fechamento de Folha - Maçaneiro e Gonzaga / Cianorte Tubos")
        self.geometry("640x560")
        self.resizable(False, False)

        self.caminhos = {
            "cadastro": tk.StringVar(),
            "afd": tk.StringVar(),
            "consignados": tk.StringVar(),
            "banco_secullum": tk.StringVar(),
            "banco_adriano": tk.StringVar(),
        }
        self.pis_adriano = tk.StringVar()

        inicio_padrao, fim_padrao = periodo_padrao()
        self.data_inicio = tk.StringVar(value=inicio_padrao.strftime("%d/%m/%Y"))
        self.data_fim = tk.StringVar(value=fim_padrao.strftime("%d/%m/%Y"))
        self.pasta_saida = tk.StringVar(value=str(Path.home() / "Desktop" / "Folha - Saida"))

        self._montar_layout()

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------
    def _montar_layout(self):
        pad = {"padx": 16, "pady": 6}

        titulo = tk.Label(self, text="Fechamento de folha de pagamento", font=("Segoe UI", 14, "bold"))
        titulo.pack(anchor="w", **pad)

        secao_obrig = tk.LabelFrame(self, text="Arquivos obrigatórios", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        secao_obrig.pack(fill="x", padx=16, pady=(4, 8))

        self._linha_arquivo(secao_obrig, "Cadastro de Colaboradores", "cadastro",
                             "Selecionar Cadastro de Colaboradores")
        self._linha_arquivo(secao_obrig, "Arquivo AFD (Ponto)", "afd",
                             "Selecionar Arquivo AFD (Ponto)")
        self._linha_arquivo(secao_obrig, "Extrato de Consignados", "consignados",
                             "Selecionar Consignados")

        secao_opc = tk.LabelFrame(self, text="Banco de horas (opcional)", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        secao_opc.pack(fill="x", padx=16, pady=(4, 8))

        self._linha_arquivo(secao_opc, "Banco de Horas (Secullum - Sidney)", "banco_secullum",
                             "Selecionar Banco de Horas")

        linha_adriano = tk.Frame(secao_opc)
        linha_adriano.pack(fill="x", pady=4)
        self._linha_arquivo(secao_opc, "Planilha semanal do Adriano", "banco_adriano",
                             "Selecionar Planilha do Adriano")
        linha_pis = tk.Frame(secao_opc)
        linha_pis.pack(fill="x", pady=(0, 4))
        tk.Label(linha_pis, text="PIS/matrícula do Adriano:", width=28, anchor="w").pack(side="left")
        tk.Entry(linha_pis, textvariable=self.pis_adriano, width=20).pack(side="left")

        secao_periodo = tk.LabelFrame(self, text="Período do fechamento", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        secao_periodo.pack(fill="x", padx=16, pady=(4, 8))
        linha_periodo = tk.Frame(secao_periodo)
        linha_periodo.pack(fill="x")
        tk.Label(linha_periodo, text="De (dd/mm/aaaa):").pack(side="left")
        tk.Entry(linha_periodo, textvariable=self.data_inicio, width=12).pack(side="left", padx=(4, 16))
        tk.Label(linha_periodo, text="Até (dd/mm/aaaa):").pack(side="left")
        tk.Entry(linha_periodo, textvariable=self.data_fim, width=12).pack(side="left", padx=4)

        secao_saida = tk.LabelFrame(self, text="Onde salvar o resultado", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        secao_saida.pack(fill="x", padx=16, pady=(4, 8))
        linha_saida = tk.Frame(secao_saida)
        linha_saida.pack(fill="x")
        tk.Entry(linha_saida, textvariable=self.pasta_saida, width=52, state="readonly").pack(side="left", padx=(0, 8))
        tk.Button(linha_saida, text="Alterar pasta...", command=self._escolher_pasta_saida).pack(side="left")

        self.botao_processar = tk.Button(
            self, text="Processar Folha de Pagamento",
            font=("Segoe UI", 12, "bold"), bg="#2F5D4E", fg="white",
            activebackground="#254A3E", activeforeground="white",
            height=2, command=self._processar_com_protecao,
        )
        self.botao_processar.pack(fill="x", padx=16, pady=(12, 4))

        self.barra_progresso = ttk.Progressbar(self, mode="determinate", maximum=len(self._ETAPAS))
        self.barra_progresso.pack(fill="x", padx=16, pady=(0, 4))

        self.barra_status = tk.Label(self, text="Pronto.", anchor="w", fg="#555555")
        self.barra_status.pack(fill="x", padx=16, pady=(0, 10))

    def _linha_arquivo(self, container, rotulo, chave, texto_botao):
        linha = tk.Frame(container)
        linha.pack(fill="x", pady=4)
        tk.Label(linha, text=rotulo + ":", width=28, anchor="w").pack(side="left")
        tk.Entry(linha, textvariable=self.caminhos[chave], width=32, state="readonly").pack(side="left", padx=(0, 8))
        tk.Button(linha, text=texto_botao, command=lambda: self._escolher_arquivo(chave)).pack(side="left")

    # ------------------------------------------------------------------
    # Ações
    # ------------------------------------------------------------------
    def _escolher_arquivo(self, chave):
        caminho = filedialog.askopenfilename(
            title="Selecionar arquivo",
            filetypes=[("Planilhas e texto", "*.xlsx *.xls *.csv *.txt"), ("Todos os arquivos", "*.*")],
        )
        if caminho:
            self.caminhos[chave].set(caminho)

    def _escolher_pasta_saida(self):
        pasta = filedialog.askdirectory(title="Escolher pasta para salvar o resultado")
        if pasta:
            self.pasta_saida.set(pasta)

    def _status(self, indice_etapa: int):
        """
        Atualiza a barra de progresso e o texto de status para a etapa
        `indice_etapa` (0-based, ver `_ETAPAS`). Seguro de chamar de dentro
        da thread de processamento em segundo plano - agenda a atualizacao
        de verdade na thread principal da UI via `self.after`.
        """
        texto = self._ETAPAS[indice_etapa]

        def atualizar():
            self.barra_progresso["value"] = indice_etapa + 1
            self.barra_status.config(text=f"Etapa {indice_etapa + 1} de {len(self._ETAPAS)}: {texto}")

        self.after(0, atualizar)

    def _processar_com_protecao(self):
        # Trava o botao pra usuaria nao clicar 2x enquanto processa
        self.botao_processar.config(state="disabled", text="Processando...")
        self.barra_progresso["value"] = 0
        self.barra_status.config(text="Processando, aguarde...")
        self.update_idletasks()
        thread = threading.Thread(target=self._processar, daemon=True)
        thread.start()

    def _validar_campos(self) -> str | None:
        if not self.caminhos["cadastro"].get():
            return "Selecione o arquivo de Cadastro de Colaboradores antes de continuar."
        if not self.caminhos["afd"].get():
            return "Selecione o Arquivo AFD (Ponto) antes de continuar."
        if not self.caminhos["consignados"].get():
            return "Selecione o Extrato de Consignados antes de continuar."
        if self.caminhos["banco_adriano"].get() and not self.pis_adriano.get().strip():
            return "Você selecionou a planilha do Adriano, mas não informou o PIS/matrícula dele."
        try:
            datetime_inicio = pd.to_datetime(self.data_inicio.get(), dayfirst=True).date()
            datetime_fim = pd.to_datetime(self.data_fim.get(), dayfirst=True).date()
        except Exception:
            return "As datas do período precisam estar no formato dd/mm/aaaa."
        if datetime_inicio >= datetime_fim:
            return "A data 'De' precisa ser anterior à data 'Até'."
        return None

    def _processar(self):
        erro_validacao = self._validar_campos()
        if erro_validacao:
            self._finalizar(erro=erro_validacao)
            return

        try:
            data_inicio = pd.to_datetime(self.data_inicio.get(), dayfirst=True).date()
            data_fim = pd.to_datetime(self.data_fim.get(), dayfirst=True).date()
            pasta_saida = Path(self.pasta_saida.get())
            pasta_saida.mkdir(parents=True, exist_ok=True)

            self._status(0)
            cadastro = carregar_cadastro_colaboradores(self.caminhos["cadastro"].get())

            self._status(1)
            consignados = carregar_consignados(self.caminhos["consignados"].get(), cadastro=cadastro)

            banco_secullum = None
            if self.caminhos["banco_secullum"].get():
                banco_secullum = ler_banco_horas_secullum(self.caminhos["banco_secullum"].get(), cadastro=cadastro)

            banco_adriano = None
            if self.caminhos["banco_adriano"].get():
                banco_adriano = ler_banco_horas_adriano(
                    self.caminhos["banco_adriano"].get(), self.pis_adriano.get().strip(), cadastro=cadastro
                )

            self._status(2)
            # Checagem pre-voo ANTES do AFD (a etapa mais pesada): confere
            # que todo PIS citado nos arquivos opcionais existe no cadastro,
            # pra usuaria descobrir um PIS digitado errado em segundos, sem
            # esperar o AFD inteiro ser processado.
            executar_checagens_preflight(
                cadastro, consignados, banco_secullum, banco_adriano,
                caminho_consignados=self.caminhos["consignados"].get(),
                caminho_banco_secullum=self.caminhos["banco_secullum"].get(),
                caminho_banco_adriano=self.caminhos["banco_adriano"].get(),
            )
            banco_horas = consolidar_banco_horas(cadastro["pis"].tolist(), banco_secullum, banco_adriano)

            self._status(3)
            apuracao, fila = rodada_1_gerar_fila_de_excecoes(
                self.caminhos["afd"].get(), cadastro, data_inicio, data_fim
            )
            exportar_fila_validacao_rh(fila, cadastro, str(pasta_saida / "fila_validacao_rh.csv"))

            self._status(4)
            # Se o RH ja' revisou as excecoes no painel de aprovacao e
            # exportou decisoes_rh.csv NESSA MESMA pasta de saida, usa as
            # decisoes reais dele. Senao, gera a versao preliminar com o
            # padrao conservador (toda excecao ainda sem decisao conta
            # como falta injustificada ate' ser revisada).
            caminho_decisoes = pasta_saida / "decisoes_rh.csv"
            decisoes = pd.read_csv(caminho_decisoes) if caminho_decisoes.exists() else None
            usou_decisoes_rh = decisoes is not None

            competencia = data_fim.strftime("%m/%Y")
            relacao = rodada_2_gerar_relacao_de_valores(
                apuracao, cadastro, consignados, banco_horas, decisoes_rh=decisoes, competencia=competencia
            )

            self._status(5)
            arquivos_gerados = exportar_por_empresa(relacao, str(pasta_saida))

            n_pendencias = len(fila)
            self._finalizar(
                sucesso=True,
                pasta_saida=str(pasta_saida),
                n_arquivos=len(arquivos_gerados),
                n_pendencias=n_pendencias,
                usou_decisoes_rh=usou_decisoes_rh,
            )
        except ArquivoInvalidoError as e:
            # Erro amigavel, ja' pronto para mostrar direto pra usuaria -
            # sem traceback tecnico.
            self._finalizar(erro=str(e))
        except FileNotFoundError as e:
            self._finalizar(erro=f"Não encontrei um dos arquivos selecionados:\n{e}")
        except ValueError as e:
            self._finalizar(erro=f"Um dos arquivos não está no formato esperado:\n{e}")
        except Exception:
            # Ultimo recurso: erro que nao encaixou em nenhuma categoria
            # conhecida. Ainda assim evita despejar o traceback inteiro na
            # tela - mostra so' a ultima linha (a causa mais provavel) e
            # sugere levar isso ao suporte tecnico.
            detalhe = traceback.format_exc().strip().splitlines()[-1]
            self._finalizar(
                erro=(
                    "Ocorreu um erro inesperado ao processar. Se o problema persistir, "
                    f"leve esta mensagem ao suporte técnico:\n\n{detalhe}"
                )
            )

    def _finalizar(self, sucesso=False, erro=None, pasta_saida="", n_arquivos=0, n_pendencias=0, usou_decisoes_rh=False):
        def atualizar_ui():
            self.botao_processar.config(state="normal", text="Processar Folha de Pagamento")
            if erro:
                self.barra_progresso["value"] = 0
                self.barra_status.config(text="Falhou. Veja a mensagem.")
                messagebox.showerror("Não foi possível processar", erro)
                return

            self.barra_progresso["value"] = len(self._ETAPAS)
            self.barra_status.config(text=f"Concluído: {n_arquivos} arquivo(s) gerado(s) em {pasta_saida}")
            if usou_decisoes_rh:
                aviso_pendencias = (
                    "\n\nEsta versão já usa as decisões do RH (decisoes_rh.csv encontrado "
                    "na pasta de saída) — pronta para enviar à contabilidade."
                )
            elif n_pendencias > 0:
                aviso_pendencias = (
                    f"\n\nEsta é a versão PRELIMINAR: {n_pendencias} dia(s) ainda precisam de "
                    "revisão manual (faltas sem batida, batida incompleta etc.) e foram tratados "
                    "como falta injustificada por padrão. Abra o dashboard_aprovacao.html, "
                    "carregue o fila_validacao_rh.csv desta pasta, classifique as exceções, "
                    "salve o decisoes_rh.csv NESTA MESMA PASTA e clique em \"Processar Folha de "
                    "Pagamento\" de novo para gerar a versão final."
                )
            else:
                aviso_pendencias = ""
            messagebox.showinfo(
                "Sucesso",
                f"Planilha gerada com sucesso!\n\n"
                f"{n_arquivos} arquivo(s) de Relação de Valores salvos em:\n{pasta_saida}"
                f"{aviso_pendencias}",
            )

        # thread em segundo plano nao pode mexer na UI direto - agenda na thread principal
        self.after(0, atualizar_ui)


def main():
    app = AplicativoFolha()
    app.mainloop()


if __name__ == "__main__":
    main()
