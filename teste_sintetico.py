"""Gera dados sinteticos e roda o pipeline inteiro, so' para validar que o
codigo funciona de ponta a ponta antes de usar com dados reais."""
from datetime import date, datetime
from pathlib import Path
import pandas as pd

from pipeline import (
    carregar_cadastro_colaboradores, rodada_1_gerar_fila_de_excecoes,
    rodada_2_gerar_relacao_de_valores, exportar_por_empresa, carregar_consignados,
    exportar_fila_validacao_rh, executar_checagens_preflight,
)
from banco_horas import consolidar_banco_horas, ler_banco_horas_secullum, ler_banco_horas_adriano
from validacao import ArquivoInvalidoError

Path("dados").mkdir(exist_ok=True)
Path("saida").mkdir(exist_ok=True)

# ---- cadastro de colaboradores (3 pessoas: normal, Sidney, Adriano) ----
cadastro = pd.DataFrame([
    {"pis": "10000000001", "matricula": "M001", "nome": "Joao Silva",
     "codigo_empresa": "01", "razao_social": "Macaneiro e Gonzaga LTDA",
     "entrada": "08:00", "saida_almoco": "12:00", "retorno_almoco": "13:00", "saida": "17:00",
     "dias_trabalho": "0,1,2,3,4"},
    {"pis": "10000000002", "matricula": "M002", "nome": "Sidney Souza",
     "codigo_empresa": "02", "razao_social": "Cianorte Tubos LTDA",
     "entrada": "07:00", "saida_almoco": "11:00", "retorno_almoco": "12:00", "saida": "16:00",
     "dias_trabalho": "0,1,2,3,4"},
    {"pis": "10000000003", "matricula": "M003", "nome": "Adriano Costa",
     "codigo_empresa": "01", "razao_social": "Macaneiro e Gonzaga LTDA",
     "entrada": "08:00", "saida_almoco": "12:00", "retorno_almoco": "13:00", "saida": "17:00",
     "dias_trabalho": "0,1,2,3,4"},
])
cadastro.to_excel("dados/cadastro_colaboradores.xlsx", index=False)

# ---- AFD sintetico ----
# Layout: NSR(9) TIPO(1)='3' PIS(12) DATAHORA(12,AAAAMMDDHHMM)
def linha_afd(nsr, pis, dt: datetime):
    return f"{nsr:09d}3{pis:>12}{dt.strftime('%Y%m%d%H%M')}"

linhas = []
nsr = 1
# Joao: dia 27/07 normal, dia 28/07 fez 1h extra, dia 29/07 faltou (sem batida)
dia_batidas = {
    date(2026, 7, 27): [(8, 0), (12, 0), (13, 0), (17, 0)],
    date(2026, 7, 28): [(8, 0), (12, 0), (13, 0), (18, 0)],  # +1h -> HE50
    # 29/07 sem batidas -> falta
}
for dia, horarios in dia_batidas.items():
    for h, m in horarios:
        linhas.append(linha_afd(nsr, "10000000001", datetime(dia.year, dia.month, dia.day, h, m)))
        nsr += 1

# Sidney: batidas normais (banco de horas dele vem de outro relatorio, nao do AFD)
for h, m in [(7, 0), (11, 0), (12, 0), (16, 0)]:
    linhas.append(linha_afd(nsr, "10000000002", datetime(2026, 7, 27, h, m)))
    nsr += 1

# Adriano: um dia com batida incompleta (3 batidas) para testar excecao
for h, m in [(8, 0), (12, 0), (13, 0)]:
    linhas.append(linha_afd(nsr, "10000000003", datetime(2026, 7, 27, h, m)))
    nsr += 1

with open("dados/ponto.txt", "w", encoding="latin-1") as f:
    f.write("\n".join(linhas) + "\n")

# ---- consignados ----
pd.DataFrame([
    {"matricula": "M001", "valor_desconto": 150.00},
    {"matricula": "M003", "valor_desconto": 80.00},
]).rename(columns={"matricula": "pis"}).assign(
    pis=lambda d: cadastro.set_index("matricula").loc[d["pis"].map({"M001": "M001", "M003": "M003"})].reset_index()["pis"] if False else d["pis"]
).to_excel("dados/consignados_tmp.xlsx", index=False)
# (simplificando: gravar direto por pis para o teste)
pd.DataFrame([
    {"pis": "10000000001", "valor_desconto": 150.00},
    {"pis": "10000000003", "valor_desconto": 80.00},
]).to_excel("dados/consignados.xlsx", index=False)

# ---- banco de horas Sidney (export Secullum) ----
pd.DataFrame([{"pis": "10000000002", "saldo": 5.5}]).to_excel("dados/banco_horas_secullum.xlsx", index=False)

# ---- planilha semanal do Adriano ----
pd.DataFrame([
    {"semana": 1, "horas_extras": 2.0, "horas_debito": 0.0, "saldo_anterior": 3.0},
    {"semana": 2, "horas_extras": 1.0, "horas_debito": 0.5, "saldo_anterior": None},
]).to_excel("dados/banco_horas_adriano.xlsx", index=False)

print("=== Rodando pipeline ===")
cadastro_lido = carregar_cadastro_colaboradores("dados/cadastro_colaboradores.xlsx")
apuracao, fila = rodada_1_gerar_fila_de_excecoes(
    "dados/ponto.txt", cadastro_lido, date(2026, 7, 26), date(2026, 8, 25)
)
print("\n-- Fila de validacao RH (excecoes) --")
print(fila.to_string(index=False))
fila_exportada = exportar_fila_validacao_rh(fila, cadastro_lido, "saida/fila_validacao_rh.csv")
print(f"\nCSV exportado com nome/empresa, ordenado por colaborador+data, {len(fila_exportada)} linha(s).")

# Simula decisao do RH: falta do Joao em 29/07 foi INJUSTIFICADA,
# a batida incompleta do Adriano foi abonada (erro de relogio)
# (datas no formato brasileiro dd/mm/aaaa, igual ao que o dashboard exporta)
decisoes = pd.DataFrame([
    {"pis": "10000000001", "data": "29/07/2026", "classificacao": "FALTA_INJUSTIFICADA"},
    {"pis": "10000000003", "data": "27/07/2026", "classificacao": "ERRO_RELOGIO_ABONAR"},
])
decisoes.to_csv("saida/decisoes_rh.csv", index=False)

consignados = carregar_consignados("dados/consignados.xlsx", cadastro=cadastro_lido)
banco_secullum = ler_banco_horas_secullum("dados/banco_horas_secullum.xlsx", cadastro=cadastro_lido)
banco_adriano = ler_banco_horas_adriano("dados/banco_horas_adriano.xlsx", "10000000003", cadastro=cadastro_lido)
banco_horas = consolidar_banco_horas(cadastro_lido["pis"].tolist(), banco_secullum, banco_adriano)
print("\n-- Banco de horas consolidado --")
print(banco_horas.to_string(index=False))

relacao = rodada_2_gerar_relacao_de_valores(
    apuracao, cadastro_lido, consignados, banco_horas, decisoes, "08/2026"
)
print("\n-- Relacao de Valores (consolidada, antes de separar por empresa) --")
print(relacao.to_string(index=False))

arquivos = exportar_por_empresa(relacao, "saida")
print("\n-- Arquivos gerados --")
for a in arquivos:
    print(a)

# ---------------------------------------------------------------------------
# Teste de regressao: extrato auxiliar identificando por MATRICULA, nao PIS
# (bug real encontrado em producao - ver validacao.normalizar_identificador_colaborador)
# ---------------------------------------------------------------------------
print("\n=== Teste de regressao: consignados identificado por matricula ===")
consignados_por_matricula = pd.DataFrame([
    {"PIS": "M001", "Valor Desconto": 150.0},  # "M001" aqui e' a MATRICULA, nao o PIS
])
consignados_por_matricula.to_excel("dados/consignados_por_matricula.xlsx", index=False)
resultado = carregar_consignados("dados/consignados_por_matricula.xlsx", cadastro=cadastro_lido)
pis_esperado = cadastro_lido.loc[cadastro_lido["matricula"] == "M001", "pis"].iloc[0]
assert resultado.iloc[0]["pis"] == pis_esperado, "normalizacao matricula->pis falhou"
print(f"OK: matricula 'M001' foi normalizada corretamente para o PIS '{pis_esperado}'")

