"""Carregador tolerante, validação por linha e adaptador do Cadastro de Colaboradores (dados fictícios)."""
from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.styles.colors import Color

from cadastro_editor import (
    cadastro_para_pipeline, ler_cadastro_para_edicao, normalizar_para_edicao, tem_erros, validar_cadastro,
)
from pipeline import carregar_cadastro_colaboradores

CABECALHO = [
    "pis", "matricula", "nome", "codigo_empresa", "razao_social", "entrada", "saida_almoco", "retorno_almoco",
    "saida", "dias_trabalho", "saida_sexta", "admissao", "cnpj", "centro_custo", "valor_va",
    "Cooper Card Assiduidade", "Alelo Alimentação", "Ticket Alimentação", "desconto_folha ", "observacao",
]
VERDE_FORTE = PatternFill("solid", fgColor="FF00B050")
VERDE_CLARO = PatternFill("solid", fgColor=Color(theme=9, tint=0.7999816888943144))


def _linha(matricula, empresa="001", razao="EMPRESA A LTDA", cnpj="17.155.411/0001-59", **extra) -> dict:
    base = dict(
        pis=str(matricula).zfill(11), matricula=matricula, nome=f"COLAB {matricula}", codigo_empresa=empresa,
        razao_social=razao, entrada="07:30", saida_almoco="12:00", retorno_almoco="13:00", saida="17:30",
        dias_trabalho="0,1,2,3,4", saida_sexta="16:30", admissao="01/03/2020", cnpj=cnpj, centro_custo="07",
    )
    base.update(extra)
    return base


def _montar_xlsx(caminho: Path, linhas: list[dict], extras_linha_a_linha: dict[int, dict] | None = None) -> Path:
    """Planilha no estilo da cliente: P/Q/R no lugar de cooper/valor_va, lixo abaixo e cores de preenchimento."""
    wb = Workbook()
    ws = wb.active
    ws.append(CABECALHO)
    for d in linhas:
        ws.append([
            d.get("pis"), d.get("matricula"), d.get("nome"), d.get("codigo_empresa"), d.get("razao_social"),
            d.get("entrada"), d.get("saida_almoco"), d.get("retorno_almoco"), d.get("saida"), d.get("dias_trabalho"),
            d.get("saida_sexta"), d.get("admissao"), d.get("cnpj"), d.get("centro_custo"), None,
            d.get("cooper"), d.get("alelo"), d.get("ticket"), None, d.get("observacao"),
        ])
    # lixo: 2 linhas em branco + legenda solta numa coluna qualquer
    ws.append([None] * 20)
    ws.append([None] * 20)
    ws.cell(row=ws.max_row, column=2, value="legenda de centros de custo")
    for (linha, coluna), preenchimento in (extras_linha_a_linha or {}).items():
        ws.cell(row=linha, column=coluna).fill = preenchimento
    wb.save(caminho)
    return caminho


# ---------------------------------------------------------------------------------- carregador
def test_cabecalhos_da_cliente_viram_colunas_canonicas_e_vale_por_empresa(tmp_path):
    caminho = _montar_xlsx(tmp_path / "c.xlsx", [
        _linha(1, "001", cooper=200, alelo=1100, ticket=None),
        _linha(2, "002", razao="EMPRESA B LTDA", cnpj="27.488.404/0001-22", cooper=140, alelo=None, ticket=900),
    ])
    res = ler_cadastro_para_edicao(str(caminho))
    df = res.df.set_index("matricula")
    assert df.at["1", "valor_va"] == 1100 and df.at["1", "beneficio"] == "Alelo Cartão"
    assert df.at["2", "valor_va"] == 900 and df.at["2", "beneficio"] == "Ticket"
    assert df.at["1", "cooper"] == 200 and df.at["2", "cooper"] == 140
    assert "desconto_folha" in df.columns and "desconto_folha " not in df.columns
    assert not any(c in df.columns for c in ("cooper card assiduidade", "alelo alimentacao", "ticket alimentacao"))


def test_valores_em_texto_da_cliente_com_milhar_americano_e_traco(tmp_path):
    """A cliente digitou " R$ 1,500.00 " (milhar com virgula) e " R$ -   " (sem vale): 1500 e vazio, nunca 1,5."""
    caminho = _montar_xlsx(tmp_path / "c.xlsx", [
        _linha(1, "001", cooper=" R$ 200.00 ", alelo=" R$ 1,500.00 ", ticket=" R$ -   "),
        _linha(2, "002", razao="EMPRESA B LTDA", cnpj="27.488.404/0001-22", cooper=" R$ 140.00 ",
               alelo=" R$ -   ", ticket=" R$ 960.00 "),
    ])
    df = ler_cadastro_para_edicao(str(caminho)).df.set_index("matricula")
    assert df.at["1", "valor_va"] == 1500 and df.at["1", "beneficio"] == "Alelo Cartão"
    assert df.at["2", "valor_va"] == 960 and df.at["2", "beneficio"] == "Ticket"
    assert df.at["1", "cooper"] == 200 and df.at["2", "cooper"] == 140


def test_linhas_sem_matricula_e_pis_sao_ignoradas_com_aviso(tmp_path):
    caminho = _montar_xlsx(tmp_path / "c.xlsx", [_linha(1), _linha(2)])
    res = ler_cadastro_para_edicao(str(caminho))
    assert list(res.df["matricula"]) == ["1", "2"]
    assert any("2 linhas ignoradas" in a and "4 a 5" in a for a in res.avisos), res.avisos


def test_normaliza_matricula_admissao_e_centro_de_custo(tmp_path):
    caminho = _montar_xlsx(tmp_path / "c.xlsx", [
        _linha(14.0, admissao="16/11/2026", centro_custo="07"),
        _linha(15, admissao=pd.Timestamp(2020, 5, 3).to_pydatetime(), centro_custo=25),
    ])
    df = ler_cadastro_para_edicao(str(caminho)).df.set_index("matricula")
    assert list(df.index) == ["14", "15"]            # sem "14.0"
    assert df.at["14", "admissao"] == date(2026, 11, 16)
    assert df.at["15", "admissao"] == date(2020, 5, 3)
    assert df.at["14", "centro_custo"] == "07"       # zero à esquerda preservado
    assert df.at["15", "centro_custo"] == "25"


def test_cores_preenchem_fixo_e_desconto_manual(tmp_path):
    caminho = _montar_xlsx(
        tmp_path / "c.xlsx",
        [_linha(1, cooper=140, alelo=1500), _linha(2, cooper=200, alelo=1000), _linha(3, cooper=200, alelo=1000)],
        extras_linha_a_linha={
            (2, 16): VERDE_FORTE,   # linha 1 do cadastro, coluna Cooper
            (2, 17): VERDE_FORTE,   # coluna Alelo
            **{(4, c): VERDE_CLARO for c in range(3, 17)},   # linha inteira do colaborador 3
        },
    )
    res = ler_cadastro_para_edicao(str(caminho))
    df = res.df.set_index("matricula")
    assert bool(df.at["1", "cooper_fixo"]) and bool(df.at["1", "va_fixo"])
    assert not df.at["2", "cooper_fixo"] and not df.at["2", "va_fixo"]
    assert bool(df.at["3", "va_desconto_manual"]) and not df.at["1", "va_desconto_manual"]
    assert not df.at["3", "va_fixo"]    # verde claro nao implica fixo


def test_falha_na_leitura_de_cores_mantem_padrao_e_avisa(tmp_path, monkeypatch):
    import cadastro_editor

    def quebra(*a, **k):
        raise RuntimeError("falha simulada")

    monkeypatch.setattr(cadastro_editor, "_ler_cores", quebra)
    caminho = _montar_xlsx(tmp_path / "c.xlsx", [_linha(1)])
    res = ler_cadastro_para_edicao(str(caminho))
    assert not res.df["va_fixo"].any()
    assert any("cores" in a.lower() for a in res.avisos)


def test_cadastro_sintetico_continua_carregando():
    caminho = Path("Dados e Docs Reais/cadastro_sintetico.xlsx")
    if not caminho.exists():
        return
    df = carregar_cadastro_colaboradores(str(caminho))
    assert len(df) > 0 and {"pis", "banco_horas", "registra_ponto"} <= set(df.columns)
    assert not [p for p in validar_cadastro(ler_cadastro_para_edicao(str(caminho)).df) if p.severidade == "erro"
                and p.campo in ("entrada", "saida", "razao_social")]


def test_carregador_antigo_tambem_entende_aliases_e_ignora_lixo(tmp_path):
    caminho = _montar_xlsx(tmp_path / "c.xlsx", [_linha(1, cooper=200, alelo=1100)])
    df = carregar_cadastro_colaboradores(str(caminho))
    assert len(df) == 1 and df.at[0, "cooper"] == 200 and df.at[0, "valor_va"] == 1100


# ---------------------------------------------------------------------------------- validação
def _df(*linhas: dict) -> pd.DataFrame:
    return normalizar_para_edicao(pd.DataFrame(list(linhas)))


def _problemas(df, **kw):
    return validar_cadastro(df, **kw)


def _achar(problemas, campo, severidade, matricula=None):
    return [p for p in problemas if p.campo == campo and p.severidade == severidade
            and (matricula is None or p.matricula == matricula)]


def test_linha_incompleta_e_erro_apontando_a_matricula():
    df = _df(_linha(1), _linha(14, razao=None, entrada=None, saida_almoco=None, retorno_almoco=None, saida=None))
    ps = _problemas(df)
    for campo in ("razao_social", "entrada", "saida_almoco", "retorno_almoco", "saida"):
        achados = _achar(ps, campo, "erro", "14")
        assert achados, campo
        assert "14" in achados[0].mensagem
    assert tem_erros(ps)
    assert not [p for p in ps if p.matricula == "1" and p.severidade == "erro"]


def test_hora_invalida_dias_trabalho_e_valores_negativos_sao_erro():
    df = _df(_linha(1, entrada="25:99", dias_trabalho="1-5", valor_va=-5, cooper=-1))
    ps = _problemas(df)
    assert _achar(ps, "entrada", "erro", "1")
    assert _achar(ps, "dias_trabalho", "erro", "1")
    assert _achar(ps, "valor_va", "erro", "1") and _achar(ps, "cooper", "erro", "1")


def test_matricula_e_pis_duplicados_sao_erro():
    df = _df(_linha(1), _linha(1, pis="99999999999"), _linha(2, pis="00000000001"))
    ps = _problemas(df)
    assert _achar(ps, "matricula", "erro", "1")
    assert _achar(ps, "pis", "erro")


def test_cnpj_invalido_e_aviso_com_sugestao_da_mesma_raiz():
    df = _df(_linha(1, cnpj="17.155.411/0001-22"), _linha(2, cnpj="17.155.411/0001-59"))
    ps = _problemas(df)
    avisos = _achar(ps, "cnpj", "aviso", "1")
    assert avisos and "17.155.411/0001-59" in avisos[0].mensagem
    assert not _achar(ps, "cnpj", "aviso", "2")


def test_pis_igual_a_matricula_com_zeros_e_aviso_sobre_cartao_ponto():
    ps = _problemas(_df(_linha(14, pis="00000000014")))
    achados = _achar(ps, "pis", "aviso", "14")
    assert achados and "Cartão Ponto" in achados[0].mensagem


def test_padrao_de_jornada_diferente_da_maioria_da_empresa_e_aviso():
    df = _df(_linha(1), _linha(2), _linha(3, entrada="09:00", saida="18:00"))
    ps = _problemas(df)
    assert _achar(ps, "entrada", "aviso", "3")
    assert not _achar(ps, "entrada", "aviso", "1")


def test_admissao_apos_o_fim_do_periodo_e_aviso():
    df = _df(_linha(1, admissao="30/11/2026"))
    assert _achar(_problemas(df, data_fim=date(2026, 10, 25)), "admissao", "aviso", "1")
    assert not _achar(_problemas(df, data_fim=date(2026, 12, 25)), "admissao", "aviso", "1")


def test_saida_sexta_ausente_e_banco_horas_ausente_sao_avisos():
    df = _df(_linha(1, saida_sexta=None))
    ps = _problemas(df)
    assert _achar(ps, "saida_sexta", "aviso", "1")
    assert any(p.campo == "banco_horas" and p.severidade == "aviso" for p in _problemas(df.drop(columns=["banco_horas"])))


def test_cadastro_correto_nao_tem_erro():
    assert not tem_erros(_problemas(_df(_linha(1), _linha(2))))


# ---------------------------------------------------------------------------------- adaptador
def test_adaptador_garante_tipos_que_o_pipeline_espera():
    df = _df(_linha(1, va_fixo=True), _linha(2, valor_va="1.100,50", admissao="16/11/2026"))
    ap = cadastro_para_pipeline(df)
    assert ap["pis"].tolist() == ["00000000001", "00000000002"]
    assert ap["banco_horas"].dtype == bool and ap["registra_ponto"].dtype == bool
    assert ap["va_fixo"].dtype == bool and ap["cooper_fixo"].dtype == bool
    assert ap["valor_va"].dtype == float and ap.at[1, "valor_va"] == 1100.5
    assert ap.at[1, "admissao"] == date(2026, 11, 16)


def test_legenda_com_codigo_na_coluna_de_matricula_tambem_e_ignorada_com_aviso(tmp_path):
    caminho = _montar_xlsx(tmp_path / "c.xlsx", [_linha(1)])
    from openpyxl import load_workbook

    wb = load_workbook(caminho)
    ws = wb.active
    ws.append([None, "07", "COSTURA"])          # legenda de centro de custo: "matrícula" e nome, sem PIS nem jornada
    wb.save(caminho)
    res = ler_cadastro_para_edicao(str(caminho))
    assert list(res.df["matricula"]) == ["1"]
    assert any("3 linhas ignoradas" in a for a in res.avisos), res.avisos


def test_planilha_baixada_pode_ser_reenviada_sem_perda(tmp_path):
    from cadastro_editor import cadastro_para_xlsx_bytes

    original = ler_cadastro_para_edicao(str(_montar_xlsx(tmp_path / "c.xlsx", [
        _linha(1, cooper=140, alelo=1500, admissao="16/11/2026"), _linha(2, "002", centro_custo="07", ticket=900),
    ], extras_linha_a_linha={(2, 16): VERDE_FORTE, (2, 17): VERDE_FORTE}))).df
    novo = tmp_path / "limpa.xlsx"
    novo.write_bytes(cadastro_para_xlsx_bytes(original))
    volta = ler_cadastro_para_edicao(str(novo))
    assert not volta.avisos or all("ignoradas" not in a for a in volta.avisos)
    pd.testing.assert_frame_equal(volta.df, original, check_dtype=False)


# ---------------------------------------------------------------------------------- datas em texto e cores
def test_admissao_em_texto_mes_dia_ano_e_detectada_pela_propria_coluna(tmp_path):
    caminho = _montar_xlsx(tmp_path / "c.xlsx", [
        _linha(1, admissao="11/16/26"),    # so' pode ser mes/dia (16 nao e' mes)
        _linha(2, admissao="3/11/26"),     # ambigua isolada: segue a convencao da coluna (11/03/2026)
        _linha(3, admissao="2/1/22"),
    ])
    res = ler_cadastro_para_edicao(str(caminho))
    df = res.df.set_index("matricula")
    assert df.at["1", "admissao"] == date(2026, 11, 16)
    assert df.at["2", "admissao"] == date(2026, 3, 11)
    assert df.at["3", "admissao"] == date(2022, 2, 1)
    assert any("mês/dia/ano" in a and "3 datas" in a for a in res.avisos), res.avisos


def test_admissao_em_texto_dia_mes_ano_nao_muda(tmp_path):
    caminho = _montar_xlsx(tmp_path / "c.xlsx", [_linha(1, admissao="16/11/2026"), _linha(2, admissao="3/11/26")])
    res = ler_cadastro_para_edicao(str(caminho))
    df = res.df.set_index("matricula")
    assert df.at["1", "admissao"] == date(2026, 11, 16) and df.at["2", "admissao"] == date(2026, 11, 3)
    assert not any("mês/dia/ano" in a for a in res.avisos)


def test_admissao_com_formatos_misturados_nao_adivinha_e_avisa(tmp_path):
    caminho = _montar_xlsx(tmp_path / "c.xlsx", [
        _linha(1, admissao="16/11/2026"), _linha(2, admissao="11/16/2026"), _linha(3, admissao="3/11/2026"),
    ])
    res = ler_cadastro_para_edicao(str(caminho))
    df = res.df.set_index("matricula")
    assert df.at["1", "admissao"] == date(2026, 11, 16) and df.at["2", "admissao"] == date(2026, 11, 16)
    assert df.at["3", "admissao"] is None          # ambigua num arquivo misto: nunca adivinhada
    assert any("misturam" in a and "matrículas 3" in a for a in res.avisos), res.avisos


def test_planilha_sem_nenhuma_cor_avisa_que_os_fixos_nao_vieram(tmp_path):
    sem_cor = ler_cadastro_para_edicao(str(_montar_xlsx(tmp_path / "a.xlsx", [_linha(1, cooper=200, alelo=1000)])))
    assert any("nenhuma cor de preenchimento" in a for a in sem_cor.avisos), sem_cor.avisos
    com_cor = ler_cadastro_para_edicao(str(_montar_xlsx(
        tmp_path / "b.xlsx", [_linha(1, cooper=200, alelo=1000)], extras_linha_a_linha={(2, 16): VERDE_FORTE})))
    assert not any("nenhuma cor" in a for a in com_cor.avisos)
    # preenchimento que nao e' verde (ex.: so' amarelo): as cores existem, entao nao ha' o aviso de perda
    outra = ler_cadastro_para_edicao(str(_montar_xlsx(
        tmp_path / "c.xlsx", [_linha(1)], extras_linha_a_linha={(2, 3): PatternFill("solid", fgColor="FFFFFF00")})))
    assert not any("nenhuma cor" in a for a in outra.avisos)


def test_ano_de_2_e_de_4_digitos_sao_avaliados_em_grupos_separados(tmp_path):
    # como na planilha da cliente: quase tudo mm/dd/aa (2 digitos) e uma data dd/mm/aaaa (4 digitos)
    caminho = _montar_xlsx(tmp_path / "c.xlsx", [
        _linha(1, admissao="3/13/26"), _linha(2, admissao="3/9/26"), _linha(3, admissao="14/10/2026"),
    ])
    res = ler_cadastro_para_edicao(str(caminho))
    df = res.df.set_index("matricula")
    assert df.at["1", "admissao"] == date(2026, 3, 13) and df.at["2", "admissao"] == date(2026, 3, 9)
    assert df.at["3", "admissao"] == date(2026, 10, 14)
    assert not any("misturam" in a for a in res.avisos)


def test_mensagem_de_validacao_cita_a_linha_da_planilha(tmp_path):
    caminho = _montar_xlsx(tmp_path / "c.xlsx", [_linha(1), _linha(2, entrada="")])
    res = ler_cadastro_para_edicao(str(caminho))
    assert res.linhas_planilha == {"1": 2, "2": 3}
    problemas = validar_cadastro(res.df, linhas_planilha=res.linhas_planilha)
    erro = next(p for p in problemas if p.campo == "entrada" and p.severidade == "erro")
    assert erro.matricula == "2" and erro.mensagem.startswith("Matrícula 2 (linha 3 da planilha): ")
    # sem o mapa (ou para linha acrescentada na tela) a mensagem segue so' com a matricula
    sem_mapa = next(p for p in validar_cadastro(res.df) if p.campo == "entrada" and p.severidade == "erro")
    assert sem_mapa.mensagem.startswith("Matrícula 2: ")


# ---------------------------------------------------------------------------------- conferencia na tela (P/Q/R e datas)
def _planilha_com_cabecalhos(caminho: Path, cabecalho: list[str], linhas: list[list]) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.append(cabecalho)
    for linha in linhas:
        ws.append(linha)
    wb.save(caminho)
    return caminho


def test_cabecalhos_com_espacos_sobrando_mapeiam_vale_e_cooper_por_empresa(tmp_path):
    cab = ["matricula", "pis", "codigo_empresa", "razao_social", "entrada", "saida_almoco", "retorno_almoco", "saida",
           "dias_trabalho", " Cooper Card Assiduidade ", " Alelo Alimentação ", "Ticket Alimentação "]
    base = ["07:30", "12:00", "13:00", "17:30", "0,1,2,3,4"]
    caminho = _planilha_com_cabecalhos(tmp_path / "c.xlsx", cab, [
        [1, "00000000001", "001", "EMPRESA A LTDA", *base, 200, 1100, 999],    # ticket preenchido, mas e' empresa 001
        [2, "00000000002", "002", "EMPRESA B LTDA", *base, 140, 888, 900],     # alelo preenchido, mas e' empresa 002
    ])
    df = ler_cadastro_para_edicao(str(caminho)).df.set_index("matricula")
    assert df.at["1", "valor_va"] == 1100 and df.at["1", "beneficio"] == "Alelo Cartão"
    assert df.at["2", "valor_va"] == 900 and df.at["2", "beneficio"] == "Ticket"
    assert df.at["1", "cooper"] == 200 and df.at["2", "cooper"] == 140
    assert not any(c.strip() != c or c in ("_vale_alelo", "_vale_ticket") for c in df.columns)


def test_empresa_vazia_ou_desconhecida_gera_aviso_claro(tmp_path):
    caminho = _montar_xlsx(tmp_path / "c.xlsx", [
        _linha(1, "003", alelo=500), _linha(2, "", razao="X", ticket=700), _linha(3, "001", alelo=100),
    ])
    res = ler_cadastro_para_edicao(str(caminho))
    assert any("Matrícula 1" in a and "'003'" in a for a in res.avisos), res.avisos
    assert any("Matrícula 2" in a and "vazio" in a for a in res.avisos), res.avisos
    assert not any("Matrícula 3" in a and "código de empresa" in a for a in res.avisos)


def test_datas_todas_ambiguas_assumem_dia_mes_e_avisam(tmp_path):
    caminho = _montar_xlsx(tmp_path / "c.xlsx", [_linha(1, admissao="3/4/26"), _linha(2, admissao="5/6/26")])
    res = ler_cadastro_para_edicao(str(caminho))
    df = res.df.set_index("matricula")
    assert df.at["1", "admissao"] == date(2026, 4, 3) and df.at["2", "admissao"] == date(2026, 6, 5)
    assert res.formato_admissao == "dia/mês/ano"
    assert any("nenhum dia acima de 12" in a and "dia/mês/ano" in a for a in res.avisos), res.avisos


def test_formato_detectado_mes_dia_ano_e_exposto(tmp_path):
    res = ler_cadastro_para_edicao(str(_montar_xlsx(tmp_path / "c.xlsx", [_linha(1, admissao="3/13/26"), _linha(2, admissao="3/4/26")])))
    assert res.formato_admissao == "mês/dia/ano"
    assert not any("nenhum dia acima de 12" in a for a in res.avisos)


def test_sem_cores_e_exposto_e_resumo_conta_vale_cooper_e_marcas(tmp_path):
    from cadastro_editor import resumo_conferencia
    caminho = _montar_xlsx(tmp_path / "c.xlsx", [
        _linha(1, "001", cooper=200, alelo=1100, admissao="3/13/26"),
        _linha(2, "002", razao="B", cooper=None, ticket=900, admissao="3/4/26"),
        _linha(3, "002", razao="B", ticket=None),
    ])
    res = ler_cadastro_para_edicao(str(caminho))
    assert res.sem_cores is True
    r = resumo_conferencia(res.df)
    assert r["total"] == 3 and r["vale_por_empresa"] == {"001": {"Alelo Cartão": 1}, "002": {"Ticket": 1}}
    assert r["sem_vale"] == 1 and r["cooper"] == 1
    assert (r["va_fixo"], r["cooper_fixo"], r["va_desconto_manual"]) == (0, 0, 0)
    assert r["amostra_admissoes"][0] == ("1", "13/03/2026")
    com_cor = ler_cadastro_para_edicao(str(_montar_xlsx(
        tmp_path / "d.xlsx", [_linha(1, cooper=200, alelo=1000)], extras_linha_a_linha={(2, 16): VERDE_FORTE})))
    assert com_cor.sem_cores is False and resumo_conferencia(com_cor.df)["cooper_fixo"] == 1
