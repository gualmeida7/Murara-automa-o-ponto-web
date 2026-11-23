"""
planilha_thales.py
==================
Gera a "Planilha que passo ao Thales": o resumo mensal do fechamento, uma tabela por empresa, no
mesmo modelo que a cliente monta a mao (Fechamento mm/aaaa, periodo, CC, colaborador, faltas com e
sem atestado, H/E 50% e 100%, desconto em folha, Alelo/Ticket, Cooper e observacoes).

O que sai do ponto (faltas, horas extras, atrasos, banco de horas) vem da apuracao; o que nao vem do
ponto vem de colunas OPCIONAIS do Cadastro de Colaboradores (vazias = celula em branco):

    centro_custo   texto livre do CC (ex.: "07 e 25")
    valor_va       valor mensal do vale (Alelo/Ticket) - R$
    va_fixo        sim/nao - vale fixo: nao sofre o desconto de 1/30 por falta
    va_desconto_manual  sim/nao - verde claro da cliente: o desconto do vale e' lancado a mao; o valor sai
                   cheio e, se houve falta injustificada, a Observacao avisa quantos dias
    beneficio      nome do vale no cabecalho da empresa (padrao "Alelo Cartão"; ex.: "Ticket ")
    cooper         valor mensal da Cooper - R$
    desconto_folha desconto em folha do mes - R$
    observacao     texto livre que vai na coluna Observacoes
    admissao       data de admissao (dd/mm/aaaa)
    nota_rodape    nota escrita abaixo da tabela da empresa, como "Nome: texto"
                   (ex.: "atestado em 28/05, detenção desde 24/06")

VA: admitido no proprio mes = proporcional aos dias ate o fim do mes (base 30); cada falta inteira
SEM justificativa desconta 1/30. Banco de horas e experiencia viram observacao automatica.
Cooper: perde o valor (e a Observacao avisa) quando as faltas do mes passam de LIMITE_FALTAS_COOPER_MIN,
salvo `cooper_fixo`; ver `pipeline.rodada_2_gerar_relacao_de_valores` (colunas `_cooper_zerado`).
Nada aqui decide se a falta e' justificada: usa a decisao do RH (padrao conservador: injustificada).
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd

from business_rules import LIMITE_FALTAS_COOPER_MIN

CABECALHO_BENEFICIO_PADRAO = "Alelo Cartão"
_FORMATO_REAL = '_-"R$"\\ * #,##0.00_-;\\-"R$"\\ * #,##0.00_-;_-"R$"\\ * "-"??_-;_-@_-'


def _hhmm(horas: float) -> str:
    """Horas decimais -> 'hh:mm' (2 digitos), '' quando zero."""
    minutos = round(abs(float(horas)) * 60)
    return f"{minutos // 60:02d}:{minutos % 60:02d}" if minutos else ""


def _texto_dias(datas: list[date]) -> str:
    return " e ".join(d.strftime("%d/%m/%Y") for d in datas)


def _numero(valor) -> float | None:
    if valor is None or pd.isna(valor) or str(valor).strip() == "":
        return None
    return float(str(valor).replace("R$", "").replace(".", "").replace(",", ".")) if isinstance(valor, str) else float(valor)


def _verdadeiro(valor) -> bool:
    return valor is not None and not pd.isna(valor) and bool(valor)


def detalhe_faltas(apuracao: pd.DataFrame, decisoes: pd.DataFrame | None) -> dict[str, dict[str, list]]:
    """
    Por colaborador: datas de falta inteira justificada / injustificada e horas (parcial + atraso) de
    falta justificada / injustificada. Mesma regra do `consolidar_mes`: sem decisao do RH, injustificada.
    """
    df = apuracao.copy()
    df["data"] = pd.to_datetime(df["data"])
    decididas = {}
    if decisoes is not None and not decisoes.empty:
        d = decisoes.copy()
        d["pis"] = d["pis"].astype(str).str.strip()
        d["data"] = pd.to_datetime(d["data"], dayfirst=True)
        decididas = {(r.pis, r.data): r.classificacao for r in d.itertuples()}

    saida: dict[str, dict[str, list]] = {}
    for r in df.itertuples():
        falta_dia = bool(r.falta)
        parcial = r.status == "FALTA_PARCIAL_CANDIDATA" and not falta_dia
        item = saida.setdefault(r.pis, {"dias_just": [], "dias_injust": [], "horas_just": 0.0, "horas_injust": 0.0})
        if falta_dia or parcial:
            classe = decididas.get((r.pis, r.data), "FALTA_INJUSTIFICADA")
            if classe == "ERRO_RELOGIO_ABONAR":
                continue
            if falta_dia:
                item["dias_just" if classe == "FALTA_JUSTIFICADA" else "dias_injust"].append(r.data.date())
            else:
                item["horas_just" if classe == "FALTA_JUSTIFICADA" else "horas_injust"] += r.minutos_falta_parcial / 60
        if r.atraso_min:
            item["horas_injust"] += r.atraso_min / 60
    return saida


def exportar_planilha_thales(
    relacao: pd.DataFrame,
    apuracao: pd.DataFrame,
    decisoes: pd.DataFrame | None,
    cadastro: pd.DataFrame,
    competencia: str,
    data_inicio: date,
    data_fim: date,
    pasta_saida: str,
) -> Path:
    """Escreve `planilha_thales_<mmaaaa>.xlsx` (uma tabela por empresa) e devolve o caminho."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    from pipeline import _nome_proprio

    mes, ano = (int(p) for p in competencia.split("/"))
    detalhes = detalhe_faltas(apuracao, decisoes)
    cad = cadastro.drop_duplicates("matricula").set_index(cadastro.drop_duplicates("matricula")["matricula"].astype(str))
    fino = Side(style="thin")
    borda = Border(left=fino, right=fino, top=fino, bottom=fino)
    amarelo = PatternFill("solid", fgColor="FFFF00")
    fonte = lambda tam=11, neg=False: Font(name="Aptos Narrow", size=tam, bold=neg)  # noqa: E731

    wb = Workbook()
    ws = wb.active
    ws.title = "Planilha1"
    ws["A1"] = f"Fechamento {competencia} "
    ws["A1"].font = fonte(14, True)
    ws["A2"] = f"Período {data_inicio:%d/%m/%Y} a {data_fim:%d/%m/%Y}"
    ws["A2"].font = fonte(12)

    linha = 4
    for (codigo, razao), grupo in relacao.groupby(["codigo_empresa", "razao_social"], sort=False):
        ws.cell(row=linha, column=2, value=f"{razao}:").font = fonte(16, True)
        linha += 1
        rotulo_va = next(
            (str(cad.at[str(m), "beneficio"]) for m in grupo["codigo_folha"]
             if str(m) in cad.index and "beneficio" in cad.columns and pd.notna(cad.at[str(m), "beneficio"])
             and str(cad.at[str(m), "beneficio"]).strip()),
            CABECALHO_BENEFICIO_PADRAO,
        )
        for col, titulo in enumerate(
            [None, "CC", "Colaborador", "Falta C/ At.", "Falta S/ At.", "H/E 50%", "H/E 100%",
             "Desc. em folha ", rotulo_va, "Cooper", "Observações"], start=1,
        ):
            c = ws.cell(row=linha, column=col, value=titulo)
            c.font, c.fill, c.border = fonte(9), amarelo, borda
            c.alignment = Alignment(horizontal=None if col == 3 else "center")
        primeira = linha + 1
        for n, (_, p) in enumerate(grupo.iterrows(), start=1):
            linha += 1
            matricula = str(p["codigo_folha"])
            ficha = cad.loc[matricula] if matricula in cad.index else pd.Series(dtype=object)
            pis = str(ficha.get("pis", "")).strip()
            det = detalhes.get(pis, {"dias_just": [], "dias_injust": [], "horas_just": 0.0, "horas_injust": 0.0})

            admissao = ficha.get("admissao")
            admissao = pd.to_datetime(admissao, dayfirst=True).date() if admissao is not None and pd.notna(admissao) else None
            valor_va = _numero(ficha.get("valor_va"))
            obs = []
            if valor_va is not None:
                valor_dia_va = valor_va / 30
                if admissao and (admissao.year, admissao.month) == (ano, mes):
                    ultimo = (pd.Timestamp(ano, mes, 1) + pd.offsets.MonthEnd(0)).day
                    valor_va = valor_dia_va * min(30, ultimo - admissao.day + 1)
                # va_fixo (verde forte da cliente): valor fixo, nunca sofre o desconto de 1/30 por falta.
                dias_calculados = int(p.get("_dias_desconto_va", 0) or 0)
                manual = _verdadeiro(ficha.get("va_desconto_manual")) and not _verdadeiro(ficha.get("va_fixo"))
                dias_desconto = 0 if _verdadeiro(ficha.get("va_fixo")) or manual else dias_calculados
                if manual and dias_calculados:
                    obs.append(f"*Desconto do vale a lançar manualmente ({dias_calculados} dia(s) de falta sem justificativa)")
                valor_va = round(valor_va - valor_dia_va * dias_desconto, 2)
            if admissao and (admissao.year, admissao.month) == (ano, mes):
                obs.append(f"*Experiencia (Admitido em {admissao:%d/%m/%Y})")
            valor_cooper = _numero(ficha.get("cooper"))
            if _verdadeiro(p.get("_cooper_zerado")):
                valor_cooper = 0.0
                obs.append(f"*Cooper zerado (faltas no mês: {_hhmm(float(p.get('_faltas_cooper_min', 0)) / 60)}, limite {_hhmm(LIMITE_FALTAS_COOPER_MIN / 60)})")
            saldo = p.get(COD_BANCO)
            if str(p.get("_banco_horas")) == "Sim" and saldo is not None and pd.notna(saldo) and _hhmm(saldo):
                obs.append(f"*Banco de Horas ({'-' if saldo < 0 else ''}{_hhmm(saldo)} {'Negativas' if saldo < 0 else 'Positivas'})")
            if pd.notna(ficha.get("observacao", None)) and str(ficha.get("observacao")).strip():
                obs.insert(0, str(ficha.get("observacao")).strip())

            # Quem esta no banco de horas leva atraso/falta parcial para o banco, nao para a folha.
            no_banco = str(p.get("_banco_horas")) == "Sim"
            horas_sem = [_texto_dias(det["dias_injust"]), "" if no_banco else _hhmm(det["horas_injust"])]
            horas_com = [_texto_dias(det["dias_just"]), "" if no_banco else _hhmm(det["horas_just"])]
            valores = [
                n, ficha.get("centro_custo") if pd.notna(ficha.get("centro_custo", None)) else None,
                _nome_proprio(p["nome_colaborador"]),
                " e ".join(x for x in horas_com if x) or None,
                " e ".join(x for x in horas_sem if x) or None,
                _hhmm(p["0150"]) or None, _hhmm(p["0200"]) or None,
                _numero(ficha.get("desconto_folha")), valor_va, valor_cooper,
                " | ".join(obs) or None,
            ]
            for col, v in enumerate(valores, start=1):
                c = ws.cell(row=linha, column=col, value=v)
                c.font, c.border = fonte(10 if 4 <= col <= 8 else 11), borda
                c.alignment = Alignment(horizontal=None if col in (3, 11) else "center")
                if col in (8, 9, 10):
                    c.number_format = _FORMATO_REAL
        linha += 1
        for col in (9, 10):
            letra = "I" if col == 9 else "J"
            c = ws.cell(row=linha, column=col, value=f"=SUM({letra}{primeira}:{letra}{linha - 1})")
            c.font, c.number_format = fonte(11, True), _FORMATO_REAL
            c.alignment = Alignment(horizontal="center")
        # Notas de rodape da empresa (coluna `nota_rodape` do cadastro): "Nome: texto", uma por linha.
        for _, p in grupo.iterrows():
            matricula = str(p["codigo_folha"])
            nota = cad.at[matricula, "nota_rodape"] if matricula in cad.index and "nota_rodape" in cad.columns else None
            if nota is not None and pd.notna(nota) and str(nota).strip():
                linha += 1
                c = ws.cell(row=linha, column=2, value=f"{_nome_proprio(p['nome_colaborador']).split()[0]}: {str(nota).strip()}")
                c.font = fonte(11)
        linha += 4

    for letra, largura in zip("ABCDEFGHIJK", (5, 17, 41.7, 19.7, 18.1, 11.4, 11.4, 14, 13.4, 12.6, 44.6)):
        ws.column_dimensions[letra].width = largura
    caminho = Path(pasta_saida)
    caminho.mkdir(parents=True, exist_ok=True)
    caminho = caminho / f"planilha_thales_{mes:02d}{ano}.xlsx"
    wb.save(caminho)
    return caminho


COD_BANCO = "0999"
