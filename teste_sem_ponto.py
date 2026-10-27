"""Cadastrado sem nenhuma batida no ponto vira aviso em `apuracao.attrs["sem_ponto"]`."""
from __future__ import annotations

from datetime import date

import pandas as pd

from pipeline import rodada_1_gerar_fila_de_excecoes


def _colaborador(pis: str, nome: str, **extra) -> dict:
    return {
        "pis": pis, "matricula": f"M{pis[-3:]}", "nome": nome, "codigo_empresa": "01",
        "razao_social": "Empresa Teste", "entrada": "08:00", "saida_almoco": "12:00",
        "retorno_almoco": "13:00", "saida": "17:00", "dias_trabalho": "0,1,2,3,4", **extra,
    }


def test_cadastrado_sem_batida_e_avisado(tmp_path):
    cadastro = pd.DataFrame([
        _colaborador("10000000001", "Com Ponto"),
        _colaborador("10000000002", "Demitido"),
        _colaborador("10000000003", "Sem Relogio", registra_ponto=False),
    ])
    afd = tmp_path / "ponto.txt"
    afd.write_text(
        "".join(f"{n:09d}3{'10000000001':>12}20260727{hm}\n" for n, hm in enumerate(["0800", "1200", "1300", "1700"], 1)),
        encoding="latin-1",
    )
    apuracao, _ = rodada_1_gerar_fila_de_excecoes(
        str(afd), cadastro, date(2026, 7, 26), date(2026, 8, 25), feriados={},
    )
    # quem nao registra ponto nem entra na apuracao, entao nao e' cobrado
    assert apuracao.attrs["sem_ponto"] == ["10000000002"]
