from datetime import date

import pytest

from datas_br import formatar_data_br, parse_data_br


def test_data_valida():
    assert parse_data_br("18/11/2026") == date(2026, 11, 18)


def test_sem_barras():
    assert parse_data_br("18112026") == date(2026, 11, 18)


def test_espacos():
    assert parse_data_br("  18/11/2026 ") == date(2026, 11, 18)


@pytest.mark.parametrize("texto", [
    "31/02/2026", "18/13/2026", "", "   ", "18-11-2026", "18.11.2026",
    "2026/11/18", "2026-11-18", "18/11/26", "181126", None,
])
def test_invalidas(texto):
    assert parse_data_br(texto) is None


def test_formatar_restaura_formatos_antigos():
    assert formatar_data_br(date(2026, 11, 18)) == "18/11/2026"
    assert formatar_data_br("2026-11-18") == "18/11/2026"
    assert formatar_data_br("18/11/2026") == "18/11/2026"
