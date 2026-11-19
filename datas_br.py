"""Entrada de datas no formato brasileiro (dd/mm/aaaa) para a camada de interface.

Parse estrito de propósito: nada de pd.to_datetime nem dayfirst, que aceitam formatos
ambíguos (ex.: 2026/11/18 ou 11/18/2026) e poderiam inverter dia e mês sem avisar.
"""
from __future__ import annotations

from datetime import date, datetime


def parse_data_br(texto: object) -> date | None:
    """Converte "dd/mm/aaaa" (ou "ddmmaaaa", só dígitos) em date; None se vazio ou inválido."""
    if not isinstance(texto, str):
        return None
    t = texto.strip()
    if t.isdigit() and len(t) == 8:
        t = f"{t[:2]}/{t[2:4]}/{t[4:]}"
    try:
        return datetime.strptime(t, "%d/%m/%Y").date()
    except ValueError:
        return None


def formatar_data_br(valor: object) -> str:
    """Formata date/datetime/string ISO como dd/mm/aaaa; devolve o texto original se já não for convertível.

    Usado para restaurar valores guardados em sessões antigas (objetos date ou strings ISO).
    """
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, date):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, str):
        try:
            return datetime.strptime(valor.strip()[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            return valor
    return ""
