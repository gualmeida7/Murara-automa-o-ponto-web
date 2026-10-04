"""
db_sessao.py
============
SQLite persistence layer for the payroll closing session.
All functions are safe to call even if the database is unavailable —
they fail silently and log errors to fechamento_folha.log.
"""

from __future__ import annotations

import logging
import pathlib
import sqlite3
from contextlib import closing
from datetime import datetime

import pandas as pd

_PASTA = pathlib.Path(__file__).parent
DB_PATH = _PASTA / "fechamento_folha.db"
LOG_PATH = _PASTA / "fechamento_folha.log"

COLUNAS_DECISOES = ["pis", "data", "classificacao", "observacao"]

_logger = logging.getLogger("db_sessao")
if not _logger.handlers:
    try:
        _handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
        _handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        _logger.addHandler(_handler)
    except Exception:  # sem permissão de escrita etc. — o app segue sem log
        _logger.addHandler(logging.NullHandler())
    _logger.setLevel(logging.INFO)
    _logger.propagate = False

_SCHEMA = """
CREATE TABLE IF NOT EXISTS sessoes_fechamento (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    competencia     TEXT NOT NULL,
    criado_em       TEXT NOT NULL,
    atualizado_em   TEXT NOT NULL,
    status          TEXT NOT NULL DEFAULT 'em_andamento',
    total_excecoes  INTEGER DEFAULT 0,
    UNIQUE(competencia)
);

CREATE TABLE IF NOT EXISTS decisoes_rh (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    competencia     TEXT NOT NULL,
    pis             TEXT NOT NULL,
    data            TEXT NOT NULL,
    classificacao   TEXT NOT NULL DEFAULT '',
    observacao      TEXT NOT NULL DEFAULT '',
    salvo_em        TEXT NOT NULL,
    UNIQUE(competencia, pis, data)
);
"""


def _agora() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _conectar() -> sqlite3.Connection:
    return sqlite3.connect(DB_PATH, timeout=5)


def _txt(valor) -> str:
    """Converte para texto seguro para o banco (None/NaN viram '')."""
    if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
        return ""
    return str(valor)


def inicializar_banco() -> None:
    """Create the database file and tables if they don't exist. Called once at app startup."""
    try:
        with closing(_conectar()) as con, con:
            con.executescript(_SCHEMA)
    except Exception:
        _logger.exception("Falha ao inicializar o banco")


def salvar_decisao(competencia: str, pis: str, data: str,
                   classificacao: str, observacao: str) -> None:
    """Upsert a single decision row (INSERT OR REPLACE)."""
    try:
        with closing(_conectar()) as con, con:
            con.execute(
                "INSERT OR REPLACE INTO decisoes_rh "
                "(competencia, pis, data, classificacao, observacao, salvo_em) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (competencia, _txt(pis), _txt(data), _txt(classificacao), _txt(observacao), _agora()),
            )
    except Exception:
        _logger.exception("Falha ao salvar decisão (%s, %s, %s)", competencia, pis, data)


def salvar_decisoes_em_lote(competencia: str, decisoes: pd.DataFrame) -> bool:
    """
    Upsert many decisions in a single transaction, writing only the rows that
    are new or differ from what is already saved (so 'last save' stays accurate
    and a Streamlit rerun with no edits touches nothing).
    `decisoes` needs columns [pis, data, classificacao, observacao].
    Returns True if anything was written.
    """
    try:
        novas = {
            (_txt(r.pis), _txt(r.data)): (_txt(r.classificacao), _txt(r.observacao))
            for r in decisoes.itertuples(index=False)
        }
        if not novas:
            return False
        with closing(_conectar()) as con, con:
            existentes = {
                (pis, data): (cls, obs)
                for pis, data, cls, obs in con.execute(
                    "SELECT pis, data, classificacao, observacao FROM decisoes_rh WHERE competencia = ?",
                    (competencia,),
                )
            }
            agora = _agora()
            mudancas = [
                (competencia, pis, data, cls, obs, agora)
                for (pis, data), (cls, obs) in novas.items()
                if existentes.get((pis, data)) != (cls, obs)
            ]
            if mudancas:
                con.executemany(
                    "INSERT OR REPLACE INTO decisoes_rh "
                    "(competencia, pis, data, classificacao, observacao, salvo_em) "
                    "VALUES (?, ?, ?, ?, ?, ?)",
                    mudancas,
                )
            return bool(mudancas)
    except Exception:
        _logger.exception("Falha ao salvar decisões em lote (%s)", competencia)
        return False


def carregar_decisoes(competencia: str) -> pd.DataFrame:
    """
    Return all saved decisions for a competencia as a DataFrame with columns:
    [pis, data, classificacao, observacao]
    Returns an empty DataFrame if nothing is saved yet.
    """
    try:
        with closing(_conectar()) as con:
            linhas = con.execute(
                "SELECT pis, data, classificacao, observacao FROM decisoes_rh WHERE competencia = ?",
                (competencia,),
            ).fetchall()
        return pd.DataFrame(linhas, columns=COLUNAS_DECISOES)
    except Exception:
        _logger.exception("Falha ao carregar decisões (%s)", competencia)
        return pd.DataFrame(columns=COLUNAS_DECISOES)


def obter_sessao(competencia: str) -> dict | None:
    """
    Return the session record for a competencia, or None if it doesn't exist.
    Dict keys: competencia, criado_em, atualizado_em, status, total_excecoes
    """
    try:
        with closing(_conectar()) as con:
            con.row_factory = sqlite3.Row
            linha = con.execute(
                "SELECT competencia, criado_em, atualizado_em, status, total_excecoes "
                "FROM sessoes_fechamento WHERE competencia = ?",
                (competencia,),
            ).fetchone()
        return dict(linha) if linha else None
    except Exception:
        _logger.exception("Falha ao obter sessão (%s)", competencia)
        return None


def criar_ou_atualizar_sessao(competencia: str, total_excecoes: int,
                               status: str = "em_andamento") -> None:
    """Upsert the session header row (keeps the original criado_em)."""
    try:
        agora = _agora()
        with closing(_conectar()) as con, con:
            con.execute(
                "INSERT INTO sessoes_fechamento (competencia, criado_em, atualizado_em, status, total_excecoes) "
                "VALUES (?, ?, ?, ?, ?) "
                "ON CONFLICT(competencia) DO UPDATE SET "
                "atualizado_em = excluded.atualizado_em, "
                "status = excluded.status, "
                "total_excecoes = excluded.total_excecoes",
                (competencia, agora, agora, status, int(total_excecoes)),
            )
    except Exception:
        _logger.exception("Falha ao gravar sessão (%s)", competencia)


def limpar_sessao(competencia: str) -> None:
    """Delete all decisions and the session row for this competencia only."""
    try:
        with closing(_conectar()) as con, con:
            con.execute("DELETE FROM decisoes_rh WHERE competencia = ?", (competencia,))
            con.execute("DELETE FROM sessoes_fechamento WHERE competencia = ?", (competencia,))
    except Exception:
        _logger.exception("Falha ao limpar sessão (%s)", competencia)


def listar_sessoes() -> pd.DataFrame:
    """
    Return all sessions ordered by most recent first.
    Used for the history view (optional, future use).
    """
    colunas = ["competencia", "criado_em", "atualizado_em", "status", "total_excecoes"]
    try:
        with closing(_conectar()) as con:
            linhas = con.execute(
                f"SELECT {', '.join(colunas)} FROM sessoes_fechamento ORDER BY atualizado_em DESC"
            ).fetchall()
        return pd.DataFrame(linhas, columns=colunas)
    except Exception:
        _logger.exception("Falha ao listar sessões")
        return pd.DataFrame(columns=colunas)
