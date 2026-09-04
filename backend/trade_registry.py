# backend/trade_registry.py
"""
Loads the CSI-anchored trade list from the kdx_reference Postgres database
(source of truth: punchai/phase0/seed/trades.json, seeded via seed_trades.py
on the host). Connection string comes from KDX_REFERENCE_URL.
"""

import os

import psycopg2


class TradeRegistryError(Exception):
    """Raised when the trade reference database is unreachable or unreadable."""


def _dsn() -> str:
    dsn = os.getenv("KDX_REFERENCE_URL")
    if not dsn:
        raise TradeRegistryError("KDX_REFERENCE_URL is not set (see .env.example).")
    return dsn


def load_trades() -> list[dict]:
    """Returns [{csi_code, name, default_prefix}], sorted by csi_code."""
    try:
        with psycopg2.connect(_dsn(), connect_timeout=5) as conn:
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT csi_code, name, default_prefix FROM trade ORDER BY csi_code"
                )
                rows = cur.fetchall()
    except psycopg2.Error as exc:
        raise TradeRegistryError(
            f"could not read trades from kdx_reference: {exc}"
        ) from exc
    if not rows:
        raise TradeRegistryError("kdx_reference.trade is empty.")
    return [{"csi_code": r[0], "name": r[1], "default_prefix": r[2]} for r in rows]
