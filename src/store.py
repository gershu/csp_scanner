"""
DuckDB persistence layer for CSP scan results.

Database file: data/csp_history.duckdb  (configurable via settings.yaml)

Tables
------
schema_version   — tracks DDL migrations
scan_runs        — one row per scanner invocation
scan_candidates  — one row per surviving CSP candidate, enriched with
                   T-Bill data and composite Risk/Reward score

Typical usage
-------------
    from src.store import ScanStore

    with ScanStore("data/csp_history.duckdb") as store:
        run_id = store.save_run(run_ts, candidates_by_ticker, tbill_matches, settings)

    # --- analysis (notebook / script) ---
    with ScanStore("data/csp_history.duckdb") as store:
        df = store.candidates(symbol="PLTR", min_yield=0.09).df()
        df = store.query(\"\"\"
            SELECT symbol, run_ts, ann_yield, score
            FROM scan_candidates
            WHERE run_ts >= now() - INTERVAL 30 DAYS
            ORDER BY score DESC
        \"\"\").df()
"""

from __future__ import annotations

import json
import logging
import math
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    import duckdb as _duckdb
    from .option_selector import CSPCandidate
    from .tbill import TBillMatch
    from .watchlist import Settings

log = logging.getLogger(__name__)

_SCHEMA_VERSION = 1

_DDL = """\
CREATE TABLE IF NOT EXISTS schema_version (
    version     INTEGER   NOT NULL,
    applied_at  TIMESTAMP NOT NULL
);

CREATE TABLE IF NOT EXISTS scan_runs (
    run_id          VARCHAR   PRIMARY KEY,
    run_ts          TIMESTAMP NOT NULL,
    n_tickers       INTEGER,
    n_candidates    INTEGER,
    settings_json   VARCHAR
);

CREATE TABLE IF NOT EXISTS scan_candidates (
    run_id          VARCHAR   NOT NULL,
    run_ts          TIMESTAMP NOT NULL,
    -- contract identity
    symbol          VARCHAR   NOT NULL,
    expiry          VARCHAR   NOT NULL,   -- YYYYMMDD
    expiry_date     DATE,                 -- for easy date arithmetic
    dte             INTEGER,
    strike          DOUBLE,
    -- market snapshot
    spot            DOUBLE,
    moneyness       DOUBLE,               -- (strike/spot - 1), negative = OTM
    bid             DOUBLE,
    ask             DOUBLE,
    mid             DOUBLE,
    spread_pct      DOUBLE,
    iv              DOUBLE,
    delta           DOUBLE,
    open_interest   DOUBLE,
    volume          DOUBLE,
    -- derived economics
    cash_required   DOUBLE,
    premium         DOUBLE,
    ann_yield       DOUBLE,
    total_yield     DOUBLE,               -- ann_yield + tbill_yield
    breakeven       DOUBLE,
    -- t-bill pairing
    tbill_bucket    INTEGER,
    tbill_yield     DOUBLE,
    tbill_interest  DOUBLE,
    -- risk/reward score
    score           INTEGER,
    rating          VARCHAR               -- 'Attraktiv' | 'Mittel' | 'Vorsicht'
);
"""

# Convenience view created on first connect (non-persistent, recreated each time)
_VIEWS = """\
CREATE OR REPLACE VIEW latest_run AS
    SELECT c.*
    FROM scan_candidates c
    JOIN (SELECT run_id FROM scan_runs ORDER BY run_ts DESC LIMIT 1) r
      ON c.run_id = r.run_id;

CREATE OR REPLACE VIEW run_summary AS
    SELECT
        r.run_id,
        r.run_ts,
        r.n_tickers,
        r.n_candidates,
        ROUND(AVG(c.ann_yield) * 100, 2)  AS avg_yield_pct,
        ROUND(AVG(c.score), 0)            AS avg_score,
        COUNT(CASE WHEN c.rating = 'Attraktiv' THEN 1 END) AS n_attraktiv,
        COUNT(CASE WHEN c.rating = 'Mittel'    THEN 1 END) AS n_mittel,
        COUNT(CASE WHEN c.rating = 'Vorsicht'  THEN 1 END) AS n_vorsicht
    FROM scan_runs r
    LEFT JOIN scan_candidates c ON c.run_id = r.run_id
    GROUP BY r.run_id, r.run_ts, r.n_tickers, r.n_candidates
    ORDER BY r.run_ts DESC;
"""


# ---------------------------------------------------------------------------
# Public class
# ---------------------------------------------------------------------------


class ScanStore:
    """Persistent DuckDB store for CSP scan results."""

    def __init__(self, db_path: str | Path = "data/csp_history.duckdb") -> None:
        import duckdb

        self._path = Path(db_path)
        self._path.parent.mkdir(parents=True, exist_ok=True)
        self._con: _duckdb.DuckDBPyConnection = duckdb.connect(str(self._path))
        self._ensure_schema()

    # ---- context manager ----------------------------------------------------

    def __enter__(self) -> "ScanStore":
        return self

    def __exit__(self, *_) -> None:
        self.close()

    def close(self) -> None:
        try:
            self._con.close()
        except Exception:
            pass

    # ---- schema management --------------------------------------------------

    def _ensure_schema(self) -> None:
        self._con.execute(_DDL)
        self._con.execute(_VIEWS)
        row = self._con.execute(
            "SELECT version FROM schema_version ORDER BY applied_at DESC LIMIT 1"
        ).fetchone()
        if row is None:
            self._con.execute(
                "INSERT INTO schema_version VALUES (?, ?)",
                [_SCHEMA_VERSION, datetime.utcnow()],
            )
            log.info("ScanStore: schema v%d initialised at %s", _SCHEMA_VERSION, self._path)
        elif row[0] < _SCHEMA_VERSION:
            self._migrate(current=row[0])

    def _migrate(self, current: int) -> None:
        """Apply incremental DDL migrations from `current` to _SCHEMA_VERSION."""
        # Placeholder — extend with ALTER TABLE statements as the schema evolves.
        log.warning(
            "ScanStore: no migration path from v%d → v%d defined.",
            current, _SCHEMA_VERSION,
        )

    # ---- write --------------------------------------------------------------

    def save_run(
        self,
        run_ts: datetime,
        candidates_by_ticker: dict[str, list["CSPCandidate"]],
        tbill_matches: dict[str, dict[str, "TBillMatch"]],
        settings: "Settings",
    ) -> str:
        """
        Persist one scanner run. Returns the run_id string ('YYYYMMDD_HHMMSS').

        Idempotent: a run_id that already exists is silently skipped so
        that re-running a notebook cell never produces duplicate rows.
        """
        run_id = run_ts.strftime("%Y%m%d_%H%M%S")

        exists = self._con.execute(
            "SELECT 1 FROM scan_runs WHERE run_id = ?", [run_id]
        ).fetchone()
        if exists:
            log.warning("ScanStore: run_id=%s already stored — skipping.", run_id)
            return run_id

        # Flatten + score
        aggregated = [
            (ticker, c, tbill_matches.get(ticker, {}).get(c.expiry))
            for ticker, candidates in candidates_by_ticker.items()
            for c in candidates
        ]
        scored = _compute_scores(aggregated)

        settings_json = json.dumps({
            "ib":      asdict(settings.ib),
            "options": asdict(settings.options),
            "tbill":   asdict(settings.tbill),
        })

        self._con.execute(
            """
            INSERT INTO scan_runs (run_id, run_ts, n_tickers, n_candidates, settings_json)
            VALUES (?, ?, ?, ?, ?)
            """,
            [run_id, run_ts, len(candidates_by_ticker), len(scored), settings_json],
        )

        if scored:
            rows = []
            for row in scored:
                try:
                    expiry_date = datetime.strptime(row["expiry"], "%Y%m%d").date()
                except ValueError:
                    expiry_date = None

                rows.append([
                    run_id, run_ts,
                    row["symbol"], row["expiry"], expiry_date, row["dte"],
                    row["strike"], row["spot"], row["moneyness"],
                    row["bid"], row["ask"], row["mid"], row["spread"],
                    row["iv"], row["delta"], row["open_int"], row["volume"],
                    row["cash_req"], row["premium"],
                    row["ann_yield"], row["total_yield"], row["breakeven"],
                    row["tbill_bucket"], row["tbill_yield_pct"], row["tbill_interest"],
                    row["score"], row["rating"],
                ])

            self._con.executemany(
                """
                INSERT INTO scan_candidates VALUES (
                    ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
                )
                """,
                rows,
            )

        log.info("ScanStore: run_id=%s saved — %d candidates.", run_id, len(scored))
        return run_id

    # ---- read / query -------------------------------------------------------

    def runs(self):
        """All scan runs as a DataFrame, newest first."""
        return self._con.execute("SELECT * FROM run_summary").df()

    def candidates(
        self,
        run_id: str | None = None,
        symbol: str | None = None,
        min_yield: float | None = None,
        min_score: int | None = None,
        since: datetime | None = None,
        rating: str | None = None,
    ):
        """
        Filtered query returning a DuckDB relation (call .df() for pandas).

        Examples
        --------
        store.candidates(symbol='PLTR', min_yield=0.09).df()
        store.candidates(rating='Attraktiv', since=datetime(2026,4,1)).df()
        """
        filters: list[str] = []
        if run_id:
            filters.append(f"run_id = '{run_id}'")
        if symbol:
            filters.append(f"symbol = '{symbol.upper()}'")
        if min_yield is not None:
            filters.append(f"ann_yield >= {min_yield}")
        if min_score is not None:
            filters.append(f"score >= {min_score}")
        if since is not None:
            filters.append(f"run_ts >= '{since.isoformat()}'")
        if rating:
            filters.append(f"rating = '{rating}'")

        where = ("WHERE " + " AND ".join(filters)) if filters else ""
        sql = f"SELECT * FROM scan_candidates {where} ORDER BY score DESC, ann_yield DESC"
        return self._con.query(sql)

    def query(self, sql: str):
        """Execute any SQL and return a DuckDB relation (call .df() for pandas)."""
        return self._con.query(sql)

    @property
    def path(self) -> Path:
        return self._path


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _nan_safe(v) -> float | None:
    """Return float or None — never NaN (DuckDB stores NaN as NULL)."""
    if v is None:
        return None
    try:
        f = float(v)
        return None if math.isnan(f) else f
    except (TypeError, ValueError):
        return None


def _percentrank(values: list[float], x: float) -> float:
    valid = [v for v in values if v is not None]
    if len(valid) <= 1:
        return 0.5
    return sum(1 for v in valid if v < x) / (len(valid) - 1)


def _compute_scores(
    aggregated: list[tuple[str, "CSPCandidate", "TBillMatch | None"]],
) -> list[dict]:
    """
    Flatten candidates + compute composite Risk/Reward score (0–100).

    Score = 50 × PERCENTRANK(ann_yield)
          + 30 × PERCENTRANK(downside_cushion)
          + 20 × (1 − PERCENTRANK(spread_pct))

    Rating: Attraktiv ≥ 70 | Mittel ≥ 45 | Vorsicht < 45
    """
    rows: list[dict] = []
    for _ticker, c, tbill in aggregated:
        spot = _nan_safe(c.underlying_price)
        cushion = (
            (spot - c.strike) / spot
            if spot and spot > 0 else None
        )
        ann = _nan_safe(c.annualized_yield)
        total = (ann or 0.0) + (tbill.yield_pct if tbill else 0.0)

        rows.append({
            "symbol":         c.symbol,
            "expiry":         c.expiry,
            "dte":            c.dte,
            "strike":         _nan_safe(c.strike),
            "spot":           spot,
            "moneyness":      _nan_safe(c.moneyness),
            "bid":            _nan_safe(c.bid),
            "ask":            _nan_safe(c.ask),
            "mid":            _nan_safe(c.mid),
            "spread":         _nan_safe(c.spread_pct),
            "iv":             _nan_safe(c.iv),
            "delta":          _nan_safe(c.delta),
            "open_int":       _nan_safe(c.open_interest),
            "volume":         _nan_safe(c.volume),
            "cash_req":       _nan_safe(c.cash_required),
            "premium":        _nan_safe(c.premium),
            "ann_yield":      ann,
            "total_yield":    _nan_safe(total),
            "breakeven":      _nan_safe(c.breakeven),
            "cushion":        cushion,
            "tbill_bucket":   tbill.bucket_days  if tbill else None,
            "tbill_yield_pct":tbill.yield_pct    if tbill else None,
            "tbill_interest": _nan_safe(tbill.interest_on(c.cash_required)) if tbill else None,
        })

    yields   = [r["ann_yield"] for r in rows if r["ann_yield"] is not None]
    cushions = [r["cushion"]   for r in rows if r["cushion"]   is not None]
    spreads  = [r["spread"]    for r in rows if r["spread"]    is not None]

    for row in rows:
        pr_y = _percentrank(yields,   row["ann_yield"]) if row["ann_yield"] is not None else 0.5
        pr_c = _percentrank(cushions, row["cushion"])   if row["cushion"]   is not None else 0.5
        pr_s = _percentrank(spreads,  row["spread"])    if row["spread"]    is not None else 0.5
        score = round(50 * pr_y + 30 * pr_c + 20 * (1 - pr_s))
        row["score"] = score
        row["rating"] = (
            "Attraktiv" if score >= 70 else
            "Mittel"    if score >= 45 else
            "Vorsicht"
        )
    return rows
