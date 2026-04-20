"""Watchlist + settings loading."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


@dataclass
class WatchlistEntry:
    symbol: str
    exchange: str = "SMART"
    currency: str = "USD"
    max_strike: float = float("inf")
    max_contracts: int = 1
    notes: str = ""


@dataclass
class IBConfig:
    host: str = "127.0.0.1"
    port: int = 7497
    client_id: int = 17
    market_data_type: int = 3
    request_timeout_s: int = 15


@dataclass
class OptionsConfig:
    dte_min: int = 7
    dte_max: int = 60
    expiry_filter: str = "all"
    max_spread_pct: float = 0.10
    require_positive_bid: bool = True
    min_annualized_yield: float = 0.05
    top_n_per_ticker: int = 25


@dataclass
class TBillConfig:
    enabled: bool = True
    buckets_days: list[int] = field(default_factory=lambda: [28, 91, 182, 364])
    fallback_yield: float = 0.045


@dataclass
class ReportConfig:
    output_dir: str = "output"
    filename_prefix: str = "csp_scan"
    open_after_run: bool = False


@dataclass
class Settings:
    ib: IBConfig = field(default_factory=IBConfig)
    options: OptionsConfig = field(default_factory=OptionsConfig)
    tbill: TBillConfig = field(default_factory=TBillConfig)
    report: ReportConfig = field(default_factory=ReportConfig)


# ---------------------------------------------------------------------------
# Loaders
# ---------------------------------------------------------------------------


def load_watchlist(path: str | Path) -> list[WatchlistEntry]:
    data = _read_yaml(path)
    entries = data.get("watchlist", [])
    out: list[WatchlistEntry] = []
    for row in entries:
        out.append(
            WatchlistEntry(
                symbol=str(row["symbol"]).upper(),
                exchange=row.get("exchange", "SMART"),
                currency=row.get("currency", "USD"),
                max_strike=float(row.get("max_strike", float("inf"))),
                max_contracts=int(row.get("max_contracts", 1)),
                notes=str(row.get("notes", "")),
            )
        )
    if not out:
        raise ValueError(f"Watchlist at {path} is empty.")
    return out


def load_settings(path: str | Path) -> Settings:
    data = _read_yaml(path)
    ib = data.get("ib", {}) or {}
    opts = data.get("options", {}) or {}
    tbill = data.get("tbill", {}) or {}
    rep = data.get("report", {}) or {}

    return Settings(
        ib=IBConfig(
            host=ib.get("host", "127.0.0.1"),
            port=int(ib.get("port", 7497)),
            client_id=int(ib.get("client_id", 17)),
            market_data_type=int(ib.get("market_data_type", 3)),
            request_timeout_s=int(ib.get("request_timeout_s", 15)),
        ),
        options=OptionsConfig(
            dte_min=int(opts.get("dte_min", 7)),
            dte_max=int(opts.get("dte_max", 60)),
            expiry_filter=str(opts.get("expiry_filter", "all")),
            max_spread_pct=float(opts.get("max_spread_pct", 0.10)),
            require_positive_bid=bool(opts.get("require_positive_bid", True)),
            min_annualized_yield=float(opts.get("min_annualized_yield", 0.05)),
            top_n_per_ticker=int(opts.get("top_n_per_ticker", 25)),
        ),
        tbill=TBillConfig(
            enabled=bool(tbill.get("enabled", True)),
            buckets_days=list(tbill.get("buckets_days", [28, 91, 182, 364])),
            fallback_yield=float(tbill.get("fallback_yield", 0.045)),
        ),
        report=ReportConfig(
            output_dir=str(rep.get("output_dir", "output")),
            filename_prefix=str(rep.get("filename_prefix", "csp_scan")),
            open_after_run=bool(rep.get("open_after_run", False)),
        ),
    )


def _read_yaml(path: str | Path) -> dict[str, Any]:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Config file not found: {p}")
    with p.open("r", encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}
