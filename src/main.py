"""
Command-line entry point for the CSP scanner.

Usage (from project root):

    python -m src.main \
        --watchlist config/watchlist.yaml \
        --settings config/settings.yaml

Run with --help for all options.
"""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path

from .ib_client import IBClient
from .option_selector import CSPCandidate, CSPSelector
from .report import write_report
from .tbill import TBillMatch, TBillMatcher
from .watchlist import Settings, WatchlistEntry, load_settings, load_watchlist


def _build_argparser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="csp-scanner",
        description="Cash-Secured Short Put scanner for US stocks via IB.",
    )
    p.add_argument(
        "--watchlist",
        type=Path,
        default=Path("config/watchlist.yaml"),
        help="Path to watchlist YAML.",
    )
    p.add_argument(
        "--settings",
        type=Path,
        default=Path("config/settings.yaml"),
        help="Path to settings YAML.",
    )
    p.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Explicit output xlsx path. Overrides settings.report.output_dir/prefix.",
    )
    p.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
    )
    return p


def main(argv: list[str] | None = None) -> int:
    args = _build_argparser().parse_args(argv)

    logging.basicConfig(
        level=args.log_level,
        format="%(asctime)s %(levelname)-7s %(name)s | %(message)s",
        datefmt="%H:%M:%S",
    )
    log = logging.getLogger("csp_scanner")

    # ---- Load config -------------------------------------------------------
    settings: Settings = load_settings(args.settings)
    watchlist: list[WatchlistEntry] = load_watchlist(args.watchlist)
    log.info("Loaded %d watchlist entries.", len(watchlist))

    # ---- Run ---------------------------------------------------------------
    run_ts = datetime.utcnow()
    candidates_by_ticker: dict[str, list[CSPCandidate]] = {}
    tbill_matches: dict[str, dict[str, TBillMatch]] = {}

    with IBClient(
        host=settings.ib.host,
        port=settings.ib.port,
        client_id=settings.ib.client_id,
        market_data_type=settings.ib.market_data_type,
        request_timeout_s=settings.ib.request_timeout_s,
    ) as client:
        selector = CSPSelector(client, settings.options)
        matcher = TBillMatcher(client, settings.tbill)

        for entry in watchlist:
            try:
                candidates = selector.scan_ticker(entry, today=run_ts)
            except Exception as e:  # noqa: BLE001
                log.exception("Scan failed for %s: %s", entry.symbol, e)
                continue

            candidates_by_ticker[entry.symbol] = candidates

            per_expiry: dict[str, TBillMatch] = {}
            for c in candidates:
                if c.expiry not in per_expiry:
                    per_expiry[c.expiry] = matcher.match(c.dte, cash_usd=c.cash_required)
            tbill_matches[entry.symbol] = per_expiry

    # ---- Report ------------------------------------------------------------
    output_path = args.output or _build_output_path(settings, run_ts)
    report_path = write_report(
        output_path=output_path,
        candidates_by_ticker=candidates_by_ticker,
        tbill_matches=tbill_matches,
        matcher=matcher,
        watchlist=watchlist,
        settings=settings,
        run_ts=run_ts,
    )
    log.info("Report written to %s", report_path)

    # ---- Console summary ---------------------------------------------------
    total = sum(len(v) for v in candidates_by_ticker.values())
    log.info("Total candidates across watchlist: %d", total)
    top = []
    for ticker, cs in candidates_by_ticker.items():
        for c in cs:
            top.append((ticker, c))
    top.sort(key=lambda x: x[1].annualized_yield, reverse=True)
    for ticker, c in top[:10]:
        log.info(
            "  %-5s %s  K=%.2f  mid=%.2f  yield=%.2f%%  DTE=%d",
            ticker,
            c.expiry,
            c.strike,
            c.mid,
            c.annualized_yield * 100,
            c.dte,
        )

    return 0


def _build_output_path(settings: Settings, ts: datetime) -> Path:
    out_dir = Path(settings.report.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = ts.strftime("%Y%m%d_%H%M%S")
    return out_dir / f"{settings.report.filename_prefix}_{stamp}.xlsx"


if __name__ == "__main__":
    sys.exit(main())
