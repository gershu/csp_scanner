"""
Option selection logic.

Given a watchlist entry and an IB connection, build the candidate list of
short puts whose strike <= max_strike, that pass the liquidity filter,
and rank them by annualized premium yield.

Yield definition (matches user spec):

    annualized_yield = (mid_premium / strike) * 365 / DTE

This is yield on cash-secured capital (Strike * 100), which is the
capital that has to be held in T-Bills.
"""

from __future__ import annotations

import logging
import math
from dataclasses import dataclass
from datetime import datetime

from .ib_client import IBClient, dte_from_expiry
from .types import OptionQuote
from .watchlist import OptionsConfig, WatchlistEntry

log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Result row
# ---------------------------------------------------------------------------


@dataclass
class CSPCandidate:
    symbol: str
    expiry: str                 # YYYYMMDD
    dte: int
    strike: float
    underlying_price: float
    bid: float
    ask: float
    mid: float
    spread_pct: float
    iv: float
    delta: float
    open_interest: float
    volume: float

    # derived
    cash_required: float        # = strike * 100  (per contract, USD)
    premium: float              # = mid * 100     (per contract, USD)
    annualized_yield: float     # decimal, on cash_required
    breakeven: float            # = strike - mid
    moneyness: float            # = strike / underlying - 1   (negative = OTM)

    notes: str = ""

    @classmethod
    def from_quote(cls, q: OptionQuote, ref: datetime | None = None) -> "CSPCandidate | None":
        """Build a candidate row from a raw quote. Returns None if unusable."""
        ref = ref or datetime.utcnow()
        dte = dte_from_expiry(q.expiry, ref)
        if dte <= 0 or q.strike <= 0:
            return None
        if math.isnan(q.mid) or q.mid <= 0:
            return None
        cash_required = q.strike * 100.0
        premium = q.mid * 100.0
        ann_yield = (q.mid / q.strike) * (365.0 / dte)
        underlying = q.underlying_price if q.underlying_price > 0 and not math.isnan(q.underlying_price) else float("nan")
        moneyness = (q.strike / underlying - 1.0) if underlying and not math.isnan(underlying) else float("nan")
        return cls(
            symbol=q.symbol,
            expiry=q.expiry,
            dte=dte,
            strike=q.strike,
            underlying_price=underlying,
            bid=q.bid,
            ask=q.ask,
            mid=q.mid,
            spread_pct=q.spread_pct,
            iv=q.iv,
            delta=q.delta,
            open_interest=q.open_interest,
            volume=q.volume,
            cash_required=cash_required,
            premium=premium,
            annualized_yield=ann_yield,
            breakeven=q.strike - q.mid,
            moneyness=moneyness,
        )


# ---------------------------------------------------------------------------
# Selector
# ---------------------------------------------------------------------------


class CSPSelector:
    """Combines IB chain access + filter/ranking config."""

    def __init__(self, client: IBClient, options_cfg: OptionsConfig) -> None:
        self.client = client
        self.cfg = options_cfg

    def scan_ticker(self, entry: WatchlistEntry, today: datetime | None = None) -> list[CSPCandidate]:
        today = today or datetime.utcnow()
        log.info("Scanning %s (max_strike=%.2f)", entry.symbol, entry.max_strike)

        stock = self.client.qualify_stock(entry.symbol, entry.exchange, entry.currency)
        spot = self.client.spot_price(stock)
        if math.isnan(spot) or spot <= 0:
            log.warning("No spot price for %s, skipping.", entry.symbol)
            return []

        # ---- Sanity check: max_strike vs. spot ----------------------------
        # Short-dated expiries only list strikes within a narrow band around
        # spot. If the user's max_strike is very far below spot, most or all
        # strike requests will fail on IB with "No security definition" —
        # that is not a bug, just the market not listing contracts that
        # deep OTM. Warn early.
        otm_pct = (spot - entry.max_strike) / spot if spot > 0 else 0.0
        if otm_pct > 0.20:
            log.warning(
                "%s: max_strike=%.2f is %.1f%% below spot=%.2f — "
                "short-dated expiries likely have no listings that deep OTM; "
                "consider raising max_strike or extending DTE window.",
                entry.symbol, entry.max_strike, otm_pct * 100, spot,
            )

        chain = self.client.option_chain_params(stock)
        if not chain["expiries"] or not chain["strikes"]:
            log.warning("No option chain for %s.", entry.symbol)
            return []

        # ---- Filter expiries by DTE window --------------------------------
        valid_expiries = [
            e for e in chain["expiries"]
            if self.cfg.dte_min <= dte_from_expiry(e, today) <= self.cfg.dte_max
        ]

        if self.cfg.expiry_filter == "monthlies":
            valid_expiries = [e for e in valid_expiries if _is_monthly(e)]

        if not valid_expiries:
            log.warning(
                "No expiries for %s in DTE [%d,%d].",
                entry.symbol, self.cfg.dte_min, self.cfg.dte_max,
            )
            return []

        # ---- Pull option quotes per expiry --------------------------------
        # Strike pre-filter is now done PER EXPIRY using a DTE-dependent
        # band around spot. This reflects the reality that short weeklies
        # only list strikes close to spot, while quarterlies/LEAPS list a
        # much wider grid. Upper bound is always min(max_strike, spot) —
        # CSPs are OTM puts, so strikes at/above spot are skipped.
        all_candidates: list[CSPCandidate] = []
        for expiry in valid_expiries:
            dte = dte_from_expiry(expiry, today)
            band = _strike_band_for_dte(dte)
            lower_bound = max(spot * (1.0 - band), 0.0)
            upper_bound = min(entry.max_strike, spot)

            candidate_strikes = [
                k for k in chain["strikes"]
                if lower_bound <= k <= upper_bound
            ]
            if not candidate_strikes:
                log.debug(
                    "%s expiry %s (DTE=%d): no strikes in band [%.2f, %.2f] "
                    "(spot=%.2f, max_strike=%.2f, band=%.0f%%).",
                    entry.symbol, expiry, dte, lower_bound, upper_bound,
                    spot, entry.max_strike, band * 100,
                )
                continue

            quotes = self.client.fetch_put_quotes(
                symbol=entry.symbol,
                expiry=expiry,
                strikes=candidate_strikes,
                exchange="SMART",
                trading_class=chain["trading_class"],
                currency=entry.currency,
            )

            for q in quotes:
                # Apply liquidity filters
                if self.cfg.require_positive_bid and (q.bid is None or q.bid <= 0 or math.isnan(q.bid)):
                    continue
                if q.spread_pct > self.cfg.max_spread_pct:
                    continue

                # Stamp underlying if greeks didn't carry it
                if math.isnan(q.underlying_price) or q.underlying_price <= 0:
                    q.underlying_price = spot

                cand = CSPCandidate.from_quote(q, today)
                if cand is None:
                    continue

                # Yield filter
                if cand.annualized_yield < self.cfg.min_annualized_yield:
                    continue

                all_candidates.append(cand)

        # ---- Rank by annualized yield -------------------------------------
        all_candidates.sort(key=lambda c: c.annualized_yield, reverse=True)
        if self.cfg.top_n_per_ticker and len(all_candidates) > self.cfg.top_n_per_ticker:
            all_candidates = all_candidates[: self.cfg.top_n_per_ticker]

        log.info("  -> %d candidates after filters.", len(all_candidates))
        return all_candidates


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _is_monthly(expiry_yyyymmdd: str) -> bool:
    """Standard US monthly = 3rd Friday of the month."""
    try:
        d = datetime.strptime(expiry_yyyymmdd, "%Y%m%d").date()
    except ValueError:
        return False
    if d.weekday() != 4:  # Friday
        return False
    return 15 <= d.day <= 21


def _strike_band_for_dte(dte: int) -> float:
    """
    How far below spot we should scan for available put strikes, given DTE.

    IB lists far fewer strikes for short expiries. Trying to qualify
    deep-OTM strikes on a 1-week weekly produces dozens of
    'No security definition' errors. Scale the band to DTE.

    Returns a fraction, e.g. 0.20 == "consider strikes down to 20% below spot".
    """
    if dte <= 7:
        return 0.10
    if dte <= 21:
        return 0.18
    if dte <= 45:
        return 0.30
    if dte <= 120:
        return 0.45
    return 0.60
