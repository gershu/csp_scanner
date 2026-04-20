"""
T-Bill matching for each CSP candidate.

For every short put we need cash = strike * 100 USD per contract, parked
until expiry. The scanner pairs each expiry with the largest available
T-Bill bucket whose maturity is <= DTE of the option; any residual days
are held overnight (assumed at the same bucket's yield as a simplification).

Live yield lookup via IB is best-effort because US T-Bill quotes on IB
require specific CUSIPs and an appropriate data subscription. When the
live lookup fails, the scanner uses `tbill.fallback_yield` from settings.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass

from .ib_client import IBClient
from .watchlist import TBillConfig

log = logging.getLogger(__name__)


@dataclass
class TBillMatch:
    dte: int                    # option DTE
    bucket_days: int            # chosen T-Bill maturity bucket (days)
    yield_pct: float            # decimal annualized yield
    source: str                 # "live" | "fallback"
    residual_days: int          # dte - bucket_days, parked overnight
    projected_interest: float   # USD interest earned on `cash_usd` over `dte`

    def interest_on(self, cash_usd: float) -> float:
        return cash_usd * self.yield_pct * (self.dte / 365.0)


class TBillMatcher:
    def __init__(self, client: IBClient, cfg: TBillConfig) -> None:
        self.client = client
        self.cfg = cfg
        # Cache live yields per bucket to avoid repeat IB calls per candidate
        self._yield_cache: dict[int, tuple[float, str]] = {}

    def match(self, dte: int, cash_usd: float = 0.0) -> TBillMatch:
        bucket = self._pick_bucket(dte)
        yld, source = self._yield_for(bucket)
        residual = max(dte - bucket, 0)
        interest = cash_usd * yld * (dte / 365.0)
        return TBillMatch(
            dte=dte,
            bucket_days=bucket,
            yield_pct=yld,
            source=source,
            residual_days=residual,
            projected_interest=interest,
        )

    # ---- internals --------------------------------------------------------

    def _pick_bucket(self, dte: int) -> int:
        """Largest bucket <= dte, falling back to the smallest if dte is tiny."""
        buckets = sorted(self.cfg.buckets_days)
        eligible = [b for b in buckets if b <= dte]
        return eligible[-1] if eligible else buckets[0]

    def _yield_for(self, bucket_days: int) -> tuple[float, str]:
        if bucket_days in self._yield_cache:
            return self._yield_cache[bucket_days]

        yld: float | None = None
        if self.cfg.enabled:
            try:
                yld = self.client.tbill_yield(bucket_days)
            except Exception as e:  # noqa: BLE001
                log.debug("T-Bill yield lookup error (%d days): %s", bucket_days, e)

        if yld is None or yld <= 0:
            result = (self.cfg.fallback_yield, "fallback")
        else:
            result = (yld, "live")

        self._yield_cache[bucket_days] = result
        return result
