"""
Pure data containers shared across modules.

Kept dependency-free (no ib_async imports) so they can be used in tests
and downstream consumers without a live IB connection.
"""

from __future__ import annotations

import math
from dataclasses import dataclass


@dataclass
class OptionQuote:
    """Flattened snapshot of a single option contract."""

    symbol: str
    expiry: str           # YYYYMMDD (IB lastTradeDateOrContractMonth)
    strike: float
    right: str            # "P" or "C"
    bid: float
    ask: float
    last: float
    mid: float
    volume: float
    open_interest: float
    iv: float             # implied vol from IB modelGreeks
    delta: float
    gamma: float
    theta: float
    vega: float
    underlying_price: float
    multiplier: int = 100

    @property
    def spread_pct(self) -> float:
        if self.mid <= 0 or math.isnan(self.mid):
            return math.inf
        return (self.ask - self.bid) / self.mid
