"""
IB connection wrapper (ib_async).

Keeps connection concerns isolated from business logic. All option-chain,
spot-price and bill-quote fetches go through this module.
"""

from __future__ import annotations

import logging
import math
from datetime import datetime
from typing import TYPE_CHECKING, Iterable

from .types import OptionQuote

if TYPE_CHECKING:
    from ib_async import IB, Stock, Option, Bond, Ticker  # noqa: F401

log = logging.getLogger(__name__)

# OptionQuote is re-exported for backwards compatibility
__all__ = ["IBClient", "OptionQuote", "dte_from_expiry", "install_ib_error_filter"]


# ---------------------------------------------------------------------------
# Log noise filter for ib_async
# ---------------------------------------------------------------------------


class _IBErrorNoiseFilter(logging.Filter):
    """
    Drop ib_async log records that report expected, non-fatal IB errors
    encountered during bulk option-chain qualification.

    Rationale
    ---------
    `reqSecDefOptParams` returns the union of strikes across all expiries
    for a given trading class. When we subsequently qualify specific
    (expiry, strike, right) triples, IB responds with Error 200
    ("No security definition has been found") for each combination that
    does not actually exist. That is expected — not every strike trades
    on every expiry — but `ib_async` logs one WARNING per failure, which
    quickly drowns legitimate output.

    This filter silences that specific class of messages; all other
    ib_async log records pass through untouched.
    """

    # IB error codes that are harmless during option-chain probing.
    _BENIGN_CODES = ("Error 200", "errorCode=200")
    _BENIGN_PHRASES = (
        "No security definition has been found",
    )

    def filter(self, record: logging.LogRecord) -> bool:  # True = keep
        try:
            msg = record.getMessage()
        except Exception:  # noqa: BLE001
            return True
        if any(code in msg for code in self._BENIGN_CODES):
            return False
        if any(phrase in msg for phrase in self._BENIGN_PHRASES):
            return False
        return True


_IB_ERROR_FILTER_INSTALLED = False


def install_ib_error_filter() -> None:
    """
    Attach the noise filter to the ib_async loggers. Idempotent.

    Call this once (IBClient.connect does it automatically). Safe to
    invoke from notebooks too if you want cleaner cell output before
    instantiating a client.
    """
    global _IB_ERROR_FILTER_INSTALLED
    if _IB_ERROR_FILTER_INSTALLED:
        return

    flt = _IBErrorNoiseFilter()
    # Cover ib_async's known logger names; the filter is cheap, so
    # attaching to several is harmless.
    for name in ("ib_async", "ib_async.wrapper", "ib_insync", "ib_insync.wrapper"):
        logging.getLogger(name).addFilter(flt)

    _IB_ERROR_FILTER_INSTALLED = True


# ---------------------------------------------------------------------------
# Client
# ---------------------------------------------------------------------------


class IBClient:
    """Thin async wrapper around ib_async.IB with project-specific helpers."""

    def __init__(
        self,
        host: str = "127.0.0.1",
        port: int = 7497,
        client_id: int = 17,
        market_data_type: int = 3,
        request_timeout_s: int = 15,
    ) -> None:
        from ib_async import IB  # local import: keeps module importable without ib_async installed

        self.host = host
        self.port = port
        self.client_id = client_id
        self.market_data_type = market_data_type
        self.timeout = request_timeout_s
        self.ib = IB()

    # ---- lifecycle --------------------------------------------------------

    def connect(self) -> None:
        # Silence expected 'Error 200 No security definition' noise from
        # bulk option qualification before we start talking to IB.
        install_ib_error_filter()

        log.info("Connecting to IB at %s:%s (clientId=%s)", self.host, self.port, self.client_id)
        self.ib.connect(self.host, self.port, clientId=self.client_id, timeout=self.timeout)
        # 1=live, 2=frozen, 3=delayed, 4=delayed-frozen
        self.ib.reqMarketDataType(self.market_data_type)

    def disconnect(self) -> None:
        if self.ib.isConnected():
            self.ib.disconnect()

    def __enter__(self) -> "IBClient":
        self.connect()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.disconnect()

    # ---- stock ------------------------------------------------------------

    def qualify_stock(
        self,
        symbol: str,
        exchange: str = "SMART",
        currency: str = "USD",
        opt_exchange: str | None = None,
        trading_class: str | None = None,
    ):
        """Qualify a stock contract for trading.
        
        Args:
            symbol: Stock ticker symbol
            exchange: Primary exchange for the stock (default: SMART)
            currency: Currency code (default: USD)
            opt_exchange: Optional options exchange (ignored, kept for API compatibility)
            trading_class: Optional trading class (ignored, kept for API compatibility)
        """
        from ib_async import Stock

        stk = Stock(symbol, exchange, currency)
        self.ib.qualifyContracts(stk)
        return stk

    def spot_price(self, stock) -> float:
        """Return the best available price for a qualified stock contract.

        Strategy (three escalating attempts):

        1. reqMktData snapshot — fast path for US stocks with delayed data
           (market_data_type 3).  Non-US exchanges (e.g. XETRA) are NOT
           covered by IB's free delayed feed, so this will return NaN for
           European stocks without a subscription.

        2. reqHistoricalData (1-day bar, TRADES) — works for virtually every
           exchange and requires no market-data subscription.  Returns the
           close of the most recent completed session.

        3. reqMktData streaming fallback — last resort with a generous wait.
        """
        # --- attempt 1: market-data snapshot --------------------------------
        ticker = self.ib.reqMktData(stock, "", snapshot=True, regulatorySnapshot=False)
        self.ib.sleep(2.0)
        price = _first_valid(ticker.marketPrice(), ticker.last, ticker.close)
        if price is not None:
            log.debug("spot_price [snapshot] %s = %.4f", stock.symbol, price)
            return float(price)

        # --- attempt 2: historical daily bar (no subscription required) -----
        log.debug("spot_price snapshot NaN for %s — trying historical bar", stock.symbol)
        try:
            bars = self.ib.reqHistoricalData(
                stock,
                endDateTime="",          # most recent available
                durationStr="5 D",       # look back up to 5 days to skip holidays
                barSizeSetting="1 day",
                whatToShow="TRADES",
                useRTH=True,
                formatDate=1,
                keepUpToDate=False,
            )
            if bars:
                price = _first_valid(bars[-1].close, bars[-1].open)
                if price is not None:
                    log.debug("spot_price [historical] %s = %.4f", stock.symbol, price)
                    return float(price)
        except Exception as exc:
            log.debug("spot_price historical bar failed for %s: %s", stock.symbol, exc)

        # --- attempt 3: streaming fallback ----------------------------------
        log.debug("spot_price historical NaN for %s — falling back to stream", stock.symbol)
        ticker = self.ib.reqMktData(stock, "", snapshot=False, regulatorySnapshot=False)
        self.ib.sleep(4.0)
        price = _first_valid(ticker.marketPrice(), ticker.last, ticker.close)
        self.ib.cancelMktData(stock)
        return float(price) if price is not None else float("nan")

    # ---- option chain -----------------------------------------------------

    def option_chain_params(self, stock) -> dict:
        """Fetch expiries & strikes for a stock via secDefOptParams.

        IB requires the stock to be qualified via SMART routing for this call
        to return the full expiry/strike universe.  Qualifying with a primary
        exchange (e.g. ARCA for SPY, NASDAQ for AAPL) can result in an empty
        or truncated response.
        """
        params = self.ib.reqSecDefOptParams(
            stock.symbol,
            "",              # futFopExchange (empty for equity options)
            stock.secType,
            stock.conId,
        )

        if not params:
            log.warning(
                "%s (conId=%s, exchange=%s): reqSecDefOptParams returned nothing. "
                "Ensure the watchlist entry uses exchange: SMART.",
                stock.symbol, stock.conId, stock.exchange,
            )
            return {"expiries": [], "strikes": [], "exchange": None, "trading_class": None}

        log.debug(
            "%s: secDefOptParams returned %d param set(s): %s",
            stock.symbol,
            len(params),
            [(p.exchange, p.tradingClass, len(p.expirations)) for p in params],
        )

        # Prefer SMART + tradingClass matching the symbol (e.g. 'SPY' over '2SPY'/'SPYW'),
        # then any SMART entry, then first available.
        smart_params = [p for p in params if p.exchange == "SMART"]
        chain = (
            next((p for p in smart_params if p.tradingClass == stock.symbol), None)
            or next(iter(smart_params), None)
            or params[0]
        )
        return {
            "expiries": sorted(chain.expirations),
            "strikes": sorted(chain.strikes),
            "exchange": chain.exchange,
            "trading_class": chain.tradingClass,
        }

    def fetch_put_quotes(
        self,
        symbol: str,
        expiry: str,
        strikes: Iterable[float],
        exchange: str = "SMART",
        trading_class: str | None = None,
        currency: str = "USD",
    ) -> list[OptionQuote]:
        """Qualify a batch of put contracts and return their market snapshot."""
        from ib_async import Option

        contracts: list[Option] = []
        for k in strikes:
            opt = Option(
                symbol=symbol,
                lastTradeDateOrContractMonth=expiry,
                strike=float(k),
                right="P",
                exchange=exchange,
                currency=currency,
                tradingClass=trading_class or symbol,
                multiplier="100",
            )
            contracts.append(opt)

        if not contracts:
            return []

        # qualifyContracts mutates in place and drops unknown ones
        self.ib.qualifyContracts(*contracts)
        valid = [c for c in contracts if c.conId]

        if not valid:
            return []

        # genericTickList 106 = option implied vol & greeks via model
        tickers = [self.ib.reqMktData(c, "106", snapshot=False, regulatorySnapshot=False) for c in valid]
        # Wait for quotes & modelGreeks to populate
        self.ib.sleep(2.5)

        out: list[OptionQuote] = []
        for ticker in tickers:
            q = _to_option_quote(ticker)
            if q is not None:
                out.append(q)
            self.ib.cancelMktData(ticker.contract)
        return out

    # ---- T-Bill -----------------------------------------------------------

    def tbill_yield(self, maturity_days: int) -> float | None:
        """
        Best-effort retrieval of an indicative US T-Bill yield for a given
        maturity bucket. IB returns bond quotes via reqTickers on a Bond
        contract — availability depends on the user's data subscription.
        Returns decimal yield (e.g. 0.0475) or None if not available.
        """
        # Generic BILL contract. IB requires a CUSIP to pin down a specific
        # issue; without that the request is indicative only. In practice
        # users either (a) supply CUSIPs in settings, or (b) rely on the
        # fallback_yield from the config file.
        try:
            from ib_async import Bond

            bond = Bond(symbol="T-BILL", exchange="SMART", currency="USD")
            self.ib.qualifyContracts(bond)
            if not bond.conId:
                return None
            ticker = self.ib.reqMktData(bond, "", snapshot=True, regulatorySnapshot=False)
            self.ib.sleep(1.0)
            yld = ticker.last if ticker.last and ticker.last > 0 else ticker.close
            self.ib.cancelMktData(bond)
            return float(yld) / 100.0 if yld else None
        except Exception as e:  # noqa: BLE001
            log.debug("T-Bill yield lookup failed: %s", e)
            return None


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _first_valid(*values) -> float | None:
    for v in values:
        if v is None:
            continue
        try:
            fv = float(v)
        except (TypeError, ValueError):
            continue
        if not math.isnan(fv) and fv > 0:
            return fv
    return None


def _to_option_quote(ticker) -> OptionQuote | None:
    c = ticker.contract
    if c is None:
        return None

    bid = _num(ticker.bid)
    ask = _num(ticker.ask)
    last = _num(ticker.last)

    mid = float("nan")
    if bid > 0 and ask > 0 and ask >= bid:
        mid = (bid + ask) / 2.0
    elif last > 0:
        mid = last

    greeks = ticker.modelGreeks or ticker.lastGreeks or ticker.bidGreeks or ticker.askGreeks

    return OptionQuote(
        symbol=c.symbol,
        expiry=c.lastTradeDateOrContractMonth,
        strike=float(c.strike),
        right=c.right,
        bid=bid,
        ask=ask,
        last=last,
        mid=mid,
        volume=_num(ticker.volume),
        open_interest=_num(getattr(ticker, "putOpenInterest", None)) or _num(getattr(ticker, "openInterest", None)),
        iv=_num(getattr(greeks, "impliedVol", None) if greeks else None),
        delta=_num(getattr(greeks, "delta", None) if greeks else None),
        gamma=_num(getattr(greeks, "gamma", None) if greeks else None),
        theta=_num(getattr(greeks, "theta", None) if greeks else None),
        vega=_num(getattr(greeks, "vega", None) if greeks else None),
        underlying_price=_num(getattr(greeks, "undPrice", None) if greeks else None),
        multiplier=int(c.multiplier) if c.multiplier else 100,
    )


def _num(x) -> float:
    if x is None:
        return float("nan")
    try:
        fv = float(x)
    except (TypeError, ValueError):
        return float("nan")
    return fv if not math.isnan(fv) else float("nan")


def dte_from_expiry(expiry: str, ref: datetime | None = None) -> int:
    """Days-to-expiration from IB expiry string YYYYMMDD."""
    ref = ref or datetime.utcnow()
    exp = datetime.strptime(expiry, "%Y%m%d")
    return max((exp - ref).days, 0)
