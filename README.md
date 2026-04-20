# CSP Scanner — Cash-Secured Short Puts via Interactive Brokers

Phase-1-Tool für eine Investment-Strategie, die Cash-Secured Short Puts auf
US-Aktien schreibt und die Cash-Besicherung parallel in US Treasury Bills
anlegt. Der Scanner liest eine Watchlist ein, fragt über die IB API Option
Chains und Quotes ab, filtert nach Liquidität und Strike, und rankt die
verbleibenden Puts nach annualisierter Prämienrendite.

> **Scope Phase 1:** reiner Selektor & Report. Es werden **keine** Orders
> abgesetzt. Backtest ist explizit ausgeklammert.

---

## Inhaltsverzeichnis

1. [Installation](#installation)
2. [IB vorbereiten (TWS / Gateway)](#ib-vorbereiten)
3. [Konfiguration](#konfiguration)
4. [Ausführen](#ausfuehren)
5. [Output-Struktur](#output-struktur)
6. [Methodik](#methodik)
7. [Bekannte Limitationen](#bekannte-limitationen)
8. [Roadmap / Phase 2+](#roadmap)

---

## Installation

```bash
# Python 3.10+ empfohlen
python -m venv .venv
source .venv/bin/activate            # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## IB vorbereiten

1. **TWS** oder **IB Gateway** starten und einloggen (Paper oder Live).
2. In **Configure → API → Settings**:
   - *Enable ActiveX and Socket Clients* aktivieren
   - *Read-Only API* deaktivieren (nicht relevant für diesen Scanner, aber
     falls später Order-Routing ergänzt wird)
   - Port notieren (Default TWS Paper: `7497`, TWS Live: `7496`,
     Gateway Paper: `4002`, Gateway Live: `4001`)
   - *Trusted IPs* bei Bedarf auf `127.0.0.1` setzen
3. **Market Data Subscriptions:** Für delayed-data-Modus (Default in
   `settings.yaml` mit `market_data_type: 3`) sind keine zusätzlichen
   Subscriptions nötig — ausreichend für einen Proof-of-Concept. Für
   produktive Nutzung Live-Feeds für Equities + US Options abonnieren.

## Konfiguration

### `config/watchlist.yaml`

Pro Aktie: Ticker + gewünschtes Einstiegsniveau (`max_strike`). Der Scanner
betrachtet nur Puts mit **Strike ≤ max_strike**, d. h. Optionen, bei deren
Assignment du die Aktie zum gewünschten Preis (oder günstiger) erhältst.

```yaml
watchlist:
  - symbol: AAPL
    max_strike: 180.0     # nur Puts mit Strike <= 180 werden betrachtet
    max_contracts: 2
    notes: "Quality compounder; would add at <=180."
```

### `config/settings.yaml`

Zentrale Stellschrauben — siehe die Kommentare in der Datei. Die wichtigsten:

| Pfad                           | Default | Bedeutung                                                                 |
| ------------------------------ | ------- | ------------------------------------------------------------------------- |
| `ib.port`                      | 7497    | TWS/Gateway-Port                                                          |
| `ib.market_data_type`          | 3       | 1=live, 3=delayed, usw.                                                   |
| `options.dte_min/dte_max`      | 7 / 60  | Fenster für Expiries in Tagen                                             |
| `options.max_spread_pct`       | 0.10    | Max. Bid/Ask-Spread als Anteil der Mid                                    |
| `options.min_annualized_yield` | 0.05    | Mindest-Prämienrendite p.a. (Filter)                                      |
| `tbill.buckets_days`           | 28/91/182/364 | Maturity-Buckets für T-Bill-Zuordnung                               |
| `tbill.fallback_yield`         | 0.045   | Wird genutzt, wenn IB keine Bill-Quotes liefert                           |

## Ausführen

Aus dem Projekt-Root:

```bash
python -m src.main \
    --watchlist config/watchlist.yaml \
    --settings config/settings.yaml
```

Der Report landet standardmäßig unter `output/csp_scan_YYYYMMDD_HHMMSS.xlsx`.
Mit `--output /pfad/zur/datei.xlsx` kann ein expliziter Pfad gesetzt werden.

## Output-Struktur

| Sheet             | Inhalt                                                                       |
| ----------------- | ---------------------------------------------------------------------------- |
| `Top Candidates`  | Aggregiertes Ranking über die gesamte Watchlist, sortiert nach Annual. Yield |
| `<TICKER>`        | Pro Aktie alle Strikes, die die Filter überstanden haben                     |
| `T-Bill Matching` | Bucket-Yields + Zuordnung jede Expiry → Bucket                               |
| `Settings`        | Snapshot der verwendeten Settings                                            |
| `Watchlist`       | Snapshot der Watchlist dieses Laufs                                          |

Jede Kandidatenzeile enthält u. a.:

- Strike, Mid, Bid, Ask, Spread %
- IV, Delta, Open Interest, Volume
- **Cash Req. (USD)** = Strike × 100 — das zu parkende Kapital
- **Premium (USD)** = Mid × 100 — die vereinnahmte Optionsprämie je Kontrakt
- **Ann. Yield** = Premium / Strike × 365 / DTE (Prämienrendite auf Cash-Basis)
- **T-Bill Yield** und **T-Bill Interest (USD)** für die gemachte Maturity-Zuordnung
- **Total Yield** = Ann. Yield + T-Bill Yield (Kombinierte Rendite, wenn Put wertlos verfällt)
- **Breakeven** = Strike − Mid

## Methodik

**Warum CSP statt Direktinvestment?**

- Entry-Disziplin: Aktie wird nur bei Erreichen des Strikes gekauft.
- Cash-Besicherung im T-Bill liefert zusätzlich Risk-Free-Rendite.
- Premium senkt den effektiven Einstiegskurs auf Breakeven = Strike − Prämie.

**Formel annualisierte Prämienrendite:**

```
annualized_yield = (mid_premium / strike) * (365 / DTE)
```

Mid-Premium wird als Mittel aus Bid/Ask verwendet — bei fehlendem Ask/Bid
fällt der Scanner auf Last zurück.

**T-Bill-Matching:**

Für jede Expiry wird die größte Bill-Maturity gewählt, die ≤ DTE liegt
(z. B. bei DTE=45 → 28-Tage-Bill; bei DTE=120 → 91-Tage-Bill). Der Rest
wird overnight zum selben Yield gerollt (vereinfachte Annahme).

## Bekannte Limitationen

- **T-Bill-Quotes auf IB** sind ohne CUSIP nur indikativ. Für produktive
  Nutzung in Phase 2 entweder (a) konkrete CUSIPs in die Settings
  hinterlegen oder (b) einen FRED-Feed (DTB4WK, DTB3, DTB6MS, DGS1) als
  Primärquelle integrieren, IB nur als Ausführungsplattform.
- **Delayed Data** kann bei engen Märkten zu Spread-Mismatches führen —
  für realistische Preise auf Market-Data-Type 1 wechseln (kostet
  Market-Data-Subscription).
- **Earnings-Blackout** ist in Phase 1 bewusst nicht enthalten (nicht
  gewünscht). Bei Bedarf trivial im Selektor nachzurüsten.
- **Dividenden-Risk:** Früh-Assignment bei Short Puts ist selten, aber bei
  tief im Geld + Dividendenzahlung vor Expiry möglich. Nicht automatisch
  modelliert.

## Roadmap

Naheliegende Erweiterungen:

- **Roll-Logik:** bestehende Short-Put-Positionen automatisch überwachen
  (Delta-Schwelle, Profit-Target 50 %).
- **Earnings-Blackout** als optionaler Filter.
- **Automatisiertes T-Bill-Ordering** bei Assignment-Cash-Aufbau.
- **Forward-Test / Live-Paper-Trading**: Orders via `placeOrder` gegen Paper.
- **Backtest-Modul** mit ORATS/Polygon-Daten, wenn belastbare Historien-Evidenz
  gewünscht wird (in Phase 1 ausgeklammert).
