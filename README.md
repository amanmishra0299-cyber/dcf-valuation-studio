# DCF & Relative Valuation Studio

A valuation workbench for Indian listed companies. Enter an NSE/BSE symbol and it pulls
twelve years of financials from **screener.in**, estimates beta against the Nifty, and
runs four valuation lenses plus sensitivity, scenarios and comparables.

    python3 -m pip install -r requirements.txt
    python3 app.py            # http://localhost:8000

## What it computes

| Method | When it applies |
|---|---|
| **Unlevered FCF DCF** | The primary model. EBIT×(1−t) + D&A − capex − ΔNWC, discounted at WACC, Gordon or exit-multiple terminal value. |
| **Residual income (Edwards-Bell-Ohlson)** | Banks, NBFCs and insurers. Their borrowings are deposits, so unlevered FCF is undefined — book value plus capitalised excess returns is the right lens. The app switches automatically and says so. |
| **Two-stage DDM** | Cross-check. Mechanically low for low-payout companies. |
| **Trading comparables** | Peer median EV/EBITDA, P/E and P/B applied to the company's own metrics, ±25%. |

Plus: a WACC × terminal-growth sensitivity grid, bear/base/bull scenarios, a
market-implied growth solver, and a football field across all methods.

## Data sources

* **screener.in** — annual P&L, balance sheet, cash flow and ratios are published in the
  company page HTML and are not masked for logged-out visitors, so they are parsed
  directly. Current price, market cap, P/E, book value, dividend yield, ROE and ROCE come
  from the same page. Verified 2026-08-28.
* **Beta** — screener.in does not publish it, so it is estimated by OLS on five years of
  weekly returns against the Nifty. Yahoo's own beta field is a fallback; every value is
  editable.
* **Comparables** — a curated sector map is the default. Screener's own
  `/api/company/{id}/peers/` endpoint is available as a button, but it returned wrong
  buckets for logged-out requests during testing (liquor stocks for Asian Paints, holding
  companies for Reliance), so it is not the default. You can also type your own tickers.

Two screener.in endpoints did **not** work anonymously and are not used:
`/company/export/` (404, login-gated) and `/api/company/{id}/quick_ratios/` (login-gated).
Because the export button is login-gated, this app rebuilds an equivalent
screener-style `.xlsx` from the parsed data and lets you download it — and that file
round-trips back in through **Import Excel**.

## Import paths

1. **Ticker** — live fetch, the normal path.
2. **Excel / pasted text** — upload the `.xls`/`.xlsx` from screener.in's Export button
   (requires being signed in there), or paste tab-separated tables. Because the export
   carries no share count, re-upload with `?shares=<crore shares>&price=<CMP>` or use
   manual entry.
3. **Manual entry** — type the numbers, nothing is fetched.

## Tests

    python3 tests.py     # engine arithmetic, parser fixtures, Excel round trip, HTTP API
    node ui_test.js      # renders the real page in jsdom and drives it against the API

`tests.py` recomputes the DCF independently and asserts the app agrees, and covers the
WACC build-up, Gordon vs exit multiple, mid-year discounting, growth fade, sensitivity
monotonicity, DDM identity, peer-multiple arithmetic, the residual income identity, the
three Excel/paste layouts, and the API surface.

## Assumptions worth knowing

* Growth starts at a blend of the 3-year and 5-year revenue CAGR, capped at 22%, and fades
  toward terminal growth.
* Working capital is set at cash-conversion-cycle ÷ 365 of revenue, clamped to
  [−5%, +20%]. Without the upper clamp, TITAN's ~211-day gold inventory cycle put NWC at
  58% of revenue and produced a valuation of ₹113 against a ₹5,150 price.
* Capex defaults to the three-year average of the absolute investing cash flow.
* Terminal value assumes capex equals depreciation in the perpetuity year.
* Every one of these is editable in the Assumptions tab; clearing a field returns it to auto.
