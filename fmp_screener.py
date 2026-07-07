"""
fmp_screener.py — Estela Capital
Busca quality compounders globales con derating significativo.

Universo: scrapeado de índices globales via index_universe.get_global_universe()
Precios/52w: yfinance (gratuito, sin límite de API)
Fundamentales opcionales: Financial Modeling Prep API (si se provee clave)

Expone:
    find_global_candidates(fmp_key=None, threshold=40.0, max_results=50) -> list[dict]
"""

from __future__ import annotations

import json
import os
import time
import urllib.request
import urllib.parse
from pathlib import Path
from typing import Optional

import yfinance as yf

# ---------------------------------------------------------------------------
# Rutas
# ---------------------------------------------------------------------------
_THIS_DIR = Path(__file__).parent
_CLAUDE_DIR = _THIS_DIR.parents[1]   # .claude/
_local_companies = _THIS_DIR / "companies.json"
_remote_companies = _CLAUDE_DIR / "AI DIGEST" / "backend" / "config" / "companies.json"
COMPANIES_JSON = _local_companies if _local_companies.exists() else _remote_companies

FMP_BASE = "https://financialmodelingprep.com/stable"

# ---------------------------------------------------------------------------
# Index universe (lazy import to avoid circular dependency)
# ---------------------------------------------------------------------------
def _get_index_universe() -> list[dict]:
    """Load the global index universe via index_universe module."""
    import importlib, sys
    # Ensure the screener directory is on the path
    screener_dir = str(_THIS_DIR)
    if screener_dir not in sys.path:
        sys.path.insert(0, screener_dir)
    try:
        iu = importlib.import_module("index_universe")
        return iu.get_global_universe(COMPANIES_JSON)
    except Exception as exc:
        print(f"  [FMP] index_universe import failed: {exc}")
        return []

# ---------------------------------------------------------------------------
# Mapeo país → geografía
# ---------------------------------------------------------------------------
COUNTRY_GEO: dict[str, str] = {
    "US": "US",
    "GB": "Europa", "DE": "Europa", "FR": "Europa", "NL": "Europa",
    "SE": "Europa", "CH": "Europa", "IT": "Europa", "ES": "Europa",
    "DK": "Europa", "NO": "Europa", "FI": "Europa", "BE": "Europa",
    "AT": "Europa", "PT": "Europa", "IE": "Europa", "LU": "Europa",
    "JP": "Asia",   "HK": "Asia",   "AU": "Asia",   "KR": "Asia",
    "TW": "Asia",   "SG": "Asia",   "IN": "Asia",   "CN": "Asia",
    "BR": "Latam",  "MX": "Latam",  "CO": "Latam",  "CL": "Latam",
    "AR": "Latam",
}

# Exchange name → ISO country (FMP returns full exchange name or code)
EXCHANGE_COUNTRY: dict[str, str] = {
    "NASDAQ": "US", "NYSE": "US", "NYSE American": "US", "NYSE Arca": "US",
    "London Stock Exchange": "GB", "LSE": "GB",
    "XETRA": "DE", "Frankfurt Stock Exchange": "DE",
    "Euronext Paris": "FR", "Paris": "FR",
    "Euronext Amsterdam": "NL", "Amsterdam": "NL",
    "Nasdaq Stockholm": "SE", "Stockholm": "SE",
    "SIX Swiss Exchange": "CH", "Swiss Exchange": "CH",
    "Borsa Italiana": "IT", "Milan": "IT",
    "Bolsa de Madrid": "ES", "Madrid": "ES",
    "Nasdaq Copenhagen": "DK",
    "Oslo Stock Exchange": "NO",
    "Nasdaq Helsinki": "FI",
    "Euronext Brussels": "BE",
    "Wiener Börse": "AT",
    "Euronext Lisbon": "PT",
    "Tokyo Stock Exchange": "JP", "TSE": "JP",
    "Hong Kong Stock Exchange": "HK", "HKEX": "HK",
    "Australian Securities Exchange": "AU", "ASX": "AU",
    "Korea Stock Exchange": "KR", "KRX": "KR",
    "Taiwan Stock Exchange": "TW", "TWSE": "TW",
    "Singapore Exchange": "SG", "SGX": "SG",
    "National Stock Exchange of India": "IN", "Bombay Stock Exchange": "IN",
    "B3": "BR",
    "Mexico Stock Exchange": "MX",
}

# ---------------------------------------------------------------------------
# Universo curado de quality compounders globales (~400 símbolos FMP)
# Incluye empresas con moats conocidos de MSCI World Quality, Stoxx600, S&P500 y Nikkei.
# Se usa como punto de partida cuando el screener endpoint está bloqueado por paywall.
# ---------------------------------------------------------------------------
GLOBAL_QUALITY_UNIVERSE: list[tuple[str, str]] = [
    # --- US Quality Compounders ---
    ("MSFT",  "US"), ("AAPL",  "US"), ("GOOGL", "US"), ("META",  "US"),
    ("AMZN",  "US"), ("V",     "US"), ("MA",    "US"), ("UNH",   "US"),
    ("JNJ",   "US"), ("PG",    "US"), ("KO",    "US"), ("PEP",   "US"),
    ("MCD",   "US"), ("SBUX",  "US"), ("NKE",   "US"), ("ADBE",  "US"),
    ("CRM",   "US"), ("NOW",   "US"), ("INTU",  "US"), ("ISRG",  "US"),
    ("TMO",   "US"), ("DHR",   "US"), ("ABT",   "US"), ("MDT",   "US"),
    ("SYK",   "US"), ("EW",    "US"), ("HOLX",  "US"), ("BDX",   "US"),
    ("ROK",   "US"), ("ROP",   "US"), ("VRSK",  "US"), ("MSCI",  "US"),
    ("SPGI",  "US"), ("MCO",   "US"), ("ICE",   "US"), ("CME",   "US"),
    ("CBOE",  "US"), ("FDS",   "US"), ("MORN",  "US"), ("VEEV",  "US"),
    ("ANSS",  "US"), ("CDNS",  "US"), ("SNPS",  "US"), ("FTNT",  "US"),
    ("PANW",  "US"), ("CSGP",  "US"), ("CBRE",  "US"), ("AMT",   "US"),
    ("CCI",   "US"), ("EQIX",  "US"), ("DLR",   "US"), ("PSA",   "US"),
    ("HSY",   "US"), ("MKC",   "US"), ("SJM",   "US"), ("CLX",   "US"),
    ("CHD",   "US"), ("ELF",   "US"), ("CPRT",  "US"), ("POOL",  "US"),
    ("WST",   "US"), ("TDY",   "US"), ("IDXX",  "US"), ("PODD",  "US"),
    ("ALGN",  "US"), ("TECH",  "US"), ("AMED",  "US"), ("GMED",  "US"),
    ("WAT",   "US"), ("A",     "US"), ("MTD",   "US"), ("CGNX",  "US"),
    ("LECO",  "US"), ("ITW",   "US"), ("GGG",   "US"), ("FAST",  "US"),
    ("NDSN",  "US"), ("AOS",   "US"), ("RBC",   "US"), ("CSL",   "US"),
    ("LFUS",  "US"), ("BWXT",  "US"), ("HXL",   "US"), ("ATR",   "US"),
    ("ODFL",  "US"), ("CHRW",  "US"), ("EXPD",  "US"), ("JBHT",  "US"),
    ("ORLY",  "US"), ("AZO",   "US"), ("BBY",   "US"), ("DG",    "US"),
    ("DLTR",  "US"), ("ROST",  "US"), ("TJX",   "US"), ("ULTA",  "US"),
    ("WSM",   "US"), ("TPR",   "US"), ("RL",    "US"), ("PVH",   "US"),
    ("HRI",   "US"), ("URI",   "US"), ("ALLE",  "US"), ("BRO",   "US"),
    ("AJG",   "US"), ("MMC",   "US"), ("AON",   "US"), ("WLTW",  "US"),
    ("EFX",   "US"), ("TRU",   "US"), ("NLSN",  "US"), ("MAN",   "US"),
    ("RHI",   "US"), ("PAYC",  "US"), ("PCTY",  "US"), ("HUBS",  "US"),
    ("ZM",    "US"), ("DOCU",  "US"), ("BILL",  "US"), ("COUP",  "US"),
    ("DDOG",  "US"), ("MDB",   "US"), ("ESTC",  "US"), ("OKTA",  "US"),
    ("NET",   "US"), ("CRWD",  "US"), ("ZS",    "US"), ("TENB",  "US"),
    ("QLYS",  "US"), ("RPM",   "US"), ("ECL",   "US"), ("SHW",   "US"),
    ("PPG",   "US"), ("IFF",   "US"), ("AVRY",  "US"), ("FMC",   "US"),
    ("CF",    "US"), ("NTR",   "US"), ("SMG",   "US"), ("AMGN",  "US"),
    ("GILD",  "US"), ("REGN",  "US"), ("VRTX",  "US"), ("ALNY",  "US"),
    ("BMRN",  "US"), ("SRPT",  "US"), ("RARE",  "US"), ("IONS",  "US"),
    ("NBIX",  "US"), ("EXEL",  "US"), ("HALO",  "US"), ("ARWR",  "US"),

    # --- Europa ---
    # UK
    ("RWS.L",   "GB"), ("SDR.L",  "GB"), ("DPLM.L", "GB"), ("HLMA.L", "GB"),
    ("IMB.L",   "GB"), ("BATS.L", "GB"), ("RMV.L",  "GB"), ("AUTO.L", "GB"),
    ("MNDI.L",  "GB"), ("SMDS.L", "GB"), ("EXPN.L", "GB"), ("RELX.L", "GB"),
    ("WPP.L",   "GB"), ("IPG.L",  "GB"), ("OMC",    "US"), ("BRBY.L", "GB"),
    ("BURBY",   "GB"), ("MONY.L", "GB"), ("WISE.L", "GB"), ("PNN.L",  "GB"),
    ("FRES.L",  "GB"), ("AHT.L",  "GB"), ("CATO",   "US"), ("INF.L",  "GB"),
    ("DCC.L",   "GB"), ("GRFS.L", "GB"), ("GSK.L",  "GB"), ("AZN.L",  "GB"),
    ("SN.L",    "GB"), ("CRH.L",  "GB"), ("ULVR.L", "GB"), ("RECKITT", "GB"),
    ("RB.L",    "GB"), ("DGE.L",  "GB"), ("BA.L",   "GB"),
    # Deutschland
    ("SAP.DE",  "DE"), ("SIE.DE", "DE"), ("BAYN.DE","DE"), ("MRK.DE", "DE"),
    ("BAS.DE",  "DE"), ("HEN3.DE","DE"), ("ALV.DE", "DE"), ("MUV2.DE","DE"),
    ("RAA.DE",  "DE"), ("HFCL",   "IN"), ("HOT.DE", "DE"), ("RHM.DE", "DE"),
    ("AIR.DE",  "DE"), ("FRE.DE", "DE"), ("MTX.DE", "DE"), ("BOY.DE", "DE"),
    ("WAF.DE",  "DE"), ("XTRA.DE","DE"), ("VNA.DE", "DE"), ("WDI.DE", "DE"),
    ("HNR1.DE", "DE"), ("2MX.DE", "DE"), ("1COV.DE","DE"), ("CON.DE", "DE"),
    # France
    ("MC.PA",   "FR"), ("OR.PA",  "FR"), ("RMS.PA", "FR"), ("KER.PA", "FR"),
    ("CDI.PA",  "FR"), ("SAN.PA", "FR"), ("AI.PA",  "FR"), ("DSY.PA", "FR"),
    ("CAP.PA",  "FR"), ("SGO.PA", "FR"), ("VIE.PA", "FR"), ("EDF.PA", "FR"),
    ("FP.PA",   "FR"), ("BNP.PA", "FR"), ("GLE.PA", "FR"), ("ACA.PA", "FR"),
    ("EN.PA",   "FR"), ("PUB.PA", "FR"), ("HO.PA",  "FR"), ("ORA.PA", "FR"),
    ("TEP.PA",  "FR"), ("BOL.PA", "FR"), ("SW.PA",  "FR"), ("AF.PA",  "FR"),
    ("ATO.PA",  "FR"), ("EDEN.PA","FR"), ("IPN.PA", "FR"), ("SOLB.BR","BE"),
    # Países Bajos
    ("WKL.AS",  "NL"), ("NN.AS",  "NL"), ("HEIA.AS","NL"), ("ASML.AS","NL"),
    ("PHIA.AS", "NL"), ("RAND.AS","NL"), ("IMCD.AS","NL"), ("BESI.AS","NL"),
    ("ASM.AS",  "NL"), ("AKZA.AS","NL"), ("TKWY.AS","NL"),
    # Suecia
    ("ATCO-A.ST","SE"),("EPAC.ST","SE"), ("HEXPB.ST","SE"),("INVE-B.ST","SE"),
    ("SAND.ST",  "SE"),("SWED-A.ST","SE"),("ESSITY-B.ST","SE"),
    ("ALFA.ST",  "SE"),("HUSQ-B.ST","SE"),("NIBE-B.ST","SE"),
    # Suiza
    ("NESN.SW",  "CH"), ("ROG.SW", "CH"), ("NOVN.SW","CH"), ("CFR.SW", "CH"),
    ("SIKA.SW",  "CH"), ("GEBN.SW","CH"), ("LOGN.SW","CH"), ("PGHN.SW","CH"),
    ("VACN.SW",  "CH"), ("STMN.SW","CH"), ("ABBN.SW","CH"), ("AMS.SW", "CH"),
    ("SRENH.SW", "CH"), ("ZURN.SW","CH"), ("CSGN.SW","CH"), ("UBS.SW", "CH"),
    ("GIVN.SW",  "CH"), ("BALN.SW","CH"), ("CPGN.SW","CH"), ("SLHN.SW","CH"),
    # Italia
    ("LDO.MI",  "IT"), ("MB.MI",  "IT"), ("FCA.MI", "IT"), ("BMED.MI","IT"),
    ("MONC.MI", "IT"), ("TOD.MI", "IT"), ("MONCLER","IT"), ("DIA.MI", "IT"),
    ("REPLY.MI","IT"), ("BAMI.MI","IT"), ("UCG.MI", "IT"),
    # España
    ("ITX.MC",  "ES"), ("ACS.MC", "ES"), ("FER.MC", "ES"), ("IAG.MC", "ES"),
    ("CABK.MC", "ES"), ("SAN.MC", "ES"), ("BBVA.MC","ES"), ("REP.MC", "ES"),
    # Dinamarca
    ("NOVO-B.CO","DK"),("NZYM-B.CO","DK"),("COLO-B.CO","DK"),("ROCK-B.CO","DK"),
    ("VWS.CO",  "DK"), ("AMBU-B.CO","DK"),("CARL-B.CO","DK"),
    # Noruega
    ("EQNR.OL", "NO"), ("DNB.OL", "NO"), ("ORK.OL", "NO"), ("TEL.OL", "NO"),
    # Finlandia
    ("FORTUM.HE","FI"),("NESTE.HE","FI"),("UPM.HE","FI"),("NOKIA.HE","FI"),
    # Bélgica
    ("LOTUS.BR","BE"), ("UCB.BR", "BE"), ("AB.BR",  "BE"), ("GBLB.BR","BE"),
    # Austria
    ("EBS.VI",  "AT"), ("OMV.VI", "AT"),

    # --- Asia ---
    # Japón
    ("7974.T",  "JP"), ("6758.T", "JP"), ("9984.T", "JP"), ("4519.T", "JP"),
    ("6861.T",  "JP"), ("9432.T", "JP"), ("6367.T", "JP"), ("7203.T", "JP"),
    ("6501.T",  "JP"), ("4568.T", "JP"), ("2914.T", "JP"), ("4661.T", "JP"),
    ("9766.T",  "JP"), ("9613.T", "JP"), ("8035.T", "JP"), ("4543.T", "JP"),
    ("6146.T",  "JP"), ("6273.T", "JP"), ("3382.T", "JP"), ("8267.T", "JP"),
    ("4901.T",  "JP"), ("7751.T", "JP"), ("6954.T", "JP"), ("6902.T", "JP"),
    ("4452.T",  "JP"), ("7741.T", "JP"), ("6762.T", "JP"), ("4716.T", "JP"),
    # Hong Kong
    ("0700.HK", "HK"), ("9988.HK","HK"), ("3690.HK","HK"), ("9618.HK","HK"),
    ("2020.HK", "HK"), ("0291.HK","HK"), ("1299.HK","HK"), ("0669.HK","HK"),
    ("1177.HK", "HK"), ("0241.HK","HK"), ("2382.HK","HK"), ("1398.HK","HK"),
    ("0005.HK", "HK"), ("3968.HK","HK"), ("2318.HK","HK"), ("0388.HK","HK"),
    # Australia
    ("CSL.AX",  "AU"), ("REA.AX", "AU"), ("CAR.AX", "AU"), ("XRO.AX", "AU"),
    ("WTC.AX",  "AU"), ("ALU.AX", "AU"), ("TNE.AX", "AU"), ("PME.AX", "AU"),
    ("BKW.AX",  "AU"), ("JBH.AX", "AU"), ("DMP.AX", "AU"), ("IEL.AX", "AU"),
    # Corea del Sur
    ("005930.KS","KR"),("000660.KS","KR"),("051910.KS","KR"),("035420.KS","KR"),
    ("018260.KS","KR"),("006400.KS","KR"),("028260.KS","KR"),
    # Taiwan
    ("2330.TW",  "TW"), ("2317.TW","TW"), ("2454.TW","TW"), ("2308.TW","TW"),
    ("3711.TW",  "TW"), ("2382.TW","TW"), ("2412.TW","TW"),
    # Singapur
    ("U11.SI",  "SG"), ("D05.SI", "SG"), ("O39.SI", "SG"), ("Z74.SI", "SG"),
    ("C07.SI",  "SG"), ("BN4.SI", "SG"),
    # India
    ("RELIANCE.NS","IN"),("TCS.NS","IN"),("INFY.NS","IN"),("HDFCBANK.NS","IN"),
    ("HINDUNILVR.NS","IN"),("NESTLEIND.NS","IN"),("PIDILITIND.NS","IN"),
    ("ASTRAL.NS","IN"),("DMART.NS","IN"),("BAJFINANCE.NS","IN"),

    # --- Latam ---
    ("WEGE3.SA","BR"),("LREN3.SA","BR"),("TOTS3.SA","BR"),("RENT3.SA","BR"),
    ("MELI",    "BR"), ("PBR",    "BR"), ("VALE",   "BR"), ("ITUB",  "BR"),
    ("GFNORTEO.MX","MX"),("CUERVO.MX","MX"),("BIMBOA.MX","MX"),
    ("CHLIE.SN","CL"), ("CENCOSUD.SN","CL"),
]

# ---------------------------------------------------------------------------
# HTTP helper
# ---------------------------------------------------------------------------
_HEADERS = {"User-Agent": "EstelaCapital-Screener/1.0"}


def _get(url: str, timeout: int = 15) -> Optional[dict | list]:
    """HTTP GET → parsed JSON, or None on error."""
    try:
        req = urllib.request.Request(url, headers=_HEADERS)
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception:
        return None


def _fmp_get(endpoint: str, fmp_key: str, params: dict | None = None) -> Optional[dict | list]:
    """Call a FMP stable endpoint. params dict is appended as query string."""
    p = {"apikey": fmp_key}
    if params:
        p.update(params)
    qs = urllib.parse.urlencode(p)
    url = f"{FMP_BASE}/{endpoint}?{qs}"
    return _get(url)


# ---------------------------------------------------------------------------
# Load existing universe tickers to avoid duplicates
# ---------------------------------------------------------------------------
def _load_existing_tickers() -> set[str]:
    """Returns set of ticker strings (upper-case) already in companies.json."""
    if not COMPANIES_JSON.exists():
        return set()
    try:
        with open(COMPANIES_JSON, encoding="utf-8") as f:
            companies = json.load(f)
        tickers: set[str] = set()
        for c in companies:
            t = c.get("ticker", "")
            if t:
                tickers.add(t.upper())
                tickers.add(t.upper().replace(".", "-"))
        return tickers
    except Exception:
        return set()


# ---------------------------------------------------------------------------
# Geography resolution
# ---------------------------------------------------------------------------
def _resolve_geo(country: str, exchange_full: str) -> str:
    """Map FMP country/exchange to Estela geo bucket."""
    if country:
        geo = COUNTRY_GEO.get(country.upper())
        if geo:
            return geo
    if exchange_full:
        for key, country_code in EXCHANGE_COUNTRY.items():
            if key.lower() in exchange_full.lower():
                return COUNTRY_GEO.get(country_code, "Global")
    return "Global"


# ---------------------------------------------------------------------------
# Parse 52-week high from FMP profile range string "low-high"
# ---------------------------------------------------------------------------
def _parse_year_high(range_str: str) -> Optional[float]:
    """Parse '199.26-317.4' → 317.4"""
    if not range_str or "-" not in range_str:
        return None
    try:
        parts = range_str.rsplit("-", 1)
        return float(parts[-1])
    except (ValueError, IndexError):
        return None


# ---------------------------------------------------------------------------
# Fetch profile for a single symbol
# ---------------------------------------------------------------------------
def _fetch_profile(symbol: str, fmp_key: str) -> Optional[dict]:
    """Returns the first item from /stable/profile?symbol=X or None."""
    data = _fmp_get("profile", fmp_key, {"symbol": symbol})
    if isinstance(data, list) and data:
        return data[0]
    if isinstance(data, dict) and data:
        return data
    return None


# ---------------------------------------------------------------------------
# Fetch ratios TTM for a single symbol (PE, gross margin, ROCE)
# ---------------------------------------------------------------------------
def _fetch_ratios(symbol: str, fmp_key: str) -> Optional[dict]:
    data = _fmp_get("ratios-ttm", fmp_key, {"symbol": symbol})
    if isinstance(data, list) and data:
        return data[0]
    return None


# ---------------------------------------------------------------------------
# yfinance derating signals
# ---------------------------------------------------------------------------
def get_derating_signals(yf_ticker: str) -> Optional[dict]:
    """
    Fetch price and 52-week high for a ticker using yfinance.
    Returns dict with price data, or None if data unavailable.
    """
    try:
        ticker_obj = yf.Ticker(yf_ticker)
        info = ticker_obj.fast_info
        # fast_info attributes: last_price, year_high, market_cap, currency
        price = getattr(info, "last_price", None)
        year_high = getattr(info, "year_high", None)
        market_cap = getattr(info, "market_cap", None)
        currency = getattr(info, "currency", "USD") or "USD"

        if price is None or year_high is None or year_high <= 0 or price <= 0:
            return None

        drop_pct = round((year_high - price) / year_high * 100, 1)

        return {
            "current_price": round(float(price), 2),
            "high_52w":      round(float(year_high), 2),
            "drop_from_high_pct": drop_pct,
            "market_cap":    market_cap,
            "currency":      str(currency),
        }
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Main public function
# ---------------------------------------------------------------------------
def find_global_candidates(
    fmp_key: Optional[str] = None,
    threshold: float = 40.0,
    max_results: int = 50,
) -> list[dict]:
    """
    Scan the global index universe (scraped from Wikipedia via index_universe)
    and return companies with significant derating that fit Estela Capital's
    investment philosophy.

    Price data is fetched via yfinance (free, no API key needed).
    FMP API is used optionally for fundamentals (PE, EV/EBITDA) if fmp_key is provided.

    Args:
        fmp_key:     Financial Modeling Prep API key (optional).
        threshold:   Minimum % drop from 52-week high to flag as derating.
        max_results: Maximum number of results to return (sorted by drop desc).

    Returns:
        List of result dicts compatible with screener.py format.
    """
    universe = _get_index_universe()
    if not universe:
        print("  [FMP] Universe is empty — check index_universe.py")
        return []

    print(f"  [FMP] Universe loaded: {len(universe)} companies to scan (threshold: {threshold}%)")

    # --- BATCH DOWNLOAD via yfinance (mucho más rápido que one-by-one) ---
    tickers_list = [c.get("yf_ticker", c.get("ticker", "")) for c in universe]
    tickers_list = [t for t in tickers_list if t]
    company_by_ticker = {c.get("yf_ticker", c.get("ticker", "")): c for c in universe if c.get("yf_ticker") or c.get("ticker")}

    print(f"  [FMP] Descargando precios en batch para {len(tickers_list)} tickers...")
    try:
        batch_data = yf.download(
            tickers_list,
            period="1y",
            interval="1d",
            group_by="ticker",
            auto_adjust=True,
            progress=False,
            threads=True,
            timeout=120,
        )
    except Exception as e:
        print(f"  [FMP] Batch download falló: {e}")
        batch_data = None

    candidates: list[dict] = []
    errors = 0
    below_threshold = 0

    if batch_data is None or batch_data.empty:
        print("  [FMP] No se pudieron descargar datos en batch")
    else:
        print(f"  [FMP] Batch completado, calculando derating...")
        # Normalise columns: yfinance returns MultiIndex (field, ticker) for multi-ticker
        import pandas as pd
        cols = batch_data.columns
        is_multi = isinstance(cols, pd.MultiIndex)

        for ticker in tickers_list:
            company = company_by_ticker.get(ticker)
            if not company:
                continue
            try:
                if is_multi:
                    # yfinance MultiIndex is (ticker, field) — level0=ticker, level1=field
                    level0 = [str(c) for c in cols.get_level_values(0)]
                    level1 = [str(c).lower() for c in cols.get_level_values(1)]
                    close_cols = [cols[i] for i, (t, f) in enumerate(zip(level0, level1))
                                  if t == ticker and f == "close"]
                    if not close_cols:
                        errors += 1
                        continue
                    close = batch_data[close_cols[0]].dropna()
                else:
                    # Single ticker — flat columns
                    if "Close" not in batch_data.columns:
                        errors += 1
                        continue
                    close = batch_data["Close"].dropna()

                if close.empty or len(close) < 2:
                    errors += 1
                    continue

                price = float(close.iloc[-1])
                year_high = float(close.max())

                if price <= 0 or year_high <= 0:
                    errors += 1
                    continue

                drop_pct = round((year_high - price) / year_high * 100, 1)

                if drop_pct < threshold:
                    below_threshold += 1
                    continue

                candidates.append({
                    "name":      company.get("name", ticker),
                    "ticker":    company.get("ticker", ticker),
                    "yf_ticker": ticker,
                    "geo":       company.get("geo", "Global"),
                    "sector":    company.get("sector", ""),
                    "signals": {
                        "current_price":          round(price, 2),
                        "high_52w":               round(year_high, 2),
                        "drop_from_high_pct":     drop_pct,
                        "pe_current":             None,
                        "pe_mean_5y":             None,
                        "pe_compression_pct":     None,
                        "ev_ebitda_current":      None,
                        "ev_ebitda_5y_mean":      None,
                        "ev_ebitda_discount_pct": None,
                        "currency":               "USD",
                        "market_cap":             None,
                        # Rich fields — populated below
                        "price_1m_ago":           None,
                        "roe":                    None,
                        "gross_margin":           None,
                        "revenue_growth_1y":      None,
                        "fcf_yield":              None,
                        "dividend_yield":         None,
                        "beta":                   None,
                        "net_debt_ebitda":        None,
                        "analyst_target":         None,
                        "analyst_upside_pct":     None,
                        "analyst_count":          None,
                        "recommendation":         "",
                        "description":            "",
                        "sector":                 company.get("sector", ""),
                    },
                    "derating_pass":   True,
                    "moat":            {},
                    "prioridad":       99,
                    "destacada":       False,
                    "razon_prioridad": "",
                    "source":          company.get("source_index", "index_universe"),
                })
            except Exception:
                errors += 1
                continue

    print(f"  [FMP] Resumen: {len(candidates)} candidatas | {below_threshold} bajo umbral | {errors} errores")

    # Sort by drop_from_high_pct descending (biggest derating first)
    candidates.sort(key=lambda r: r["signals"]["drop_from_high_pct"], reverse=True)
    top = candidates[:max_results]

    # Enrich top candidates with full yfinance info (ROE, margins, analysts, etc.)
    print(f"  [FMP] Enriqueciendo {len(top)} candidatas con datos fundamentales...")
    for cand in top:
        try:
            tk = yf.Ticker(cand.get("yf_ticker") or cand["ticker"])
            info = tk.info or {}
            sig = cand["signals"]

            # Price 1 month ago
            try:
                hist = tk.history(period="2mo", interval="1d")
                if len(hist) >= 22:
                    sig["price_1m_ago"] = round(float(hist["Close"].iloc[-22]), 2)
            except Exception:
                pass

            # Currency override
            sig["currency"] = info.get("currency", sig.get("currency", "USD")) or "USD"

            # PE & EV/EBITDA
            pe = info.get("trailingPE") or info.get("forwardPE")
            if pe:
                sig["pe_current"] = round(float(pe), 1)
            ev_ebitda = info.get("enterpriseToEbitda")
            if ev_ebitda:
                sig["ev_ebitda_current"] = round(float(ev_ebitda), 2)

            # Quality metrics
            roe = info.get("returnOnEquity")
            if roe:
                sig["roe"] = round(roe * 100, 1)
            gm = info.get("grossMargins")
            if gm:
                sig["gross_margin"] = round(gm * 100, 1)
            rg = info.get("revenueGrowth")
            if rg:
                sig["revenue_growth_1y"] = round(rg * 100, 1)
            fcf = info.get("freeCashflow")
            mkt = info.get("marketCap")
            if fcf and mkt and mkt > 0:
                sig["fcf_yield"] = round(fcf / mkt * 100, 1)
            dy = info.get("dividendYield")
            if dy:
                sig["dividend_yield"] = round(dy * 100, 2)
            beta = info.get("beta")
            if beta:
                sig["beta"] = round(float(beta), 2)

            # Debt/EBITDA
            total_debt = info.get("totalDebt") or 0
            cash = info.get("totalCash") or 0
            ebitda = info.get("ebitda")
            if ebitda and ebitda != 0:
                sig["net_debt_ebitda"] = round((total_debt - cash) / ebitda, 1)

            # Analyst targets
            at = info.get("targetMeanPrice")
            cp = sig.get("current_price")
            if at and cp and cp > 0:
                sig["analyst_target"] = round(float(at), 2)
                sig["analyst_upside_pct"] = round((at - cp) / cp * 100, 1)
            sig["analyst_count"] = info.get("numberOfAnalystOpinions")
            sig["recommendation"] = info.get("recommendationKey", "") or ""

            # Description & sector
            raw_desc = info.get("longBusinessSummary", "") or ""
            sig["description"] = (raw_desc[:297] + "...") if len(raw_desc) > 300 else raw_desc
            if not cand.get("sector"):
                cand["sector"] = info.get("sector", "") or info.get("industry", "")
            sig["sector"] = cand["sector"]

            # Market cap update
            if mkt:
                sig["market_cap"] = mkt

        except Exception:
            pass  # Leave signals as-is if enrichment fails

    return top


# ---------------------------------------------------------------------------
# __main__ — test rápido
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import sys

    fmp_key = os.environ.get("FMP_API_KEY") or "RwcpOZu3SNxBrH7qlHcvKVMe3SESFpnA"
    threshold = float(sys.argv[1]) if len(sys.argv) > 1 else 30.0

    print(f"\n=== fmp_screener TEST — umbral: {threshold}% ===\n")

    results = find_global_candidates(fmp_key, threshold=threshold, max_results=20)

    if not results:
        print("No se encontraron candidatas.")
    else:
        print(f"\n{'='*80}")
        print(f"{'Empresa':<35} {'Ticker':<14} {'Geo':<8} {'Sector':<22} {'Caída':>8} {'P/E':>7} {'EV/EBITDA':>10}")
        print(f"{'='*80}")
        for r in results:
            sig = r["signals"]
            print(
                f"{r['name'][:34]:<35} {r['ticker']:<14} {r['geo']:<8} "
                f"{r['sector'][:21]:<22} {sig['drop_from_high_pct']:>7.1f}% "
                f"{str(sig['pe_current'] or 'N/D'):>7} "
                f"{str(sig['ev_ebitda_current'] or 'N/D'):>10}"
            )
        print(f"\nTotal: {len(results)} candidatas encontradas")
