import yfinance as yf
from datetime import datetime

TICKERS = [
    {"ticker": "RENT3", "nome": "Localiza"},
    {"ticker": "MOVI3", "nome": "Movida"},
    {"ticker": "LCAM3", "nome": "Unidas"},
]

INDICADORES = [
    ("pl",            "P/L"),
    ("pvp",           "P/VP"),
    ("roe",           "ROE"),
    ("margem_bruta",  "M. Bruta"),
    ("margem_ebitda", "M. EBITDA"),
    ("margem_ebit",   "M. EBIT"),
    ("margem_liquida","M. Líquida"),
    ("div_ebitda",    "Dív. Líq./EBITDA"),
    ("div_ebit",      "Dív. Líq./EBIT"),
]

_cache: dict = {}


def _nd(v, decimals=2):
    if v is None:
        return "N/D"
    try:
        return f"{float(v):.{decimals}f}".replace(".", ",")
    except (TypeError, ValueError):
        return "N/D"


def _pct(v):
    if v is None:
        return "N/D"
    try:
        return f"{float(v) * 100:.2f}%".replace(".", ",")
    except (TypeError, ValueError):
        return "N/D"


def _ratio(num, denom):
    try:
        n, d = float(num), float(denom)
        if not d:
            return "N/D"
        return f"{n / d:.2f}".replace(".", ",")
    except (TypeError, ValueError, ZeroDivisionError):
        return "N/D"


def scrape_ticker(ticker: str) -> dict:
    try:
        info = yf.Ticker(f"{ticker}.SA").info

        total_debt = float(info.get("totalDebt") or 0)
        total_cash = float(info.get("totalCash") or 0)
        div_liq = total_debt - total_cash

        ebitda = info.get("ebitda")
        revenue = info.get("totalRevenue")
        op_margins = info.get("operatingMargins")
        ebit = (float(op_margins) * float(revenue)) if (op_margins and revenue) else None

        return {
            "ticker":         ticker,
            "erro":           None,
            "pl":             _nd(info.get("trailingPE")),
            "pvp":            _nd(info.get("priceToBook")),
            "roe":            _pct(info.get("returnOnEquity")),
            "margem_bruta":   _pct(info.get("grossMargins")),
            "margem_ebitda":  _pct(info.get("ebitdaMargins")),
            "margem_ebit":    _pct(info.get("operatingMargins")),
            "margem_liquida": _pct(info.get("profitMargins")),
            "div_ebitda":     _ratio(div_liq, ebitda),
            "div_ebit":       _ratio(div_liq, ebit),
        }
    except Exception as e:
        return {"ticker": ticker, "erro": str(e),
                **{k: "N/D" for k in ["pl","pvp","roe","margem_bruta","margem_ebitda",
                                       "margem_ebit","margem_liquida","div_ebitda","div_ebit"]}}


def obter_dados(forcar: bool = False) -> list[dict]:
    agora = datetime.now()
    if not forcar and _cache.get("dados"):
        return _cache["dados"]

    dados = []
    for t in TICKERS:
        d = scrape_ticker(t["ticker"])
        d["nome"] = t["nome"]
        dados.append(d)

    _cache["dados"] = dados
    _cache["ts"] = agora
    _cache["atualizado_em"] = agora.strftime("%d/%m/%Y %H:%M")
    return dados


def cache_info() -> str:
    return _cache.get("atualizado_em", "Nunca")
