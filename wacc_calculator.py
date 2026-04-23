import yfinance as yf


def get_risk_free_rate():
    """10-year Treasury yield as risk-free rate."""
    try:
        tnx = yf.Ticker("^TNX").history(period="5d")
        return tnx["Close"].iloc[-1] / 100
    except Exception:
        return 0.042  # fallback


def get_market_return():
    """Long-term S&P 500 average return."""
    return 0.10


def calculate_cost_of_equity(beta, rf, rm):
    """CAPM: Re = Rf + Beta * (Rm - Rf)"""
    return rf + beta * (rm - rf)


def calculate_cost_of_debt(info):
    """Approximate cost of debt = Interest Expense / Total Debt."""
    try:
        interest_expense = abs(info.get("interestExpense", 0) or 0)
        total_debt = info.get("totalDebt", 0) or 0
        if total_debt > 0 and interest_expense > 0:
            return interest_expense / total_debt
    except Exception:
        pass
    return 0.05  # fallback


def calculate_wacc(info, tax_rate=0.21):
    """WACC = (E/V)*Re + (D/V)*Rd*(1-T)"""
    try:
        beta = info.get("beta", 1.0) or 1.0
        market_cap = info.get("marketCap", 0) or 0
        total_debt = info.get("totalDebt", 0) or 0

        rf = get_risk_free_rate()
        rm = get_market_return()

        re = calculate_cost_of_equity(beta, rf, rm)
        rd = calculate_cost_of_debt(info)

        V = market_cap + total_debt
        if V == 0:
            return None, {}

        we = market_cap / V
        wd = total_debt / V

        wacc = we * re + wd * rd * (1 - tax_rate)

        breakdown = {
            "Risk-Free Rate": rf,
            "Market Return": rm,
            "Beta": beta,
            "Cost of Equity": re,
            "Cost of Debt": rd,
            "Weight of Equity": we,
            "Weight of Debt": wd,
            "Tax Rate": tax_rate,
            "WACC": wacc,
        }
        return wacc, breakdown
    except Exception as e:
        return None, {"error": str(e)}