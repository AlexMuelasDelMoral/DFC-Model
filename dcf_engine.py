import numpy as np
import pandas as pd

def run_dcf(base_fcf, growth_rates, terminal_growth, wacc, years):
    """
    growth_rates: list of annual growth rates (len = years) OR single float
    """
    if isinstance(growth_rates, (int, float)):
        growth_rates = [growth_rates] * years
    
    projected_fcf = []
    fcf = base_fcf
    for g in growth_rates:
        fcf = fcf * (1 + g)
        projected_fcf.append(fcf)
    
    discounted_fcf = [
        fcf / ((1 + wacc) ** (i + 1)) for i, fcf in enumerate(projected_fcf)
    ]
    
    if wacc <= terminal_growth:
        terminal_value = 0
    else:
        terminal_value = (projected_fcf[-1] * (1 + terminal_growth)) / (wacc - terminal_growth)
    
    discounted_terminal = terminal_value / ((1 + wacc) ** years)
    enterprise_value = sum(discounted_fcf) + discounted_terminal
    
    return {
        "projected_fcf": projected_fcf,
        "discounted_fcf": discounted_fcf,
        "terminal_value": terminal_value,
        "discounted_terminal": discounted_terminal,
        "enterprise_value": enterprise_value,
    }

def enterprise_to_equity(ev, cash, debt, shares):
    equity = ev + cash - debt
    price = equity / shares if shares else 0
    return equity, price

def sensitivity_analysis(base_fcf, growth, wacc_range, tg_range, years, cash, debt, shares):
    sens = pd.DataFrame(
        index=[f"{w*100:.2f}%" for w in wacc_range],
        columns=[f"{t*100:.2f}%" for t in tg_range]
    )
    for w in wacc_range:
        for t in tg_range:
            if w <= t:
                sens.loc[f"{w*100:.2f}%", f"{t*100:.2f}%"] = np.nan
                continue
            result = run_dcf(base_fcf, growth, t, w, years)
            _, price = enterprise_to_equity(result["enterprise_value"], cash, debt, shares)
            sens.loc[f"{w*100:.2f}%", f"{t*100:.2f}%"] = round(price, 2)
    return sens

    def build_full_projection(
    base_revenue, sales_growth, ebitda_margin, da_pct_sales,
    wc_pct_sales, capex_pct_sales, tax_rate, wacc, terminal_growth
):
    """
    Build a full line-by-line DCF projection schedule.
    All list inputs should be same length (projection years).
    Returns dict of lists.
    """
    years = len(sales_growth)
    
    revenue = []
    rev = base_revenue
    for g in sales_growth:
        rev = rev * (1 + g)
        revenue.append(rev)
    
    ebitda = [r * m for r, m in zip(revenue, ebitda_margin)]
    da = [r * d for r, d in zip(revenue, da_pct_sales)]
    ebit = [e - d for e, d in zip(ebitda, da)]
    nopat = [e * (1 - tax_rate) for e in ebit]
    
    # Working capital as % of sales, change in WC = delta
    wc = [r * w for r, w in zip(revenue, wc_pct_sales)]
    prev_wc = base_revenue * wc_pct_sales[0]  # base year WC approx
    change_wc = []
    for w in wc:
        change_wc.append(w - prev_wc)
        prev_wc = w
    
    capex = [r * c for r, c in zip(revenue, capex_pct_sales)]
    
    # Operating CF = NOPAT + D&A - Change in WC
    operating_cf = [n + d - cw for n, d, cw in zip(nopat, da, change_wc)]
    
    # Unlevered FCF = Operating CF - Capex
    unlevered_fcf = [o - cx for o, cx in zip(operating_cf, capex)]
    
    # Discounted FCF
    discounted_fcf = [
        f / ((1 + wacc) ** (i + 1)) for i, f in enumerate(unlevered_fcf)
    ]
    
    # Terminal value (Gordon growth on last unlevered FCF)
    if wacc > terminal_growth:
        terminal_value = (unlevered_fcf[-1] * (1 + terminal_growth)) / (wacc - terminal_growth)
    else:
        terminal_value = 0
    discounted_terminal = terminal_value / ((1 + wacc) ** years)
    
    enterprise_value = sum(discounted_fcf) + discounted_terminal
    
    return {
        "revenue": revenue,
        "ebitda": ebitda,
        "da": da,
        "ebit": ebit,
        "nopat": nopat,
        "wc": wc,
        "change_wc": change_wc,
        "capex": capex,
        "operating_cf": operating_cf,
        "unlevered_fcf": unlevered_fcf,
        "discounted_fcf": discounted_fcf,
        "terminal_value": terminal_value,
        "discounted_terminal": discounted_terminal,
        "enterprise_value": enterprise_value,
    }