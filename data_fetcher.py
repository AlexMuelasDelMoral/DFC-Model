import yfinance as yf
import pandas as pd
import streamlit as st

@st.cache_data(ttl=3600)
def fetch_company_data(ticker):
    """Fetch all needed financial data for a ticker."""
    stock = yf.Ticker(ticker)
    data = {
        "info": stock.info,
        "cashflow": stock.cashflow,
        "financials": stock.financials,
        "balance_sheet": stock.balance_sheet,
        "quarterly_cashflow": stock.quarterly_cashflow,
        "history": stock.history(period="5y"),
    }
    return data

def get_historical_fcf(cashflow, years=5):
    """Get historical FCF for the last N years."""
    try:
        ocf = cashflow.loc["Operating Cash Flow"]
        capex = cashflow.loc["Capital Expenditure"]
        fcf = ocf + capex  # capex negative
        return fcf.head(years)
    except Exception:
        return None

def get_historical_revenue(financials, years=5):
    try:
        return financials.loc["Total Revenue"].head(years)
    except Exception:
        return None

def calculate_fcf_growth_rate(historical_fcf):
    """CAGR of historical FCF."""
    try:
        fcf_values = historical_fcf.dropna().values
        if len(fcf_values) < 2 or fcf_values[-1] <= 0:
            return None
        years = len(fcf_values) - 1
        # fcf is ordered newest -> oldest
        start, end = fcf_values[-1], fcf_values[0]
        cagr = (end / start) ** (1 / years) - 1
        return cagr
    except Exception:
        return None