import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px

from data_fetcher import (
    fetch_company_data, get_historical_fcf,
    get_historical_revenue, calculate_fcf_growth_rate
)
from wacc_calculator import calculate_wacc
from dcf_engine import run_dcf, enterprise_to_equity, sensitivity_analysis
from monte_carlo import run_monte_carlo
from utils import to_excel

st.set_page_config(page_title="DCF Pro", layout="wide")

st.title("DCF Pro — Professional Valuation Model")
st.caption("Discounted Cash Flow analysis with auto-WACC, Monte Carlo, and sensitivity analysis.")

# ---------- SIDEBAR ----------
with st.sidebar:
    st.header("Ticker")
    ticker_input = st.text_input("Ticker(s) — comma-separated for compare", "AAPL").upper()
    tickers = [t.strip() for t in ticker_input.split(",") if t.strip()]
    
    st.header("Assumptions")
    
    auto_wacc = st.checkbox("Auto-calculate WACC", value=True)
    manual_wacc = st.slider("Manual WACC", 0.05, 0.20, 0.09, 0.005, disabled=auto_wacc)
    
    use_historical_growth = st.checkbox("Use historical FCF CAGR", value=True)
    manual_growth = st.slider("Manual FCF Growth", -0.10, 0.40, 0.08, 0.01,
                              disabled=use_historical_growth)
    
    terminal_growth = st.slider("Terminal Growth Rate", 0.0, 0.05, 0.025, 0.005)
    projection_years = st.slider("Projection Years", 5, 10, 5)
    tax_rate = st.slider("Tax Rate", 0.0, 0.40, 0.21, 0.01)
    
    st.header("Monte Carlo")
    run_mc = st.checkbox("Run Monte Carlo Simulation", value=True)
    n_sims = st.number_input("Simulations", 100, 10000, 1000, step=100)
    
    st.header("Base Year")
    fcf_method = st.radio("Base FCF method",
                          ["Most Recent", "3-Year Average", "5-Year Average"])

run_btn = st.button("Run Analysis", type="primary", use_container_width=True)

# ---------- SINGLE COMPANY FULL ANALYSIS ----------
def run_full_analysis(ticker):
    data = fetch_company_data(ticker)
    info = data["info"]
    cashflow = data["cashflow"]
    financials = data["financials"]
    history = data["history"]
    
    company_name = info.get("longName", ticker)
    current_price = info.get("currentPrice") or info.get("regularMarketPrice") or 0
    shares = info.get("sharesOutstanding", 0) or 0
    market_cap = info.get("marketCap", 0) or 0
    cash = info.get("totalCash", 0) or 0
    debt = info.get("totalDebt", 0) or 0
    
    st.header(f"{company_name} ({ticker})")
    
    # Overview
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Price", f"${current_price:.2f}")
    c2.metric("Market Cap", f"${market_cap/1e9:.2f}B")
    c3.metric("Cash", f"${cash/1e9:.2f}B")
    c4.metric("Debt", f"${debt/1e9:.2f}B")
    c5.metric("Beta", f"{info.get('beta', 'N/A')}")
    
    # Historical FCF
    hist_fcf = get_historical_fcf(cashflow, 5)
    if hist_fcf is None or hist_fcf.empty:
        st.error("Insufficient cash flow data.")
        return
    
    # Base FCF
    if fcf_method == "Most Recent":
        base_fcf = hist_fcf.iloc[0]
    elif fcf_method == "3-Year Average":
        base_fcf = hist_fcf.head(3).mean()
    else:
        base_fcf = hist_fcf.head(5).mean()
    
    # Growth rate
    hist_growth = calculate_fcf_growth_rate(hist_fcf)
    growth = hist_growth if use_historical_growth and hist_growth else manual_growth
    if use_historical_growth and hist_growth is None:
        st.warning("Couldn't compute historical CAGR → using manual growth.")
        growth = manual_growth
    # sanity caps
    growth = max(min(growth, 0.30), -0.10)
    
    # WACC
    if auto_wacc:
        wacc, wacc_breakdown = calculate_wacc(info, tax_rate)
        if wacc is None:
            st.warning("Auto WACC failed → using manual.")
            wacc = manual_wacc
            wacc_breakdown = {}
    else:
        wacc = manual_wacc
        wacc_breakdown = {}
    
    # Tabs
    tabs = st.tabs([
        "Summary", "Historicals", "DCF",
        "Sensitivity", "Monte Carlo", "WACC", "Export"
    ])
    
    # DCF
    dcf = run_dcf(base_fcf, growth, terminal_growth, wacc, projection_years)
    equity_value, intrinsic_price = enterprise_to_equity(
        dcf["enterprise_value"], cash, debt, shares)
    upside = ((intrinsic_price - current_price) / current_price * 100) if current_price else 0
    
    # ---- SUMMARY ----
    with tabs[0]:
        st.subheader("Valuation Summary")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Enterprise Value", f"${dcf['enterprise_value']/1e9:.2f}B")
        c2.metric("Equity Value", f"${equity_value/1e9:.2f}B")
        c3.metric("Intrinsic Price", f"${intrinsic_price:.2f}")
        c4.metric("Upside/Downside", f"{upside:.1f}%", delta=f"{upside:.1f}%")
        
        st.info(f"""
        **Base FCF:** ${base_fcf/1e9:.2f}B ({fcf_method})  
        **Growth Rate Applied:** {growth*100:.2f}%  
        **WACC:** {wacc*100:.2f}%  
        **Terminal Growth:** {terminal_growth*100:.2f}%  
        **Projection Horizon:** {projection_years} years
        """)
        
        # Verdict
        if upside > 20:
            st.success(f"Undervalued by {upside:.1f}% — Potential BUY")
        elif upside < -20:
            st.error(f"Overvalued by {abs(upside):.1f}% — Potential SELL")
        else:
            st.warning(f"Fairly valued ({upside:.1f}%)")
    
    # ---- HISTORICALS ----
    with tabs[1]:
        st.subheader("Historical FCF")
        fcf_df = pd.DataFrame({
            "Year": [str(d.year) for d in hist_fcf.index],
            "FCF ($B)": [v/1e9 for v in hist_fcf.values]
        })
        st.dataframe(fcf_df, use_container_width=True)
        
        fig = px.bar(fcf_df, x="Year", y="FCF ($B)", title="Historical Free Cash Flow")
        st.plotly_chart(fig, use_container_width=True)
        
        rev = get_historical_revenue(financials, 5)
        if rev is not None:
            rev_df = pd.DataFrame({
                "Year": [str(d.year) for d in rev.index],
                "Revenue ($B)": [v/1e9 for v in rev.values]
            })
            st.subheader("Historical Revenue")
            st.plotly_chart(px.bar(rev_df, x="Year", y="Revenue ($B)"),
                            use_container_width=True)
        
        st.subheader("5-Year Price History")
        st.line_chart(history["Close"])
    
    # ---- DCF ----
    with tabs[2]:
        st.subheader("Projected Cash Flows")
        proj_df = pd.DataFrame({
            "Year": [f"Year {i+1}" for i in range(projection_years)],
            "Projected FCF ($B)": [p/1e9 for p in dcf["projected_fcf"]],
            "Discounted FCF ($B)": [d/1e9 for d in dcf["discounted_fcf"]]
        })
        st.dataframe(proj_df.style.format({
            "Projected FCF ($B)": "{:.2f}",
            "Discounted FCF ($B)": "{:.2f}"
        }), use_container_width=True)
        
        fig = go.Figure()
        fig.add_trace(go.Bar(x=proj_df["Year"], y=proj_df["Projected FCF ($B)"],
                             name="Projected"))
        fig.add_trace(go.Bar(x=proj_df["Year"], y=proj_df["Discounted FCF ($B)"],
                             name="Discounted"))
        fig.update_layout(barmode="group", title="Projected vs Discounted FCF")
        st.plotly_chart(fig, use_container_width=True)
        
        st.metric("Terminal Value", f"${dcf['terminal_value']/1e9:.2f}B")
        st.metric("Discounted Terminal Value", f"${dcf['discounted_terminal']/1e9:.2f}B")
        
        # EV composition pie
        pie = go.Figure(data=[go.Pie(
            labels=["Sum of Discounted FCF", "Discounted Terminal Value"],
            values=[sum(dcf["discounted_fcf"]), dcf["discounted_terminal"]],
            hole=0.4
        )])
        pie.update_layout(title="Enterprise Value Composition")
        st.plotly_chart(pie, use_container_width=True)
    
    # ---- SENSITIVITY ----
    with tabs[3]:
        st.subheader("Sensitivity Analysis — Intrinsic Price")
        wacc_range = np.arange(wacc - 0.02, wacc + 0.025, 0.005)
        tg_range = np.arange(max(0.0, terminal_growth - 0.01),
                             terminal_growth + 0.015, 0.005)
        sens = sensitivity_analysis(base_fcf, growth, wacc_range, tg_range,
                                    projection_years, cash, debt, shares)
        sens.index.name = "WACC ↓ / Terminal g →"
        
        styled = sens.style.background_gradient(cmap="RdYlGn", axis=None)\
                           .format("{:.2f}")
        st.dataframe(styled, use_container_width=True)
        st.caption(f"Current market price: ${current_price:.2f}")
    
    # ---- MONTE CARLO ----
    with tabs[4]:
        if run_mc:
            st.subheader(f"Monte Carlo Simulation ({n_sims} runs)")
            with st.spinner("Running simulations..."):
                prices = run_monte_carlo(
                    base_fcf,
                    growth_mean=growth, growth_std=0.03,
                    wacc_mean=wacc, wacc_std=0.01,
                    tg_mean=terminal_growth, tg_std=0.005,
                    years=projection_years,
                    cash=cash, debt=debt, shares=shares,
                    n_sims=int(n_sims)
                )
            prices = prices[np.isfinite(prices)]
            prices = prices[(prices > 0) & (prices < current_price * 10)]
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Mean", f"${prices.mean():.2f}")
            c2.metric("Median", f"${np.median(prices):.2f}")
            c3.metric("5th %ile", f"${np.percentile(prices, 5):.2f}")
            c4.metric("95th %ile", f"${np.percentile(prices, 95):.2f}")
            
            fig = px.histogram(prices, nbins=50,
                               title="Distribution of Intrinsic Price")
            fig.add_vline(x=current_price, line_dash="dash", line_color="red",
                          annotation_text=f"Current: ${current_price:.2f}")
            fig.add_vline(x=np.median(prices), line_dash="dash", line_color="green",
                          annotation_text=f"Median: ${np.median(prices):.2f}")
            st.plotly_chart(fig, use_container_width=True)
            
            prob_undervalued = (prices > current_price).mean() * 100
            st.info(f"Probability stock is undervalued: **{prob_undervalued:.1f}%**")
        else:
            st.info("Enable Monte Carlo in sidebar to run this.")
    
    # ---- WACC ----
    with tabs[5]:
        st.subheader("WACC Breakdown")
        if wacc_breakdown:
            for key, val in wacc_breakdown.items():
                if isinstance(val, float):
                    st.write(f"**{key}:** {val*100:.2f}%" if abs(val) < 10
                             else f"**{key}:** {val:.2f}")
                else:
                    st.write(f"**{key}:** {val}")
        else:
            st.info("Using manual WACC.")
    
 # ---- EXPORT ----
    with tabs[6]:
        st.subheader("Professional DCF Excel Export")
        st.caption("Exports a full DCF model matching institutional template format.")
        
        from excel_exporter import create_dcf_excel
        from dcf_engine import build_full_projection
        import datetime
        
        # Default projection assumptions (editable)
        col1, col2 = st.columns(2)
        with col1:
            sales_growth_default = st.text_input(
                "Sales Growth % (comma-separated, 6 years)",
                "16.7, 18.0, 18.0, 14.0, 10.0, 7.0"
            )
            ebitda_margin_default = st.text_input(
                "EBITDA Margin %",
                "12.5, 12.5, 12.5, 12.5, 12.5, 12.5"
            )
            da_default = st.text_input(
                "D&A % of Sales",
                "1.7, 1.7, 1.6, 1.6, 1.6, 1.6"
            )
        with col2:
            wc_default = st.text_input(
                "Working Capital % Sales",
                "30.0, 26.0, 22.0, 18.0, 17.0, 17.0"
            )
            capex_default = st.text_input(
                "Capex % of Sales",
                "5.0, 6.0, 6.0, 2.2, 2.0, 2.0"
            )
            guidance_text = st.text_area("Guidance notes", "")
        
        def parse_pct_list(s):
            return [float(x.strip()) / 100 for x in s.split(",")]
        
        try:
            sg = parse_pct_list(sales_growth_default)
            em = parse_pct_list(ebitda_margin_default)
            da = parse_pct_list(da_default)
            wc_p = parse_pct_list(wc_default)
            cx = parse_pct_list(capex_default)
            
            # Base revenue from financials
            try:
                base_rev = financials.loc["Total Revenue"].iloc[0]
            except Exception:
                base_rev = info.get("totalRevenue", 0) or 0
            
            # Build full projection
            projection = build_full_projection(
                base_rev, sg, em, da, wc_p, cx,
                tax_rate, wacc, terminal_growth
            )
            
            # Scale to billions for display
            scale = 1e9
            years_list = [datetime.datetime.now().year + i + 1
                          for i in range(len(sg))]
            
            # Sensitivity matrices
            sens_wacc = [wacc - 0.01, wacc, wacc + 0.005,
                         wacc + 0.01, wacc + 0.02]
            sens_tg = [terminal_growth - 0.01, terminal_growth - 0.005,
                       terminal_growth, terminal_growth + 0.005]
            
            wacc_matrix = []
            price_matrix = []
            for t in sens_tg:
                wacc_row = []
                price_row = []
                for w in sens_wacc:
                    if w <= t:
                        wacc_row.append(0)
                        price_row.append(0)
                        continue
                    proj = build_full_projection(
                        base_rev, sg, em, da, wc_p, cx, tax_rate, w, t
                    )
                    ev = proj["enterprise_value"] / scale
                    eq = ev + (cash - debt) / scale
                    price_val = eq * scale / shares if shares else 0
                    wacc_row.append(ev)
                    price_row.append(price_val)
                wacc_matrix.append(wacc_row)
                price_matrix.append(price_row)
            
            # Valuation multiples
            pe = info.get("trailingPE")
            pb = info.get("priceToBook")
            ev_ebitda = info.get("enterpriseToEbitda")
            
            model_data = {
                "company_name": company_name,
                "ticker": ticker,
                "current_price": current_price,
                "target_price": intrinsic_price,
                "tax_rate": tax_rate,
                "wacc": wacc,
                "terminal_growth": terminal_growth,
                "shares_outstanding": shares,
                "net_debt": (debt - cash) / scale,
                "pe": pe, "pb": pb, "ev_ebitda": ev_ebitda,
                "years": years_list,
                "sales_growth": sg,
                "ebitda_margin": em,
                "da_pct_sales": da,
                "wc_pct_sales": wc_p,
                "capex_pct_sales": cx,
                "base_revenue": base_rev / scale,
                "revenue": [v / scale for v in projection["revenue"]],
                "ebitda": [v / scale for v in projection["ebitda"]],
                "da": [v / scale for v in projection["da"]],
                "ebit": [v / scale for v in projection["ebit"]],
                "nopat": [v / scale for v in projection["nopat"]],
                "wc": [v / scale for v in projection["wc"]],
                "change_wc": [v / scale for v in projection["change_wc"]],
                "capex": [v / scale for v in projection["capex"]],
                "operating_cf": [v / scale for v in projection["operating_cf"]],
                "unlevered_fcf": [v / scale for v in projection["unlevered_fcf"]],
                "discounted_fcf": [v / scale for v in projection["discounted_fcf"]],
                "enterprise_value": projection["enterprise_value"] / scale,
                "equity_value": (projection["enterprise_value"] + cash - debt) / scale,
                "terminal_value": projection["terminal_value"] / scale,
                "sens_wacc": sens_wacc,
                "sens_tg": sens_tg,
                "sens_wacc_matrix": wacc_matrix,
                "sens_price_matrix": price_matrix,
                "guidance": guidance_text,
                "context": f"DCF model for {company_name} generated on "
                           f"{datetime.date.today()}",
            }
            
            excel_bytes = create_dcf_excel(model_data)
            
            st.success("✅ Excel model ready!")
            st.download_button(
                "📥 Download Professional DCF Model (.xlsx)",
                data=excel_bytes,
                file_name=f"{ticker}_DCF_Model.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            # Preview of projections
            st.subheader("Preview — Full Projection Schedule (in $B)")
            preview = pd.DataFrame({
                "Year": years_list,
                "Revenue": [f"{v:.1f}" for v in model_data["revenue"]],
                "EBITDA": [f"{v:.1f}" for v in model_data["ebitda"]],
                "EBIT": [f"{v:.1f}" for v in model_data["ebit"]],
                "NOPAT": [f"{v:.1f}" for v in model_data["nopat"]],
                "Unlevered FCF": [f"{v:.1f}" for v in model_data["unlevered_fcf"]],
                "Discounted FCF": [f"{v:.1f}" for v in model_data["discounted_fcf"]],
            })
            st.dataframe(preview, use_container_width=True)
            
        except Exception as e:
            st.error(f"Error building Excel: {e}")
            import traceback
            st.code(traceback.format_exc())

# ---------- MAIN ----------
if run_btn:
    if len(tickers) == 1:
        try:
            run_full_analysis(tickers[0])
        except Exception as e:
            st.error(f"Error analyzing {tickers[0]}: {e}")
    else:
        st.header("Multi-Company Comparison")
        results = []
        for t in tickers:
            try:
                st.divider()
                r = run_full_analysis(t)
                if r:
                    results.append(r)
            except Exception as e:
                st.error(f"Error on {t}: {e}")
        
        if results:
            st.divider()
            st.header("Comparison Table")
            comp_df = pd.DataFrame(results)
            st.dataframe(comp_df.style.format({
                "Current Price": "${:.2f}",
                "Intrinsic Price": "${:.2f}",
                "Upside %": "{:.1f}%",
                "WACC": "{:.2%}",
                "Growth": "{:.2%}",
            }).background_gradient(subset=["Upside %"], cmap="RdYlGn"),
                         use_container_width=True)
else:
    st.info("Enter a ticker and click **Run Analysis**.")
    with st.expander("About this app"):
        st.markdown("""
        **Features:**
        - Auto-WACC via CAPM (live 10-yr Treasury)
        - Historical FCF CAGR or custom growth
        - Base FCF = most recent, 3-yr, or 5-yr avg
        - Sensitivity heatmap
        - Monte Carlo simulation
        - Multi-company comparison (comma-separated tickers)
        - Excel export
        """)