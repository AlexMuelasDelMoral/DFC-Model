DCF Intelligence
Institutional‑grade discounted cash flow valuation platform built in Python, designed to replicate buy‑side analytical workflows used in equity research, asset management, and private investments.
This project focuses on valuation rigor, scenario analysis, and decision‑oriented outputs, not dashboards or demos.

Overview
DCF Intelligence is a fully self‑contained valuation engine that automates the end‑to‑end DCF process:

Data ingestion from financial statements and market data
Cash flow normalization and growth estimation
Cost of capital modeling (manual and CAPM‑based WACC)
Explicit forecast modeling and terminal value calculation
Sensitivity analysis across valuation drivers
Probabilistic valuation via Monte Carlo simulation
Professional Excel export matching institutional templates

The UI is intentionally minimal and dark‑themed, prioritizing readability, decision focus, and professional aesthetics rather than consumer styling.

Key Features
Valuation Engine

Unlevered DCF based on free cash flow to firm (FCFF)
Flexible base FCF selection (most recent, 3‑year average, 5‑year average)
Growth modeling using historical CAGR or manual assumptions
Terminal value using perpetual growth method
Enterprise‑to‑equity bridge including cash, debt, and shares outstanding

Cost of Capital

Automatic WACC calculation using CAPM assumptions
Manual WACC override for scenario testing
Transparent breakdown of individual cost of capital drivers

Risk & Scenario Analysis

Sensitivity matrix across WACC and terminal growth
Monte Carlo simulation for intrinsic value distribution
Probabilistic upside and downside assessment

Professional Outputs

Decision‑oriented summary (BUY / HOLD / SELL style interpretation)
Clean tabular views for analyst review
Fully structured Excel export consistent with sell‑side / buy‑side DCF models


Architecture
The project is deliberately modular to reflect production‑quality analytical codebases.
DCF-Model/
│
├── app.py                # Streamlit interface and orchestration
├── data_fetcher.py       # Financial and market data ingestion
├── dcf_engine.py         # Core DCF and valuation logic
├── wacc_calculator.py    # Cost of capital modeling
├── monte_carlo.py        # Probabilistic valuation simulation
├── excel_exporter.py     # Institutional-format Excel output
├── utils.py              # Shared helper functions
└── .streamlit/
    └── config.toml       # Runtime and UI configuration

Each component is decoupled, testable, and designed to be reusable outside the UI context.

Technology Stack

Python (financial modeling, analytics, simulation)
Pandas / NumPy (data transformation and numerical computation)
Streamlit (lightweight professional UI layer)
Plotly (interactive visualizations)
OpenPyXL / Excel tooling (exportable institutional models)

No proprietary platforms or services are required.

Design Philosophy
This project deliberately avoids:

Marketing dashboards
Over‑styled consumer UI patterns
Black‑box models

Instead, it emphasizes:

Transparency of assumptions
Analyst‑style workflows
Decision‑first presentation
Strong separation between logic and presentation

The UI is dark, neutral, and accent‑driven to mirror professional financial software environments.

Intended Use Cases

Equity research and valuation practice
Investment analysis prototyping
Financial modeling portfolio project
Interviews for roles in:

Asset management
Equity research
Private markets
Fintech analytics
Quantitative finance (applied)




Running the App
Shellpip install -r requirements.txtstreamlit run app.py``Show more lines
Requires Python 3.10+.

What This Project Demonstrates

Strong understanding of corporate valuation mechanics
Practical implementation of capital markets theory
Ability to translate financial concepts into maintainable code
Attention to decision‑driven UI design
Real‑world thinking beyond toy examples


Notes for Recruiters
This is not a tutorial project and not a generic dashboard clone.
It is designed to reflect how valuation is actually performed and stress‑tested in professional settings, with an emphasis on clarity, robustness, and interpretability.
