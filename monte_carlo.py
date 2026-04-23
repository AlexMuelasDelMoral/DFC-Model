import numpy as np
from dcf_engine import run_dcf, enterprise_to_equity


def run_monte_carlo(base_fcf, growth_mean, growth_std, wacc_mean, wacc_std,
                    tg_mean, tg_std, years, cash, debt, shares, n_sims=1000):
    """Run Monte Carlo simulation to get distribution of intrinsic prices."""
    prices = []
    for _ in range(n_sims):
        g = np.random.normal(growth_mean, growth_std)
        w = np.random.normal(wacc_mean, wacc_std)
        t = np.random.normal(tg_mean, tg_std)

        # Keep reasonable bounds
        w = max(w, t + 0.005)  # WACC must exceed terminal growth
        g = max(g, -0.5)

        result = run_dcf(base_fcf, g, t, w, years)
        _, price = enterprise_to_equity(
            result["enterprise_value"], cash, debt, shares
        )
        prices.append(price)
    return np.array(prices)