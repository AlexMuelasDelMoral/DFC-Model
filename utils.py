import pandas as pd
from io import BytesIO

def format_currency(value, unit="B"):
    divisor = {"B": 1e9, "M": 1e6, "K": 1e3}[unit]
    return f"${value/divisor:,.2f}{unit}"

def to_excel(dfs_dict):
    """Export multiple dataframes to Excel (one sheet each)."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for name, df in dfs_dict.items():
            df.to_excel(writer, sheet_name=name[:31])
    return output.getvalue()