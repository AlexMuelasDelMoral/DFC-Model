import xlsxwriter
from io import BytesIO


def create_dcf_excel(model_data):
    """
    Build a professional DCF Excel file matching the template format.
    
    model_data dict keys:
        company_name, ticker, current_price, target_price,
        tax_rate, wacc, terminal_growth, shares_outstanding, net_debt,
        pe, pb, ev_ebitda,
        years (list, e.g. [2024, 2025, ...]),
        sales_growth (list of %),
        ebitda_margin (list of %),
        da_pct_sales (list of %),
        wc_pct_sales (list of %),
        capex_pct_sales (list of %),
        base_revenue (float, most recent),
        revenue (list), ebitda (list), da (list), ebit (list),
        nopat (list), wc (list), change_wc (list),
        operating_cf (list), capex (list), unlevered_fcf (list),
        discounted_fcf (list),
        enterprise_value, equity_value, terminal_value,
        wacc_sensitivity (dict: {wacc: {tg: price}}),
    """
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})
    ws = wb.add_worksheet("DCF Model")
    ws.hide_gridlines(2)

    # ---------- FORMATS ----------
    NAVY = "#1F4E79"
    LIGHT_BLUE = "#D9E1F2"
    YELLOW = "#FFF2CC"

    title_fmt = wb.add_format({
        "bold": True, "font_size": 16, "align": "center",
        "valign": "vcenter", "bottom": 2
    })
    banner_fmt = wb.add_format({
        "bold": True, "font_color": "white", "bg_color": NAVY,
        "align": "center", "valign": "vcenter", "font_size": 11
    })
    section_fmt = wb.add_format({
        "bold": True, "font_color": "white", "bg_color": NAVY,
        "align": "left", "valign": "vcenter", "border": 1
    })
    section_center_fmt = wb.add_format({
        "bold": True, "font_color": "white", "bg_color": NAVY,
        "align": "center", "valign": "vcenter", "border": 1
    })
    label_fmt = wb.add_format({
        "bold": True, "bg_color": LIGHT_BLUE, "border": 1, "align": "left"
    })
    sublabel_fmt = wb.add_format({"align": "left", "border": 1})
    italic_label = wb.add_format({
        "italic": True, "align": "left", "border": 1, "font_color": "#404040"
    })
    input_pct = wb.add_format({
        "bg_color": YELLOW, "border": 1, "align": "right",
        "num_format": "0.0%"
    })
    input_num = wb.add_format({
        "bg_color": YELLOW, "border": 1, "align": "right",
        "num_format": "#,##0.00"
    })
    pct_fmt = wb.add_format({
        "num_format": "0.0%", "border": 1, "align": "right"
    })
    num_fmt = wb.add_format({
        "num_format": "#,##0", "border": 1, "align": "right"
    })
    num_decimal_fmt = wb.add_format({
        "num_format": "#,##0.0", "border": 1, "align": "right"
    })
    bold_num_fmt = wb.add_format({
        "bold": True, "num_format": "#,##0.0", "border": 1, "align": "right",
        "top": 2, "bg_color": LIGHT_BLUE
    })
    total_fmt = wb.add_format({
        "bold": True, "num_format": "#,##0", "border": 1,
        "align": "right", "bg_color": LIGHT_BLUE
    })
    guidance_header = wb.add_format({
        "bold": True, "align": "left", "bottom": 2, "font_size": 11
    })
    guidance_box = wb.add_format({"border": 1, "text_wrap": True, "valign": "top"})

    # ---------- COLUMN WIDTHS ----------
    ws.set_column("A:A", 26)
    ws.set_column("B:B", 14)
    ws.set_column("C:C", 2)
    ws.set_column("D:D", 28)
    ws.set_column("E:J", 11)
    ws.set_column("K:K", 2)
    ws.set_column("L:L", 22)
    ws.set_column("M:M", 22)

    # ---------- TITLE + BANNER ----------
    ws.merge_range("A1:J1", model_data.get("company_name", "Company"), title_fmt)
    ws.merge_range("A2:J2", "Discounted Cash Flow Model (DCF)", banner_fmt)
    ws.set_row(0, 24)
    ws.set_row(1, 20)

    # Guidance boxes (right side)
    ws.write("L1", "Guidance:", guidance_header)
    ws.merge_range("L2:M6", model_data.get("guidance", ""), guidance_box)
    ws.write("L8", "Context:", guidance_header)
    ws.merge_range("L9:M14", model_data.get("context", ""), guidance_box)

    # ---------- COMPANY NAME ----------
    ws.write("A4", "Company Name", label_fmt)
    ws.write("B4", model_data.get("company_name", ""), sublabel_fmt)

    # ---------- INPUTS BLOCK (left) ----------
    ws.merge_range("A6:B6", "Inputs", section_fmt)
    inputs = [
        ("Effective Tax Rate", model_data["tax_rate"], "pct"),
        ("WACC", model_data["wacc"], "pct"),
        ("Growth Rate", model_data["terminal_growth"], "pct"),
        ("No Shares Outstanding", model_data["shares_outstanding"] / 1e9, "num"),
        ("Target Stock Price", model_data["target_price"], "num"),
        ("Current Stock Price", model_data["current_price"], "num"),
    ]
    for i, (lbl, val, kind) in enumerate(inputs):
        r = 6 + i  # row index (0-based)
        ws.write(r, 0, lbl, label_fmt)
        fmt = input_pct if kind == "pct" else input_num
        ws.write(r, 1, val, fmt)

    # ---------- VALUATION MULTIPLES ----------
    ws.merge_range("A14:B14", "Valuation Multiples", section_fmt)
    multiples = [
        ("P/E", model_data.get("pe")),
        ("P/B", model_data.get("pb")),
        ("EV/EBITDA", model_data.get("ev_ebitda")),
    ]
    for i, (lbl, val) in enumerate(multiples):
        r = 14 + i
        ws.write(r, 0, lbl, label_fmt)
        ws.write(r, 1, val if val is not None else "N/A", num_decimal_fmt)

    # ---------- MANUAL OVERRIDES (center) ----------
    years = model_data["years"]
    n = len(years)
    ws.merge_range(5, 3, 5, 3 + n, "Manual overrides for forecasts*",
                   section_center_fmt)
    ws.write(6, 3, "Year", label_fmt)
    for i, y in enumerate(years):
        ws.write(6, 4 + i, y, label_fmt)

    override_rows = [
        ("Sales Growth %", model_data["sales_growth"]),
        ("EBITDA margin %", model_data["ebitda_margin"]),
        ("D&A as % of Sales", model_data["da_pct_sales"]),
        ("Working Capital % Sales", model_data["wc_pct_sales"]),
        ("Capex % of Sales", model_data["capex_pct_sales"]),
    ]
    for i, (lbl, values) in enumerate(override_rows):
        r = 7 + i
        ws.write(r, 3, lbl, italic_label)
        for j, v in enumerate(values):
            ws.write(r, 4 + j, v, input_pct)

    # ---------- FORECASTS TABLE ----------
    fc_start = 13  # row 14 (0-based 13)
    ws.merge_range(fc_start, 3, fc_start, 3 + n - 1, "Forecasts", section_fmt)
    ws.write(fc_start, 3 + n, "Term Year", section_center_fmt)

    def write_row(r, label, values, fmt=num_decimal_fmt, label_format=sublabel_fmt):
        ws.write(r, 3, label, label_format)
        for i, v in enumerate(values):
            if v is None or v == "":
                ws.write_blank(r, 4 + i, None, fmt)
            else:
                ws.write(r, 4 + i, v, fmt)

    # Revenue row (bold, with base year)
    base_rev_fmt = wb.add_format({
        "bold": True, "num_format": "#,##0", "border": 1, "align": "right"
    })
    ws.write(fc_start + 1, 3, "Revenue", label_fmt)
    ws.write(fc_start + 1, 4, model_data["base_revenue"], base_rev_fmt)
    for i, v in enumerate(model_data["revenue"]):
        ws.write(fc_start + 1, 5 + i, v, num_fmt)

    # EBITDA
    ws.write(fc_start + 3, 3, "EBITDA", label_fmt)
    ws.write(fc_start + 3, 4, model_data["ebitda"][0] * 0.8, num_decimal_fmt)  # placeholder base
    for i, v in enumerate(model_data["ebitda"]):
        ws.write(fc_start + 3, 5 + i, v, num_decimal_fmt)

    # D&A
    ws.write(fc_start + 4, 3, "Depreciation/amortization", italic_label)
    for i, v in enumerate(model_data["da"]):
        ws.write(fc_start + 4, 5 + i, v, num_decimal_fmt)

    # EBIT
    ws.write(fc_start + 5, 3, "Operating profit (EBIT)", label_fmt)
    ws.write(fc_start + 5, 4, model_data["ebit"][0] * 0.8, num_decimal_fmt)
    for i, v in enumerate(model_data["ebit"]):
        ws.write(fc_start + 5, 5 + i, v, num_decimal_fmt)

    # Tax Rate
    ws.write(fc_start + 6, 3, "Tax Rate", italic_label)
    for i in range(n):
        ws.write(fc_start + 6, 5 + i, model_data["tax_rate"], pct_fmt)
    ws.write(fc_start + 6, 4, model_data["tax_rate"], pct_fmt)

    # NOPAT
    ws.write(fc_start + 8, 3, "NOPAT", label_fmt)
    ws.write(fc_start + 8, 4, model_data["nopat"][0] * 0.75, num_decimal_fmt)
    for i, v in enumerate(model_data["nopat"]):
        ws.write(fc_start + 8, 5 + i, v, num_decimal_fmt)

    # Change in WC
    ws.write(fc_start + 9, 3, "Change in (WC)", italic_label)
    for i, v in enumerate(model_data["change_wc"]):
        ws.write(fc_start + 9, 5 + i, v, num_decimal_fmt)

    # WC
    ws.write(fc_start + 10, 3, "WC", italic_label)
    ws.write(fc_start + 10, 4, 0, num_decimal_fmt)
    for i, v in enumerate(model_data["wc"]):
        ws.write(fc_start + 10, 5 + i, v, num_decimal_fmt)

    # Operating Cash Flow
    ws.write(fc_start + 11, 3, "Operating Cash Flow", label_fmt)
    for i, v in enumerate(model_data["operating_cf"]):
        ws.write(fc_start + 11, 5 + i, v, num_decimal_fmt)

    # Capex
    ws.write(fc_start + 12, 3, "Capex", italic_label)
    for i, v in enumerate(model_data["capex"]):
        ws.write(fc_start + 12, 5 + i, v, num_decimal_fmt)

    # Unlevered FCF
    ws.write(fc_start + 13, 3, "Unlevered FCF", label_fmt)
    for i, v in enumerate(model_data["unlevered_fcf"]):
        ws.write(fc_start + 13, 5 + i, v, num_decimal_fmt)

    # Discounted FCF (bold top border)
    ws.write(fc_start + 15, 3, "Discounted FCF", italic_label)
    for i, v in enumerate(model_data["discounted_fcf"]):
        ws.write(fc_start + 15, 5 + i, v, bold_num_fmt)

    # EV / Net Debt / Equity Value
    ws.write(fc_start + 16, 3, "Enterprise Value", label_fmt)
    ws.write(fc_start + 16, 4, model_data["enterprise_value"], total_fmt)

    ws.write(fc_start + 17, 3, "Net Debt", italic_label)
    ws.write(fc_start + 17, 4, model_data["net_debt"], num_fmt)

    ws.write(fc_start + 18, 3, "Equity Value", label_fmt)
    ws.write(fc_start + 18, 4, model_data["equity_value"], total_fmt)

    # ---------- SENSITIVITY TABLES ----------
    sens_start = fc_start + 22
    wacc_values = model_data.get("sens_wacc", [])
    tg_values = model_data.get("sens_tg", [])
    wacc_matrix = model_data.get("sens_wacc_matrix", [])
    price_matrix = model_data.get("sens_price_matrix", [])

    # WACC sensitivity
    if wacc_values and tg_values and wacc_matrix:
        ws.merge_range(sens_start, 3, sens_start,
                       3 + len(wacc_values), "Wacc", section_center_fmt)
        ws.write(sens_start + 1, 3, "", label_fmt)
        for i, w in enumerate(wacc_values):
            ws.write(sens_start + 1, 4 + i, w, pct_fmt)
        for i, t in enumerate(tg_values):
            ws.write(sens_start + 2 + i, 3, t, pct_fmt)
            for j, val in enumerate(wacc_matrix[i]):
                ws.write(sens_start + 2 + i, 4 + j, val, num_decimal_fmt)

    # Price Table sensitivity
    price_start = sens_start + 2 + len(tg_values) + 2
    if wacc_values and tg_values and price_matrix:
        ws.merge_range(price_start, 3, price_start,
                       3 + len(wacc_values), "Price Table", section_center_fmt)
        ws.write(price_start + 1, 3, "", label_fmt)
        for i, w in enumerate(wacc_values):
            ws.write(price_start + 1, 4 + i, w, pct_fmt)
        for i, t in enumerate(tg_values):
            ws.write(price_start + 2 + i, 3, t, pct_fmt)
            for j, val in enumerate(price_matrix[i]):
                ws.write(price_start + 2 + i, 4 + j, val, num_decimal_fmt)

    wb.close()
    return output.getvalue()