import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

st.set_page_config(layout="wide", page_title="Reconciliation Engine")

st.title("📘 Multi-Source Excel Reconciliation Engine – FINAL STRICT INDEX VERSION")
st.write("Upload all the required files as instructed below.")

# ============================================================
# Helper Functions
# ============================================================

def extract_month_year(date_value):
    """Convert Excel date to datetime if needed, return (month, year)."""
    if pd.isna(date_value):
        return None, None

    dt = pd.to_datetime(date_value, errors="coerce")
    if dt is None:
        return None, None

    return dt.month, dt.year

def is_quarter_end(dt):
    """Check if a date is a quarter-end."""
    return (dt.month, dt.day) in [(3, 31), (6, 30), (9, 30), (12, 31)]


# ============================================================
# 1️⃣ MAIN RECONCILIATION SHEET
# ============================================================

main_file = st.file_uploader("Upload MAIN RECONCILIATION EXCEL", type=["xlsx"])

if main_file:
    main_xls = pd.ExcelFile(main_file)

    metadata = pd.read_excel(main_xls, "metadata")
    exclusion_stocks = pd.read_excel(main_xls, "exclusion_stocks")
    closing_stocks_existing = pd.read_excel(main_xls, "closing_stocks")
    dividends_existing = pd.read_excel(main_xls, "dividends")
    holdings_existing = pd.read_excel(main_xls, "holdings")

    # ⭐ NEW CASH BALANCE PART (read from main file)
    cash_balance_existing = pd.read_excel(main_xls, "cash_balance")

    st.success("Main reconciliation file loaded successfully.")

    # ============================================================
    # 2️⃣ CLOSING STOCKS
    # ============================================================
    closing_file = st.file_uploader("Upload CLOSING STOCKS Excel", type=["xlsx"])
    if closing_file:
        closing_stocks = pd.read_excel(closing_file)

    # ============================================================
    # 3️⃣ DIVIDENDS
    # ============================================================
    dividends_file = st.file_uploader("Upload DIVIDENDS Excel", type=["xlsx"])
    if dividends_file:
        dividends = pd.read_excel(dividends_file)

    # ============================================================
    # 4️⃣ HOLDINGS
    # ============================================================
    holdings_file = st.file_uploader("Upload HOLDINGS Excel", type=["xlsx"])
    if holdings_file:
        holdings = pd.read_excel(holdings_file)

    # ============================================================
    # ⭐ 5️⃣ NEW CASH BALANCE UPLOADER
    # ============================================================
    st.subheader("Upload CASH BALANCE Excel")

    cash_balance_file = st.file_uploader("Upload CASH BALANCE Excel", type=["xlsx"])
    if cash_balance_file:
        cash_balance = pd.read_excel(cash_balance_file)
    else:
        cash_balance = cash_balance_existing


    # ============================================================
    # 6️⃣ FUND INFUSION → metadata Column F
    # ============================================================
    st.subheader("FUND INFUSION → Populate metadata!F")

    fund_file = st.file_uploader("Upload FUND INFUSION Excel", type=["xlsx"])

    if fund_file:
        fund_df = pd.read_excel(fund_file)

        infusion_month_cols = fund_df.columns[3:]
        infusion_header_dates = {
            col: pd.to_datetime(col, errors="coerce")
            for col in infusion_month_cols
        }

        metadata["amount_added"] = 0

        for idx, row in metadata.iterrows():
            client_main = row["client_id"]    # NOW MUST BE COLUMN C
            main_dt = pd.to_datetime(row["closing_end_date"])
            m_month, m_year = extract_month_year(main_dt)

            fund_row = fund_df[fund_df.iloc[:, 0] == client_main]

            if fund_row.empty:
                continue

            for col, hdr_date in infusion_header_dates.items():
                if hdr_date is None:
                    continue

                h_month, h_year = extract_month_year(hdr_date)

                if (h_month == m_month) and (h_year == m_year):
                    val = fund_row[col].values[0]
                    metadata.at[idx, "amount_added"] = val

        st.success("Column F (amount_added) populated using strict index matching.")


    # ============================================================
    # 7️⃣ QUARTERLY SETTLEMENT → metadata Column I
    # ============================================================
    st.subheader("QUARTERLY SETTLEMENT → Populate metadata!I")

    qset_file = st.file_uploader("Upload QUARTERLY SETTLEMENT Excel", type=["xlsx"])

    if qset_file:
        qdf = pd.read_excel(qset_file)

        q_month_cols = qdf.columns[3:]
        q_header_dates = {col: pd.to_datetime(col, errors="coerce") for col in q_month_cols}

        metadata["quarterly_settlement"] = ""

        for idx, row in metadata.iterrows():
            client_main = row["client_id"]
            main_dt = pd.to_datetime(row["closing_end_date"])

            if not is_quarter_end(main_dt):
                continue

            if main_dt.month == 3:
                q_months = [1, 2, 3]
            elif main_dt.month == 6:
                q_months = [4, 5, 6]
            elif main_dt.month == 9:
                q_months = [7, 8, 9]
            else:
                q_months = [10, 11, 12]

            q_row = qdf[qdf.iloc[:, 1] == client_main]

            if q_row.empty:
                continue

            total_val = 0

            for col, hdr_date in q_header_dates.items():
                if hdr_date is None:
                    continue

                if (hdr_date.month in q_months) and (hdr_date.year == main_dt.year):
                    val = q_row[col].values[0]
                    total_val += val

            metadata.at[idx, "quarterly_settlement"] = total_val

        st.success("Column I populated successfully.")


    # ============================================================
    # 8️⃣ PART 2 — EXCLUSION STOCK PRICE UPLOAD & N/O CALCULATION
    # ============================================================
    st.subheader("📉 Upload EXCLUSION STOCK PRICE file (ISIN + Price)")

    excl_price_file = st.file_uploader("Upload EXCLUSION STOCK PRICE Excel", type=["xlsx"])

    if excl_price_file:
        excl_price_df = pd.read_excel(excl_price_file)

        price_map = dict(zip(excl_price_df.iloc[:, 0], excl_price_df.iloc[:, 1]))

        col_M_vals = []
        col_M_calc_vals = []

        for idx, row in exclusion_stocks.iterrows():
            isin = row.iloc[3]
            col_I_val = row.iloc[8]

            if pd.isna(col_I_val) or col_I_val == "":
                price = price_map.get(isin, 0)
                col_M_vals.append(price)
                col_M_calc_vals.append(price)
            else:
                col_M_vals.append("")
                col_M_calc_vals.append(0)

        exclusion_stocks.iloc[:, 12] = col_M_vals
        exclusion_stocks.iloc[:, 13] = exclusion_stocks.iloc[:, 5] * col_M_calc_vals
        exclusion_stocks.iloc[:, 14] = exclusion_stocks.iloc[:, 13] - exclusion_stocks.iloc[:, 7]

        st.success("Exclusion Stock Prices updated with conditional M + N/O recalculated.")

    # ============================================================
    # 9️⃣ metadata blanks → 0
    # ============================================================
    cols_to_fill = [
        "amount_added",
        "total_capital_on_the_date_of_withdrawal",
        "actual_withdrawal",
        "quarterly_settlement"
    ]

    for col in cols_to_fill:
        if col in metadata.columns:
            metadata[col] = metadata[col].fillna(0)


    # ============================================================
    # 🔟 FINAL CHECK FILE (NO metadata changes except Column E)
    # ============================================================
    st.subheader("Upload FINAL CHECK FILE (Client Code | Date | Amount)")

    final_check_file = st.file_uploader("Upload FINAL CHECK Excel", type=["xlsx"])

    if final_check_file:
        final_df = pd.read_excel(final_check_file)

        st.write("### 🔍 Final Check Results (PASS / FLAG)")

        results = []

        for idx2, row2 in final_df.iterrows():

            client = row2.iloc[0]
            date_given = row2.iloc[1]
            amt_given = row2.iloc[2]

            meta_row = metadata[metadata["client_id"] == client]

            if meta_row.empty:
                results.append({
                    "client_id": client,
                    "status": "CLIENT NOT FOUND",
                    "difference": None
                })
                continue

            meta_idx = meta_row.index[0]

            metadata.iat[meta_idx, 4] = date_given

            amt_stored = metadata.at[meta_idx, "amount_added"]

            if amt_stored == amt_given:
                results.append({
                    "client_id": client,
                    "date": date_given,
                    "amount_given": amt_given,
                    "amount_in_metadata": amt_stored,
                    "status": "PASS",
                    "difference": 0
                })
            else:
                results.append({
                    "client_id": client,
                    "date": date_given,
                    "amount_given": amt_given,
                    "amount_in_metadata": amt_stored,
                    "status": "FLAG ❗",
                    "difference": amt_given - amt_stored
                })

        results_df = pd.DataFrame(results)
        st.dataframe(results_df, use_container_width=True)

        st.success("Final check completed. Results shown above.")


    # ============================================================
    # 1️⃣1️⃣ DOWNLOAD FINAL OUTPUT
    # ============================================================
    if st.button("📥 DOWNLOAD FINAL RECONCILED FILE"):

        output = pd.ExcelWriter("Reconciled_Final.xlsx", engine="xlsxwriter")

        metadata.to_excel(output, sheet_name="metadata", index=False)
        exclusion_stocks.to_excel(output, sheet_name="exclusion_stocks", index=False)

        if closing_file:
            closing_stocks.to_excel(output, sheet_name="closing_stocks", index=False)
        else:
            closing_stocks_existing.to_excel(output, sheet_name="closing_stocks", index=False)

        if dividends_file:
            dividends.to_excel(output, sheet_name="dividends", index=False)
        else:
            dividends_existing.to_excel(output, sheet_name="dividends", index=False)

        if holdings_file:
            holdings.to_excel(output, sheet_name="holdings", index=False)
        else:
            holdings_existing.to_excel(output, sheet_name="holdings", index=False)

        # ⭐ NEW CASH BALANCE EXPORT
        cash_balance.to_excel(output, sheet_name="cash_balance", index=False)

        output.close()

        with open("Reconciled_Final.xlsx", "rb") as f:
            st.download_button("⬇️ DOWNLOAD FINAL RECONCILED FILE", f, file_name="Reconciled_Final.xlsx")

        st.success("🎉 FINAL FILE CREATED SUCCESSFULLY!")
