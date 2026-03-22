import io
import re

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Kurier Analysis", layout="centered")

st.title("Kurier Analysis")


month_map = {
    "jan": "jan", "january": "jan",
    "feb": "feb", "february": "feb", "febrauray": "feb", "febraury": "feb",
    "mar": "mar", "march": "mar",
    "apr": "apr", "april": "apr",
    "may": "may",
    "jun": "jun", "june": "jun",
    "jul": "jul", "july": "jul",
    "aug": "aug", "august": "aug",
    "sep": "sep", "september": "sep",
    "oct": "oct", "october": "oct",
    "nov": "nov", "november": "nov",
    "dec": "dec", "december": "dec",
}


def is_valid_year(y):
    y = int(y)
    if 2020 <= y <= 2030:
        return True
    if 20 <= y <= 30:
        return True
    return False


def is_valid_month_year_sheet(name):
    name = name.strip().lower()

    if not re.fullmatch(r"[a-z0-9\s]+", name):
        return False

    parts = name.split()

    if len(parts) != 2:
        return False

    part1, part2 = parts

    if part1 in month_map and part2.isdigit() and is_valid_year(part2):
        return True

    if part2 in month_map and part1.isdigit() and is_valid_year(part1):
        return True

    return False


def load_all_sheets(excel_url, uploaded_file):
    if uploaded_file is not None:
        file_bytes = io.BytesIO(uploaded_file.getvalue())
        return pd.read_excel(file_bytes, sheet_name=None, engine="openpyxl")

    if not excel_url:
        raise ValueError("Please paste a Google Sheet / Excel URL or upload an Excel file.")

    excel_source = excel_url.strip()
    if "docs.google.com/spreadsheets" in excel_source and "/edit" in excel_source:
        excel_source = excel_source.split("/edit")[0] + "/export?format=xlsx"

    return pd.read_excel(excel_source, sheet_name=None, engine="openpyxl")


def get_monthly_sheets_filtered(all_sheets):
    return {
        name: df for name, df in all_sheets.items()
        if is_valid_month_year_sheet(name)
    }


def get_date_wise_packet_count(sheet_name, monthly_sheets_filtered):
    df = monthly_sheets_filtered[sheet_name].copy()

    df.columns = df.columns.astype(str).str.strip().str.upper()
    df = df.loc[:, ~df.columns.duplicated()]

    if "DATE" in df.columns:
        date_col = "DATE"
    elif "AHU" in df.columns:
        date_col = "AHU"
    else:
        date_col = df.columns[0]

    df["DATE"] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
    df = df[df["DATE"].notna()]

    if "AWB NO." not in df.columns:
        return pd.DataFrame()

    result = (
        df.groupby("DATE")["AWB NO."]
        .count()
        .reset_index(name="Packet Count")
        .sort_values("DATE", ascending=True)
    )

    result["DATE"] = result["DATE"].dt.date
    return result


def get_packets_booked_per_sender(sheet_name, monthly_sheets_filtered):
    df = monthly_sheets_filtered[sheet_name].copy()

    df.columns = df.columns.astype(str).str.strip().str.upper()
    df = df.loc[:, ~df.columns.duplicated()]

    if "AWB NO." not in df.columns:
        return pd.DataFrame()

    if "SENDER NAME" in df.columns:
        sender_col = "SENDER NAME"
    elif "MODE" in df.columns:
        sender_col = "MODE"
    elif "BILLING DETAILS" in df.columns:
        sender_col = "BILLING DETAILS"
    else:
        return pd.DataFrame()

    result = (
        df.groupby(sender_col)["AWB NO."]
        .count()
        .reset_index(name="Packet Count")
        .sort_values("Packet Count", ascending=False)
    )

    return result.rename(columns={sender_col: "SENDER NAME"})


def get_sender_wise_packets_for_each_date(sheet_name, monthly_sheets_filtered):
    df = monthly_sheets_filtered[sheet_name].copy()

    df.columns = df.columns.astype(str).str.strip().str.upper()
    df = df.loc[:, ~df.columns.duplicated()]

    if "DATE" in df.columns:
        date_col = "DATE"
    elif "AHU" in df.columns:
        date_col = "AHU"
    else:
        date_col = df.columns[0]

    df["DATE"] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
    df = df[df["DATE"].notna()]

    required_cols = ["DATE", "SENDER NAME", "AWB NO."]
    if not all(col in df.columns for col in required_cols):
        return pd.DataFrame()

    result = (
        df.groupby(["DATE", "SENDER NAME"])["AWB NO."]
        .count()
        .reset_index(name="Packet Count")
        .sort_values(["DATE", "Packet Count", "SENDER NAME"], ascending=[True, False, True])
    )

    result["DATE"] = result["DATE"].dt.date
    return result


def get_packets_booked_per_mode(sheet_name, monthly_sheets_filtered):
    df = monthly_sheets_filtered[sheet_name].copy()

    df.columns = df.columns.astype(str).str.strip().str.upper()
    df = df.loc[:, ~df.columns.duplicated()]

    if "AWB NO." not in df.columns or "MODE" not in df.columns:
        return pd.DataFrame()

    df["MODE"] = df["MODE"].astype(str).str.upper().str.strip()

    return (
        df.groupby("MODE")["AWB NO."]
        .count()
        .reset_index(name="Packet Count")
        .sort_values("Packet Count", ascending=False)
    )


def get_payment_received_per_month(sheet_name, monthly_sheets_filtered):
    df = monthly_sheets_filtered[sheet_name].copy()

    df.columns = df.columns.astype(str).str.strip().str.upper()
    df = df.loc[:, ~df.columns.duplicated()]

    required_cols = ["CREDIT OR CASH", "AMOUNT", "SENDER NAME"]
    if not all(col in df.columns for col in required_cols):
        return None

    df["CREDIT OR CASH"] = df["CREDIT OR CASH"].astype(str).str.upper().str.strip()
    df["AMOUNT"] = df["AMOUNT"].astype(str).str.strip()
    df["AMOUNT_NUM"] = pd.to_numeric(df["AMOUNT"], errors="coerce")

    cash_total = df[df["CREDIT OR CASH"] == "CASH"]["AMOUNT_NUM"].sum()
    upi_total = df[df["CREDIT OR CASH"] == "UPI"]["AMOUNT_NUM"].sum()

    credit_numeric = df[df["CREDIT OR CASH"] == "CREDIT"]
    credit_total = credit_numeric["AMOUNT_NUM"].sum()

    credit_monthly = df[
        (df["CREDIT OR CASH"] == "CREDIT") &
        (df["AMOUNT"].str.lower() == "monthly")
    ]

    monthly_sender_count = (
        credit_monthly.groupby("SENDER NAME")["AMOUNT"]
        .count()
        .reset_index(name="Monthly Count")
        .sort_values("Monthly Count", ascending=False)
    )

    return {
        "cash_total": int(cash_total) if pd.notna(cash_total) else 0,
        "upi_total": int(upi_total) if pd.notna(upi_total) else 0,
        "credit_total": int(credit_total) if pd.notna(credit_total) else 0,
        "monthly_sender_count": monthly_sender_count,
    }


if "monthly_sheets_filtered" not in st.session_state:
    st.session_state.monthly_sheets_filtered = None

if "confirmed_month" not in st.session_state:
    st.session_state.confirmed_month = None

if "confirmed_report" not in st.session_state:
    st.session_state.confirmed_report = None


excel_url = st.text_input("Paste Google Sheet / Excel URL")
uploaded_file = st.file_uploader("Or upload Excel file", type=["xlsx", "xls"])

if st.button("Analyze", key="source_analyze"):
    try:
        all_sheets = load_all_sheets(excel_url, uploaded_file)
        monthly_sheets_filtered = get_monthly_sheets_filtered(all_sheets)

        if not monthly_sheets_filtered:
            st.session_state.monthly_sheets_filtered = None
            st.session_state.confirmed_month = None
            st.session_state.confirmed_report = None
            st.error("No valid month sheets were found in the file.")
        else:
            st.session_state.monthly_sheets_filtered = monthly_sheets_filtered
            st.session_state.confirmed_month = None
            st.session_state.confirmed_report = None
            st.success("File analyzed. Select a month below.")
    except Exception as exc:
        st.session_state.monthly_sheets_filtered = None
        st.session_state.confirmed_month = None
        st.session_state.confirmed_report = None
        st.error(f"Could not read the sheet: {exc}")


if st.session_state.monthly_sheets_filtered:
    month_options = list(st.session_state.monthly_sheets_filtered.keys())
    selected_month = st.selectbox("Select Month", month_options)

    if st.button("Analyze", key="month_analyze"):
        st.session_state.confirmed_month = selected_month
        st.session_state.confirmed_report = None


if st.session_state.confirmed_month:
    report_options = [
        "date wise packet count",
        "packets booked per sender",
        "packets booked per mode",
        "payment received per month",
    ]

    selected_report = st.selectbox("Select Report", report_options)

    if st.button("Show Result"):
        st.session_state.confirmed_report = selected_report


if st.session_state.confirmed_month and st.session_state.confirmed_report:
    month_name = st.session_state.confirmed_month
    report_name = st.session_state.confirmed_report
    monthly_sheets_filtered = st.session_state.monthly_sheets_filtered

    if report_name == "date wise packet count":
        st.subheader("Date Wise Packet Count")
        result = get_date_wise_packet_count(month_name, monthly_sheets_filtered)
        if result.empty:
            st.warning("Required columns were not found for this report.")
        else:
            sender_date_result = get_sender_wise_packets_for_each_date(month_name, monthly_sheets_filtered)
            event = st.dataframe(
                result,
                use_container_width=True,
                hide_index=True,
                on_select="rerun",
                selection_mode="single-row",
                key="date_wise_table",
            )

            if not sender_date_result.empty:
                selected_rows = event.selection.rows
                if selected_rows:
                    selected_row = result.iloc[selected_rows[0]]
                    selected_date = selected_row["DATE"]
                    day_result = sender_date_result[sender_date_result["DATE"] == selected_date]

                    st.write(f"Sender-wise packet count for {selected_date}")
                    st.dataframe(
                        day_result[["SENDER NAME", "Packet Count"]],
                        use_container_width=True,
                        hide_index=True,
                    )
                else:
                    st.caption("Click a date row in the table to see sender-wise packet count for that day.")

    elif report_name == "packets booked per sender":
        st.subheader("Packets Booked Per Sender")
        result = get_packets_booked_per_sender(month_name, monthly_sheets_filtered)
        if result.empty:
            st.warning("Required columns were not found for this report.")
        else:
            st.metric("Total Packets Booked", int(result["Packet Count"].sum()))
            st.dataframe(result, use_container_width=True)

    elif report_name == "packets booked per mode":
        st.subheader("Packets Booked Per Mode")
        result = get_packets_booked_per_mode(month_name, monthly_sheets_filtered)
        if result.empty:
            st.warning("Required columns were not found for this report.")
        else:
            st.dataframe(result, use_container_width=True)

    elif report_name == "payment received per month":
        st.subheader("Payment Received Per Month")
        result = get_payment_received_per_month(month_name, monthly_sheets_filtered)

        if result is None:
            st.warning("Required columns were not found for this report.")
        else:
            col1, col2, col3 = st.columns(3)
            col1.metric("Cash Booking", f"Rs {result['cash_total']}")
            col2.metric("UPI Booking", f"Rs {result['upi_total']}")
            col3.metric("Credit Booking", f"Rs {result['credit_total']}")

            st.write("Monthly Credit (per Sender)")
            if result["monthly_sender_count"].empty:
                st.info("No monthly credit entries")
            else:
                st.dataframe(result["monthly_sender_count"], use_container_width=True)
