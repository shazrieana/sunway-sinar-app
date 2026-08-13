import streamlit as st
import pandas as pd
from supabase import create_client, Client
from datetime import datetime
from dateutil.relativedelta import relativedelta
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. DATABASE & APP INITIALIZATION
# ==========================================
SUPABASE_URL = st.secrets["SUPABASE_URL"]
SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

st.set_page_config(page_title="Neighbourhood Maintenance Portal", layout="wide")
st.title("🏘️ Neighbourhood Maintenance Fee Management System")

# Define payment year scope
SELECTED_YEAR = 2024
MONTHS = [datetime(SELECTED_YEAR, m, 1) for m in range(1, 13)]

# ==========================================
# 2. HELPER EXPORT & DATA FUNCTIONS
# ==========================================
def get_unit_payments(unit_id):
    """Retrieves all payment records for a specific unit."""
    response = supabase.table("payments").select("*").eq("unit_id", unit_id).execute()
    return response.data

def get_paid_months_set(unit_id):
    """Calculates all distinct YYYY-MM months paid by a unit."""
    records = get_unit_payments(unit_id)
    paid_months = set()
    for rec in records:
        curr = datetime.strptime(rec["start_month"], "%Y-%m-%d")
        end = datetime.strptime(rec["end_month"], "%Y-%m-%d")
        while curr <= end:
            paid_months.add(curr.strftime("%Y-%m"))
            curr += relativedelta(months=1)
    return paid_months

def generate_monthly_excel(records, month_name, year):
    """Generates a styled Excel workbook for a single month's collection log."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{month_name} {year} Log"

    # Header Title Banner (Merged A1:I1)
    ws.merge_cells("A1:I1")
    title_cell = ws["A1"]
    title_cell.value = f"NEIGHBOURHOOD MAINTENANCE FEE COLLECTION - {month_name.upper()} {year}"
    title_cell.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    headers = ["No", "Collection Date", "OR NO.", "Unit ID", "Collector Name", "Payment Method", "Amount Paid (RM)", "Coverage Start", "Coverage End"]
    ws.append([]) # Blank row
    ws.append(headers)

    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    ws.row_dimensions[3].height = 22
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    total_amount = 0.0
    for idx, rec in enumerate(records, start=1):
        row_num = idx + 3
        amt = float(rec["amount_paid"])
        total_amount += amt
        
        row_data = [
            idx,
            rec["created_at"][:10],
            rec.get("or_no", "-"),
            rec["unit_id"],
            rec.get("collector_name", "-"),
            rec.get("payment_method", "CASH"),
            amt,
            datetime.strptime(rec["start_month"], "%Y-%m-%d").strftime("%b %Y"),
            datetime.strptime(rec["end_month"], "%Y-%m-%d").strftime("%b %Y")
        ]
        ws.append(row_data)
        ws.row_dimensions[row_num].height = 20
        
        for c_idx in range(1, 10):
            cell = ws.cell(row=row_num, column=c_idx)
            cell.border = thin_border
            cell.font = Font(name="Arial", size=10)
            if c_idx == 7:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c_idx in [3, 4, 5]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # Add Summary Total Row
    tot_row = len(records) + 4
    ws.cell(row=tot_row, column=6, value="TOTAL COLLECTED:").font = Font(name="Arial", size=10, bold=True)
    ws.cell(row=tot_row, column=6).alignment = Alignment(horizontal="right", vertical="center")
    tot_cell = ws.cell(row=tot_row, column=7, value=f"=SUM(G4:G{tot_row-1})")
    tot_cell.font = Font(name="Arial", size=10, bold=True)
    tot_cell.number_format = '#,##0.00'

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def generate_annual_excel(matrix_df, year):
    """Generates a styled Excel matrix workbook for an entire year."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Summary {year}"

    ws.merge_cells("A1:M1")
    t_cell = ws["A1"]
    t_cell.value = f"NEIGHBOURHOOD MAINTENANCE FEE - ANNUAL SUMMARY ({year})"
    t_cell.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    t_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    t_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    headers = list(matrix_df.columns)
    ws.append([])
    ws.append(headers)

    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    paid_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    unpaid_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    paid_font = Font(name="Arial", size=9, bold=True, color="375623")
    unpaid_font = Font(name="Arial", size=9, bold=True, color="C65911")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    ws.row_dimensions[3].height = 22
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for idx, row in matrix_df.iterrows():
        r_num = idx + 4
        ws.append(list(row))
        ws.row_dimensions[r_num].height = 20
        ws.cell(row=r_num, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r_num, column=1).font = Font(name="Arial", size=10, bold=True)
        ws.cell(row=r_num, column=1).border = thin_border

        for c_idx, val in enumerate(row[1:], start=2):
            cell = ws.cell(row=r_num, column=c_idx)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if val == "PAID":
                cell.fill = paid_fill
                cell.font = paid_font
            else:
                cell.fill = unpaid_fill
                cell.font = unpaid_font

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = 11

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# 3. INTERFACE TABS
# ==========================================
tab_entry, tab_monthly, tab_annual = st.tabs([
    "📝 Record Payment", 
    "📅 Monthly Collection Exports", 
    "📊 Annual Payment Summary"
])

# --- TAB 1: RECORD PAYMENT ---
with tab_entry:
    st.header("Log Maintenance Collection")
    
    col1, col2 = st.columns(2)
    with col1:
        block = st.selectbox("Select Block", ["S1", "S2", "S3"])
        floor = st.selectbox("Select Floor Level", [0, 1, 2, 3, 4], format_func=lambda x: "Ground Floor" if x == 0 else f"Level {x}")
    
    with col2:
        units_resp = supabase.table("units").select("unit_id").eq("block", block).eq("floor_level", floor).execute()
        unit_list = [u["unit_id"] for u in units_resp.data] if units_resp.data else []
        selected_unit = st.selectbox("Select Unit Number", unit_list) if unit_list else None
        collector = st.text_input("Collector Name", value="Committee Member")

    st.markdown("---")
    
    if selected_unit:
        paid_set = get_paid_months_set(selected_unit)
        st.subheader(f"Status for Unit: {selected_unit}")
        
        status_cols = st.columns(6)
        for idx, m_date in enumerate(MONTHS):
            m_str = m_date.strftime("%Y-%m")
            col_idx = idx % 6
            if m_str in paid_set:
                status_cols[col_idx].success(f"✅ {m_date.strftime('%b')}")
            else:
                status_cols[col_idx].error(f"❌ {m_date.strftime('%b')}")

        st.markdown("---")
        st.subheader("New Entry")
        
        f_col1, f_col2 = st.columns(2)
        with f_col1:
            or_no = st.text_input("Official Receipt (OR NO.)", placeholder="e.g. 2239")
            payment_method = st.selectbox("Payment Method", ["CASH", "Online Transfer", "DuitNow QR", "CHEQUE"])
        with f_col2:
            start_m = st.selectbox("Coverage Start Month", MONTHS, format_func=lambda x: x.strftime("%b %Y"))
            end_m = st.selectbox("Coverage End Month", MONTHS, format_func=lambda x: x.strftime("%b %Y"), index=0)
        
        amount = st.number_input("Total Amount Received (RM)", min_value=0.0, step=35.0, value=35.0)

        if st.button("Submit Payment", type="primary"):
            if start_m > end_m:
                st.error("Error: Start Month cannot be later than End Month.")
            else:
                payment_payload = {
                    "unit_id": selected_unit,
                    "or_no": or_no,
                    "collector_name": collector,
                    "payment_method": payment_method,
                    "amount_paid": amount,
                    "start_month": start_m.strftime("%Y-%m-%d"),
                    "end_month": end_m.strftime("%Y-%m-%d")
                }
                supabase.table("payments").insert(payment_payload).execute()
                st.success(f"Payment recorded successfully for {selected_unit}!")
                st.rerun()

# --- TAB 2: MONTHLY COLLECTION EXPORT ---
with tab_monthly:
    st.header("Monthly Collection Logs")
    
    m_col1, m_col2, m_col3, m_col4 = st.columns(4)
    with m_col1:
        selected_month_num = st.selectbox("Select Month", list(range(1, 13)), format_func=lambda x: datetime(2024, x, 1).strftime("%B"))
    with m_col2:
        selected_year = st.number_input("Select Year", min_value=2023, max_value=2030, value=2024)
    with m_col3:
        sort_by_m = st.selectbox("Sort By", ["Collection Date", "Receipt Number (OR NO.)", "Unit ID", "Block"])
    with m_col4:
        sort_order_m = st.selectbox("Order", ["Ascending (A-Z / Oldest First)", "Descending (Z-A / Newest First)"])

    m_start_str = f"{selected_year}-{selected_month_num:02d}-01"
    next_m = datetime(selected_year, selected_month_num, 1) + relativedelta(months=1)
    m_end_str = next_m.strftime("%Y-%m-%d")

    records = supabase.table("payments").select("*").gte("created_at", m_start_str).lt("created_at", m_end_str).execute().data
    
    if len(records) == 0:
        st.info("No payments were logged for this selected month/year.")
    else:
        st.write(f"Found **{len(records)}** payment logs.")
        df_logs = pd.DataFrame(records)[["created_at", "or_no", "unit_id", "collector_name", "payment_method", "amount_paid", "start_month", "end_month"]]
        
        # Add Block column for sorting
        df_logs["Block"] = df_logs["unit_id"].apply(lambda x: x.split("-")[0] if "-" in str(x) else "")
        
        # Mapping sort selections to columns
        sort_map = {
            "Collection Date": "created_at",
            "Receipt Number (OR NO.)": "or_no",
            "Unit ID": "unit_id",
            "Block": "Block"
        }
        
        is_asc = True if "Ascending" in sort_order_m else False
        df_logs = df_logs.sort_values(by=sort_map[sort_by_m], ascending=is_asc)
        
        st.dataframe(df_logs.drop(columns=["Block"]), use_container_width=True)
        
        # Pass sorted records to Excel generator
        sorted_records = df_logs.to_dict("records")
        m_name = datetime(2024, selected_month_num, 1).strftime("%B")
        excel_file = generate_monthly_excel(sorted_records, m_name, selected_year)
        
        st.download_button(
            label=f"📥 Download {m_name} {selected_year} Collection Log (.xlsx)",
            data=excel_file,
            file_name=f"Collection_Log_{m_name}_{selected_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- TAB 3: ANNUAL SUMMARY EXPORT ---
with tab_annual:
    st.header("Annual Payment Matrix")
    
    a_col1, a_col2, a_col3 = st.columns(3)
    with a_col1:
        ann_year = st.number_input("Select Annual Year Scope", min_value=2023, max_value=2030, value=2024, key="annual_year")
    with a_col2:
        sort_by_a = st.selectbox("Sort By", ["Unit ID", "Block", "Floor Level"])
    with a_col3:
        sort_order_a = st.selectbox("Order", ["Ascending (A-Z / Lowest First)", "Descending (Z-A / Highest First)"], key="ann_order")
    
    all_units = supabase.table("units").select("unit_id", "block", "floor_level").execute().data
    all_payments = supabase.table("payments").select("*").execute().data
    
    payment_map = {}
    for p in all_payments:
        u = p["unit_id"]
        if u not in payment_map:
            payment_map[u] = set()
        c = datetime.strptime(p["start_month"], "%Y-%m-%d")
        e = datetime.strptime(p["end_month"], "%Y-%m-%d")
        while c <= e:
            payment_map[u].add(c.strftime("%Y-%m"))
            c += relativedelta(months=1)
            
    months_list = [datetime(ann_year, m, 1) for m in range(1, 13)]
    matrix_data = []
    for u_obj in all_units:
        u_id = u_obj["unit_id"]
        row = {
            "Unit ID": u_id,
            "Block": u_obj["block"],
            "Floor": u_obj["floor_level"]
        }
        u_paid = payment_map.get(u_id, set())
        for m_date in months_list:
            m_str = m_date.strftime("%Y-%m")
            row[m_date.strftime("%b")] = "PAID" if m_str in u_paid else "UNPAID"
        matrix_data.append(row)
        
    df_matrix = pd.DataFrame(matrix_data)
    
    # Sort Matrix Data
    is_asc_a = True if "Ascending" in sort_order_a else False
    if sort_by_a == "Block":
        df_matrix = df_matrix.sort_values(by=["Block", "Unit ID"], ascending=[is_asc_a, True])
    elif sort_by_a == "Floor Level":
        df_matrix = df_matrix.sort_values(by=["Floor", "Unit ID"], ascending=[is_asc_a, True])
    else:
        df_matrix = df_matrix.sort_values(by="Unit ID", ascending=is_asc_a)
        
    # Drop helper sorting columns before displaying and exporting
    df_display = df_matrix.drop(columns=["Block", "Floor"])
    st.dataframe(df_display, use_container_width=True, height=400)
    
    ann_excel = generate_annual_excel(df_display, ann_year)
    st.download_button(
        label=f"📥 Download Full {ann_year} Annual Summary (.xlsx)",
        data=ann_excel,
        file_name=f"Maintenance_Fee_Summary_{ann_year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
