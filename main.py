import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import io
from typing import Tuple, List, Dict, Optional
import time

st.set_page_config(
    page_title="Medical Reimbursement Enrichment",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🏥 Medical Reimbursement Data Enrichment")
st.markdown("Enrich Medical Reimbursement data with Payroll Area information (Monthly Quarter)")

MONTH_NAMES = {
    1: "Bulan 1 (Beginning)",
    2: "Bulan 2 (Middle)",
    3: "Bulan 3 (Ending)"
}


def get_worksheet_names(file_bytes: bytes) -> List[str]:
    """Extract worksheet names from Excel file."""
    try:
        workbook = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        return workbook.sheetnames
    except Exception as e:
        st.error(f"Error reading Excel workbook: {str(e)}")
        return []


def filter_worksheet_names(sheet_names: List[str]) -> Tuple[List[str], List[str]]:
    """Separate valid sheets from BA-prefixed sheets."""
    ba_sheets = [name for name in sheet_names if name.startswith("BA")]
    valid_sheets = [name for name in sheet_names if not name.startswith("BA")]
    return valid_sheets, ba_sheets


def remove_rows_after_total(df: pd.DataFrame) -> pd.DataFrame:
    """Remove rows after the first row that contains 'total' (case-insensitive) in first column."""
    if df.empty:
        return df
    
    first_col = df.iloc[:, 0]
    total_indices = [
        idx for idx, val in enumerate(first_col)
        if isinstance(val, str)
        and any(keyword in val.lower() for keyword in ["total", "jumlah"])
    ]
    
    if total_indices:
        first_total_idx = total_indices[0]
        df = df.iloc[:first_total_idx]
    
    return df


def get_transaction_month(date_value) -> Optional[int]:
    """Extract month from date value and map to quarter month (1-3)."""
    try:
        if pd.isna(date_value):
            return None
        
        # Handle different date formats
        if isinstance(date_value, str):
            # Try to parse string dates
            try:
                dt = pd.to_datetime(date_value)
            except:
                return None
        else:
            dt = pd.to_datetime(date_value)
        
        month = dt.month
        return month  # Return actual month (1-12)
    except:
        return None


def detect_quarter_month(month: int) -> Optional[int]:
    """Convert calendar month to quarter position (1-3).
    
    Common quarters:
    Jan-Feb-Mar (Q1): 1,2,3
    Apr-May-Jun (Q2): 4,5,6 -> map to 1,2,3
    Jul-Aug-Sep (Q3): 7,8,9 -> map to 1,2,3
    Oct-Nov-Dec (Q4): 10,11,12 -> map to 1,2,3
    """
    if month is None:
        return None
    
    # Determine which quarter
    quarter = (month - 1) // 3
    quarter_position = (month - 1) % 3 + 1  # 1, 2, or 3
    
    return quarter_position


def get_amount_column(sheet_name: str, df: pd.DataFrame) -> Optional[str]:
    """Get amount column based on sheet name prefix.
    
    Rules:
    - If sheet starts with "prov" -> look for PAID_TO_PROVIDER
    - If sheet contains "reimb" -> look for TOTAL REALISASI
    - Fallback to any numeric column with keywords
    """
    if df.empty:
        return None
    
    cols_lower = {col.lower(): col for col in df.columns}
    
    if sheet_name.lower().startswith("prov"):
        # Look for PAID_TO_PROVIDER
        for col_lower, col_orig in cols_lower.items():
            if "paid_to_provider" in col_lower or "paid_to_provide" in col_lower:
                return col_orig
    
    if "reimb" in sheet_name.lower():
        # Look for total realisasi
        for col_lower, col_orig in cols_lower.items():
            if "total realisasi" in col_lower or "total_realisasi" in col_lower:
                return col_orig
    
    # Fallback to any numeric column that looks like amount
    for col_lower, col_orig in cols_lower.items():
        if any(keyword in col_lower for keyword in ["amount", "biaya", "paid", "total", "nilai", "cost"]):
            return col_orig
    
    return None


def select_payroll_area_by_month(row: pd.Series, detected_month: Optional[int], quarter_start_month: Optional[int] = None) -> str:
    """Select appropriate Payroll Area based on detected transaction month.
    
    Logic:
    - If transaction month is within current quarter -> use corresponding Bulan (1/2/3)
    - If transaction month is from different quarter -> use Bulan 1 as fallback
    - If month not detected -> use priority: 3 > 2 > 1
    
    Args:
        row: Series containing Payroll Area columns
        detected_month: Transaction month (1-12)
        quarter_start_month: Start month of current quarter (e.g., 4 for Q2 Apr-Jun)
    """

    if detected_month is None or pd.isna(detected_month):
        # If we can't detect month, use priority: 3 > 2 > 1
        for month in [3, 2, 1]:
            col_name = f"Payroll Area Bulan {month}"
            if col_name in row and pd.notna(row[col_name]):
                return row[col_name]
        return "NOT FOUND"
    
    # Get quarter position of detected month
    quarter_position = detect_quarter_month(detected_month)
    
    # Determine current quarter start month (if not provided, assume current date)
    if quarter_start_month is None:
        current_month = datetime.now().month
        quarter_start_month = ((current_month - 1) // 3) * 3 + 1
    
    # Get quarter position of current quarter
    current_quarter = (quarter_start_month - 1) // 3
    transaction_quarter = (detected_month - 1) // 3
    
    # If transaction is from same quarter, use the corresponding month
    if transaction_quarter == current_quarter and quarter_position is not None:
        col_name = f"Payroll Area Bulan {quarter_position}"
        if col_name in row and pd.notna(row[col_name]):
            return row[col_name]
    
    # If transaction is from different quarter (previous/future), use Bulan 1
    # Bulan 1 is the beginning month of the current quarter
    col_name = "Payroll Area Bulan 1"
    if col_name in row and pd.notna(row[col_name]):
        return row[col_name]
    
    # Fallback: use other months if Bulan 1 not available
    for month in [2, 3]:
        col_name = f"Payroll Area Bulan {month}"
        if col_name in row and pd.notna(row[col_name]):
            return row[col_name]
    
    return "NOT FOUND"


def load_medical_reimbursement_sheets(file_bytes: bytes, selected_sheets: List[str]) -> Tuple[pd.DataFrame, Dict]:
    """Load only selected worksheets from Medical Reimbursement file one by one.
    
    Returns: (combined_dataframe, sheet_metadata_dict)
    """
    if not selected_sheets:
        st.error("Please select at least one worksheet.")
        return pd.DataFrame(), {}
    
    try:
        dataframes = []
        sheet_metadata = {}
        
        for sheet_name in selected_sheets:
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=1)
            
            if df.empty:
                continue
            
            # Remove rows after total row
            df = remove_rows_after_total(df)
            
            if df.empty:
                continue
            
            # Remove rows where all columns are NaN
            df = df.dropna(how='all')
            
            # Detect amount column based on sheet name
            amount_col = get_amount_column(sheet_name, df)
            df['Source Sheet'] = sheet_name
            df['Amount Column'] = amount_col if amount_col else "N/A"
            
            sheet_metadata[sheet_name] = {
                'amount_column': amount_col,
                'record_count': len(df)
            }
            
            dataframes.append(df)
        
        if not dataframes:
            st.error("No valid data found in selected worksheets.")
            return pd.DataFrame(), {}
        
        combined_df = pd.concat(dataframes, ignore_index=True)
        return combined_df, sheet_metadata
    except Exception as e:
        st.error(f"Error loading worksheets: {str(e)}")
        return pd.DataFrame(), {}


def load_employee_master(file_bytes: bytes, month: int) -> pd.DataFrame:
    """Load Employee Master file with required columns for specific month."""
    try:
        excel = pd.ExcelFile(io.BytesIO(file_bytes))

        hr_sheets = [s for s in excel.sheet_names if "HR" in s]

        # Baca semua sheet yang mengandung "HR"
        df_dict = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=hr_sheets
        )

        # Jika ingin digabung menjadi satu DataFrame
        df = pd.concat(
            [v.assign(SourceSheet=k) for k, v in df_dict.items()],
            ignore_index=True
        )
        
        # Check for required columns
        required_columns = ["NIK", "Payroll Area"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"{MONTH_NAMES[month]}: Missing required columns: {', '.join(missing_columns)}")
            return pd.DataFrame()
        
        # Select and rename columns
        df = df[["NIK", "Payroll Area"]].copy()
        df.rename(columns={"NIK": "NIK", "Payroll Area": f"Payroll Area Bulan {month}"}, inplace=True)
        
        return df
    except Exception as e:
        st.error(f"Error loading {MONTH_NAMES[month]}: {str(e)}, pastikan nama sheet pegawai mengandung 'HR'")
        return pd.DataFrame()


def merge_employee_masters(month1_df: pd.DataFrame, month2_df: pd.DataFrame, month3_df: pd.DataFrame) -> pd.DataFrame:
    """Merge three monthly Employee Master files."""
    dfs = [month1_df, month2_df, month3_df]
    
    if all(df.empty for df in dfs):
        st.error("All Employee Master files are empty.")
        return pd.DataFrame()
    
    # Remove duplicates from each dataframe
    for i, df in enumerate(dfs, 1):
        if not df.empty:
            dfs[i-1] = df.drop_duplicates(subset=["NIK"], keep="last")
    
    # Start with first non-empty dataframe
    merged = None
    for df in dfs:
        if not df.empty:
            merged = df
            break
    
    if merged is None:
        return pd.DataFrame()
    
    # Merge remaining dataframes
    for df in dfs:
        if not df.empty and not merged.equals(df):
            merged = pd.merge(merged, df, on="NIK", how="outer")
    
    # Determine final Payroll Area based on priority: Bulan 3 > Bulan 2 > Bulan 1
    def get_payroll_area(row):
        for month in [3, 2, 1]:
            col_name = f"Payroll Area Bulan {month}"
            if col_name in row and pd.notna(row[col_name]):
                return row[col_name]
        return None
    
    merged["Payroll Area"] = merged.apply(get_payroll_area, axis=1)
    
    # Select required columns
    cols_to_keep = ["NIK", "Payroll Area"]
    for month in [1, 2, 3]:
        col_name = f"Payroll Area Bulan {month}"
        if col_name in merged.columns:
            cols_to_keep.insert(1 + (month - 1), col_name)
    
    merged = merged[[col for col in cols_to_keep if col in merged.columns]]
    
    return merged


def join_medical_to_master(medical_df: pd.DataFrame, master_df: pd.DataFrame) -> pd.DataFrame:
    """LEFT JOIN medical data to employee master with transaction-month aware Payroll Area.
    
    For transactions outside the current quarter, uses Payroll Area Bulan 1 as fallback.
    """
    if "NIK" not in medical_df.columns:
        st.error("Medical data does not contain NIK column.")
        return pd.DataFrame()
    
    # Ensure NIK is int type for consistent joining
    medical_df_copy = medical_df.copy()
    medical_df_copy["NIK"] = medical_df_copy["NIK"].astype(int)
    master_df_copy = master_df.copy()
    master_df_copy["NIK"] = master_df_copy["NIK"].astype(int)
    
    # LEFT JOIN
    result = pd.merge(medical_df_copy, master_df_copy, on="NIK", how="left")
    
    # Detect transaction month from various date columns
    date_column = None
    for col in result.columns:
        if any(keyword in col.lower() for keyword in ['admission', 'date', 'tanggal', 'tgl']):
            date_column = col
            break
    
    # Extract and convert month, determine quarter
    if date_column:
        result['Transaction_Month'] = result[date_column].apply(get_transaction_month)
        result['Quarter_Position'] = result['Transaction_Month'].apply(detect_quarter_month)
        
        # Determine quarter start month from the data (use earliest date to infer quarter)
        valid_dates = pd.to_datetime(result[date_column], errors='coerce')
        if not valid_dates.isna().all():
            earliest_date = valid_dates.min()
            quarter_start_month = ((earliest_date.month - 1) // 3) * 3 + 1
        else:
            # Fallback to current quarter
            quarter_start_month = ((datetime.now().month - 1) // 3) * 3 + 1
    else:
        result['Quarter_Position'] = None
        quarter_start_month = ((datetime.now().month - 1) // 3) * 3 + 1
    
    # Select Payroll Area based on transaction month and quarter
    result['Payroll Area'] = result.apply(
        lambda row: select_payroll_area_by_month(row, row.get('Transaction_Month'), quarter_start_month),
        axis=1
    )
    
    # Fill remaining NaN with "NOT FOUND"
    result['Payroll Area'] = result['Payroll Area'].fillna("NOT FOUND")
    
    return result


def calculate_statistics(df: pd.DataFrame, selected_sheets: List[str]) -> Dict:
    """Calculate KPI metrics."""
    total_records = len(df)
    unique_nik = df["NIK"].nunique()
    matched = len(df[df["Payroll Area"] != "NOT FOUND"])
    not_found = len(df[df["Payroll Area"] == "NOT FOUND"])
    match_pct = (matched / total_records * 100) if total_records > 0 else 0
    not_found_pct = (not_found / total_records * 100) if total_records > 0 else 0
    
    return {
        "total_sheets": len(selected_sheets),
        "total_records": total_records,
        "unique_nik": unique_nik,
        "matched": matched,
        "not_found": not_found,
        "match_pct": match_pct,
        "not_found_pct": not_found_pct
    }


def display_kpi_metrics(stats: Dict):
    """Display KPI metrics in columns."""
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Selected Sheets", stats["total_sheets"])
    with col2:
        st.metric("Total Records", f"{stats['total_records']:,}")
    with col3:
        st.metric("Unique NIK", f"{stats['unique_nik']:,}")
    with col4:
        st.metric("Matched Records", f"{stats['matched']:,}")
    
    col5, col6, col7 = st.columns(3)
    
    with col5:
        st.metric("Not Found Records", f"{stats['not_found']:,}")
    with col6:
        st.metric("Match Percentage", f"{stats['match_pct']:.2f}%")
    with col7:
        st.metric("Not Found Percentage", f"{stats['not_found_pct']:.2f}%")


def filter_dataframe(df: pd.DataFrame, search_term: str = "") -> pd.DataFrame:
    """Apply search filtering to dataframe."""
    result = df.copy()
    
    if search_term:
        mask = result.astype(str).apply(lambda x: x.str.contains(search_term, case=False)).any(axis=1)
        result = result[mask]
    
    return result


def create_summary_table(df: pd.DataFrame) -> pd.DataFrame:
    """Create summary grouped by Payroll Area."""
    summary = df.groupby("Payroll Area", as_index=False).size()
    summary.rename(columns={"size": "Total Records"}, inplace=True)
    
    total = summary["Total Records"].sum()
    summary["Percentage"] = (summary["Total Records"] / total * 100).round(2)
    
    summary = summary.sort_values("Total Records", ascending=False)
    
    return summary


def aggregate_by_payroll_area(df: pd.DataFrame, amount_columns: Dict[str, str]) -> pd.DataFrame:
    """Aggregate records by Payroll Area based on transaction month, summing amounts.
    
    Args:
        df: enriched dataframe with transaction-aware Payroll Area
        amount_columns: dict mapping sheet_name to amount_column_name
    
    Returns:
        Aggregated dataframe grouped by Payroll Area with totals
    """
    if df.empty or "Payroll Area" not in df.columns:
        return pd.DataFrame()
    
    # Create a copy for aggregation
    agg_df = df.copy()
    
    # Prepare amount for summing
    numeric_amounts = []
    for idx, row in agg_df.iterrows():
        sheet_name = row.get("Source Sheet", "")
        amount_col = amount_columns.get(sheet_name)
        
        if amount_col and amount_col in row:
            try:
                # Handle numeric values with various formats
                val = row[amount_col]
                if isinstance(val, str):
                    # Remove common separators and convert
                    val = val.replace('.', '').replace(',', '.')
                amount = float(val)
                numeric_amounts.append(amount)
            except (ValueError, TypeError):
                numeric_amounts.append(0)
        else:
            numeric_amounts.append(0)
    
    agg_df['Numeric Amount'] = numeric_amounts
    
    # Group by Payroll Area
    summary = agg_df.groupby("Payroll Area", as_index=False).agg({
        "NIK": "count",
        "Numeric Amount": "sum"
    }).round(2)
    
    summary.rename(columns={"NIK": "Total Records"}, inplace=True)
    summary.rename(columns={"Numeric Amount": "Total Amount"}, inplace=True)
    
    total_records = summary["Total Records"].sum()
    total_amount = summary["Total Amount"].sum()
    
    summary["Percentage Records"] = (summary["Total Records"] / total_records * 100).round(2)
    summary["Percentage Amount"] = (summary["Total Amount"] / total_amount * 100).round(2)
    
    summary = summary.sort_values("Total Amount", ascending=False)
    
    return summary


def display_aggregation_summary(enriched_df: pd.DataFrame, agg_summary_df: pd.DataFrame):
    """Display aggregation summary with visual breakdown."""
    if agg_summary_df.empty:
        st.info("No aggregation data available.")
        return
    
    # Show total metrics
    col1, col2, col3 = st.columns(3)
    
    with col1:
        total_records = agg_summary_df["Total Records"].sum()
        st.metric("Total Records", f"{int(total_records):,}")
    
    with col2:
        total_amount = agg_summary_df["Total Amount"].sum()
        st.metric("Total Amount", f"Rp {total_amount:,.0f}")
    
    with col3:
        avg_amount = total_amount / total_records if total_records > 0 else 0
        st.metric("Average per Record", f"Rp {avg_amount:,.0f}")
    
    st.divider()
    
    # Show detailed breakdown
    st.subheader("Breakdown by Payroll Area")
    
    display_cols = ["Payroll Area", "Total Records", "Total Amount", "Percentage Records", "Percentage Amount"]
    display_df = agg_summary_df[display_cols].copy()
    display_df["Total Amount"] = display_df["Total Amount"].apply(lambda x: f"Rp {x:,.0f}")
    display_df["Percentage Records"] = display_df["Percentage Records"].apply(lambda x: f"{x:.2f}%")
    display_df["Percentage Amount"] = display_df["Percentage Amount"].apply(lambda x: f"{x:.2f}%")
    
    st.dataframe(
        display_df,
        use_container_width=True,
        height=min(400, len(display_df) * 35 + 35),
        hide_index=True
    )
    
    # Show top Payroll Areas
    st.subheader("Top 5 Payroll Areas by Amount")
    top_5 = agg_summary_df.head(5)[["Payroll Area", "Total Amount"]].copy()
    top_5["Total Amount"] = top_5["Total Amount"].apply(lambda x: f"Rp {x:,.0f}")
    
    col1, col2 = st.columns([2, 1])
    with col1:
        st.bar_chart(
            agg_summary_df.head(10).set_index("Payroll Area")["Total Amount"],
            use_container_width=True
        )
    
    with col2:
        st.dataframe(top_5, use_container_width=True, hide_index=True)


def create_processing_log(
    medical_filename: str,
    month1_filename: str,
    month2_filename: str,
    month3_filename: str,
    selected_sheets: List[str],
    ignored_sheets: List[str],
    processing_time: float,
    stats: Dict
) -> pd.DataFrame:
    """Create processing log dataframe."""
    log_data = {
        "Medical filename": [medical_filename],
        "Bulan 1 filename": [month1_filename],
        "Bulan 2 filename": [month2_filename],
        "Bulan 3 filename": [month3_filename],
        "Selected worksheets": [", ".join(selected_sheets)],
        "Ignored worksheets": [", ".join(ignored_sheets)],
        "Processing date": [datetime.now().strftime("%Y-%m-%d")],
        "Processing time (seconds)": [round(processing_time, 2)],
        "Total selected sheets": [stats["total_sheets"]],
        "Total records": [stats["total_records"]],
        "Unique NIK": [stats["unique_nik"]],
        "Matched": [stats["matched"]],
        "Not Found": [stats["not_found"]],
        "Match percentage": [f"{stats['match_pct']:.2f}%"]
    }
    return pd.DataFrame(log_data)


def export_to_excel(
    all_data: pd.DataFrame,
    matched_data: pd.DataFrame,
    not_found_data: pd.DataFrame,
    summary_data: pd.DataFrame,
    agg_summary_data: pd.DataFrame,
    log_data: pd.DataFrame
) -> bytes:
    """Create Excel workbook with multiple sheets."""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        all_data.to_excel(writer, sheet_name="All Data", index=False)
        matched_data.to_excel(writer, sheet_name="Matched", index=False)
        not_found_data.to_excel(writer, sheet_name="Not Found", index=False)
        summary_data.to_excel(writer, sheet_name="Summary Count", index=False)
        if not agg_summary_data.empty:
            agg_summary_data.to_excel(writer, sheet_name="Summary Aggregated", index=False)
        log_data.to_excel(writer, sheet_name="Processing Log", index=False)
    
    output.seek(0)
    return output.getvalue()


def display_data_table(df: pd.DataFrame, key_suffix: str):
    """Display searchable, sortable dataframe."""
    if df.empty:
        st.info("No data to display.")
        return
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        search_term = st.text_input("🔍 Search", key=f"search_{key_suffix}", placeholder="Search all columns...")
    
    with col2:
        row_count = st.selectbox(
            "Rows per page",
            [10, 25, 50, 100],
            key=f"rows_{key_suffix}"
        )
    
    filtered_df = filter_dataframe(df, search_term)
    
    st.markdown(f"**Showing {len(filtered_df)} of {len(df)} rows**")
    
    st.dataframe(
        filtered_df,
        use_container_width=True,
        height=min(400, len(filtered_df) * 35 + 35)
    )


# Initialize session state
if "processing_complete" not in st.session_state:
    st.session_state.processing_complete = False
if "enriched_data" not in st.session_state:
    st.session_state.enriched_data = pd.DataFrame()
if "statistics" not in st.session_state:
    st.session_state.statistics = {}
if "processing_log" not in st.session_state:
    st.session_state.processing_log = pd.DataFrame()
if "sheet_metadata" not in st.session_state:
    st.session_state.sheet_metadata = {}

# Sidebar
with st.sidebar:
    st.header("📁 Upload Files")
    
    medical_file = st.file_uploader("Medical Reimbursement", type=["xlsx", "xls"], key="medical")
    
    month1_file = st.file_uploader("Employee Master - Bulan 1", type=["xlsx", "xls"], key="month1")
    month2_file = st.file_uploader("Employee Master - Bulan 2", type=["xlsx", "xls"], key="month2")
    month3_file = st.file_uploader("Employee Master - Bulan 3", type=["xlsx", "xls"], key="month3")
    
    valid_sheets = []
    ignored_sheets = []
    
    if medical_file:
        st.subheader("📋 Worksheet Selection")
        sheet_names = get_worksheet_names(medical_file.getvalue())
        valid_sheets, ignored_sheets = filter_worksheet_names(sheet_names)
        
        if ignored_sheets:
            with st.expander(f"ℹ️ Ignored worksheets ({len(ignored_sheets)})"):
                st.write(", ".join(ignored_sheets))
        
        selected_sheets = st.multiselect(
            "Select worksheets to process",
            options=valid_sheets,
            default=valid_sheets,
            key="sheet_selector"
        )
    else:
        selected_sheets = []
    
    st.divider()
    
    process_button = st.button(
        "🚀 Process",
        type="primary",
        use_container_width=True,
        key="process_btn"
    )

# Main content
if process_button:
    if not medical_file:
        st.error("Please upload Medical Reimbursement file.")
    elif not month1_file:
        st.error("Please upload Employee Master - Bulan 1 file.")
    elif not month2_file:
        st.error("Please upload Employee Master - Bulan 2 file.")
    elif not month3_file:
        st.error("Please upload Employee Master - Bulan 3 file.")
    elif not selected_sheets:
        st.error("Please select at least one worksheet.")
    else:
        start_time = time.time()
        
        with st.spinner("⏳ Processing data..."):
            # Load medical reimbursement data
            medical_df, sheet_metadata = load_medical_reimbursement_sheets(medical_file.getvalue(), selected_sheets)
            
            if medical_df.empty:
                st.error("Failed to load medical data.")
            else:
                # Load employee masters (3 months)
                month1_df = load_employee_master(month1_file.getvalue(), 1)
                month2_df = load_employee_master(month2_file.getvalue(), 2)
                month3_df = load_employee_master(month3_file.getvalue(), 3)
                
                if month1_df.empty and month2_df.empty and month3_df.empty:
                    st.error("Failed to load all Employee Master files.")
                else:
                    # Merge employee masters
                    merged_master = merge_employee_masters(month1_df, month2_df, month3_df)
                    
                    # Join to medical data
                    enriched_df = join_medical_to_master(medical_df, merged_master)
                    
                    # Calculate statistics
                    stats = calculate_statistics(enriched_df, selected_sheets)
                    
                    # Create processing log
                    processing_time = time.time() - start_time
                    log_df = create_processing_log(
                        medical_filename=medical_file.name,
                        month1_filename=month1_file.name,
                        month2_filename=month2_file.name,
                        month3_filename=month3_file.name,
                        selected_sheets=selected_sheets,
                        ignored_sheets=ignored_sheets,
                        processing_time=processing_time,
                        stats=stats
                    )
                    
                    # Store in session state
                    st.session_state.processing_complete = True
                    st.session_state.enriched_data = enriched_df
                    st.session_state.statistics = stats
                    st.session_state.processing_log = log_df
                    st.session_state.selected_sheets = selected_sheets
                    st.session_state.ignored_sheets = ignored_sheets
                    st.session_state.sheet_metadata = sheet_metadata
                    
                    st.success(f"✅ Processing completed in {processing_time:.2f} seconds")

# Display results if processing is complete
if st.session_state.processing_complete:
    enriched_df = st.session_state.enriched_data
    stats = st.session_state.statistics
    sheet_metadata = st.session_state.sheet_metadata
    
    # KPI Metrics
    st.subheader("📊 KPI Metrics")
    display_kpi_metrics(stats)
    
    st.divider()
    
    # Prepare data for tabs
    matched_df = enriched_df[enriched_df["Payroll Area"] != "NOT FOUND"]
    not_found_df = enriched_df[enriched_df["Payroll Area"] == "NOT FOUND"]
    summary_df = create_summary_table(enriched_df)
    
    # Build amount columns dict for aggregation
    amount_columns = {sheet: metadata.get('amount_column') for sheet, metadata in sheet_metadata.items()}
    agg_summary_df = aggregate_by_payroll_area(enriched_df, amount_columns)
    
    log_df = st.session_state.processing_log
    
    # Create tabs
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["All Data", "Matched", "Not Found", "Summary", "Aggregated by PArea", "Processing Log"])
    
    with tab1:
        st.subheader("All Data")
        display_data_table(enriched_df, "all_data")
    
    with tab2:
        st.subheader("Matched Records")
        st.metric("Count", f"{len(matched_df):,}")
        display_data_table(matched_df, "matched")
    
    with tab3:
        st.subheader("Not Found Records")
        st.metric("Count", f"{len(not_found_df):,}")
        display_data_table(not_found_df, "not_found")
    
    with tab4:
        st.subheader("Summary by Payroll Area (Count)")
        display_data_table(summary_df, "summary")
    
    with tab5:
        st.subheader("💰 Aggregated Summary by Payroll Area")
        display_aggregation_summary(enriched_df, agg_summary_df)
    
    with tab6:
        st.subheader("Processing Log")
        st.dataframe(log_df, use_container_width=True)
    
    st.divider()
    
    # Export section
    st.subheader("💾 Export")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        csv_all = enriched_df.to_csv(index=False)
        st.download_button(
            label="📥 All Data (CSV)",
            data=csv_all,
            file_name="medical_all_data.csv",
            mime="text/csv"
        )
    
    with col2:
        csv_matched = matched_df.to_csv(index=False)
        st.download_button(
            label="📥 Matched (CSV)",
            data=csv_matched,
            file_name="medical_matched.csv",
            mime="text/csv"
        )
    
    with col3:
        csv_not_found = not_found_df.to_csv(index=False)
        st.download_button(
            label="📥 Not Found (CSV)",
            data=csv_not_found,
            file_name="medical_not_found.csv",
            mime="text/csv"
        )
    
    with col4:
        csv_summary = summary_df.to_csv(index=False)
        st.download_button(
            label="📥 Summary (CSV)",
            data=csv_summary,
            file_name="medical_summary.csv",
            mime="text/csv"
        )
    
    # Excel export
    st.markdown("---")
    excel_bytes = export_to_excel(enriched_df, matched_df, not_found_df, summary_df, agg_summary_df, log_df)
    st.download_button(
        label="📊 Download Complete Excel Workbook",
        data=excel_bytes,
        file_name="Medical_Reimbursement_Result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👈 Upload files and click Process to begin.")