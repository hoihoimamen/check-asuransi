import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
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
st.markdown("Enrich Medical Reimbursement data with Payroll Area information")


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
    
    # Get the first column
    first_col = df.iloc[:, 0]
    
    # Find rows where first column contains 'total' (case-insensitive)
    total_indices = [
        idx for idx, val in enumerate(first_col)
        if isinstance(val, str)
        and any(keyword in val.lower() for keyword in ["total", "jumlah"])
    ]
    
    if total_indices:
        # Keep only rows before the first total row
        first_total_idx = total_indices[0]
        df = df.iloc[:first_total_idx]
    
    return df


def load_medical_reimbursement_sheets(file_bytes: bytes, selected_sheets: List[str]) -> pd.DataFrame:
    """Load only selected worksheets from Medical Reimbursement file one by one."""
    if not selected_sheets:
        st.error("Please select at least one worksheet.")
        return pd.DataFrame()
    
    try:
        dataframes = []
        
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
            df['Source Sheet'] = sheet_name
            dataframes.append(df)
        
        if not dataframes:
            st.error("No valid data found in selected worksheets.")
            return pd.DataFrame()
        
        combined_df = pd.concat(dataframes, ignore_index=True)
        return combined_df
    except Exception as e:
        st.error(f"Error loading worksheets: {str(e)}")
        return pd.DataFrame()


def load_employee_master(file_bytes: bytes, file_type: str) -> pd.DataFrame:
    """Load Employee Master file with required columns."""
    try:
        df = pd.read_excel(io.BytesIO(file_bytes))
        
        # Check for required columns
        required_columns = ["NIK", "Payroll Area"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"{file_type}: Missing required columns: {', '.join(missing_columns)}")
            return pd.DataFrame()
        
        # Select and rename columns
        df = df[["NIK", "Payroll Area"]].copy()
        df.rename(columns={"NIK": "NIK"}, inplace=True)
        
        return df
    except Exception as e:
        st.error(f"Error loading {file_type}: {str(e)}")
        return pd.DataFrame()


def merge_employee_masters(beginning_df: pd.DataFrame, ending_df: pd.DataFrame) -> pd.DataFrame:
    """Merge Beginning and Ending Quarter employee masters."""
    if beginning_df.empty and ending_df.empty:
        st.error("Both Employee Master files are empty.")
        return pd.DataFrame()
    
    # Rename Payroll Area columns
    beginning_df.rename(columns={"Payroll Area": "Payroll Area Awal"}, inplace=True)
    ending_df.rename(columns={"Payroll Area": "Payroll Area Akhir"}, inplace=True)
    
    # Remove duplicates, keeping the last occurrence
    beginning_df = beginning_df.drop_duplicates(subset=["NIK"], keep="last")
    ending_df = ending_df.drop_duplicates(subset=["NIK"], keep="last")
    
    # Outer merge to capture NIK from both files
    merged = pd.merge(beginning_df, ending_df, on="NIK", how="outer")
    
    # Determine final Payroll Area based on rules
    def get_payroll_area(row):
        awal = row["Payroll Area Awal"]
        akhir = row["Payroll Area Akhir"]

        if pd.notna(awal) and pd.notna(akhir):
            if awal == akhir:
                return akhir
            else:
                return f"{awal} move to {akhir}"

        elif pd.notna(awal):
            return awal

        elif pd.notna(akhir):
            return akhir

        else:
            return "NOT FOUND"
    
    merged["Payroll Area"] = merged.apply(get_payroll_area, axis=1)
    
    # Select required columns
    merged = merged[["NIK", "Payroll Area Awal", "Payroll Area Akhir", "Payroll Area"]]
    
    return merged


def join_medical_to_master(medical_df: pd.DataFrame, master_df: pd.DataFrame) -> pd.DataFrame:
    """LEFT JOIN medical data to employee master."""
    if "NIK" not in medical_df.columns:
        st.error("Medical data does not contain NIK column.")
        return pd.DataFrame()
    
    # Ensure NIK is string type for consistent joining
    medical_df["NIK"] = medical_df["NIK"].astype(str).str.strip()
    master_df["NIK"] = master_df["NIK"].astype(str).str.strip()
    
    # LEFT JOIN
    result = pd.merge(medical_df, master_df, on="NIK", how="left")
    
    # Fill missing Payroll Area with "NOT FOUND"
    if "Payroll Area" in result.columns:
        result["Payroll Area"] = result["Payroll Area"].fillna("NOT FOUND")
    
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


def filter_dataframe(df: pd.DataFrame, search_term: str = "", filters: Dict = None) -> pd.DataFrame:
    """Apply search and filtering to dataframe."""
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


def create_processing_log(
    medical_filename: str,
    beginning_filename: str,
    ending_filename: str,
    selected_sheets: List[str],
    ignored_sheets: List[str],
    processing_time: float,
    stats: Dict
) -> pd.DataFrame:
    """Create processing log dataframe."""
    log_data = {
        "Medical filename": [medical_filename],
        "Beginning Quarter filename": [beginning_filename],
        "Ending Quarter filename": [ending_filename],
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
    not_found_data: pd.DataFrame
) -> bytes:
    """Create Excel workbook with multiple sheets."""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        all_data.to_excel(writer, sheet_name="All Data", index=False)
        matched_data.to_excel(writer, sheet_name="Matched", index=False)
        not_found_data.to_excel(writer, sheet_name="Matched", index=False)
    
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
if "medical_data" not in st.session_state:
    st.session_state.medical_data = pd.DataFrame()
if "enriched_data" not in st.session_state:
    st.session_state.enriched_data = pd.DataFrame()
if "statistics" not in st.session_state:
    st.session_state.statistics = {}
if "processing_log" not in st.session_state:
    st.session_state.processing_log = pd.DataFrame()

# Sidebar
with st.sidebar:
    st.header("📁 Upload Files")
    
    medical_file = st.file_uploader("Medical Reimbursement", type=["xlsx", "xls"], key="medical")
    beginning_file = st.file_uploader("Employee Master - Beginning Quarter", type=["xlsx", "xls"], key="beginning")
    ending_file = st.file_uploader("Employee Master - Ending Quarter", type=["xlsx", "xls"], key="ending")
    
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
    elif not beginning_file:
        st.error("Please upload Employee Master - Beginning Quarter file.")
    elif not ending_file:
        st.error("Please upload Employee Master - Ending Quarter file.")
    elif not selected_sheets:
        st.error("Please select at least one worksheet.")
    else:
        start_time = time.time()
        
        with st.spinner("⏳ Processing data..."):
            # Load medical reimbursement data
            medical_df = load_medical_reimbursement_sheets(medical_file.getvalue(), selected_sheets)

            if medical_df.empty:
                st.error("Failed to load medical data.")
            else:
                # Load employee masters
                beginning_df = load_employee_master(beginning_file.getvalue(), "Beginning Quarter")
                ending_df = load_employee_master(ending_file.getvalue(), "Ending Quarter")
                
                if beginning_df.empty and ending_df.empty:
                    st.error("Failed to load employee master files.")
                else:
                    # Merge employee masters
                    merged_master = merge_employee_masters(beginning_df, ending_df)
                    medical_df['NIK'] = medical_df['NIK'].astype(int)
                    # Join to medical data
                    enriched_df = join_medical_to_master(medical_df, merged_master)
                    
                    # Calculate statistics
                    stats = calculate_statistics(enriched_df, selected_sheets)
                    
                    # Create processing log
                    processing_time = time.time() - start_time
                    log_df = create_processing_log(
                        medical_filename=medical_file.name,
                        beginning_filename=beginning_file.name,
                        ending_filename=ending_file.name,
                        selected_sheets=selected_sheets,
                        ignored_sheets=ignored_sheets,
                        processing_time=processing_time,
                        stats=stats
                    )
                    
                    # Store in session state
                    st.session_state.processing_complete = True
                    st.session_state.medical_data = medical_df
                    st.session_state.enriched_data = enriched_df
                    st.session_state.statistics = stats
                    st.session_state.processing_log = log_df
                    st.session_state.selected_sheets = selected_sheets
                    st.session_state.ignored_sheets = ignored_sheets
                    
                    st.success(f"✅ Processing completed in {processing_time:.2f} seconds")

# Display results if processing is complete
if st.session_state.processing_complete:
    enriched_df = st.session_state.enriched_data
    stats = st.session_state.statistics
    
    # KPI Metrics
    st.subheader("📊 KPI Metrics")
    display_kpi_metrics(stats)
    
    st.divider()
    
    # Prepare data for tabs
    matched_df = enriched_df[enriched_df["Payroll Area"] != "NOT FOUND"]
    not_found_df = enriched_df[enriched_df["Payroll Area"] == "NOT FOUND"]
    summary_df = create_summary_table(enriched_df)
    
    # Create tabs
    tab1, tab2, tab3, tab4 = st.tabs(["All Data", "Matched", "Not Found", "Summary"])
    
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
        st.subheader("Summary by Payroll Area")
        display_data_table(summary_df, "summary")
    
    
    st.divider()

    # Excel export
    st.markdown("---")
    excel_bytes = export_to_excel(enriched_df, matched_df, not_found_df)
    st.download_button(
        label="📊 Download Complete Excel Workbook",
        data=excel_bytes,
        file_name="Medical_Reimbursement_Result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👈 Upload files and click Process to begin.")