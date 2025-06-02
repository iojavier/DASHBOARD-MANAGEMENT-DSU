import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from pandas import ExcelWriter
import re

st.set_page_config(layout="wide", page_title="Daily Remark Summary", page_icon="📊", initial_sidebar_state="expanded")
st.title('Daily Remark Summary')

# File uploader for multiple Excel files
uploaded_files = st.file_uploader("Upload Excel Files", type=["xlsx"], accept_multiple_files=True)

@st.cache_data
def load_data(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip().str.upper()
        if 'TIME' not in df.columns:
            st.error("No 'TIME' column found in the Excel file. Available columns: " + str(df.columns.tolist()))
        # Convert date column to date-only and exclude Sundays
        df['DATE'] = pd.to_datetime(df['DATE'].dt.date, errors='coerce')
        df = df[df['DATE'].dt.weekday != 6]  # Exclude Sundays
        # Optimize data types early
        for col in ['REMARK BY', 'DEBTOR', 'STATUS', 'REMARK', 'CLIENT']:
            if col in df.columns:
                df[col] = df[col].astype('category')
        return df
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return pd.DataFrame()

def to_excel(summary_df, cumulative_df):
    output = BytesIO()
    with ExcelWriter(output, engine='xlsxwriter', date_format='yyyy-mm-dd') as writer:
        workbook = writer.book
        formats = {
            'title': workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFFF00'}),
            'center': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1}),
            'header': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bg_color': 'red', 'font_color': 'white', 'bold': True}),
            'comma': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0'}),
        }

        # Write Combined Summary
        worksheet = workbook.add_worksheet('Combined Summary')
        current_row = 0
        worksheet.merge_range(current_row, 0, current_row, len(summary_df.columns) - 1, 'Combined Summary', formats['title'])
        current_row += 1
        for col_num, col_name in enumerate(summary_df.columns):
            worksheet.write(current_row, col_num, col_name, formats['header'])
            max_len = max(summary_df[col_name].astype(str).str.len().max(), len(col_name)) + 2
            worksheet.set_column(col_num, col_num, max_len)
        current_row += 1
        for row_num in range(len(summary_df)):
            for col_num, col_name in enumerate(summary_df.columns):
                value = summary_df.iloc[row_num, col_num]
                if col_name == 'DATE RANGE':
                    worksheet.write(current_row + row_num, col_num, value, formats['center'])
                elif col_name in ['ACCOUNTS', 'BANK ESCALATION', 'CONNECTED #']:
                    worksheet.write(current_row + row_num, col_num, value, formats['comma'])

        # Write Cumulative Account Summary
        worksheet = workbook.add_worksheet('Cumulative Account Summary')
        current_row = 0
        worksheet.merge_range(current_row, 0, current_row, len(cumulative_df.columns) - 1, 'Cumulative Account Summary', formats['title'])
        current_row += 1
        for col_num, col_name in enumerate(cumulative_df.columns):
            worksheet.write(current_row, col_num, col_name, formats['header'])
            max_len = max(cumulative_df[col_name].astype(str).str.len().max(), len(col_name)) + 2
            worksheet.set_column(col_num, col_num, max_len)
        current_row += 1
        for row_num in range(len(cumulative_df)):
            for col_num, col_name in enumerate(cumulative_df.columns):
                value = cumulative_df.iloc[row_num, col_num]
                if col_name == 'DATE RANGE':
                    worksheet.write(current_row + row_num, col_num, value, formats['center'])
                elif col_name == 'ACCOUNT NO.':
                    worksheet.write(current_row + row_num, col_num, value, formats['comma'])

    return output.getvalue()

def process_file(df, selected_clients, chunksize=50000):
    try:
        # Initialize list to store processed chunks
        processed_chunks = []
        
        # Apply client filter first to reduce DataFrame size
        if selected_clients and "All Clients" not in selected_clients:
            df = df[df['CLIENT'].isin(selected_clients)]
        
        # Process DataFrame in chunks
        for start in range(0, len(df), chunksize):
            chunk = df.iloc[start:start + chunksize].copy()
            # Ensure categorical types are preserved
            for col in ['REMARK BY', 'DEBTOR', 'STATUS', 'REMARK']:
                if col in chunk.columns:
                    chunk[col] = chunk[col].astype('category')
            # Apply filters
            chunk = chunk[chunk['REMARK BY'] != 'SPMADRID']
            chunk = chunk[~chunk['DEBTOR'].str.contains("DEFAULT_LEAD_", case=False, na=False)]
            chunk = chunk[~chunk['STATUS'].str.contains('ABORT', na=False)]
            chunk = chunk[~chunk['REMARK'].str.contains(r'1_\d{11} - PTP NEW', case=False, na=False, regex=True)]
            chunk = chunk[~chunk['REMARK'].str.contains('Broadcast', case=False, na=False)]
            excluded_remarks = ["Broken Promise", "New files imported", "Updates when case reassign to another collector", 
                                "NDF IN ICS", "FOR PULL OUT (END OF HANDLING PERIOD)", "END OF HANDLING PERIOD", 
                                "New Assignment -", "File Unhold"]
            escaped_remarks = [re.escape(remark) for remark in excluded_remarks]
            chunk = chunk[~chunk['REMARK'].str.contains('|'.join(escaped_remarks), case=False, na=False)]
            chunk['CARD NO.'] = chunk['CARD NO.'].astype(str)
            processed_chunks.append(chunk)
        
        # Concatenate chunks
        return pd.concat(processed_chunks, ignore_index=True) if processed_chunks else pd.DataFrame()
    except MemoryError:
        st.error("Memory error occurred while processing the file. Try reducing chunksize or increasing system memory.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Error processing file: {e}")
        return pd.DataFrame()

def apply_client_filter(df, selected_clients):
    try:
        if selected_clients and "All Clients" not in selected_clients:
            df = df[df['CLIENT'].isin(selected_clients)]
        return df
    except Exception as e:
        st.error(f"Error applying client filter: {e}")
        return pd.DataFrame()

def get_connected_accounts_per_file(df, remark_types):
    try:
        df_filtered = df[df['REMARK TYPE'].isin(remark_types)].copy()
        df_filtered = df_filtered[~((df_filtered['CALL STATUS'].str.contains('OTHERS', case=False, na=False)) & 
                                    (df_filtered['REMARK'].str.contains('@', case=False, na=False)))]
        connected_accounts = df_filtered[pd.to_numeric(df_filtered['TALK TIME DURATION'], errors='coerce') > 0]['ACCOUNT NO.'].unique()
        return set(connected_accounts)
    except Exception as e:
        st.error(f"Error calculating connected accounts: {e}")
        return set()

def get_bank_escalation_accounts_per_file(df):
    try:
        return set(df[df['STATUS'].str.contains('BANK ESCALATION', case=False, na=False)]['ACCOUNT NO.'].unique())
    except Exception as e:
        st.error(f"Error calculating bank escalation accounts: {e}")
        return set()

def calculate_summary(all_dfs, start_date, end_date):
    summary_columns = ['DATE RANGE', 'ACCOUNTS', 'BANK ESCALATION', 'CONNECTED #']
    summary_table = pd.DataFrame(columns=summary_columns)
    
    all_accounts = set()
    all_bank_escalations = set()
    all_connected_accounts = set()
    
    for df in all_dfs:
        df_filtered = df[(df['DATE'].dt.date >= start_date) & (df['DATE'].dt.date <= end_date)]
        if df_filtered.empty:
            continue
        accounts = set(df_filtered['ACCOUNT NO.'].unique())
        all_accounts.update(accounts)
        bank_escalation_accounts = get_bank_escalation_accounts_per_file(df_filtered)
        all_bank_escalations.update(bank_escalation_accounts)
        connected_accounts = get_connected_accounts_per_file(df_filtered, ['Predictive', 'Follow Up', 'Outgoing'])
        all_connected_accounts.update(connected_accounts)
    
    if not all_accounts:
        st.write("No data available for the specified date range.")
        return summary_table

    date_range_str = f"{start_date.strftime('%Y-%m-%d')} - {end_date.strftime('%Y-%m-%d')}"
    
    summary_rows = [{
        'DATE RANGE': date_range_str,
        'ACCOUNTS': len(all_accounts),
        'BANK ESCALATION': len(all_bank_escalations),
        'CONNECTED #': len(all_connected_accounts)
    }]
    
    return pd.DataFrame(summary_rows)

def calculate_cumulative_account_summary(all_dfs, start_date, end_date):
    summary_columns = ['DATE RANGE', 'ACCOUNT NO.']
    summary_table = pd.DataFrame(columns=summary_columns)
    
    all_accounts = set()
    
    for df in all_dfs:
        df_filtered = df[(df['DATE'].dt.date >= start_date) & (df['DATE'].dt.date <= end_date)]
        if df_filtered.empty:
            continue
        accounts = set(df_filtered['ACCOUNT NO.'].unique())
        all_accounts.update(accounts)
    
    if not all_accounts:
        st.write("No data available for the specified date range.")
        return summary_table

    date_range_str = f"{start_date.strftime('%Y-%m-%d')} - {end_date.strftime('%Y-%m-%d')}"
    
    summary_rows = [{
        'DATE RANGE': date_range_str,
        'ACCOUNT NO.': len(all_accounts)
    }]
    
    return pd.DataFrame(summary_rows)

# Collect unique clients and dates from uploaded files
all_clients = set()
all_dates = []
if uploaded_files:
    for file in uploaded_files:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip().str.upper()
        if 'CLIENT' in df.columns:
            all_clients.update(df['CLIENT'].dropna().unique())
        if 'DATE' in df.columns:
            df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
            all_dates.extend(df['DATE'].dt.date.dropna().unique())
all_clients = sorted(list(all_clients))
all_clients.insert(0, "All Clients")

# Determine min and max dates, with fallback
min_date = min(all_dates) if all_dates else datetime.date.today()
max_date = max(all_dates) if all_dates else datetime.date.today()

# Sidebar inputs
st.sidebar.header("Select Filters")
selected_clients = st.sidebar.multiselect("Select Clients", all_clients, default=["All Clients"])
start_date = st.sidebar.date_input("Start Date", value=min_date, min_value=min_date, max_value=max_date)
end_date = st.sidebar.date_input("End Date", value=max_date, min_value=min_date, max_value=max_date)

if uploaded_files:
    all_dfs_processed = []
    all_dfs_raw = []
    all_dfs_raw_filtered = []
    for idx, file in enumerate(uploaded_files, 1):
        # Load raw data
        df_raw = load_data(file)
        if df_raw.empty:
            continue
        all_dfs_raw.append(df_raw)
        # Apply client filter to raw data for cumulative summary
        df_raw_filtered = apply_client_filter(df_raw.copy(), selected_clients)
        all_dfs_raw_filtered.append(df_raw_filtered)
        # Process data for combined summary with all filters
        df_processed = process_file(df_raw.copy(), selected_clients, chunksize=50000)
        if df_processed.empty:
            continue
        all_dfs_processed.append(df_processed)
        st.write(f"Processed file {idx}")

    # Calculate summaries
    combined_summary = calculate_summary(all_dfs_processed, start_date, end_date)
    cumulative_summary = calculate_cumulative_account_summary(all_dfs_raw_filtered, start_date, end_date)

    st.write("## Overall Combined Summary Table")
    st.write(combined_summary)
    
    st.write("## Cumulative Account Summary Table")
    st.write(cumulative_summary)

    # Prepare Excel download
    try:
        excel_data = to_excel(combined_summary, cumulative_summary)
        st.download_button(
            label="Download Summaries as Excel",
            data=excel_data,
            file_name=f"Remark_Summaries_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error generating Excel file: {e}")
else:
    st.info("Please upload at least one Excel file.")