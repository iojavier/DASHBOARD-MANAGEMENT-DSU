import streamlit as st
import pandas as pd
import datetime
import re
from io import BytesIO
from pandas import ExcelWriter
import uuid
import itertools

st.set_page_config(layout="wide", page_title="Daily Remark Summary", page_icon="📊", initial_sidebar_state="expanded")
st.title('Daily Remark Summary')

@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip().str.upper()
    if 'TIME' not in df.columns:
        return pd.DataFrame()
    df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
    df = df[df['DATE'].dt.weekday != 6]  # Exclude Sundays
    return df

def to_excel(summary_groups):
    output = BytesIO()
    with ExcelWriter(output, engine='xlsxwriter', date_format='yyyy-mm-dd') as writer:
        workbook = writer.book
        formats = {
            'title': workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFFF00'}),
            'center': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1}),
            'header': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bg_color': 'red', 'font_color': 'white', 'bold': True}),
            'comma': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0'}),
            'percent': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.00%'}),
            'date': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': 'yyyy-mm-dd'}),
            'time': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': 'hh:mm:ss'})
        }

        for sheet_name, df_dict in summary_groups.items():
            worksheet = workbook.add_worksheet(str(sheet_name))
            current_row = 0

            for title, df in df_dict.items():
                if df.empty:
                    continue
                df_for_excel = df.copy()
                for col in ['PENETRATION RATE (%)', 'CONNECTED RATE (%)', 'CONTACT RATE (%)', 'PTP RATE', 'CALL DROP RATIO #']:
                    if col in df_for_excel.columns:
                        df_for_excel[col] = df_for_excel[col].str.rstrip('%').astype(float) / 100

                worksheet.merge_range(current_row, 0, current_row, len(df.columns) - 1, title, formats['title'])
                current_row += 1

                for col_num, col_name in enumerate(df_for_excel.columns):
                    worksheet.write(current_row, col_num, col_name, formats['header'])
                    max_len = max(df_for_excel[col_name].astype(str).str.len().max(), len(col_name)) + 2
                    worksheet.set_column(col_num, col_num, max_len)

                current_row += 1

                for row_num in range(len(df_for_excel)):
                    for col_num, col_name in enumerate(df_for_excel.columns):
                        value = df_for_excel.iloc[row_num, col_num]
                        if col_name == 'DATE':
                            worksheet.write_datetime(current_row + row_num, col_num, value, formats['date'])
                        elif col_name in ['TOTAL PTP AMOUNT', 'TOTAL BALANCE', 'PTP AMOUNT', 'PTP COUNT', 'MANUAL DIALED', 
                                         'MANUAL CONNECTED', 'PD CONNECTED', 'MANUAL PTP COUNT', 'PD PTP COUNT', 
                                         'MANUAL PTP AMOUNT', 'PD PTP AMOUNT', 'RPC #', 'TPC', 'VM', 'CONFIRMED COUNT', 
                                         'CONFIRMED AMOUNT', 'PTP CONNECTED', 'RPC MANUAL', 'RPC PD', 'CONNECTED COUNT',
                                         'PTP STATUS COUNT', 'PTP STATUS AMOUNT', 'CONFIRMED CONNECTED', 'NEGATIVE CONNECTED',
                                         'CALL DROP #', 'SYSTEM DROP']:
                            worksheet.write(current_row + row_num, col_num, value, formats['comma'])
                        elif col_name in ['PENETRATION RATE (%)', 'CONNECTED RATE (%)', 'CONTACT RATE (%)', 'PTP RATE', 'CALL DROP RATIO #']:
                            worksheet.write(current_row + row_num, col_num, value, formats['percent'])
                        elif col_name in ['TOTAL TALK TIME', 'TALK TIME AVE', 'TIME', 'MANUAL TALKTIME', 'PD TALK TIME']:
                            worksheet.write(current_row + row_num, col_num, value, formats['time'])
                        else:
                            worksheet.write(current_row + row_num, col_num, value, formats['center'])

                current_row += len(df_for_excel) + 2

    return output.getvalue()

uploaded_files = st.sidebar.file_uploader("Upload Daily Remark Files", type="xlsx", accept_multiple_files=True)

def process_file(df):
    if df.empty:
        return df
    df = df[df['REMARK BY'] != 'SPMADRID']
    df = df[~df['DEBTOR'].str.contains("DEFAULT_LEAD_", case=False, na=False)]
    df = df[~df['STATUS'].str.contains('ABORT', na=False)]
    df = df[~df['REMARK'].str.contains(r'1_\d{11} - PTP NEW', case=False, na=False, regex=True)]
    df = df[~df['REMARK'].str.contains('Broadcast', case=False, na=False)]
    excluded_remarks = ["Broken Promise", "New files imported", "Updates when case reassign to another collector", 
                        "NDF IN ICS", "FOR PULL OUT (END OF HANDLING PERIOD)", "END OF HANDLING PERIOD", "New Assignment -", "File Unhold"]
    escaped_remarks = [re.escape(remark) for remark in excluded_remarks]
    df = df[~df['REMARK'].str.contains('|'.join(escaped_remarks), case=False, na=False)]
    df['CARD NO.'] = df['CARD NO.'].astype(str)
    return df

def format_seconds_to_hms(seconds):
    seconds = int(seconds)
    hours, minutes = seconds // 3600, (seconds % 3600) // 60
    return f"{hours:02d}:{minutes:02d}:{seconds % 60:02d}"

def calculate_summary(df, remark_types, manual_correction=False):
    summary_columns = ['DATE', 'CLIENT', 'COLLECTORS', 'ACCOUNTS', 'TOTAL DIALED', 'TOTAL TALK TIME', 'TALK TIME AVE', 
                      'CONNECTED #', 'CONNECTED ACC', 'CONNECTED AVE', 'RPC #', 'TPC', 'VM', 'PTP CONNECTED', 
                      'CONFIRMED CONNECTED', 'NEGATIVE CONNECTED', 'CALL DROP #', 'SYSTEM DROP', 'PTP ACC', 
                      'TOTAL PTP AMOUNT', 'CONFIRMED COUNT', 'CONFIRMED AMOUNT', 'TOTAL BALANCE', 
                      'PENETRATION RATE (%)', 'CONNECTED RATE (%)', 'CONTACT RATE (%)', 'PTP RATE', 'CALL DROP RATIO #']
    summary_table = pd.DataFrame(columns=summary_columns)
    
    if df.empty:
        return summary_table

    df_filtered = df[df['REMARK TYPE'].isin(remark_types)].copy()
    df_filtered['DATE'] = pd.to_datetime(df_filtered['DATE']).dt.date

    if 'CALL STATUS' in df_filtered.columns and 'REMARK' in df_filtered.columns:
        df_excluded = df_filtered[~((df_filtered['CALL STATUS'].str.contains('OTHERS', case=False, na=False)) & 
                                    (df_filtered['REMARK'].str.contains('@', case=False, na=False)))]
    else:
        df_excluded = df_filtered.copy()

    summary_rows = []
    for (date, client), group in df_excluded.groupby(['DATE', 'CLIENT']):
        group_for_connected = df_filtered[(df_filtered['DATE'] == date) & (df_filtered['CLIENT'] == client)]
        collectors = group[group['CALL DURATION'].notna()]['REMARK BY'].nunique()
        if collectors == 0:
            continue
        accounts = group['ACCOUNT NO.'].nunique()
        total_dialed = group['ACCOUNT NO.'].count()
        connected_acc = group_for_connected[(pd.to_numeric(group_for_connected['TALK TIME DURATION'], errors='coerce') > 0) | 
                                           ((group_for_connected['STATUS'].str.contains('DROPPED', na=False)) & 
                                            (group_for_connected['REMARK BY'] == 'SYSTEM'))]['ACCOUNT NO.'].count()
        connected = group_for_connected[(pd.to_numeric(group_for_connected['TALK TIME DURATION'], errors='coerce') > 0) | 
                                       ((group_for_connected['STATUS'].str.contains('DROPPED', na=False)) & 
                                        (group_for_connected['REMARK BY'] == 'SYSTEM'))]['ACCOUNT NO.'].nunique()
        connected_group = group_for_connected[(pd.to_numeric(group_for_connected['TALK TIME DURATION'], errors='coerce') > 0) | 
                                             ((group_for_connected['STATUS'].str.contains('DROPPED', na=False)) & 
                                              (group_for_connected['REMARK BY'] == 'SYSTEM'))]
        rpc = connected_group[connected_group['STATUS'].str.contains('RPC', case=False, na=False)]['ACCOUNT NO.'].nunique()
        tpc = connected_group[connected_group['STATUS'].str.contains('TPC', case=False, na=False)]['ACCOUNT NO.'].nunique()
        vm = connected_group[connected_group['CALL STATUS'].str.contains('VM', case=False, na=False)]['ACCOUNT NO.'].nunique()
        ptp_connected = connected_group[(connected_group['STATUS'].str.contains('PTP', na=False)) & 
                                       (pd.to_numeric(group_for_connected['PTP AMOUNT'], errors='coerce') > 0)]['ACCOUNT NO.'].nunique()
        confirmed_connected = connected_group[pd.to_numeric(group_for_connected['CLAIM PAID AMOUNT'], errors='coerce') > 0]['ACCOUNT NO.'].nunique() if 'CLAIM PAID AMOUNT' in group_for_connected.columns else 0
        system_drop = group[(group['STATUS'].str.contains('DROPPED', na=False)) & (group['REMARK BY'] == 'SYSTEM')]['ACCOUNT NO.'].count()
        call_drop_count = group[(group['STATUS'].str.contains('NEGATIVE CALLOUTS - DROP CALL|NEGATIVE_CALLOUTS - DROPPED_CALL', na=False)) & 
                              (~group['REMARK BY'].str.upper().isin(['SYSTEM']))]['ACCOUNT NO.'].count()
        negative_connected = connected_acc - (rpc + tpc + vm + ptp_connected + confirmed_connected + call_drop_count + system_drop)
        penetration_rate = f"{(total_dialed / accounts * 100):.2f}%" if accounts else "0.00%"
        connected_rate = f"{(connected_acc / total_dialed * 100):.2f}%" if total_dialed else "0.00%"
        contact_rate = f"{(connected_acc / accounts * 100):.2f}%" if accounts else "0.00%"
        ptp_acc = group[(group['STATUS'].str.contains('PTP', na=False)) & (pd.to_numeric(group['PTP AMOUNT'], errors='coerce') > 0)]['ACCOUNT NO.'].nunique()
        ptp_rate = f"{(ptp_acc / connected * 100):.2f}%" if connected else "0.00%"
        total_ptp_amount = pd.to_numeric(group[(group['STATUS'].str.contains('PTP', na=False)) & (pd.to_numeric(group['PTP AMOUNT'], errors='coerce') > 0)]['PTP AMOUNT'], errors='coerce').sum()
        total_balance = pd.to_numeric(group[(pd.to_numeric(group['PTP AMOUNT'], errors='coerce') > 0)]['BALANCE'], errors='coerce').sum()
        call_drop_ratio = f"{(call_drop_count / connected_acc * 100):.2f}%" if manual_correction and connected_acc else \
                         f"{(system_drop / connected_acc * 100):.2f}%" if connected_acc else "0.00%"
        total_talk_seconds = pd.to_numeric(group['TALK TIME DURATION'], errors='coerce').sum()
        total_talk_time = format_seconds_to_hms(total_talk_seconds)
        talk_time_ave = format_seconds_to_hms(total_talk_seconds / collectors) if collectors else "00:00:00"
        connected_ave = round(connected_acc / collectors, 2) if collectors else 0

        confirmed_count = 0
        confirmed_amount = 0
        if 'CLAIM PAID AMOUNT' in group.columns:
            confirmed_df = connected_group[pd.to_numeric(connected_group['CLAIM PAID AMOUNT'], errors='coerce') > 0][['DATE', 'ACCOUNT NO.', 'CLAIM PAID AMOUNT']].drop_duplicates()
            confirmed_count = len(confirmed_df)
            confirmed_amount = pd.to_numeric(confirmed_df['CLAIM PAID AMOUNT'], errors='coerce').sum()

        summary_rows.append({
            'DATE': date,
            'CLIENT': client,
            'COLLECTORS': collectors,
            'ACCOUNTS': accounts,
            'TOTAL DIALED': total_dialed,
            'TOTAL TALK TIME': total_talk_time,
            'TALK TIME AVE': talk_time_ave,
            'CONNECTED #': connected,
            'CONNECTED ACC': connected_acc,
            'CONNECTED AVE': connected_ave,
            'RPC #': rpc,
            'TPC': tpc,
            'VM': vm,
            'PTP CONNECTED': ptp_connected,
            'CONFIRMED CONNECTED': confirmed_connected,
            'NEGATIVE CONNECTED': negative_connected,
            'CALL DROP #': call_drop_count,
            'SYSTEM DROP': system_drop,
            'PTP ACC': ptp_acc,
            'TOTAL PTP AMOUNT': total_ptp_amount,
            'CONFIRMED COUNT': confirmed_count,
            'CONFIRMED AMOUNT': confirmed_amount,
            'TOTAL BALANCE': total_balance,
            'PENETRATION RATE (%)': penetration_rate,
            'CONNECTED RATE (%)': connected_rate,
            'CONTACT RATE (%)': contact_rate,
            'PTP RATE': ptp_rate,
            'CALL DROP RATIO #': call_drop_ratio
        })
    
    if summary_rows:
        summary_table = pd.DataFrame(summary_rows, columns=summary_columns)
    return summary_table.sort_values(by=['DATE'])

def calculate_productivity_per_agent(df):
    summary_columns = ['DATE', 'COLLECTOR', 'CAMPAIGN', 'MANUAL DIALED', 'MANUAL CONNECTED', 'PD CONNECTED', 
                      'MANUAL TALKTIME', 'PD TALK TIME', 'RPC MANUAL', 'RPC PD', 'MANUAL PTP COUNT', 'PD PTP COUNT', 
                      'MANUAL PTP AMOUNT', 'PD PTP AMOUNT']
    summary_table = pd.DataFrame(columns=summary_columns)
    
    if df.empty:
        return summary_table

    excluded_remarks = ["Broken Promise", "New files imported", "Updates when case reassign to another collector", 
                        "NDF IN ICS", "FOR PULL OUT (END OF HANDLING PERIOD)", "END OF HANDLING PERIOD", "New Assignment -", "File Unhold"]
    escaped_remarks = [re.escape(remark) for remark in excluded_remarks]
    
    collectors_with_excluded_remarks = df[df['REMARK'].str.contains('|'.join(escaped_remarks), case=False, na=False)]['REMARK BY'].unique()
    
    df_filtered = df[~df['REMARK BY'].isin(collectors_with_excluded_remarks)]
    
    if 'CALL STATUS' in df_filtered.columns and 'REMARK' in df_filtered.columns:
        df_filtered = df_filtered[~((df_filtered['CALL STATUS'].str.contains('OTHERS', case=False, na=False)) & 
                                   (df_filtered['REMARK'].str.contains('@', case=False, na=False)))]
    
    df_filtered = df_filtered[~df_filtered['REMARK BY'].str.upper().eq('SYSTEM')]
    
    df_filtered['DATE'] = pd.to_datetime(df_filtered['DATE']).dt.date

    summary_rows = []
    for (date, collector, campaign), group in df_filtered.groupby(['DATE', 'REMARK BY', 'CLIENT']):
        manual_dialed = group[group['REMARK TYPE'] == 'Outgoing']['ACCOUNT NO.'].count()
        manual_connected = group[(group['REMARK TYPE'] == 'Outgoing') & (pd.to_numeric(group['TALK TIME DURATION'], errors='coerce') > 0)]['ACCOUNT NO.'].count()
        pd_connected = group[(group['REMARK TYPE'].isin(['Predictive', 'Follow Up'])) & (pd.to_numeric(group['TALK TIME DURATION'], errors='coerce') > 0)]['ACCOUNT NO.'].count()
        manual_talktime_seconds = pd.to_numeric(group[group['REMARK TYPE'] == 'Outgoing']['TALK TIME DURATION'], errors='coerce').sum()
        pd_talktime_seconds = pd.to_numeric(group[group['REMARK TYPE'].isin(['Predictive', 'Follow Up'])]['TALK TIME DURATION'], errors='coerce').sum()
        manual_talktime = format_seconds_to_hms(manual_talktime_seconds) if manual_talktime_seconds > 0 else "00:00:00"
        pd_talk_time = format_seconds_to_hms(pd_talktime_seconds) if pd_talktime_seconds > 0 else "00:00:00"
        rpc_manual = group[(group['REMARK TYPE'] == 'Outgoing') & (pd.to_numeric(group['TALK TIME DURATION'], errors='coerce') > 0) & 
                          (group['STATUS'].str.contains('RPC', case=False, na=False))]['ACCOUNT NO.'].nunique()
        rpc_pd = group[(group['REMARK TYPE'].isin(['Predictive', 'Follow Up'])) & (pd.to_numeric(group['TALK TIME DURATION'], errors='coerce') > 0) & 
                       (group['STATUS'].str.contains('RPC', case=False, na=False))]['ACCOUNT NO.'].nunique()
        manual_ptp_count = group[(group['REMARK TYPE'] == 'Outgoing') & (pd.to_numeric(group['PTP AMOUNT'], errors='coerce') > 0)]['ACCOUNT NO.'].nunique()
        pd_ptp_count = group[(group['REMARK TYPE'].isin(['Predictive', 'Follow Up'])) & (pd.to_numeric(group['PTP AMOUNT'], errors='coerce') > 0)]['ACCOUNT NO.'].nunique()
        manual_ptp_amount = pd.to_numeric(group[(group['REMARK TYPE'] == 'Outgoing') & (pd.to_numeric(group['PTP AMOUNT'], errors='coerce') > 0)]['PTP AMOUNT'], errors='coerce').sum()
        pd_ptp_amount = pd.to_numeric(group[(group['REMARK TYPE'].isin(['Predictive', 'Follow Up'])) & (pd.to_numeric(group['PTP AMOUNT'], errors='coerce') > 0)]['PTP AMOUNT'], errors='coerce').sum()

        summary_rows.append({
            'DATE': date,
            'COLLECTOR': collector,
            'CAMPAIGN': campaign,
            'MANUAL DIALED': manual_dialed,
            'MANUAL CONNECTED': manual_connected,
            'PD CONNECTED': pd_connected,
            'MANUAL TALKTIME': manual_talktime,
            'PD TALK TIME': pd_talk_time,
            'RPC MANUAL': rpc_manual,
            'RPC PD': rpc_pd,
            'MANUAL PTP COUNT': manual_ptp_count,
            'PD PTP COUNT': pd_ptp_count,
            'MANUAL PTP AMOUNT': manual_ptp_amount,
            'PD PTP AMOUNT': pd_ptp_amount
        })
    
    if summary_rows:
        summary_table = pd.DataFrame(summary_rows)
    return summary_table.sort_values(by=['DATE', 'COLLECTOR', 'CAMPAIGN'])

def calculate_ptp_hourly_summary(df, year, month):
    summary_columns = ['CAMPAIGN', 'TIME RANGE', 'PTP COUNT', 'PTP AMOUNT']
    summary_table = pd.DataFrame(columns=summary_columns)
    
    if df.empty:
        return summary_table

    ptp_df = df[(df['STATUS'].str.contains('PTP', case=False, na=False)) & (pd.to_numeric(df['PTP AMOUNT'], errors='coerce') > 0)].copy()
    if ptp_df.empty:
        return summary_table

    required_columns = ['DATE', 'ACCOUNT NO.', 'PTP AMOUNT', 'CLIENT', 'TIME']
    missing_columns = [col for col in required_columns if col not in ptp_df.columns]
    if missing_columns:
        return summary_table

    ptp_df = ptp_df[required_columns].drop_duplicates(subset=['DATE', 'ACCOUNT NO.', 'PTP AMOUNT', 'CLIENT', 'TIME'])
    ptp_df['DATE'] = pd.to_datetime(ptp_df['DATE'], errors='coerce')
    ptp_df = ptp_df[ptp_df['DATE'].notna()]
    ptp_df = ptp_df[(ptp_df['DATE'].dt.year == year) & (ptp_df['DATE'].dt.month == month)]
    if ptp_df.empty:
        return summary_table

    if 'TIME' not in ptp_df.columns:
        return summary_table

    if all(isinstance(x, datetime.time) for x in ptp_df['TIME'].dropna()):
        ptp_df['TIME'] = ptp_df['TIME']
    elif pd.to_numeric(ptp_df['TIME'], errors='coerce').notna().any():
        ptp_df['TIME'] = pd.to_datetime(ptp_df['TIME'].astype(float), unit='d', origin='1899-12-30', errors='coerce').dt.time
    else:
        time_formats = ['%I:%M:%S %p', '%H:%M:%S', '%I:%M %p', '%H:%M']
        for fmt in time_formats:
            try:
                ptp_df['TIME'] = pd.to_datetime(ptp_df['TIME'], format=fmt, errors='coerce').dt.time
                if not ptp_df['TIME'].isna().all():
                    break
            except ValueError:
                continue

    ptp_df = ptp_df[ptp_df['TIME'].notna()]
    if ptp_df.empty:
        return summary_table

    ptp_df['HOUR'] = ptp_df['TIME'].apply(lambda x: x.hour)
    ptp_df = ptp_df[(ptp_df['HOUR'] >= 7) & (ptp_df['HOUR'] < 20)]
    if ptp_df.empty:
        return summary_table

    ptp_df = ptp_df[ptp_df['CLIENT'].notna() & (ptp_df['CLIENT'].str.strip() != '')]

    time_range_order = []
    for hour in range(7, 20):
        start_hour_12 = hour if hour <= 12 else hour - 12
        end_hour_12 = (hour + 1) if (hour + 1) <= 12 else (hour + 1) - 12
        start_suffix = "AM" if hour < 12 else "PM"
        end_suffix = "AM" if (hour + 1) < 12 else "PM"
        if hour == 12:
            start_suffix = "PM"
        if (hour + 1) == 12:
            end_suffix = "PM"
        time_range = f"{start_hour_12:02d}:00 {start_suffix} - {end_hour_12:02d}:00 {end_suffix}"
        time_range_order.append(time_range)

    summary_data = []
    for hour in range(7, 20):
        start_hour_12 = hour if hour <= 12 else hour - 12
        end_hour_12 = (hour + 1) if (hour + 1) <= 12 else (hour + 1) - 12
        start_suffix = "AM" if hour < 12 else "PM"
        end_suffix = "AM" if (hour + 1) < 12 else "PM"
        if hour == 12:
            start_suffix = "PM"
        if (hour + 1) == 12:
            end_suffix = "PM"
        time_range = f"{start_hour_12:02d}:00 {start_suffix} - {end_hour_12:02d}:00 {end_suffix}"

        hour_df = ptp_df[ptp_df['HOUR'] == hour]
        if hour_df.empty:
            continue

        for client, group in hour_df.groupby('CLIENT'):
            ptp_count = group['ACCOUNT NO.'].nunique()
            ptp_amount = pd.to_numeric(group['PTP AMOUNT'], errors='coerce').sum()

            if pd.isna(client) or not client.strip():
                continue

            summary_data.append({
                'CAMPAIGN': client,
                'TIME RANGE': time_range,
                'PTP COUNT': ptp_count,
                'PTP AMOUNT': ptp_amount
            })

    if summary_data:
        summary_table = pd.DataFrame(summary_data, columns=summary_columns)
        summary_table = summary_table.dropna(subset=['CAMPAIGN', 'TIME RANGE'])
        summary_table = summary_table[summary_table['TIME RANGE'].isin(time_range_order)]
        summary_table = summary_table[summary_table['CAMPAIGN'].str.strip() != '']
        summary_table = summary_table.groupby(['CAMPAIGN', 'TIME RANGE'], as_index=False).agg({
            'PTP COUNT': 'sum',
            'PTP AMOUNT': 'sum'
        })
        summary_table['TIME RANGE'] = pd.Categorical(summary_table['TIME RANGE'], categories=time_range_order, ordered=True)
        summary_table = summary_table.sort_values(by=['CAMPAIGN', 'TIME RANGE'])
    return summary_table

def calculate_connected_hourly_summary(df, year, month):
    summary_columns = ['CAMPAIGN', 'TIME RANGE', 'CONNECTED COUNT']
    summary_table = pd.DataFrame(columns=summary_columns)
    
    if df.empty:
        return summary_table

    connected_df = df[pd.to_numeric(df['TALK TIME DURATION'], errors='coerce') > 0].copy()
    if connected_df.empty:
        return summary_table

    connected_df['DATE'] = pd.to_datetime(connected_df['DATE'], errors='coerce')
    connected_df = connected_df[connected_df['DATE'].notna()]
    connected_df = connected_df[(connected_df['DATE'].dt.year == year) & (connected_df['DATE'].dt.month == month)]
    if connected_df.empty:
        return summary_table

    if 'TIME' not in connected_df.columns:
        return summary_table

    if all(isinstance(x, datetime.time) for x in connected_df['TIME'].dropna()):
        connected_df['TIME'] = connected_df['TIME']
    elif pd.to_numeric(connected_df['TIME'], errors='coerce').notna().any():
        connected_df['TIME'] = pd.to_datetime(connected_df['TIME'].astype(float), unit='d', origin='1899-12-30', errors='coerce').dt.time
    else:
        time_formats = ['%I:%M:%S %p', '%H:%M:%S', '%I:%M %p', '%H:%M']
        for fmt in time_formats:
            try:
                connected_df['TIME'] = pd.to_datetime(connected_df['TIME'], format=fmt, errors='coerce').dt.time
                if not connected_df['TIME'].isna().all():
                    break
            except ValueError:
                continue

    connected_df = connected_df[connected_df['TIME'].notna()]
    if connected_df.empty:
        return summary_table

    connected_df['HOUR'] = connected_df['TIME'].apply(lambda x: x.hour)
    connected_df = connected_df[(connected_df['HOUR'] >= 7) & (connected_df['HOUR'] < 20)]
    if connected_df.empty:
        return summary_table

    connected_df = connected_df[connected_df['CLIENT'].notna() & (connected_df['CLIENT'].str.strip() != '')]

    time_range_order = []
    for hour in range(7, 20):
        start_hour_12 = hour if hour <= 12 else hour - 12
        end_hour_12 = (hour + 1) if (hour + 1) <= 12 else (hour + 1) - 12
        start_suffix = "AM" if hour < 12 else "PM"
        end_suffix = "AM" if (hour + 1) < 12 else "PM"
        if hour == 12:
            start_suffix = "PM"
        if (hour + 1) == 12:
            end_suffix = "PM"
        time_range = f"{start_hour_12:02d}:00 {start_suffix} - {end_hour_12:02d}:00 {end_suffix}"
        time_range_order.append(time_range)

    summary_data = []
    for hour in range(7, 20):
        start_hour_12 = hour if hour <= 12 else hour - 12
        end_hour_12 = (hour + 1) if (hour + 1) <= 12 else (hour + 1) - 12
        start_suffix = "AM" if hour < 12 else "PM"
        end_suffix = "AM" if (hour + 1) < 12 else "PM"
        if hour == 12:
            start_suffix = "PM"
        if (hour + 1) == 12:
            end_suffix = "PM"
        time_range = f"{start_hour_12:02d}:00 {start_suffix} - {end_hour_12:02d}:00 {end_suffix}"

        hour_df = connected_df[connected_df['HOUR'] == hour]
        if hour_df.empty:
            continue

        for client, group in hour_df.groupby('CLIENT'):
            connected_count = len(group)

            if pd.isna(client) or not client.strip():
                continue

            summary_data.append({
                'CAMPAIGN': client,
                'TIME RANGE': time_range,
                'CONNECTED COUNT': connected_count
            })

    if summary_data:
        summary_table = pd.DataFrame(summary_data, columns=summary_columns)
        summary_table = summary_table.dropna(subset=['CAMPAIGN', 'TIME RANGE'])
        summary_table = summary_table[summary_table['TIME RANGE'].isin(time_range_order)]
        summary_table = summary_table[summary_table['CAMPAIGN'].str.strip() != '']
        summary_table = summary_table.groupby(['CAMPAIGN', 'TIME RANGE'], as_index=False).agg({
            'CONNECTED COUNT': 'sum'
        })
        summary_table['TIME RANGE'] = pd.Categorical(summary_table['TIME RANGE'], categories=time_range_order, ordered=True)
        summary_table = summary_table.sort_values(by=['CAMPAIGN', 'TIME RANGE'])
    return summary_table

def calculate_ptp_trend_summary(df):
    summary_columns = ['DATE RANGE', 'RANK', 'CAMPAIGN', 'PTP STATUS', 'PTP STATUS COUNT', 'PTP STATUS AMOUNT', 'BALANCE']
    summary_table = pd.DataFrame(columns=summary_columns)
    
    if df.empty:
        return summary_table

    ptp_df = df[pd.to_numeric(df['PTP AMOUNT'], errors='coerce') > 0].copy()
    ptp_df = ptp_df[ptp_df['STATUS'].str.contains('PTP', case=False, na=False)]
    ptp_df = ptp_df[~ptp_df['STATUS'].str.contains('PTP FF UP', case=False, na=False)]
    
    if ptp_df.empty:
        return summary_table

    required_columns = ['DATE', 'ACCOUNT NO.', 'PTP AMOUNT', 'CLIENT', 'STATUS', 'BALANCE']
    missing_columns = [col for col in required_columns if col not in ptp_df.columns]
    if missing_columns:
        return summary_table

    ptp_df = ptp_df[required_columns].drop_duplicates(subset=['DATE', 'ACCOUNT NO.', 'PTP AMOUNT'])
    ptp_df['DATE'] = pd.to_datetime(ptp_df['DATE'], errors='coerce').dt.date
    ptp_df = ptp_df[ptp_df['DATE'].notna()]
    if ptp_df.empty:
        return summary_table

    ptp_df = ptp_df[ptp_df['CLIENT'].notna() & (ptp_df['CLIENT'].str.strip() != '')]
    ptp_df = ptp_df[ptp_df['STATUS'].notna() & (ptp_df['STATUS'].str.strip() != '')]
    if ptp_df.empty:
        return summary_table

    min_date = ptp_df['DATE'].min()
    max_date = ptp_df['DATE'].max()
    if pd.isna(min_date) or pd.isna(max_date):
        return summary_table
    date_range = f"{min_date.strftime('%b %d %Y')} - {max_date.strftime('%b %d %Y')}"

    summary_data = ptp_df.groupby(['CLIENT', 'STATUS']).agg({
        'ACCOUNT NO.': 'count',
        'PTP AMOUNT': 'sum',
        'BALANCE': 'sum'
    }).reset_index()
    summary_data.columns = ['CAMPAIGN', 'PTP STATUS', 'PTP STATUS COUNT', 'PTP STATUS AMOUNT', 'BALANCE']

    summary_data['RANK'] = summary_data.groupby('CAMPAIGN')['PTP STATUS COUNT'].rank(method='min', ascending=False).astype(int)
    summary_data['DATE RANGE'] = date_range

    summary_data = summary_data[summary_columns]
    summary_data = summary_data.sort_values(by=['RANK', 'CAMPAIGN', 'PTP STATUS'])
    return summary_data

if uploaded_files:
    all_combined = []
    all_predictive = []
    all_manual = []
    all_productivity = []
    all_ptp_trend = []
    monthly_ptp_hourly = {}
    monthly_connected_hourly = {}

    time_range_order = []
    for hour in range(7, 20):
        start_hour_12 = hour if hour <= 12 else hour - 12
        end_hour_12 = (hour + 1) if (hour + 1) <= 12 else (hour + 1) - 12
        start_suffix = "AM" if hour < 12 else "PM"
        end_suffix = "AM" if (hour + 1) < 12 else "PM"
        if hour == 12:
            start_suffix = "PM"
        if (hour + 1) == 12:
            end_suffix = "PM"
        time_range = f"{start_hour_12:02d}:00 {start_suffix} - {end_hour_12:02d}:00 {end_suffix}"
        time_range_order.append(time_range)

    for idx, file in enumerate(uploaded_files, 1):
        df = load_data(file)
        df = process_file(df)
        df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
        df = df[df['DATE'].notna()]
        
        combined_summary = calculate_summary(df, ['Predictive', 'Follow Up', 'Outgoing'])
        predictive_summary = calculate_summary(df, ['Predictive', 'Follow Up'])
        manual_summary = calculate_summary(df, ['Outgoing'], manual_correction=True)
        productivity_summary = calculate_productivity_per_agent(df)
        ptp_trend_summary = calculate_ptp_trend_summary(df)

        all_combined.append(combined_summary)
        all_predictive.append(predictive_summary)
        all_manual.append(manual_summary)
        all_productivity.append(productivity_summary)
        all_ptp_trend.append(ptp_trend_summary)

        # Group by year and month for PTP and Connected Hourly summaries
        df['YEAR'] = df['DATE'].dt.year
        df['MONTH'] = df['DATE'].dt.month
        for (year, month), month_df in df.groupby(['YEAR', 'MONTH']):
            month_start = pd.Timestamp(year=year, month=month, day=1)
            month_end = month_start + pd.offsets.MonthEnd(0)
            month_name = month_start.strftime('%b %Y')
            date_range = f"{month_start.strftime('%B %d, %Y')} - {month_end.strftime('%B %d, %Y')}"

            ptp_hourly_summary = calculate_ptp_hourly_summary(month_df, year, month)
            connected_hourly_summary = calculate_connected_hourly_summary(month_df, year, month)

            month_key = (year, month, month_name, date_range)
            if month_key not in monthly_ptp_hourly:
                monthly_ptp_hourly[month_key] = []
                monthly_connected_hourly[month_key] = []
            
            monthly_ptp_hourly[month_key].append(ptp_hourly_summary)
            monthly_connected_hourly[month_key].append(connected_hourly_summary)

    # Combine summaries for non-monthly tables
    combined_summary = pd.concat(all_combined, ignore_index=True).sort_values(by=['DATE'])
    predictive_summary = pd.concat(all_predictive, ignore_index=True).sort_values(by=['DATE'])
    manual_summary = pd.concat(all_manual, ignore_index=True).sort_values(by=['DATE'])
    productivity_summary = pd.concat(all_productivity, ignore_index=True).sort_values(by=['DATE', 'COLLECTOR', 'CAMPAIGN'])
    ptp_trend_summary = pd.concat(all_ptp_trend, ignore_index=True).sort_values(by=['RANK', 'CAMPAIGN', 'PTP STATUS'])

    # Process monthly PTP and Connected Hourly summaries
    summary_groups = {
        'Combined': {'Combined Summary': combined_summary},
        'Predictive': {'Predictive Summary': predictive_summary},
        'Manual': {'Manual Summary': manual_summary},
        'Productivity': {'Productivity Per Agent Summary': productivity_summary},
        'PTP Trend': {'PTP Trend Summary': ptp_trend_summary}
    }

    for month_key, ptp_dfs in monthly_ptp_hourly.items():
        year, month, month_name, date_range = month_key
        all_ptp_hourly_df = pd.concat(ptp_dfs, ignore_index=True)
        
        all_ptp_hourly_df = all_ptp_hourly_df.dropna(subset=['CAMPAIGN', 'TIME RANGE'])
        all_ptp_hourly_df = all_ptp_hourly_df[all_ptp_hourly_df['TIME RANGE'].isin(time_range_order)]
        all_ptp_hourly_df = all_ptp_hourly_df.drop_duplicates(subset=['CAMPAIGN', 'TIME RANGE'])
        all_ptp_hourly_df['PTP COUNT'] = pd.to_numeric(all_ptp_hourly_df['PTP COUNT'], errors='coerce').fillna(0)
        all_ptp_hourly_df['PTP AMOUNT'] = pd.to_numeric(all_ptp_hourly_df['PTP AMOUNT'], errors='coerce').fillna(0)
        all_ptp_hourly_df['TIME RANGE'] = all_ptp_hourly_df['TIME RANGE'].astype(str)

        all_ptp_hourly_df = all_ptp_hourly_df.groupby(['CAMPAIGN', 'TIME RANGE'], as_index=False).agg({
            'PTP COUNT': 'sum',
            'PTP AMOUNT': 'sum'
        })

        campaigns = all_ptp_hourly_df['CAMPAIGN'].dropna().unique()
        complete_index = pd.DataFrame(
            list(itertools.product(campaigns, time_range_order)),
            columns=['CAMPAIGN', 'TIME RANGE']
        )

        ptp_hourly_summary = complete_index.merge(
            all_ptp_hourly_df,
            on=['CAMPAIGN', 'TIME RANGE'],
            how='left'
        ).fillna({'PTP COUNT': 0, 'PTP AMOUNT': 0})
        ptp_hourly_summary['TIME RANGE'] = pd.Categorical(ptp_hourly_summary['TIME RANGE'], categories=time_range_order, ordered=True)
        ptp_hourly_summary = ptp_hourly_summary.sort_values(by=['CAMPAIGN', 'TIME RANGE'])

        # Add to summary_groups with month-specific sheet name
        summary_groups[f'PTP Hourly {month_name}'] = {f'PTP Hourly Summary ({date_range})': ptp_hourly_summary}

        st.write(f"## PTP Hourly Summary Table ({date_range})")
        st.write(ptp_hourly_summary)

    for month_key, connected_dfs in monthly_connected_hourly.items():
        year, month, month_name, date_range = month_key
        all_connected_hourly_df = pd.concat(connected_dfs, ignore_index=True)
        
        all_connected_hourly_df = all_connected_hourly_df.dropna(subset=['CAMPAIGN', 'TIME RANGE'])
        all_connected_hourly_df = all_connected_hourly_df[all_connected_hourly_df['TIME RANGE'].isin(time_range_order)]
        all_connected_hourly_df = all_connected_hourly_df.drop_duplicates(subset=['CAMPAIGN', 'TIME RANGE'])
        all_connected_hourly_df['CONNECTED COUNT'] = pd.to_numeric(all_connected_hourly_df['CONNECTED COUNT'], errors='coerce').fillna(0)
        all_connected_hourly_df['TIME RANGE'] = all_connected_hourly_df['TIME RANGE'].astype(str)

        all_connected_hourly_df = all_connected_hourly_df.groupby(['CAMPAIGN', 'TIME RANGE'], as_index=False).agg({
            'CONNECTED COUNT': 'sum'
        })

        campaigns = all_connected_hourly_df['CAMPAIGN'].dropna().unique()
        complete_index = pd.DataFrame(
            list(itertools.product(campaigns, time_range_order)),
            columns=['CAMPAIGN', 'TIME RANGE']
        )

        connected_hourly_summary = complete_index.merge(
            all_connected_hourly_df,
            on=['CAMPAIGN', 'TIME RANGE'],
            how='left'
        ).fillna({'CONNECTED COUNT': 0})
        connected_hourly_summary['TIME RANGE'] = pd.Categorical(connected_hourly_summary['TIME RANGE'], categories=time_range_order, ordered=True)
        connected_hourly_summary = connected_hourly_summary.sort_values(by=['CAMPAIGN', 'TIME RANGE'])

        # Add to summary_groups with month-specific sheet name
        summary_groups[f'Connected Hourly {month_name}'] = {f'Connected Hourly Summary ({date_range})': connected_hourly_summary}

        st.write(f"## Connected Hourly Summary Table ({date_range})")
        st.write(connected_hourly_summary)

    # Display non-monthly summaries
    st.write("## Overall Combined Summary Table")
    st.write(combined_summary)
    st.write("## Overall Predictive Summary Table")
    st.write(predictive_summary)
    st.write("## Overall Manual Summary Table")
    st.write(manual_summary)
    st.write("## Productivity Per Agent Summary Table")
    st.write(productivity_summary)
    st.write("## PTP Trend Summary Table")
    st.write(ptp_trend_summary)

    st.download_button(
        label="Download All Summaries as Excel",
        data=to_excel(summary_groups),
        file_name=f"Daily_Remark_Summary_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )