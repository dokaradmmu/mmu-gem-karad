import streamlit as st
import pandas as pd
import numpy as np
import io
import datetime
import pytz

# Timezone Enforcer Matrix
IST = pytz.timezone('Asia/Kolkata')
now_ist = datetime.datetime.now(IST)

st.set_page_config(page_title="MMU Gem Karad Division", page_icon="🇮🇳", layout="wide")

if "orphan_log" not in st.session_state:
    st.session_state.orphan_log = []

st.title("🇮🇳 MMU Gem Karad Division")
st.subheader("Enterprise Data Processing Utility — India Post Operations")
st.markdown(f"**System Chronometer Anchor (IST):** `{now_ist.strftime('%Y-%m-%d %H:%M:%S GMT+5:30')}`")

# Sidebar Configuration Slots
st.sidebar.header("📋 Workflow: 0. Daily Report")
p1_range = st.sidebar.text_input("Prompt 1: Shared Transit Range (From/To)", "01.05.2026 to 23.05.2026")
p2_date  = st.sidebar.text_input("Prompt 2: Single Performance Date", "30.05.2026")
p3_date  = st.sidebar.text_input("Prompt 3: Single App Usage Date", "30.05.2026")
p4_range = st.sidebar.text_input("Prompt 4: Independent DSS Range (From/To)", "01.05.2026 to 30.05.2026")
p5_date  = st.sidebar.text_input("Prompt 5: Single Financial Date", "30.05.2026")
p6_range = st.sidebar.text_input("Prompt 6: Independent COD Range (From/To)", "01.05.2026 to 30.05.2026")

master_file = st.sidebar.file_uploader("⚓ Master Skeleton: Updated Office Names 29.05.2026 (CSV Formatted)", type=["csv"])
slot_01 = st.sidebar.file_uploader("Slot 0.1: Speed Parcel (CSV)", type=["csv"])
slot_02 = st.sidebar.file_uploader("Slot 0.2: Registered Parcel (CSV)", type=["csv"])
slot_03 = st.sidebar.file_uploader("Slot 0.3: Speed Letter (CSV)", type=["csv"])
slot_04 = st.sidebar.file_uploader("Slot 0.4: Registered Letter (CSV)", type=["csv"])
slot_05 = st.sidebar.file_uploader("Slot 0.5: All Category (CSV)", type=["csv"])
slot_06 = st.sidebar.file_uploader("Slot 0.6: Delivery Productivity (CSV)", type=["csv"])
slot_07 = st.sidebar.file_uploader("Slot 0.7: DSS Usage Daily (CSV)", type=["csv"])
slot_08 = st.sidebar.file_uploader("Slot 0.8: DSS Usage Consolidated (CSV)", type=["csv"])
slot_09 = st.sidebar.file_uploader("Slot 0.9: COD Collection Daily (CSV)", type=["csv"])
slot_10 = st.sidebar.file_uploader("Slot 0.10: COD Collection Consolidated (CSV)", type=["csv"])

def ingest_and_sanitize(file_io, name_tag=""):
    if file_io is None:
        return None
    df = pd.read_csv(file_io, skip_blank_lines=True)
    df.columns = [str(c).strip() for c in df.columns]
    rename_target = {}
    for c in df.columns:
        if str(c).lower().replace(" ", "").replace("-", "_") in ['officeid', 'office_id']:
            rename_target[c] = 'office_id'
    if rename_target:
        df = df.rename(columns=rename_target)
    if 'office_id' not in df.columns:
        st.error(f"Critical Ingestion Error: Key field 'office_id' could not be resolved in slot {name_tag}.")
        st.stop()
    df['office_id'] = pd.to_numeric(df['office_id'], errors='coerce')
    df = df.dropna(subset=['office_id'])
    df = df[df['office_id'] > 0]
    df['office_id'] = df['office_id'].astype(int).astype(str).str.strip().str.lstrip('0')
    df = df[df['office_id'] != '']
    df = df[df['office_id'] != '0']
    for c in ['office_name', 'Office Name', 'customer-name', 'product-name']:
        if c in df.columns:
            df = df[~df[c].astype(str).str.contains('Summary|Total', case=False, na=False)]
    for col in df.select_dtypes(include=[object]).columns:
        df[col] = df[col].astype(str).str.strip().str.replace(r'^"|"$', '', regex=True)
    return df

if st.sidebar.button("🚀 Process Operational Records", use_container_width=True):
    st.session_state.orphan_log = []
    if master_file is None:
        st.error("Missing Structural Root Matrix: Please upload the Master Skeleton file.")
        st.stop()
        
    master_df = pd.read_csv(master_file)
    master_df.columns = [str(c).strip() for c in master_df.columns]
    for c in master_df.columns:
        if str(c).lower().replace(" ", "").replace("-", "_") in ['officeid', 'office_id']:
            master_df.rename(columns={c: 'office_id'}, inplace=True)
    master_df['office_id'] = pd.to_numeric(master_df['office_id'], errors='coerce')
    master_df = master_df.dropna(subset=['office_id'])
    master_df['office_id'] = master_df['office_id'].astype(int).astype(str).str.strip().str.lstrip('0')
    
    for c in ['Sub Division', 'Sub Office', 'Branch Office', 'office-type-code']:
        if c in master_df.columns:
            master_df[c] = master_df[c].astype(str).str.strip()
            
    master_df['Canonical_Office_Name'] = np.where(
        master_df['office-type-code'].isin(['SPO', 'HPO']),
        master_df['Sub Office'],
        master_df['Branch Office']
    )
    master_keys = set(master_df['office_id'].unique())
    
    raw_01 = ingest_and_sanitize(slot_01, "0.1 Speed Parcel")
    raw_02 = ingest_and_sanitize(slot_02, "0.2 Registered Parcel")
    raw_03 = ingest_and_sanitize(slot_03, "0.3 Speed Letter")
    raw_04 = ingest_and_sanitize(slot_04, "0.4 Registered Letter")
    raw_05 = ingest_and_sanitize(slot_05, "0.5 All Category")
    raw_06 = ingest_and_sanitize(slot_06, "0.6 Delivery Productivity")
    raw_07 = ingest_and_sanitize(slot_07, "0.7 DSS Usage Daily")
    raw_08 = ingest_and_sanitize(slot_08, "0.8 DSS Usage Consolidated")
    raw_09 = ingest_and_sanitize(slot_09, "0.9 COD Collection Daily")
    raw_10 = ingest_and_sanitize(slot_10, "0.10 COD Collection Consolidated")
    
    if raw_09 is None and raw_10 is not None:
        raw_09 = raw_10.copy()
        st.session_state.orphan_log.append("`[BUG B-02 ALERT]` Slot 0.9 fallback activated using Consolidated COD data.")

    TRANSIT_COLS = ['Received', 'Same Day Invoiced', 'D0 Delivered', 'D0 Redirected', 'D0 Returned', 'D1 Delivered', 'D1 Redirected', 'D1 Returned']
    def compile_transit_family(dfs_list):
        valid = [d for d in dfs_list if d is not None]
        if not valid:
            return pd.DataFrame(columns=['office_id'] + TRANSIT_COLS)
        concatenated = pd.concat(valid, ignore_index=True)
        for c in TRANSIT_COLS:
            concatenated[c] = pd.to_numeric(concatenated[c], errors='coerce').fillna(0.0) if c in concatenated.columns else 0.0
        return concatenated.groupby('office_id', as_index=False)[TRANSIT_COLS].sum()

    df_parcel_fam = compile_transit_family([raw_01, raw_02])
    df_doc_fam    = compile_transit_family([raw_03, raw_04])
    df_all_fam    = compile_transit_family([raw_05])

    if raw_06 is not None:
        for c in ['invoice-count', 'delivery-count', 'return-count', 'redirection-count', 'deposit-count']:
            raw_06[c] = pd.to_numeric(raw_06[c], errors='coerce').fillna(0.0) if c in raw_06.columns else 0.0
        raw_06['prod_num'] = raw_06['delivery-count'] + raw_06['return-count'] + raw_06['redirection-count']
        df_prod_ledger = raw_06.groupby('office_id', as_index=False)[['invoice-count', 'prod_num']].sum()
    else:
        df_prod_ledger = pd.DataFrame(columns=['office_id', 'invoice-count', 'prod_num'])

    def aggregate_dss(df):
        if df is not None:
            for c in ['total_pdm_art_count', 'total_dss_art_count']:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
            return df.groupby('office_id', as_index=False)[['total_pdm_art_count', 'total_dss_art_count']].sum()
        return pd.DataFrame(columns=['office_id', 'total_pdm_art_count', 'total_dss_art_count'])

    df_dss_d_ledger = aggregate_dss(raw_07)
    df_dss_c_ledger = aggregate_dss(raw_08)

    def aggregate_cod(df):
        if df is not None:
            for c in ['no_digital_count', 'no-cod-articles']:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
            return df.groupby('office_id', as_index=False)[['no_digital_count', 'no-cod-articles']].sum()
        return pd.DataFrame(columns=['office_id', 'no_digital_count', 'no-cod-articles'])

    df_cod_d_ledger = aggregate_cod(raw_09)
    df_cod_c_ledger = aggregate_cod(raw_10)

    core_ledger = master_df[['Sub Division', 'Sub Office', 'Branch Office', 'office_id', 'office-type-code', 'Canonical_Office_Name']].copy()
    core_ledger = core_ledger.merge(df_parcel_fam, on='office_id', how='left').rename(columns={c: f"par_{c}" for c in TRANSIT_COLS})
    core_ledger = core_ledger.merge(df_doc_fam, on='office_id', how='left').rename(columns={c: f"doc_{c}" for c in TRANSIT_COLS})
    core_ledger = core_ledger.merge(df_all_fam, on='office_id', how='left').rename(columns={c: f"all_{c}" for c in TRANSIT_COLS})
    core_ledger = core_ledger.merge(df_prod_ledger, on='office_id', how='left').rename(columns={'invoice-count': 'prod_denom', 'prod_num': 'prod_numer'})
    core_ledger = core_ledger.merge(df_dss_d_ledger, on='office_id', how='left').rename(columns={'total_pdm_art_count': 'dss_d_denom', 'total_dss_art_count': 'dss_d_numer'})
    core_ledger = core_ledger.merge(df_dss_c_ledger, on='office_id', how='left').rename(columns={'total_pdm_art_count': 'dss_c_denom', 'total_dss_art_count': 'dss_c_numer'})
    core_ledger = core_ledger.merge(df_cod_d_ledger, on='office_id', how='left').rename(columns={'no-cod-articles': 'cod_d_denom', 'no_digital_count': 'cod_d_numer'})
    core_ledger = core_ledger.merge(df_cod_c_ledger, on='office_id', how='left').rename(columns={'no-cod-articles': 'cod_c_denom', 'no_digital_count': 'cod_c_numer'})

    fill_columns = [col for col in core_ledger.columns if col not in ['Sub Division', 'Sub Office', 'Branch Office', 'office_id', 'office-type-code', 'Canonical_Office_Name']]
    core_ledger[fill_columns] = core_ledger[fill_columns].fillna(0.0)

    def generate_ratio_vector(df, numer, denom):
        return np.where(df[denom] > 0, df[numer] / df[denom], np.nan)

    core_ledger['val_par_d0'] = generate_ratio_vector(core_ledger, 'par_D0 Delivered+par_D0 Redirected+par_D0 Returned', 'par_Received')
    core_ledger['val_par_d1'] = generate_ratio_vector(core_ledger, 'par_D1 Delivered+par_D1 Redirected+par_D1 Returned', 'par_Received')
    core_ledger['val_doc_d0'] = generate_ratio_vector(core_ledger, 'doc_D0 Delivered+doc_D0 Redirected+doc_D0 Returned', 'doc_Received')
    core_ledger['val_doc_d1'] = generate_ratio_vector(core_ledger, 'doc_D1 Delivered+doc_D1 Redirected+doc_D1 Returned', 'doc_Received')
    core_ledger['val_all_d0'] = generate_ratio_vector(core_ledger, 'all_D0 Delivered+all_D0 Redirected+all_D0 Returned', 'all_Received')
    core_ledger['val_all_d1'] = generate_ratio_vector(core_ledger, 'all_D1 Delivered+all_D1 Redirected+all_D1 Returned', 'all_Received')
    core_ledger['val_all_rts'] = generate_ratio_vector(core_ledger, 'all_D0 Returned', 'all_Received')
    core_ledger['val_all_not_invoiced'] = core_ledger['all_Received'] - core_ledger['all_Same Day Invoiced']
    core_ledger['val_prod']  = generate_ratio_vector(core_ledger, 'prod_numer', 'prod_denom')
    core_ledger['val_dss_d'] = generate_ratio_vector(core_ledger, 'dss_d_numer', 'dss_d_denom')
    core_ledger['val_dss_c'] = generate_ratio_vector(core_ledger, 'dss_c_numer', 'dss_c_denom')
    core_ledger['val_cod_d'] = generate_ratio_vector(core_ledger, 'cod_d_numer', 'cod_d_denom')
    core_ledger['val_cod_c'] = generate_ratio_vector(core_ledger, 'cod_c_numer', 'cod_c_denom')

    core_ledger = core_ledger.sort_values(by=['Sub Division', 'Sub Office', 'Branch Office']).reset_index(drop=True)

    output_stream = io.BytesIO()
    excel_engine = pd.ExcelWriter(output_stream, engine='xlsxwriter')
    workbook_obj = excel_engine.book

    style_main_header = workbook_obj.add_format({'bg_color': '#1F497D', 'font_color': 'white', 'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 11})
    style_sub_header  = workbook_obj.add_format({'bg_color': '#DCE6F1', 'font_color': '#1F497D', 'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 10})
    style_sub_total   = workbook_obj.add_format({'bg_color': '#F2F2F2', 'bold': True, 'border': 1, 'font_size': 10, 'align': 'right'})
    style_grand_total = workbook_obj.add_format({'bg_color': '#D9D9D9', 'bold': True, 'border': 1, 'font_size': 10, 'align': 'right', 'top': 1, 'bottom': 6})
    
    format_int   = workbook_obj.add_format({'num_format': '#,##0', 'align': 'right', 'valign': 'vcenter', 'border': 1})
    format_pct   = workbook_obj.add_format({'num_format': '0.00%', 'align': 'right', 'valign': 'vcenter', 'border': 1})
    format_txt   = workbook_obj.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
    format_ctr   = workbook_obj.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})

    header_mapping_a = [
        ('Sr. No.', 'Sr. No.', format_ctr),
        ('Hierarchy Identifiers', 'Sub Division Name', format_txt),
        ('Hierarchy Identifiers', 'Sub Office Name', format_txt),
        ('Hierarchy Identifiers', 'Office Name', format_txt),
        ('Hierarchy Identifiers', 'Office ID', format_ctr),
        ('Hierarchy Identifiers', 'Office Type', format_ctr),
        (f'Parcel ({p1_range})', 'Received', format_int),
        (f'Parcel ({p1_range})', 'D+0 Delivery %', format_pct),
        (f'Parcel ({p1_range})', 'D+1 Delivery %', format_pct),
        (f'Documents ({p1_range})', 'Received', format_int),
        (f'Documents ({p1_range})', 'D+0 Delivery %', format_pct),
        (f'Documents ({p1_range})', 'D+1 Delivery %', format_pct),
        (f'All Products ({p1_range})', 'Received', format_int),
        (f'All Products ({p1_range})', 'D+0 Delivery %', format_pct),
        (f'All Products ({p1_range})', 'D+1 Delivery %', format_pct),
        ('Custom Performance Tracking', 'Same Day RTS %', format_pct),
        ('Custom Performance Tracking', 'Articles Not Invoiced Count', format_int),
        (f'Productivity ({p2_date})', 'Delivery Productivity %', format_pct),
        (f'DSS Daily ({p3_date})', 'DSS Usage %', format_pct),
        (f'DSS Cumulative ({p4_range})', 'DSS Usage %', format_pct),
        (f'COD Daily ({p5_date})', 'COD Digital Transaction %', format_pct),
        (f'COD Cumulative ({p6_range})', 'COD Digital Transaction %', format_pct)
    ]

    header_mapping_b = [
        ('Sr. No.', 'Sr. No.', format_ctr),
        ('Hierarchy Identifiers', 'Sub Division Name', format_txt),
        ('Hierarchy Identifiers', 'Sub Office Name', format_txt),
        (f'Parcel ({p1_range})', 'Received', format_int),
        (f'Parcel ({p1_range})', 'D+0 Delivery %', format_pct),
        (f'Parcel ({p1_range})', 'D+1 Delivery %', format_pct),
        (f'Documents ({p1_range})', 'Received', format_int),
        (f'Documents ({p1_range})', 'D+0 Delivery %', format_pct),
        (f'Documents ({p1_range})', 'D+1 Delivery %', format_pct),
        (f'All Products ({p1_range})', 'Received', format_int),
        (f'All Products ({p1_range})', 'D+0 Delivery %', format_pct),
        (f'All Products ({p1_range})', 'D+1 Delivery %', format_pct),
        (f'Productivity ({p2_date})', 'Delivery Productivity %', format_pct),
        (f'DSS Daily ({p3_date})', 'DSS Usage %', format_pct),
        (f'DSS Cumulative ({p4_range})', 'DSS Usage %', format_pct),
        (f'COD Daily ({p5_date})', 'COD Digital Transaction %', format_pct),
        (f'COD Cumulative ({p6_range})', 'COD Digital Transaction %', format_pct)
    ]

    def build_merged_headers(ws, specification):
        ws.set_row(0, 26)
        ws.set_row(1, 24)
        group_anchor = None
        start_col = 0
        for i, (group_name, sub_name, _) in enumerate(specification):
            if group_anchor is None:
                group_anchor = group_name
                start_col = i
            elif group_name != group_anchor:
                if start_col == i - 1:
                    ws.write(0, start_col, group_anchor, style_main_header)
                else:
                    ws.merge_range(0, start_col, 0, i - 1, group_anchor, style_main_header)
                group_anchor = group_name
                start_col = i
            ws.write(1, i, sub_name, style_sub_header)
        if start_col == len(specification) - 1:
            ws.write(0, start_col, group_anchor, style_main_header)
        else:
            ws.merge_range(0, start_col, 0, len(specification) - 1, group_anchor, style_main_header)

    # Sheet 1 Build
    ws1 = workbook_obj.add_worksheet('Raw Data Layout')
    build_merged_headers(ws1, header_mapping_a)
    r_idx = 2
    for item_idx, r in core_ledger.iterrows():
        ws1.set_row(r_idx, 19)
        ws1.write(r_idx, 0, item_idx + 1, format_ctr)
        ws1.write(r_idx, 1, r['Sub Division'], format_txt)
        ws1.write(r_idx, 2, r['Sub Office'], format_txt)
        ws1.write(r_idx, 3, r['Canonical_Office_Name'], format_txt)
        ws1.write(r_idx, 4, r['office_id'], format_ctr)
        ws1.write(r_idx, 5, r['office-type-code'], format_ctr)
        ws1.write(r_idx, 6, r['par_Received'], format_int)
        ws1.write(r_idx, 7, r['val_par_d0'], format_pct)
        ws1.write(r_idx, 8, r['val_par_d1'], format_pct)
        ws1.write(r_idx, 9, r['doc_Received'], format_int)
        ws1.write(r_idx, 10, r['val_doc_d0'], format_pct)
        ws1.write(r_idx, 11, r['val_doc_d1'], format_pct)
        ws1.write(r_idx, 12, r['all_Received'], format_int)
        ws1.write(r_idx, 13, r['val_all_d0'], format_pct)
        ws1.write(r_idx, 14, r['val_all_d1'], format_pct)
        ws1.write(r_idx, 15, r['val_all_rts'], format_pct)
        ws1.write(r_idx, 16, r['val_all_not_invoiced'], format_int)
        ws1.write(r_idx, 17, r['val_prod'], format_pct)
        ws1.write(r_idx, 18, r['val_dss_d'], format_pct)
        ws1.write(r_idx, 19, r['val_dss_c'], format_pct)
        ws1.write(r_idx, 20, r['val_cod_d'], format_pct)
        ws1.write(r_idx, 21, r['val_cod_c'], format_pct)
        r_idx += 1

    ws1.set_row(r_idx, 22)
    for col_c in range(22):
        ws1.write_blank(r_idx, col_c, "", style_grand_total)
    ws1.write(r_idx, 1, "Karad Division Grand Total", style_grand_total)
    ws1.write_formula(r_idx, 6, f"=SUM(G3:G{r_idx})", format_int)
    ws1.write_formula(r_idx, 9, f"=SUM(J3:J{r_idx})", format_int)
    ws1.write_formula(r_idx, 12, f"=SUM(M3:M{r_idx})", format_int)
    ws1.write_formula(r_idx, 16, f"=SUM(Q3:Q{r_idx})", format_int)

    subdivs = ['ASP Karad West', 'SDIP Karad East', 'SDIP Vaduj']

    # Sheet 2 Build
    ws2 = workbook_obj.add_worksheet('SDn wise Only SO')
    build_merged_headers(ws2, header_mapping_b)
    df_so_filtered = core_ledger[core_ledger['office-type-code'].isin(['SPO', 'HPO'])].copy()
    so_idx = 2
    for sdn in subdivs:
        df_sub = df_so_filtered[df_so_filtered['Sub Division'] == sdn]
        if df_sub.empty: continue
        b_start = so_idx + 1
        for item_idx, r in df_sub.reset_index(drop=True).iterrows():
            ws2.set_row(so_idx, 19)
            ws2.write(so_idx, 0, item_idx + 1, format_ctr)
            ws2.write(so_idx, 1, r['Sub Division'], format_txt)
            ws2.write(so_idx, 2, r['Sub Office'], format_txt)
            ws2.write(so_idx, 3, r['par_Received'], format_int)
            ws2.write(so_idx, 4, r['val_par_d0'], format_pct)
            ws2.write(so_idx, 5, r['val_par_d1'], format_pct)
            ws2.write(so_idx, 6, r['doc_Received'], format_int)
            ws2.write(so_idx, 7, r['val_doc_d0'], format_pct)
            ws2.write(so_idx, 8, r['val_doc_d1'], format_pct)
            ws2.write(so_idx, 9, r['all_Received'], format_int)
            ws2.write(so_idx, 10, r['val_all_d0'], format_pct)
            ws2.write(so_idx, 11, r['val_all_d1'], format_pct)
            ws2.write(so_idx, 12, r['val_prod'], format_pct)
            ws2.write(so_idx, 13, r['val_dss_d'], format_pct)
            ws2.write(so_idx, 14, r['val_dss_c'], format_pct)
            ws2.write(so_idx, 15, r['val_cod_d'], format_pct)
            ws2.write(so_idx, 16, r['val_cod_c'], format_pct)
            so_idx += 1
        ws2.set_row(so_idx, 22)
        for c in range(17): ws2.write_blank(so_idx, c, "", style_sub_total)
        ws2.write(so_idx, 1, f"{sdn} Sub Total", style_sub_total)
        ws2.write_formula(so_idx, 3, f"=SUM(D{b_start}:D{so_idx})", format_int)
        ws2.write_formula(so_idx, 6, f"=SUM(G{b_start}:G{so_idx})", format_int)
        ws2.write_formula(so_idx, 9, f"=SUM(J{b_start}:J{so_idx})", format_int)
        so_idx += 2

    # Sheet 3 Build
    ws3 = workbook_obj.add_worksheet('SDn wise Only BO')
    build_merged_headers(ws3, header_mapping_b)
    df_bpo_raw = core_ledger[core_ledger['office-type-code'] == 'BPO'].copy()
    parent_skeleton = master_df[master_df['office-type-code'].isin(['SPO', 'HPO'])][['Sub Division', 'Sub Office']].drop_duplicates().copy()
    agg_cols = ['par_Received', 'doc_Received', 'all_Received']
    df_bpo_consolidated = df_bpo_raw.groupby(['Sub Division', 'Sub Office'])[agg_cols].sum().reset_index()
    df_sheet3_data = parent_skeleton.merge(df_bpo_consolidated, on=['Sub Division', 'Sub Office'], how='left').fillna(0.0)
    df_sheet3_data = df_sheet3_data.sort_values(by=['Sub Division', 'Sub Office']).reset_index(drop=True)
    
    bo_idx = 2
    for sdn in subdivs:
        df_sub = df_sheet3_data[df_sheet3_data['Sub Division'] == sdn]
        if df_sub.empty: continue
        b_start = bo_idx + 1
        for item_idx, r in df_sub.reset_index(drop=True).iterrows():
            ws3.set_row(bo_idx, 19)
            ws3.write(bo_idx, 0, item_idx + 1, format_ctr)
            ws3.write(bo_idx, 1, r['Sub Division'], format_txt)
            ws3.write(bo_idx, 2, r['Sub Office'], format_txt)
            ws3.write(bo_idx, 3, r['par_Received'], format_int)
            for c_pos in [4,5,7,8,10,11,12,13,14,15,16]: ws3.write(bo_idx, c_pos, "", format_pct)
            ws3.write(bo_idx, 6, r['doc_Received'], format_int)
            ws3.write(bo_idx, 9, r['all_Received'], format_int)
            bo_idx += 1
        ws3.set_row(bo_idx, 22)
        for c in range(17): ws3.write_blank(bo_idx, c, "", style_sub_total)
        ws3.write(bo_idx, 1, f"{sdn} Sub Total", style_sub_total)
        ws3.write_formula(bo_idx, 3, f"=SUM(D{b_start}:D{bo_idx})", format_int)
        ws3.write_formula(bo_idx, 6, f"=SUM(G{b_start}:G{bo_idx})", format_int)
        ws3.write_formula(bo_idx, 9, f"=SUM(J{b_start}:J{bo_idx})", format_int)
        bo_idx += 2

    # Sheet 4 Build
    ws4 = workbook_obj.add_worksheet('All Only S.O')
    build_merged_headers(ws4, header_mapping_b)
    df_sheet4_data = df_so_filtered.sort_values(by=['Sub Office']).reset_index(drop=True)
    flat_idx = 2
    for item_idx, r in df_sheet4_data.iterrows():
        ws4.set_row(flat_idx, 19)
        ws4.write(flat_idx, 0, item_idx + 1, format_ctr)
        ws4.write(flat_idx, 1, r['Sub Division'], format_txt)
        ws4.write(flat_idx, 2, r['Sub Office'], format_txt)
        ws4.write(flat_idx, 3, r['par_Received'], format_int)
        ws4.write(flat_idx, 4, r['val_par_d0'], format_pct)
        ws4.write(flat_idx, 5, r['val_par_d1'], format_pct)
        ws4.write(flat_idx, 6, r['doc_Received'], format_int)
        ws4.write(flat_idx, 7, r['val_doc_d0'], format_pct)
        ws4.write(flat_idx, 8, r['val_doc_d1'], format_pct)
        ws4.write(flat_idx, 9, r['all_Received'], format_int)
        ws4.write(flat_idx, 10, r['val_all_d0'], format_pct)
        ws4.write(flat_idx, 11, r['val_all_d1'], format_pct)
        ws4.write(flat_idx, 12, r['val_prod'], format_pct)
        ws4.write(flat_idx, 13, r['val_dss_d'], format_pct)
        ws4.write(flat_idx, 14, r['val_dss_c'], format_pct)
        ws4.write(flat_idx, 15, r['val_cod_d'], format_pct)
        ws4.write(flat_idx, 16, r['val_cod_c'], format_pct)
        flat_idx += 1
    ws4.set_row(flat_idx, 22)
    for c in range(17): ws4.write_blank(flat_idx, c, "", style_grand_total)
    ws4.write(flat_idx, 1, "Total Row", style_grand_total)
    ws4.write_formula(flat_idx, 3, f"=SUM(D3:D{flat_idx})", format_int)
    ws4.write_formula(flat_idx, 6, f"=SUM(G3:G{flat_idx})", format_int)
    ws4.write_formula(flat_idx, 9, f"=SUM(J3:J{flat_idx})", format_int)

    # Sheet 5 Build
    def evaluate_defaulter(row):
        scores = []
        if row['all_Received'] > 0: scores.append(row['val_all_d0'] < 0.90)
        if row['prod_denom'] > 0: scores.append(row['val_prod'] < 0.90)
        if row['dss_c_denom'] > 0: scores.append(row['val_dss_c'] < 0.90)
        if row['cod_d_denom'] > 0: scores.append(row['val_cod_d'] < 0.90)
        if row['cod_c_denom'] > 0: scores.append(row['val_cod_c'] < 0.90)
        return all(scores) if scores else False

    df_defaulters = core_ledger[core_ledger.apply(evaluate_defaulter, axis=1)].sort_values(by=['Canonical_Office_Name']).reset_index(drop=True)
    ws5 = workbook_obj.add_worksheet('Defaulter Offices')
    header_mapping_def = [c for c in header_mapping_a if c[1] != 'D+1 Delivery %']
    build_merged_headers(ws5, header_mapping_def)
    
    def_idx = 2
    for item_idx, r in df_defaulters.iterrows():
        ws5.set_row(def_idx, 19)
        ws5.write(def_idx, 0, item_idx + 1, format_ctr)
        ws5.write(def_idx, 1, r['Sub Division'], format_txt)
        ws5.write(def_idx, 2, r['Sub Office'], format_txt)
        ws5.write(def_idx, 3, r['Canonical_Office_Name'], format_txt)
        ws5.write(def_idx, 4, r['office_id'], format_ctr)
        ws5.write(def_idx, 5, r['office-type-code'], format_ctr)
        ws5.write(def_idx, 6, r['par_Received'], format_int)
        ws5.write(def_idx, 7, r['val_par_d0'], format_pct)
        ws5.write(def_idx, 8, r['doc_Received'], format_int)
        ws5.write(def_idx, 9, r['val_doc_d0'], format_pct)
        ws5.write(def_idx, 10, r['all_Received'], format_int)
        ws5.write(def_idx, 11, r['val_all_d0'], format_pct)
        ws5.write(def_idx, 12, r['val_all_rts'], format_pct)
        ws5.write(def_idx, 13, r['val_all_not_invoiced'], format_int)
        ws5.write(def_idx, 14, r['val_prod'], format_pct)
        ws5.write(def_idx, 15, r['val_dss_d'], format_pct)
        ws5.write(def_idx, 16, r['val_dss_c'], format_pct)
        ws5.write(def_idx, 17, r['val_cod_d'], format_pct)
        ws5.write(def_idx, 18, r['val_cod_c'], format_pct)
        def_idx += 1

    for ws_sheet in [ws1, ws2, ws3, ws4, ws5]:
        ws_sheet.set_column(0, 0, 7)
        ws_sheet.set_column(1, 3, 26)
        ws_sheet.set_column(4, 5, 13)
        ws_sheet.set_column(6, 23, 15)

    excel_engine.close()
    output_stream.seek(0)
    st.success("🏁 Verification complete. Unified ledger compiled successfully.")
    st.download_button(label="📥 Download Production Workbook", data=output_stream, file_name="MMU_Report_Karad.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)