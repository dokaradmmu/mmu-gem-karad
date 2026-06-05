"""
Application: MMU Gem Karad Division Data Processing Utility
Engineers: Senior Full-Stack Engineer / Data Architect / Python Developer
Dependencies: streamlit, pandas, openpyxl, xlsxwriter
Validation Framework: Fully compliant with Bug Fixes B-01 through B-15
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import datetime

# ==========================================
# 1. TIMEZONE OVERRIDE ANCHOR (NO PYTZ DEPENDENCY)
# ==========================================
# UTC+5:30 for India Standard Time using Python standard library
IST = datetime.timezone(datetime.timedelta(hours=5, minutes=30))
now_ist = datetime.datetime.now(IST)

st.set_page_config(
    page_title="MMU Gem Karad Division", 
    page_icon="🇮🇳", 
    layout="wide"
)

# Initialize System Warnings and Diagnostics Ledger
if "system_diagnostics" not in st.session_state:
    st.session_state.system_diagnostics = []

st.title("🇮🇳 MMU Gem Karad Division")
st.subheader("Enterprise Data Processing Utility — India Post Operations")
st.markdown(
    f"**Application Environment Temporal Anchor (IST):** `{now_ist.strftime('%Y-%m-%d %H:%M:%S GMT+5:30')}` "
    f"| Authentication: `Public Wide-Open`"
)

# ==========================================
# 2. USER INTERFACE GENERATION MATRIX
# ==========================================
st.sidebar.header("📋 Workflow: 0. Daily Report")
st.sidebar.markdown("---")

st.sidebar.subheader("📅 Dynamic Header Labels Configuration")
p1_range = st.sidebar.text_input("Prompt 1: Shared Transit Range (From/To)", "01.05.2026 to 23.05.2026")
p2_date  = st.sidebar.text_input("Prompt 2: Single Performance Date", "30.05.2026")
p3_date  = st.sidebar.text_input("Prompt 3: Single App Usage Date", "30.05.2026")
p4_range = st.sidebar.text_input("Prompt 4: Independent DSS Range (From/To)", "01.05.2026 to 30.05.2026")
p5_date  = st.sidebar.text_input("Prompt 5: Single Financial Date", "30.05.2026")
p6_range = st.sidebar.text_input("Prompt 6: Independent COD Range (From/To)", "01.05.2026 to 30.05.2026")

st.sidebar.markdown("---")
st.sidebar.subheader("📂 Ingestion Matrix Slots")

master_file = st.sidebar.file_uploader("⚓ Master Skeleton: Updated Office Names 29.05.2026 (CSV)", type=["csv"])
slot_01 = st.sidebar.file_uploader("Slot 0.1: Speed Parcel (CSV)", type=["csv"])
slot_02 = st.sidebar.file_uploader("Slot 0.2: Registered Parcel (CSV)", type=["csv"])
slot_03 = st.sidebar.file_uploader("Slot 0.3: Speed Letter (CSV)", type=["csv"])
slot_04 = st.sidebar.file_uploader("Slot 0.4: Registered Letter (CSV)", type=["csv"])
slot_05 = st.sidebar.file_uploader("Slot 0.5: All Category (CSV)", type=["csv"])
slot_06 = st.sidebar.file_uploader("Slot 0.6: Delivery Productivity (CSV)", type=["csv"])
slot_07 = st.sidebar.file_uploader("Slot 0.7: DSS Usage Daily (CSV)", type=["csv"])
slot_08 = st.sidebar.file_uploader("Slot 0.8: DSS Usage Consolidated (CSV)", type=["csv"])
slot_09 = st.sidebar.file_uploader("Slot 0.9: COD Collection Daily (CSV / Optional)", type=["csv"])
slot_10 = st.sidebar.file_uploader("Slot 0.10: COD Collection Consolidated (CSV)", type=["csv"])

# ==========================================
# 3. DATA SANITIZATION CONTEXT PIPELINE (BUG B-03, B-05)
# ==========================================
def ingest_and_sanitize(file_io, name_tag=""):
    if file_io is None:
        return None
    df = pd.read_csv(file_io, skip_blank_lines=True)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Key Normalization Rule (First Phase - Bug B-03)
    rename_target = {}
    for c in df.columns:
        if str(c).lower().replace(" ", "").replace("-", "_") in ['officeid', 'office_id']:
            rename_target[c] = 'office_id'
    if rename_target:
        df = df.rename(columns=rename_target)
        
    if 'office_id' not in df.columns:
        st.error(f"Critical Ingestion Error: Key field 'office_id' could not be resolved in slot {name_tag}.")
        st.stop()
        
    # Strict Key Normalization Order (Second Phase - Bug B-05 Fix)
    df['office_id'] = pd.to_numeric(df['office_id'], errors='coerce')
    df = df.dropna(subset=['office_id'])
    df = df[df['office_id'] > 0]
    df['office_id'] = df['office_id'].astype(int).astype(str).str.strip().str.lstrip('0')
    df = df[df['office_id'] != '']
    df = df[df['office_id'] != '0']
    
    # Evict System Pre-Aggregated Summary Sentinels
    for c in ['office_name', 'Office Name', 'customer-name', 'product-name', 'office-name']:
        if c in df.columns:
            df = df[~df[c].astype(str).str.contains('Summary|Total', case=False, na=False)]
            
    # Universal metadata whitespace and quote stripping
    for col in df.select_dtypes(include=[object]).columns:
        df[col] = df[col].astype(str).str.strip().str.replace(r'^"|"$', '', regex=True)
    return df

# ==========================================
# 4. RUNTIME DATA PROCESSOR ENGINE
# ==========================================
if st.sidebar.button("🚀 Process Operational Records", use_container_width=True):
    st.session_state.system_diagnostics = []
    
    if master_file is None:
        st.error("Missing Structural Root Matrix: Please upload the Master Skeleton file to authorize joins.")
        st.stop()
        
    # Ingest and verify structural backbone (Bug B-15 verification)
    master_df = pd.read_csv(master_file)
    master_df.columns = [str(c).strip() for c in master_df.columns]
    for c in master_df.columns:
        if str(c).lower().replace(" ", "").replace("-", "_") in ['officeid', 'office_id']:
            master_df.rename(columns={c: 'office_id'}, inplace=True)
            
    master_df['office_id'] = pd.to_numeric(master_df['office_id'], errors='coerce')
    master_df = master_df.dropna(subset=['office_id'])
    master_df['office_id'] = master_df['office_id'].astype(int).astype(str).str.strip().str.lstrip('0')
    
    # Enforce strict display data sanitization (Bug B-12 canonical rule)
    for c in ['Sub Division', 'Sub Office', 'Branch Office', 'office-type-code']:
        if c in master_df.columns:
            master_df[c] = master_df[c].astype(str).str.strip()
            
    # Map cleaned canonical display name field
    master_df['Canonical_Office_Name'] = np.where(
        master_df['office-type-code'].isin(['SPO', 'HPO']),
        master_df['Sub Office'],
        master_df['Branch Office']
    )
    master_keys = set(master_df['office_id'].unique())
    
    # Ingest Upload Slots
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
    
    # Bug Fix B-02: Decommission slot 0.9 errors via Consolidated fallback matching
    if raw_09 is None and raw_10 is not None:
        raw_09 = raw_10.copy()
        st.session_state.system_diagnostics.append("ℹ️ **[BUG B-02 HANDLED]** Slot 0.9 (Daily COD) empty. Deployed matching Consolidated COD structure as real-time fallback.")

    # Diagnostic Orphan Verification Auditing (Bug B-08)
    files_matrix = [
        (raw_01, "Speed Parcel"), (raw_02, "Registered Parcel"), (raw_03, "Speed Letter"), 
        (raw_04, "Registered Letter"), (raw_05, "All Category"), (raw_06, "Delivery Productivity"), 
        (raw_07, "DSS Daily"), (raw_08, "DSS Consolidated"), (raw_10, "COD Consolidated")
    ]
    for target_df, label in files_matrix:
        if target_df is not None:
            orphans = set(target_df['office_id'].unique()) - master_keys
            if orphans:
                st.session_state.system_diagnostics.append(
                    f"⚠️ **[BUG B-08 ALERT]** File `[{label}]` contains {len(orphans)} Office ID keys missing "
                    f"from structural master file (Dropped during left-join pipeline): `{list(orphans)[:2]}`"
                )

    # ----------------------------------------------------
    # VECTOR REDUCTIONS & IN-MEMORY LEDGERS (BUG B-14)
    # ----------------------------------------------------
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

    # Delivery Productivity Calculations (Bug B-10 compliance)
    if raw_06 is not None:
        for c in ['invoice-count', 'delivery-count', 'return-count', 'redirection-count', 'deposit-count']:
            raw_06[c] = pd.to_numeric(raw_06[c], errors='coerce').fillna(0.0) if c in raw_06.columns else 0.0
        raw_06['prod_num'] = raw_06['delivery-count'] + raw_06['return-count'] + raw_06['redirection-count']
        df_prod_ledger = raw_06.groupby('office_id', as_index=False)[['invoice-count', 'prod_num']].sum()
    else:
        df_prod_ledger = pd.DataFrame(columns=['office_id', 'invoice-count', 'prod_num'])

    # DSS System Aggregation Ledger Maps
    def aggregate_dss(df):
        if df is not None:
            for c in ['total_pdm_art_count', 'total_dss_art_count']:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
            return df.groupby('office_id', as_index=False)[['total_pdm_art_count', 'total_dss_art_count']].sum()
        return pd.DataFrame(columns=['office_id', 'total_pdm_art_count', 'total_dss_art_count'])

    df_dss_d_ledger = aggregate_dss(raw_07)
    df_dss_c_ledger = aggregate_dss(raw_08)

    # COD Transaction level reduction maps (Bug B-03 compliance)
    def aggregate_cod(df):
        if df is not None:
            for c in ['no_digital_count', 'no-cod-articles']:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
            return df.groupby('office_id', as_index=False)[['no_digital_count', 'no-cod-articles']].sum()
        return pd.DataFrame(columns=['office_id', 'no_digital_count', 'no-cod-articles'])

    df_cod_d_ledger = aggregate_cod(raw_09)
    df_cod_c_ledger = aggregate_cod(raw_10)

    # Compile Unified Structural Skeleton
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

    # Flexible Input Verification (KeyError Resolution Matrix)
    def generate_ratio_vector(df, numer, denom):
        n_val = df[numer] if isinstance(numer, str) else numer
        return np.where(df[denom] > 0, n_val / df[denom], np.nan)

    core_ledger['val_par_d0'] = generate_ratio_vector(core_ledger, core_ledger['par_D0 Delivered'] + core_ledger['par_D0 Redirected'] + core_ledger['par_D0 Returned'], 'par_Received')
    core_ledger['val_par_d1'] = generate_ratio_vector(core_ledger, core_ledger['par_D1 Delivered'] + core_ledger['par_D1 Redirected'] + core_ledger['par_D1 Returned'], 'par_Received')
    core_ledger['val_doc_d0'] = generate_ratio_vector(core_ledger, core_ledger['doc_D0 Delivered'] + core_ledger['doc_D0 Redirected'] + core_ledger['doc_D0 Returned'], 'doc_Received')
    core_ledger['val_doc_d1'] = generate_ratio_vector(core_ledger, core_ledger['doc_D1 Delivered'] + core_ledger['doc_D1 Redirected'] + core_ledger['doc_D1 Returned'], 'doc_Received')
    core_ledger['val_all_d0'] = generate_ratio_vector(core_ledger, core_ledger['all_D0 Delivered'] + core_ledger['all_D0 Redirected'] + core_ledger['all_D0 Returned'], 'all_Received')
    core_ledger['val_all_d1'] = generate_ratio_vector(core_ledger, core_ledger['all_D1 Delivered'] + core_ledger['all_D1 Redirected'] + core_ledger['all_D1 Returned'], 'all_Received')
    core_ledger['val_all_rts'] = generate_ratio_vector(core_ledger, 'all_D0 Returned', 'all_Received')
    core_ledger['val_all_not_invoiced'] = core_ledger['all_Received'] - core_ledger['all_Same Day Invoiced']
    core_ledger['val_prod']  = generate_ratio_vector(core_ledger, 'prod_numer', 'prod_denom')
    core_ledger['val_dss_d'] = generate_ratio_vector(core_ledger, 'dss_d_numer', 'dss_d_denom')
    core_ledger['val_dss_c'] = generate_ratio_vector(core_ledger, 'dss_c_numer', 'dss_c_denom')
    core_ledger['val_cod_d'] = generate_ratio_vector(core_ledger, 'cod_d_numer', 'cod_d_denom')
    core_ledger['val_cod_c'] = generate_ratio_vector(core_ledger, 'cod_c_numer', 'cod_c_denom')

    core_ledger = core_ledger.sort_values(by=['Sub Division', 'Sub Office', 'Branch Office']).reset_index(drop=True)

    # ==========================================
    # 5. EXCEL FORMAT STRUCTURE GENERATION DEFS
    # ==========================================
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

    # ----------------------------------------------------
    # SHEET 1: RAW DATA LAYOUT (TYPE A ARCHETYPE)
    # ----------------------------------------------------
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
        
        # Injected Data Metrics
        ws1.write(r_idx, 6, r['par_Received'], format_int)
        
        # Live interactive row formulas pointed straight to hidden matrix elements (Bug B-01)
        ws1.write_formula(r_idx, 7, f"=IFERROR((W{r_idx+1})/G{r_idx+1}, \"\")", format_pct)
        ws1.write_formula(r_idx, 8, f"=IFERROR((X{r_idx+1})/G{r_idx+1}, \"\")", format_pct)
        
        ws1.write(r_idx, 9, r['doc_Received'], format_int)
        ws1.write_formula(r_idx, 10, f"=IFERROR((Y{r_idx+1})/J{r_idx+1}, \"\")", format_pct)
        ws1.write_formula(r_idx, 11, f"=IFERROR((Z{r_idx+1})/J{r_idx+1}, \"\")", format_pct)
        
        ws1.write(r_idx, 12, r['all_Received'], format_int)
        ws1.write_formula(r_idx, 13, f"=IFERROR((AA{r_idx+1})/M{r_idx+1}, \"\")", format_pct)
        ws1.write_formula(r_idx, 14, f"=IFERROR((AB{r_idx+1})/M{r_idx+1}, \"\")", format_pct)
        ws1.write_formula(r_idx, 15, f"=IFERROR((AC{r_idx+1})/M{r_idx+1}, \"\")", format_pct)
        ws1.write(r_idx, 16, r['val_all_not_invoiced'], format_int)
        
        ws1.write_formula(r_idx, 17, f"=IFERROR(AD{r_idx+1}/AE{r_idx+1}, \"\")", format_pct)
        ws1.write_formula(r_idx, 18, f"=IFERROR(AF{r_idx+1}/AG{r_idx+1}, \"\")", format_pct)
        ws1.write_formula(r_idx, 19, f"=IFERROR(AH{r_idx+1}/AI{r_idx+1}, \"\")", format_pct)
        ws1.write_formula(r_idx, 20, f"=IFERROR(AJ{r_idx+1}/AK{r_idx+1}, \"\")", format_pct)
        ws1.write_formula(r_idx, 21, f"=IFERROR(AL{r_idx+1}/AM{r_idx+1}, \"\")", format_pct)
        
        # Populate hidden row-level processing cells on the same sheet (Bug B-13)
        ws1.write(r_idx, 22, (r['par_D0 Delivered'] + r['par_D0 Redirected'] + r['par_D0 Returned']), format_int)
        ws1.write(r_idx, 23, (r['par_D1 Delivered'] + r['par_D1 Redirected'] + r['par_D1 Returned']), format_int)
        ws1.write(r_idx, 24, (r['doc_D0 Delivered'] + r['doc_D0 Redirected'] + r['doc_D0 Returned']), format_int)
        ws1.write(r_idx, 25, (r['doc_D1 Delivered'] + r['doc_D1 Redirected'] + r['doc_D1 Returned']), format_int)
        ws1.write(r_idx, 26, (r['all_D0 Delivered'] + r['all_D0 Redirected'] + r['all_D0 Returned']), format_int)
        ws1.write(r_idx, 27, (r['all_D1 Delivered'] + r['all_D1 Redirected'] + r['all_D1 Returned']), format_int)
        ws1.write(r_idx, 28, r['all_D0 Returned'], format_int)
        ws1.write(r_idx, 29, r['prod_numer'], format_int)
        ws1.write(r_idx, 30, r['prod_denom'], format_int)
        ws1.write(r_idx, 31, r['dss_d_numer'], format_int)
        ws1.write(r_idx, 32, r['dss_d_denom'], format_int)
        ws1.write(r_idx, 33, r['dss_c_numer'], format_int)
        ws1.write(r_idx, 34, r['dss_c_denom'], format_int)
        ws1.write(r_idx, 35, r['cod_d_numer'], format_int)
        ws1.write(r_idx, 36, r['cod_d_denom'], format_int)
        ws1.write(r_idx, 37, r['cod_c_numer'], format_int)
        ws1.write(r_idx, 38, r['cod_c_denom'], format_int)
        r_idx += 1

    # Grand Total Injection (Spec §6 compliance)
    ws1.set_row(r_idx, 22)
    for col_c in range(39):
        ws1.write_blank(r_idx, col_c, "", style_grand_total)
    ws1.write(r_idx, 1, "Karad Division Grand Total", style_grand_total)
    
    ws1.write_formula(r_idx, 6, f"=SUM(G3:G{r_idx})", format_int)
    ws1.write_formula(r_idx, 7, f"=IFERROR(SUM(W3:W{r_idx})/G{r_idx+1}, \"\")", format_pct)
    ws1.write_formula(r_idx, 8, f"=IFERROR(SUM(X3:X{r_idx})/G{r_idx+1}, \"\")", format_pct)
    ws1.write_formula(r_idx, 9, f"=SUM(J3:J{r_idx})", format_int)
    ws1.write_formula(r_idx, 10, f"=IFERROR(SUM(Y3:Y{r_idx})/J{r_idx+1}, \"\")", format_pct)
    ws1.write_formula(r_idx, 11, f"=IFERROR(SUM(Z3:Z{r_idx})/J{r_idx+1}, \"\")", format_pct)
    ws1.write_formula(r_idx, 12, f"=SUM(M3:M{r_idx})", format_int)
    ws1.write_formula(r_idx, 13, f"=IFERROR(SUM(AA3:AA{r_idx})/M{r_idx+1}, \"\")", format_pct)
    ws1.write_formula(r_idx, 14, f"=IFERROR(SUM(AB3:AB{r_idx})/M{r_idx+1}, \"\")", format_pct)
    ws1.write_formula(r_idx, 15, f"=IFERROR(SUM(AC3:AC{r_idx})/M{r_idx+1}, \"\")", format_pct)
    ws1.write_formula(r_idx, 16, f"=SUM(Q3:Q{r_idx})", format_int)
    ws1.write_formula(r_idx, 17, f"=IFERROR(SUM(AD3:AD{r_idx})/SUM(AE3:AE{r_idx}), \"\")", format_pct)
    ws1.write_formula(r_idx, 18, f"=IFERROR(SUM(AF3:AF{r_idx})/SUM(AG3:AG{r_idx}), \"\")", format_pct)
    ws1.write_formula(r_idx, 19, f"=IFERROR(SUM(AH3:AH{r_idx})/SUM(AI3:AI{r_idx}), \"\")", format_pct)
    ws1.write_formula(r_idx, 20, f"=IFERROR(SUM(AJ3:AJ{r_idx})/SUM(AK3:AK{r_idx}), \"\")", format_pct)
    ws1.write_formula(r_idx, 21, f"=IFERROR(SUM(AL3:AL{r_idx})/SUM(AM3:AM{r_idx}), \"\")", format_pct)
    ws1.set_column(22, 38, None, None, {'hidden': True})

    # Sub-Divisions Constant Definition Loop Matrix
    subdivs = ['ASP Karad West', 'SDIP Karad East', 'SDIP Vaduj']

    # ----------------------------------------------------
    # AUXILIARY BLOCK COMPILER (FOR TYPE B CONFIGURATIONS)
    # ----------------------------------------------------
    def populate_type_b_sheet(ws, dataset):
        build_merged_headers(ws, header_mapping_b)
        cursor = 2
        block_references = {}
        
        for sdn in subdivs:
            df_sub = dataset[dataset['Sub Division'] == sdn]
            if df_sub.empty:
                continue
            b_start = cursor + 1
            item_seq = 1
            
            for _, r in df_sub.reset_index(drop=True).iterrows():
                ws.set_row(cursor, 19)
                ws.write(cursor, 0, item_seq, format_ctr)
                ws.write(cursor, 1, r['Sub Division'], format_txt)
                ws.write(cursor, 2, r['Sub Office'], format_txt)
                
                ws.write(cursor, 3, r['par_Received'], format_int)
                ws.write_formula(cursor, 4, f"=IFERROR(R{cursor+1}/D{cursor+1}, \"\")", format_pct)
                ws.write_formula(cursor, 5, f"=IFERROR(S{cursor+1}/D{cursor+1}, \"\")", format_pct)
                
                ws.write(cursor, 6, r['doc_Received'], format_int)
                ws.write_formula(cursor, 7, f"=IFERROR(T{cursor+1}/G{cursor+1}, \"\")", format_pct)
                ws.write_formula(cursor, 8, f"=IFERROR(U{cursor+1}/G{cursor+1}, \"\")", format_pct)
                
                ws.write(cursor, 9, r['all_Received'], format_int)
                ws.write_formula(cursor, 10, f"=IFERROR(V{cursor+1}/J{cursor+1}, \"\")", format_pct)
                ws.write_formula(cursor, 11, f"=IFERROR(W{cursor+1}/J{cursor+1}, \"\")", format_pct)
                
                ws.write_formula(cursor, 12, f"=IFERROR(X{cursor+1}/Y{cursor+1}, \"\")", format_pct)
                ws.write_formula(cursor, 13, f"=IFERROR(Z{cursor+1}/AA{cursor+1}, \"\")", format_pct)
                ws.write_formula(cursor, 14, f"=IFERROR(AB{cursor+1}/AC{cursor+1}, \"\")", format_pct)
                ws.write_formula(cursor, 15, f"=IFERROR(AD{cursor+1}/AE{cursor+1}, \"\")", format_pct)
                ws.write_formula(cursor, 16, f"=IFERROR(AF{cursor+1}/AG{cursor+1}, \"\")", format_pct)
                
                # Hidden vector mapping inside layout boundaries (Bug B-01/B-13 setup)
                ws.write(cursor, 17, (r['par_D0 Delivered'] + r['par_D0 Redirected'] + r['par_D0 Returned']), format_int)
                ws.write(cursor, 18, (r['par_D1 Delivered'] + r['par_D1 Redirected'] + r['par_D1 Returned']), format_int)
                ws.write(cursor, 19, (r['doc_D0 Delivered'] + r['doc_D0 Redirected'] + r['doc_D0 Returned']), format_int)
                ws.write(cursor, 20, (r['doc_D1 Delivered'] + r['doc_D1 Redirected'] + r['doc_D1 Returned']), format_int)
                ws.write(cursor, 21, (r['all_D0 Delivered'] + r['all_D0 Redirected'] + r['all_D0 Returned']), format_int)
                ws.write(cursor, 22, (r['all_D1 Delivered'] + r['all_D1 Redirected'] + r['all_D1 Returned']), format_int)
                ws.write(cursor, 23, r['prod_numer'], format_int)
                ws.write(cursor, 24, r['prod_denom'], format_int)
                ws.write(cursor, 25, r['dss_d_numer'], format_int)
                ws.write(cursor, 26, r['dss_d_denom'], format_int)
                ws.write(cursor, 27, r['dss_c_numer'], format_int)
                ws.write(cursor, 28, r['dss_c_denom'], format_int)
                ws.write(cursor, 29, r['cod_d_numer'], format_int)
                ws.write(cursor, 30, r['cod_d_denom'], format_int)
                ws.write(cursor, 31, r['cod_c_numer'], format_int)
                ws.write(cursor, 32, r['cod_c_denom'], format_int)
                
                cursor += 1
                item_seq += 1
                
            # Block Sub Total Row Formulations
            ws.set_row(cursor, 22)
            for col_c in range(33):
                ws.write_blank(cursor, col_c, "", style_sub_total)
            ws.write(cursor, 1, f"{sdn} Sub Total", style_sub_total)
            
            ws.write_formula(cursor, 3, f"=SUM(D{b_start}:D{cursor})", format_int)
            ws.write_formula(cursor, 4, f"=IFERROR(SUM(R{b_start}:R{cursor})/D{cursor+1}, \"\")", format_pct)
            ws.write_formula(cursor, 5, f"=IFERROR(SUM(S{b_start}:S{cursor})/D{cursor+1}, \"\")", format_pct)
            ws.write_formula(cursor, 6, f"=SUM(G{b_start}:G{cursor})", format_int)
            ws.write_formula(cursor, 7, f"=IFERROR(SUM(T{b_start}:T{cursor})/G{cursor+1}, \"\")", format_pct)
            ws.write_formula(cursor, 8, f"=IFERROR(SUM(U{b_start}:U{cursor})/G{cursor+1}, \"\")", format_pct)
            ws.write_formula(cursor, 9, f"=SUM(J{b_start}:J{cursor})", format_int)
            ws.write_formula(cursor, 10, f"=IFERROR(SUM(V{b_start}:V{cursor})/J{cursor+1}, \"\")", format_pct)
            ws.write_formula(cursor, 11, f"=IFERROR(SUM(W{b_start}:W{cursor})/J{cursor+1}, \"\")", format_pct)
            ws.write_formula(cursor, 12, f"=IFERROR(SUM(X{b_start}:X{cursor})/SUM(Y{b_start}:Y{cursor}), \"\")", format_pct)
            ws.write_formula(cursor, 13, f"=IFERROR(SUM(Z{b_start}:Z{cursor})/SUM(AA{b_start}:AA{cursor}), \"\")", format_pct)
            ws.write_formula(cursor, 14, f"=IFERROR(SUM(AB{b_start}:AB{cursor})/SUM(AC{b_start}:AC{cursor}), \"\")", format_pct)
            ws.write_formula(cursor, 15, f"=IFERROR(SUM(AD{b_start}:AD{cursor})/SUM(AE{b_start}:AE{cursor}), \"\")", format_pct)
            ws.write_formula(cursor, 16, f"=IFERROR(SUM(AF{b_start}:AF{cursor})/SUM(AG{b_start}:AG{cursor}), \"\")", format_pct)
            
            # Populate block subtotal aggregations inside hidden columns for outer-block summaries
            ws.write(cursor, 17, f"=SUM(R{b_start}:R{cursor})", format_int)
            ws.write(cursor, 18, f"=SUM(S{b_start}:S{cursor})", format_int)
            ws.write(cursor, 19, f"=SUM(T{b_start}:T{cursor})", format_int)
            ws.write(cursor, 20, f"=SUM(U{b_start}:U{cursor})", format_int)
            ws.write(cursor, 21, f"=SUM(V{b_start}:V{cursor})", format_int)
            ws.write(cursor, 22, f"=SUM(W{b_start}:W{cursor})", format_int)
            ws.write(cursor, 23, f"=SUM(X{b_start}:X{cursor})", format_int)
            ws.write(cursor, 24, f"=SUM(Y{b_start}:Y{cursor})", format_int)
            ws.write(cursor, 25, f"=SUM(Z{b_start}:Z{cursor})", format_int)
            ws.write(cursor, 26, f"=SUM(AA{b_start}:AA{cursor})", format_int)
            ws.write(cursor, 27, f"=SUM(AB{b_start}:AB{cursor})", format_int)
            ws.write(cursor, 28, f"=SUM(AC{b_start}:AC{cursor})", format_int)
            ws.write(cursor, 29, f"=SUM(AD{b_start}:AD{cursor})", format_int)
            ws.write(cursor, 30, f"=SUM(AE{b_start}:AE{cursor})", format_int)
            ws.write(cursor, 31, f"=SUM(AF{b_start}:AF{cursor})", format_int)
            ws.write(cursor, 32, f"=SUM(AG{b_start}:AG{cursor})", format_int)
            
            block_references[sdn] = cursor + 1 # Save 1-based index position
            cursor += 2 # Leave whitespace row gap
            
        # ----------------------------------------------------
        # TRAILING SUMMARY BLOCK INJECTION (BUG B-07 FIX)
        # ----------------------------------------------------
        ws.write(cursor, 1, "Karad Division Summary Block", style_sub_header)
        cursor += 1
        summary_start_row = cursor + 1
        
        for sdn in subdivs:
            if sdn not in block_references:
                continue
            ws.set_row(cursor, 19)
            for col_c in range(33):
                ws.write_blank(cursor, col_c, "", format_txt)
            ws.write(cursor, 1, sdn, format_txt)
            ws.write(cursor, 2, "Summary Row", format_ctr)
            
            target_line = block_references[sdn]
            # Map single cell tracking pointers straight to the computed block subtotals
            ws.write_formula(cursor, 3, f"=D{target_line}", format_int)
            ws.write_formula(cursor, 4, f"=E{target_line}", format_pct)
            ws.write_formula(cursor, 5, f"=F{target_line}", format_pct)
            ws.write_formula(cursor, 6, f"=G{target_line}", format_int)
            ws.write_formula(cursor, 7, f"=H{target_line}", format_pct)
            ws.write_formula(cursor, 8, f"=I{target_line}", format_pct)
            ws.write_formula(cursor, 9, f"=J{target_line}", format_int)
            ws.write_formula(cursor, 10, f"=K{target_line}", format_pct)
            ws.write_formula(cursor, 11, f"=L{target_line}", format_pct)
            ws.write_formula(cursor, 12, f"=M{target_line}", format_pct)
            ws.write_formula(cursor, 13, f"=N{target_line}", format_pct)
            ws.write_formula(cursor, 14, f"=O{target_line}", format_pct)
            ws.write_formula(cursor, 15, f"=P{target_line}", format_pct)
            ws.write_formula(cursor, 16, f"=Q{target_line}", format_pct)
            
            # Track values through the hidden matrix block
            for hidden_idx, col_char in enumerate(['R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG']):
                ws.write_formula(cursor, 17 + hidden_idx, f"={col_char}{target_line}", format_int)
            cursor += 1
            
        summary_end_row = cursor
        ws.set_row(cursor, 22)
        for col_c in range(33):
            ws.write_blank(cursor, col_c, "", style_grand_total)
        ws.write(cursor, 1, "Karad Division Grand Total", style_grand_total)
        
        # Grand total sums the three trailing division summary records
        ws.write_formula(cursor, 3, f"=SUM(D{summary_start_row}:D{summary_end_row})", format_int)
        ws.write_formula(cursor, 4, f"=IFERROR(SUM(R{summary_start_row}:R{summary_end_row})/D{cursor+1}, \"\")", format_pct)
        ws.write_formula(cursor, 5, f"=IFERROR(SUM(S{summary_start_row}:S{summary_end_row})/D{cursor+1}, \"\")", format_pct)
        ws.write_formula(cursor, 6, f"=SUM(G{summary_start_row}:G{summary_end_row})", format_int)
        ws.write_formula(cursor, 7, f"=IFERROR(SUM(T{summary_start_row}:T{summary_end_row})/G{cursor+1}, \"\")", format_pct)
        ws.write_formula(cursor, 8, f"=IFERROR(SUM(U{summary_start_row}:U{summary_end_row})/G{cursor+1}, \"\")", format_pct)
        ws.write_formula(cursor, 9, f"=SUM(J{summary_start_row}:J{summary_end_row})", format_int)
        ws.write_formula(cursor, 10, f"=IFERROR(SUM(V{summary_start_row}:V{summary_end_row})/J{cursor+1}, \"\")", format_pct)
        ws.write_formula(cursor, 11, f"=IFERROR(SUM(W{summary_start_row}:W{summary_end_row})/J{cursor+1}, \"\")", format_pct)
        ws.write_formula(cursor, 12, f"=IFERROR(SUM(X{summary_start_row}:X{summary_end_row})/SUM(Y{summary_start_row}:Y{summary_end_row}), \"\")", format_pct)
        ws.write_formula(cursor, 13, f"=IFERROR(SUM(Z{summary_start_row}:Z{summary_end_row})/SUM(AA{summary_start_row}:AA{summary_end_row}), \"\")", format_pct)
        ws.write_formula(cursor, 14, f"=IFERROR(SUM(AB{summary_start_row}:AB{summary_end_row})/SUM(AC{summary_start_row}:AC{summary_end_row}), \"\")", format_pct)
        ws.write_formula(cursor, 15, f"=IFERROR(SUM(AD{summary_start_row}:AD{summary_end_row})/SUM(AE{summary_start_row}:AE{summary_end_row}), \"\")", format_pct)
        ws.write_formula(cursor, 16, f"=IFERROR(SUM(AF{summary_start_row}:AF{summary_end_row})/SUM(AG{summary_start_row}:AG{summary_end_row}), \"\")", format_pct)
        ws.set_column(17, 32, None, None, {'hidden': True})

    # ----------------------------------------------------
    # SHEET 2: SDn wise Only SO (SPO/HPO FILTER — BUG B-09)
    # ----------------------------------------------------
    ws2 = workbook_obj.add_worksheet('SDn wise Only SO')
    df_so_filtered = core_ledger[core_ledger['office-type-code'].isin(['SPO', 'HPO'])].copy()
    populate_type_b_sheet(ws2, df_so_filtered)

    # ----------------------------------------------------
    # SHEET 3: SDn wise Only BO (PARENT ANCHOR JOIN — BUG B-04)
    # ----------------------------------------------------
    ws3 = workbook_obj.add_worksheet('SDn wise Only BO')
    df_bpo_raw = core_ledger[core_ledger['office-type-code'] == 'BPO'].copy()
    
    # Establish parent skeleton to map metrics without dropping office records
    parent_skeleton = master_df[master_df['office-type-code'].isin(['SPO', 'HPO'])][['Sub Division', 'Sub Office']].drop_duplicates().copy()
    
    aggregation_targets = [
        'par_Received', 'par_D0 Delivered', 'par_D0 Redirected', 'par_D0 Returned', 'par_D1 Delivered', 'par_D1 Redirected', 'par_D1 Returned',
        'doc_Received', 'doc_D0 Delivered', 'doc_D0 Redirected', 'doc_D0 Returned', 'doc_D1 Delivered', 'doc_D1 Redirected', 'doc_D1 Returned',
        'all_Received', 'all_Same Day Invoiced', 'all_D0 Delivered', 'all_D0 Redirected', 'all_D0 Returned', 'all_D1 Delivered', 'all_D1 Redirected', 'all_D1 Returned',
        'prod_denom', 'prod_numer', 'dss_d_denom', 'dss_d_numer', 'dss_c_denom', 'dss_c_numer', 'cod_d_denom', 'cod_d_numer', 'cod_c_denom', 'cod_c_numer'
    ]
    df_bpo_consolidated = df_bpo_raw.groupby(['Sub Division', 'Sub Office'])[aggregation_targets].sum().reset_index()
    
    # Map back to the parent skeletal system (Bug B-04 Fix)
    df_sheet3_final = parent_skeleton.merge(df_bpo_consolidated, on=['Sub Division', 'Sub Office'], how='left').fillna(0.0)
    df_sheet3_final = df_sheet3_final.sort_values(by=['Sub Division', 'Sub Office']).reset_index(drop=True)
    populate_type_b_sheet(ws3, df_sheet3_final)

    # ----------------------------------------------------
    # SHEET 4: All Only S.O (FLAT TYPE B LISTING)
    # ----------------------------------------------------
    ws4 = workbook_obj.add_worksheet('All Only S.O')
    build_merged_headers(ws4, header_mapping_b)
    df_sheet4_final = df_so_filtered.sort_values(by=['Sub Office']).reset_index(drop=True)
    
    flat_idx = 2
    for item_idx, r in df_sheet4_final.iterrows():
        ws4.set_row(flat_idx, 19)
        ws4.write(flat_idx, 0, item_idx + 1, format_ctr)
        ws4.write(flat_idx, 1, r['Sub Division'], format_txt)
        ws4.write(flat_idx, 2, r['Sub Office'], format_txt)
        
        ws4.write(flat_idx, 3, r['par_Received'], format_int)
        ws4.write_formula(flat_idx, 4, f"=IFERROR(R{flat_idx+1}/D{flat_idx+1}, \"\")", format_pct)
        ws4.write_formula(flat_idx, 5, f"=IFERROR(S{flat_idx+1}/D{flat_idx+1}, \"\")", format_pct)
        
        ws4.write(flat_idx, 6, r['doc_Received'], format_int)
        ws4.write_formula(flat_idx, 7, f"=IFERROR(T{flat_idx+1}/G{flat_idx+1}, \"\")", format_pct)
        ws4.write_formula(flat_idx, 8, f"=IFERROR(U{flat_idx+1}/G{flat_idx+1}, \"\")", format_pct)
        
        ws4.write(flat_idx, 9, r['all_Received'], format_int)
        ws4.write_formula(flat_idx, 10, f"=IFERROR(V{flat_idx+1}/J{flat_idx+1}, \"\")", format_pct)
        ws4.write_formula(flat_idx, 11, f"=IFERROR(W{flat_idx+1}/J{flat_idx+1}, \"\")", format_pct)
        
        ws4.write_formula(flat_idx, 12, f"=IFERROR(X{flat_idx+1}/Y{flat_idx+1}, \"\")", format_pct)
        ws4.write_formula(flat_idx, 13, f"=IFERROR(Z{flat_idx+1}/AA{flat_idx+1}, \"\")", format_pct)
        ws4.write_formula(flat_idx, 14, f"=IFERROR(AB{flat_idx+1}/AC{flat_idx+1}, \"\")", format_pct)
        ws4.write_formula(flat_idx, 15, f"=IFERROR(AD{flat_idx+1}/AE{flat_idx+1}, \"\")", format_pct)
        ws4.write_formula(flat_idx, 16, f"=IFERROR(AF{flat_idx+1}/AG{flat_idx+1}, \"\")", format_pct)
        
        # Hidden metrics grid maps
        ws4.write(flat_idx, 17, (r['par_D0 Delivered'] + r['par_D0 Redirected'] + r['par_D0 Returned']), format_int)
        ws4.write(flat_idx, 18, (r['par_D1 Delivered'] + r['par_D1 Redirected'] + r['par_D1 Returned']), format_int)
        ws4.write(flat_idx, 19, (r['doc_D0 Delivered'] + r['doc_D0 Redirected'] + r['doc_D0 Returned']), format_int)
        ws4.write(flat_idx, 20, (r['doc_D1 Delivered'] + r['doc_D1 Redirected'] + r['doc_D1 Returned']), format_int)
        ws4.write(flat_idx, 21, (r['all_D0 Delivered'] + r['all_D0 Redirected'] + r['all_D0 Returned']), format_int)
        ws4.write(flat_idx, 22, (r['all_D1 Delivered'] + r['all_D1 Redirected'] + r['all_D1 Returned']), format_int)
        ws4.write(flat_idx, 23, r['prod_numer'], format_int)
        ws4.write(flat_idx, 24, r['prod_denom'], format_int)
        ws4.write(flat_idx, 25, r['dss_d_numer'], format_int)
        ws4.write(flat_idx, 26, r['dss_d_denom'], format_int)
        ws4.write(flat_idx, 27, r['dss_c_numer'], format_int)
        ws4.write(flat_idx, 28, r['dss_c_denom'], format_int)
        ws4.write(flat_idx, 29, r['cod_d_numer'], format_int)
        ws4.write(flat_idx, 30, r['cod_d_denom'], format_int)
        ws4.write(flat_idx, 31, r['cod_c_numer'], format_int)
        ws4.write(flat_idx, 32, r['cod_c_denom'], format_int)
        flat_idx += 1
        
    ws4.set_row(flat_idx, 22)
    for col_c in range(33):
        ws4.write_blank(flat_idx, col_c, "", style_grand_total)
    ws4.write(flat_idx, 1, "Total Row", style_grand_total)
    ws4.write_formula(flat_idx, 3, f"=SUM(D3:D{flat_idx})", format_int)
    ws4.write_formula(flat_idx, 4, f"=IFERROR(SUM(R3:R{flat_idx})/D{flat_idx+1}, \"\")", format_pct)
    ws4.write_formula(flat_idx, 5, f"=IFERROR(SUM(S3:S{flat_idx})/D{flat_idx+1}, \"\")", format_pct)
    ws4.write_formula(flat_idx, 6, f"=SUM(G3:G{flat_idx})", format_int)
    ws4.write_formula(flat_idx, 7, f"=IFERROR(SUM(T3:T{flat_idx})/G{flat_idx+1}, \"\")", format_pct)
    ws4.write_formula(flat_idx, 8, f"=IFERROR(SUM(U3:U{flat_idx})/G{flat_idx+1}, \"\")", format_pct)
    ws4.write_formula(flat_idx, 9, f"=SUM(J3:J{flat_idx})", format_int)
    ws4.write_formula(flat_idx, 10, f"=IFERROR(SUM(V3:V{flat_idx})/J{flat_idx+1}, \"\")", format_pct)
    ws4.write_formula(flat_idx, 11, f"=IFERROR(SUM(W3:W{flat_idx})/J{flat_idx+1}, \"\")", format_pct)
    ws4.write_formula(flat_idx, 12, f"=IFERROR(SUM(X3:X{flat_idx})/SUM(Y3:Y{flat_idx}), \"\")", format_pct)
    ws4.write_formula(flat_idx, 13, f"=IFERROR(SUM(Z3:Z{flat_idx})/SUM(AA3:AA{flat_idx}), \"\")", format_pct)
    ws4.write_formula(flat_idx, 14, f"=IFERROR(SUM(AB3:AB{flat_idx})/SUM(AC3:AC{flat_idx}), \"\")", format_pct)
    ws4.write_formula(flat_idx, 15, f"=IFERROR(SUM(AD3:AD{flat_idx})/SUM(AE3:AE{flat_idx}), \"\")", format_pct)
    ws4.write_formula(flat_idx, 16, f"=IFERROR(SUM(AF3:AF{flat_idx})/SUM(AG3:AG{flat_idx}), \"\")", format_pct)
    ws4.set_column(17, 32, None, None, {'hidden': True})

    # ----------------------------------------------------
    # SHEET 5: DEFAULTER OFFICES (STRICT PASS — BUG B-06)
    # ----------------------------------------------------
    def evaluate_defaulter(row):
        scores = []
        if row['all_Received'] > 0: scores.append(row['val_all_d0'] < 0.90)
        if row['prod_denom'] > 0: scores.append(row['val_prod'] < 0.90)
        if row['dss_c_denom'] > 0: scores.append(row['val_dss_c'] < 0.90)
        if row['cod_d_denom'] > 0: scores.append(row['val_cod_d'] < 0.90)
        if row['cod_c_denom'] > 0: scores.append(row['val_cod_c'] < 0.90)
        
        # Bug Fix B-06 compliance validation check
        return all(scores) if scores else False

    df_defaulters = core_ledger[core_ledger.apply(evaluate_defaulter, axis=1)].sort_values(by=['Canonical_Office_Name']).reset_index(drop=True)
    ws5 = workbook_obj.add_worksheet('Defaulter Offices')
    
    # Exclude All Products D+1% column from display matrix mapping (21 Columns)
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
        ws5.write_formula(def_idx, 7, f"=IFERROR(T{def_idx+1}/G{def_idx+1}, \"\")", format_pct)
        ws5.write(def_idx, 8, r['doc_Received'], format_int)
        ws5.write_formula(def_idx, 9, f"=IFERROR(U{def_idx+1}/I{def_idx+1}, \"\")", format_pct)
        ws5.write(def_idx, 10, r['all_Received'], format_int)
        ws5.write_formula(def_idx, 11, f"=IFERROR(V{def_idx+1}/K{def_idx+1}, \"\")", format_pct)
        ws5.write_formula(def_idx, 12, f"=IFERROR(W{def_idx+1}/K{def_idx+1}, \"\")", format_pct)
        ws5.write(def_idx, 13, r['val_all_not_invoiced'], format_int)
        
        ws5.write_formula(def_idx, 14, f"=IFERROR(X{def_idx+1}/Y{def_idx+1}, \"\")", format_pct)
        ws5.write_formula(def_idx, 15, f"=IFERROR(Z{def_idx+1}/AA{def_idx+1}, \"\")", format_pct)
        ws5.write_formula(def_idx, 16, f"=IFERROR(AB{def_idx+1}/AC{def_idx+1}, \"\")", format_pct)
        ws5.write_formula(def_idx, 17, f"=IFERROR(AD{def_idx+1}/AE{def_idx+1}, \"\")", format_pct)
        ws5.write_formula(def_idx, 18, f"=IFERROR(AF{def_idx+1}/AG{def_idx+1}, \"\")", format_pct)
        
        # Populate processing structures inside hidden columns
        ws5.write(def_idx, 19, (r['par_D0 Delivered'] + r['par_D0 Redirected'] + r['par_D0 Returned']), format_int)
        ws5.write(def_idx, 20, (r['doc_D0 Delivered'] + r['doc_D0 Redirected'] + r['doc_D0 Returned']), format_int)
        ws5.write(def_idx, 21, (r['all_D0 Delivered'] + r['all_D0 Redirected'] + r['all_D0 Returned']), format_int)
        ws5.write(def_idx, 22, r['all_D0 Returned'], format_int)
        ws5.write(def_idx, 23, r['prod_numer'], format_int)
        ws5.write(def_idx, 24, r['prod_denom'], format_int)
        ws5.write(def_idx, 25, r['dss_d_numer'], format_int)
        ws5.write(def_idx, 26, r['dss_d_denom'], format_int)
        ws5.write(def_idx, 27, r['dss_c_numer'], format_int)
        ws5.write(def_idx, 28, r['dss_c_denom'], format_int)
        ws5.write(def_idx, 29, r['cod_d_numer'], format_int)
        ws5.write(def_idx, 30, r['cod_d_denom'], format_int)
        ws5.write(def_idx, 31, r['cod_c_numer'], format_int)
        ws5.write(def_idx, 32, r['cod_c_denom'], format_int)
        def_idx += 1
    ws5.set_column(19, 32, None, None, {'hidden': True})

    # Formatting Passes (Prevents Truncation / Column Width Issues)
    for ws_target in [ws1, ws2, ws3, ws4, ws5]:
        ws_target.set_column(0, 0, 7)
        ws_target.set_column(1, 3, 26)
        ws_target.set_column(4, 5, 13)
        ws_target.set_column(6, 18, 15)

    excel_engine.close()
    output_stream.seek(0)
    
    st.success("🏁 Verification complete. Unified ledger compiled successfully.")
    
    if st.session_state.system_diagnostics:
        with st.expander("🔬 System Ingestion Logs & Alerts", expanded=True):
            for entry in st.session_state.system_diagnostics:
                st.markdown(entry)
                
    st.download_button(
        label="📥 Download Production Workbook", 
        data=output_stream, 
        file_name="MMU_Report_Karad.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
        use_container_width=True
    )
