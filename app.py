"""
Application: MMU Gem Karad Division Data Processing Utility
Engineers: Senior Full-Stack Engineer / Data Architect / Python Developer
Dependencies: streamlit, pandas, openpyxl, xlsxwriter (Automated Self-Heal bootstrap enabled)
Validation Framework: Fully compliant with Bug Fixes B-01 through B-15
"""

import subprocess
import sys

# ==========================================
# 0. SELF-HEALING BOOTSTRAP LOADER MATRIX
# ==========================================
# Forcefully hooks the host virtual environment and resolves missing modules on the fly
for package in ["xlsxwriter", "openpyxl"]:
    try:
        __import__(package)
    except ModuleNotFoundError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

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
    style_sub_header  = workbook_obj.add_format({'bg_color': '#DCE6F1', 'font_color': '#1F497D', 'bold': True, 'align': 'center', 'valign': 'vcenter', 'border
