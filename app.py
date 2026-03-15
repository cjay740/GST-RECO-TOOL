"""
GST ITC Reconciliation Tool
============================
Reconciles GST ITC as per Books vs GST Portal (GSTR-2B/2A).
Built for Indian Chartered Accountants — monthly reconciliation workflow.

Two matching modes:
  1. GSTIN-Level Summary — aggregates all invoices per GSTIN and compares totals
  2. GSTIN + Invoice No — line-by-line invoice matching

Run:  streamlit run app.py
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

# ─────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────
st.set_page_config(page_title="GST ITC Reconciliation", page_icon="📊", layout="wide")

# ─────────────────────────────────────────────
# Custom CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        font-size: 2rem; font-weight: 700; color: #1a5276;
        border-bottom: 3px solid #2e86c1; padding-bottom: 10px; margin-bottom: 20px;
    }
    .metric-card {
        background: #f8f9fa; border-radius: 10px; padding: 15px;
        border-left: 4px solid #2e86c1; margin: 5px 0;
    }
    .match-ok { color: #27ae60; font-weight: bold; }
    .match-diff { color: #e74c3c; font-weight: bold; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #f0f2f6; border-radius: 6px 6px 0 0; padding: 8px 20px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">GST ITC Reconciliation Tool</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# Sidebar — Column Mapping
# ─────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Column Mapping")
    st.caption("Map your file columns below. Defaults are pre-set for the standard format.")

    st.subheader("📘 Books File Columns")
    b_gstin = st.text_input("GSTIN Column (Books)", value="GSTIN Number")
    b_supplier = st.text_input("Supplier Name Column (Books)", value="Supplier")
    b_invoice = st.text_input("Invoice No Column (Books)", value="Invoice No")
    b_taxable = st.text_input("Taxable Value Column (Books)", value="Sum of Taxable Value")
    b_igst = st.text_input("IGST Column (Books)", value="Sum of Integrated Tax")
    b_cgst = st.text_input("CGST Column (Books)", value="Sum of Central Tax")
    b_sgst = st.text_input("SGST Column (Books)", value="Sum of State UT Tax")
    b_cess = st.text_input("Cess Column (Books)", value="Sum of CESS Tax")

    st.divider()
    st.subheader("🌐 Portal File Columns")
    p_gstin = st.text_input("GSTIN Column (Portal)", value="GSTIN of supplier")
    p_supplier = st.text_input("Supplier Name Column (Portal)", value="Trade/Legal name")
    p_invoice = st.text_input("Invoice No Column (Portal)", value="Invoice number")
    p_taxable = st.text_input("Taxable Value Column (Portal)", value="Sum of Taxable Value (₹)")
    p_igst = st.text_input("IGST Column (Portal)", value="Sum of Integrated Tax(₹)")
    p_cgst = st.text_input("CGST Column (Portal)", value="Sum of Central Tax(₹)")
    p_sgst = st.text_input("SGST Column (Portal)", value="Sum of State/UT Tax(₹)")
    p_cess = st.text_input("Cess Column (Portal)", value="Sum of Cess(₹)")
    p_itc_avail = st.text_input("ITC Availability Column (Portal)", value="ITC Availability")

    st.divider()
    st.subheader("🔧 Tolerance")
    tolerance = st.number_input("Amount tolerance (₹)", min_value=0.0, value=1.0, step=0.5,
                                help="Differences within this amount are treated as matched.")

# ─────────────────────────────────────────────
# Helper functions
# ─────────────────────────────────────────────
def clean_gstin(s):
    """Normalise GSTIN: strip, uppercase, treat blanks as NaN."""
    if pd.isna(s):
        return np.nan
    s = str(s).strip().upper()
    return np.nan if s == "" else s

def clean_invoice(s):
    """Normalise invoice numbers: strip, uppercase, remove leading zeros."""
    if pd.isna(s):
        return ""
    s = str(s).strip().upper()
    # Remove common prefixes/suffixes that differ between books and portal
    s = s.lstrip("0")
    return s

def safe_float(col):
    return pd.to_numeric(col, errors="coerce").fillna(0.0)

def load_and_normalise_books(df):
    """Normalise Books dataframe to standard internal columns."""
    out = pd.DataFrame()
    out["GSTIN"] = df[b_gstin].apply(clean_gstin)
    out["Supplier_Name"] = df[b_supplier].astype(str).str.strip()
    out["Invoice_No"] = df[b_invoice].apply(clean_invoice)
    out["Taxable_Value"] = safe_float(df[b_taxable])
    out["IGST"] = safe_float(df[b_igst])
    out["CGST"] = safe_float(df[b_cgst])
    out["SGST"] = safe_float(df[b_sgst])
    out["Cess"] = safe_float(df[b_cess])
    out["Total_Tax"] = out["IGST"] + out["CGST"] + out["SGST"] + out["Cess"]
    out["Total_Value"] = out["Taxable_Value"] + out["Total_Tax"]
    return out

def load_and_normalise_portal(df):
    """Normalise Portal dataframe to standard internal columns."""
    out = pd.DataFrame()
    out["GSTIN"] = df[p_gstin].apply(clean_gstin)
    out["Supplier_Name"] = df[p_supplier].astype(str).str.strip()
    out["Invoice_No"] = df[p_invoice].apply(clean_invoice)
    out["Taxable_Value"] = safe_float(df[p_taxable])
    out["IGST"] = safe_float(df[p_igst])
    out["CGST"] = safe_float(df[p_cgst])
    out["SGST"] = safe_float(df[p_sgst])
    out["Cess"] = safe_float(df[p_cess])
    out["Total_Tax"] = out["IGST"] + out["CGST"] + out["SGST"] + out["Cess"]
    out["Total_Value"] = out["Taxable_Value"] + out["Total_Tax"]
    if p_itc_avail in df.columns:
        out["ITC_Availability"] = df[p_itc_avail].astype(str).str.strip()
    return out

def to_excel_download(sheets_dict, index=False):
    """Convert dict of {sheet_name: dataframe} to downloadable Excel bytes."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets_dict.items():
            # Excel sheet names max 31 chars
            sheet_name = name[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=index)
    buf.seek(0)
    return buf

# ─────────────────────────────────────────────
# GSTIN-Level Summary Reconciliation
# ─────────────────────────────────────────────
def reconcile_gstin_level(books_df, portal_df, tol):
    """Aggregate by GSTIN and compare totals."""
    agg_cols = ["Taxable_Value", "IGST", "CGST", "SGST", "Cess", "Total_Tax", "Total_Value"]

    b_agg = books_df.dropna(subset=["GSTIN"]).groupby("GSTIN", as_index=False).agg(
        Supplier_Name_Books=("Supplier_Name", "first"),
        Invoice_Count_Books=("Invoice_No", "count"),
        Taxable_Books=("Taxable_Value", "sum"),
        IGST_Books=("IGST", "sum"),
        CGST_Books=("CGST", "sum"),
        SGST_Books=("SGST", "sum"),
        Cess_Books=("Cess", "sum"),
        TotalTax_Books=("Total_Tax", "sum"),
    )

    p_agg = portal_df.dropna(subset=["GSTIN"]).groupby("GSTIN", as_index=False).agg(
        Supplier_Name_Portal=("Supplier_Name", "first"),
        Invoice_Count_Portal=("Invoice_No", "count"),
        Taxable_Portal=("Taxable_Value", "sum"),
        IGST_Portal=("IGST", "sum"),
        CGST_Portal=("CGST", "sum"),
        SGST_Portal=("SGST", "sum"),
        Cess_Portal=("Cess", "sum"),
        TotalTax_Portal=("Total_Tax", "sum"),
    )

    # Merge
    merged = pd.merge(b_agg, p_agg, on="GSTIN", how="outer", indicator=True)

    # Compute differences
    for col_suffix in ["Taxable", "IGST", "CGST", "SGST", "Cess", "TotalTax"]:
        b_col = f"{col_suffix}_Books"
        p_col = f"{col_suffix}_Portal"
        merged[b_col] = merged[b_col].fillna(0)
        merged[p_col] = merged[p_col].fillna(0)
        merged[f"Diff_{col_suffix}"] = merged[b_col] - merged[p_col]

    merged["Invoice_Count_Books"] = merged["Invoice_Count_Books"].fillna(0).astype(int)
    merged["Invoice_Count_Portal"] = merged["Invoice_Count_Portal"].fillna(0).astype(int)

    # Classify
    def classify(row):
        if row["_merge"] == "left_only":
            return "Only in Books"
        elif row["_merge"] == "right_only":
            return "Only in Portal"
        elif abs(row["Diff_TotalTax"]) <= tol and abs(row["Diff_Taxable"]) <= tol:
            return "Matched"
        else:
            return "Mismatched"

    merged["Status"] = merged.apply(classify, axis=1)
    merged.drop(columns=["_merge"], inplace=True)

    # Round
    num_cols = merged.select_dtypes(include=[np.number]).columns
    merged[num_cols] = merged[num_cols].round(2)

    # Split
    matched = merged[merged["Status"] == "Matched"].copy()
    mismatched = merged[merged["Status"] == "Mismatched"].copy()
    only_books = merged[merged["Status"] == "Only in Books"].copy()
    only_portal = merged[merged["Status"] == "Only in Portal"].copy()

    # Also capture books entries with no GSTIN
    no_gstin = books_df[books_df["GSTIN"].isna()].copy()

    return merged, matched, mismatched, only_books, only_portal, no_gstin

# ─────────────────────────────────────────────
# GSTIN + Invoice Level Reconciliation
# ─────────────────────────────────────────────
def reconcile_invoice_level(books_df, portal_df, tol):
    """Match on GSTIN + Invoice No."""
    b = books_df.dropna(subset=["GSTIN"]).copy()
    p = portal_df.dropna(subset=["GSTIN"]).copy()

    # Merge key
    b["_key"] = b["GSTIN"] + "|" + b["Invoice_No"]
    p["_key"] = p["GSTIN"] + "|" + p["Invoice_No"]

    # Rename columns for clarity
    b_renamed = b.rename(columns={
        "Supplier_Name": "Supplier_Books",
        "Taxable_Value": "Taxable_Books",
        "IGST": "IGST_Books", "CGST": "CGST_Books",
        "SGST": "SGST_Books", "Cess": "Cess_Books",
        "Total_Tax": "TotalTax_Books", "Total_Value": "TotalValue_Books",
    })
    p_renamed = p.rename(columns={
        "Supplier_Name": "Supplier_Portal",
        "Taxable_Value": "Taxable_Portal",
        "IGST": "IGST_Portal", "CGST": "CGST_Portal",
        "SGST": "SGST_Portal", "Cess": "Cess_Portal",
        "Total_Tax": "TotalTax_Portal", "Total_Value": "TotalValue_Portal",
    })

    b_cols = ["_key", "GSTIN", "Invoice_No", "Supplier_Books",
              "Taxable_Books", "IGST_Books", "CGST_Books", "SGST_Books",
              "Cess_Books", "TotalTax_Books", "TotalValue_Books"]
    p_cols = ["_key", "Supplier_Portal",
              "Taxable_Portal", "IGST_Portal", "CGST_Portal", "SGST_Portal",
              "Cess_Portal", "TotalTax_Portal", "TotalValue_Portal"]
    if "ITC_Availability" in p_renamed.columns:
        p_cols.append("ITC_Availability")

    merged = pd.merge(b_renamed[b_cols], p_renamed[p_cols], on="_key", how="outer", indicator=True)

    # Fill GSTIN/Invoice from portal side for right_only
    right_mask = merged["_merge"] == "right_only"
    if right_mask.any():
        # Get GSTIN and Invoice from portal for right_only rows
        right_keys = merged.loc[right_mask, "_key"]
        for idx in right_keys.index:
            key = merged.loc[idx, "_key"]
            if pd.notna(key) and "|" in str(key):
                parts = str(key).split("|", 1)
                merged.loc[idx, "GSTIN"] = parts[0]
                merged.loc[idx, "Invoice_No"] = parts[1]

    # Differences
    for col_suffix in ["Taxable", "IGST", "CGST", "SGST", "Cess", "TotalTax"]:
        b_col = f"{col_suffix}_Books"
        p_col = f"{col_suffix}_Portal"
        merged[b_col] = merged[b_col].fillna(0)
        merged[p_col] = merged[p_col].fillna(0)
        merged[f"Diff_{col_suffix}"] = merged[b_col] - merged[p_col]

    # Classify
    def classify(row):
        if row["_merge"] == "left_only":
            return "Only in Books"
        elif row["_merge"] == "right_only":
            return "Only in Portal"
        elif abs(row["Diff_TotalTax"]) <= tol and abs(row["Diff_Taxable"]) <= tol:
            return "Matched"
        else:
            return "Mismatched"

    merged["Status"] = merged.apply(classify, axis=1)
    merged.drop(columns=["_merge", "_key"], inplace=True)

    # Round
    num_cols = merged.select_dtypes(include=[np.number]).columns
    merged[num_cols] = merged[num_cols].round(2)

    matched = merged[merged["Status"] == "Matched"].copy()
    mismatched = merged[merged["Status"] == "Mismatched"].copy()
    only_books = merged[merged["Status"] == "Only in Books"].copy()
    only_portal = merged[merged["Status"] == "Only in Portal"].copy()

    no_gstin = books_df[books_df["GSTIN"].isna()].copy()

    return merged, matched, mismatched, only_books, only_portal, no_gstin

# ─────────────────────────────────────────────
# Main UI
# ─────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    books_file = st.file_uploader("📘 Upload **Books** file (.xlsx)", type=["xlsx", "xls"])
with col2:
    portal_file = st.file_uploader("🌐 Upload **Portal** file (.xlsx)", type=["xlsx", "xls"])

recon_mode = st.radio(
    "🔍 Reconciliation Mode",
    ["GSTIN-Level Summary", "GSTIN + Invoice No (Detailed)"],
    horizontal=True,
    help="Summary mode aggregates per GSTIN. Detailed mode matches each invoice."
)

run_btn = st.button("🚀 Run Reconciliation", type="primary", use_container_width=True)

# ─────────────────────────────────────────────
if run_btn:
    if not books_file or not portal_file:
        st.error("Please upload both Books and Portal files.")
        st.stop()

    with st.spinner("Reading files..."):
        try:
            raw_books = pd.read_excel(books_file)
            raw_portal = pd.read_excel(portal_file)
        except Exception as e:
            st.error(f"Error reading files: {e}")
            st.stop()

    # Validate columns exist
    missing_b = [c for c in [b_gstin, b_supplier, b_invoice, b_taxable, b_igst, b_cgst, b_sgst, b_cess]
                 if c not in raw_books.columns]
    missing_p = [c for c in [p_gstin, p_supplier, p_invoice, p_taxable, p_igst, p_cgst, p_sgst, p_cess]
                 if c not in raw_portal.columns]
    if missing_b:
        st.error(f"Books file missing columns: {missing_b}. Please check sidebar column mapping.")
        st.stop()
    if missing_p:
        st.error(f"Portal file missing columns: {missing_p}. Please check sidebar column mapping.")
        st.stop()

    with st.spinner("Normalising data..."):
        books_df = load_and_normalise_books(raw_books)
        portal_df = load_and_normalise_portal(raw_portal)

    with st.spinner("Running reconciliation..."):
        if recon_mode == "GSTIN-Level Summary":
            full, matched, mismatched, only_books, only_portal, no_gstin = \
                reconcile_gstin_level(books_df, portal_df, tolerance)
        else:
            full, matched, mismatched, only_books, only_portal, no_gstin = \
                reconcile_invoice_level(books_df, portal_df, tolerance)

    # ─── Summary metrics ───
    st.divider()
    st.subheader("📊 Reconciliation Summary")
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total Records", len(full))
    m2.metric("✅ Matched", len(matched))
    m3.metric("⚠️ Mismatched", len(mismatched))
    m4.metric("📘 Only in Books", len(only_books))
    m5.metric("🌐 Only in Portal", len(only_portal))

    if len(no_gstin) > 0:
        st.info(f"ℹ️ {len(no_gstin)} entries in Books have no GSTIN (e.g., Petty Cash) — shown in a separate sheet.")

    # ─── Amount summary ───
    st.subheader("💰 Amount Summary")
    amt_cols = {
        "Source": ["Books (with GSTIN)", "Portal", "Difference"],
    }
    bk = books_df.dropna(subset=["GSTIN"])
    for label, b_col, p_col in [
        ("Taxable Value", "Taxable_Value", "Taxable_Value"),
        ("IGST", "IGST", "IGST"),
        ("CGST", "CGST", "CGST"),
        ("SGST", "SGST", "SGST"),
        ("Cess", "Cess", "Cess"),
        ("Total Tax", "Total_Tax", "Total_Tax"),
    ]:
        b_sum = bk[b_col].sum()
        p_sum = portal_df[p_col].sum()
        amt_cols[label] = [round(b_sum, 2), round(p_sum, 2), round(b_sum - p_sum, 2)]

    amt_df = pd.DataFrame(amt_cols)
    st.dataframe(amt_df, use_container_width=True, hide_index=True)

    # ─── Detailed tabs ───
    st.divider()
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        f"✅ Matched ({len(matched)})",
        f"⚠️ Mismatched ({len(mismatched)})",
        f"📘 Only in Books ({len(only_books)})",
        f"🌐 Only in Portal ({len(only_portal)})",
        f"🚫 No GSTIN ({len(no_gstin)})",
    ])

    with tab1:
        if len(matched) > 0:
            st.dataframe(matched, use_container_width=True, hide_index=True)
        else:
            st.info("No matched records found.")
    with tab2:
        if len(mismatched) > 0:
            st.dataframe(mismatched, use_container_width=True, hide_index=True)
        else:
            st.success("No mismatches — everything balances!")
    with tab3:
        if len(only_books) > 0:
            st.dataframe(only_books, use_container_width=True, hide_index=True)
        else:
            st.info("No records found only in Books.")
    with tab4:
        if len(only_portal) > 0:
            st.dataframe(only_portal, use_container_width=True, hide_index=True)
        else:
            st.info("No records found only in Portal.")
    with tab5:
        if len(no_gstin) > 0:
            st.dataframe(no_gstin, use_container_width=True, hide_index=True)
        else:
            st.info("All books entries have a GSTIN.")

    # ─── Download ───
    st.divider()
    st.subheader("📥 Download Results")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    mode_tag = "GSTIN" if recon_mode == "GSTIN-Level Summary" else "Invoice"
    filename = f"GST_Reconciliation_{mode_tag}_{timestamp}.xlsx"

    sheets = {
        "Summary": amt_df,
        "Full Reconciliation": full,
        "Matched": matched,
        "Mismatched": mismatched,
        "Only in Books": only_books,
        "Only in Portal": only_portal,
        "No GSTIN (Books)": no_gstin,
    }

    excel_bytes = to_excel_download(sheets)

    st.download_button(
        label=f"⬇️ Download Full Report — {filename}",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

# ─── Footer ───
st.divider()
st.caption("GST ITC Reconciliation Tool • Built for CA professionals • Books vs Portal (GSTR-2B/2A)")
