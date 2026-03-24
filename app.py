"""
GST ITC Reconciliation Tool — Advanced Version
================================================
New in v2:
  • Fuzzy invoice matching — catches near-matches like 'INV-123' vs 'INV123'
  • Colour-coded Excel download — green/red/yellow/purple traffic-light rows
  • Supplier drill-down — pick any GSTIN and see all its invoices side by side

Run:  streamlit run app.py
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from difflib import SequenceMatcher
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────
st.set_page_config(page_title="GST ITC Reconciliation", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .main-header {
        font-size: 2rem; font-weight: 700; color: #1a5276;
        border-bottom: 3px solid #2e86c1; padding-bottom: 10px; margin-bottom: 20px;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #f0f2f6; border-radius: 6px 6px 0 0; padding: 8px 20px;
    }
    .badge-matched   { background:#d4edda; color:#155724; padding:2px 10px; border-radius:12px; font-weight:600; }
    .badge-mismatch  { background:#f8d7da; color:#721c24; padding:2px 10px; border-radius:12px; font-weight:600; }
    .badge-books     { background:#cce5ff; color:#004085; padding:2px 10px; border-radius:12px; font-weight:600; }
    .badge-portal    { background:#e2d9f3; color:#432874; padding:2px 10px; border-radius:12px; font-weight:600; }
    .badge-fuzzy     { background:#fff3cd; color:#856404; padding:2px 10px; border-radius:12px; font-weight:600; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">📊 GST ITC Reconciliation Tool</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# Sidebar
# ─────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")

    with st.expander("📘 Books Column Names", expanded=False):
        b_gstin    = st.text_input("GSTIN",         value="GSTIN Number",          key="b_gstin")
        b_supplier = st.text_input("Supplier Name", value="Supplier",              key="b_sup")
        b_invoice  = st.text_input("Invoice No",    value="Invoice No",            key="b_inv")
        b_taxable  = st.text_input("Taxable Value", value="Sum of Taxable Value",  key="b_tax")
        b_igst     = st.text_input("IGST",          value="Sum of Integrated Tax", key="b_igst")
        b_cgst     = st.text_input("CGST",          value="Sum of Central Tax",    key="b_cgst")
        b_sgst     = st.text_input("SGST",          value="Sum of State UT Tax",   key="b_sgst")
        b_cess     = st.text_input("Cess",          value="Sum of CESS Tax",       key="b_cess")

    with st.expander("🌐 Portal Column Names", expanded=False):
        p_gstin    = st.text_input("GSTIN",         value="GSTIN of supplier",          key="p_gstin")
        p_supplier = st.text_input("Supplier Name", value="Trade/Legal name",           key="p_sup")
        p_invoice  = st.text_input("Invoice No",    value="Invoice number",             key="p_inv")
        p_taxable  = st.text_input("Taxable Value", value="Sum of Taxable Value (₹)",   key="p_tax")
        p_igst     = st.text_input("IGST",          value="Sum of Integrated Tax(₹)",   key="p_igst")
        p_cgst     = st.text_input("CGST",          value="Sum of Central Tax(₹)",      key="p_cgst")
        p_sgst     = st.text_input("SGST",          value="Sum of State/UT Tax(₹)",     key="p_sgst")
        p_cess     = st.text_input("Cess",          value="Sum of Cess(₹)",             key="p_cess")
        p_itc_avail= st.text_input("ITC Availability", value="ITC Availability",        key="p_itc")

    st.divider()
    st.subheader("🔧 Matching Settings")
    tolerance = st.number_input("Amount tolerance (₹)", min_value=0.0, value=1.0, step=0.5,
                                help="Differences within this ₹ amount are treated as matched.")
    fuzzy_threshold = st.slider("Fuzzy match sensitivity", min_value=50, max_value=100, value=80,
                                help="How similar two invoice numbers must be to be fuzzy-matched. 100 = exact only, 80 = allows minor differences.")

# ─────────────────────────────────────────────
# Helper functions
# ─────────────────────────────────────────────
def clean_gstin(s):
    if pd.isna(s): return np.nan
    s = str(s).strip().upper()
    return np.nan if s == "" else s

def clean_invoice(s):
    if pd.isna(s): return ""
    s = str(s).strip().upper().lstrip("0")
    return s

def safe_float(col):
    return pd.to_numeric(col, errors="coerce").fillna(0.0)

def similarity(a, b):
    """Return 0-100 similarity score between two strings."""
    return round(SequenceMatcher(None, str(a), str(b)).ratio() * 100, 1)

def load_books(df):
    out = pd.DataFrame()
    out["GSTIN"]         = df[b_gstin].apply(clean_gstin)
    out["Supplier_Name"] = df[b_supplier].astype(str).str.strip()
    out["Invoice_No"]    = df[b_invoice].apply(clean_invoice)
    out["Taxable_Value"] = safe_float(df[b_taxable])
    out["IGST"]          = safe_float(df[b_igst])
    out["CGST"]          = safe_float(df[b_cgst])
    out["SGST"]          = safe_float(df[b_sgst])
    out["Cess"]          = safe_float(df[b_cess])
    out["Total_Tax"]     = out["IGST"] + out["CGST"] + out["SGST"] + out["Cess"]
    out["Total_Value"]   = out["Taxable_Value"] + out["Total_Tax"]
    return out

def load_portal(df):
    out = pd.DataFrame()
    out["GSTIN"]         = df[p_gstin].apply(clean_gstin)
    out["Supplier_Name"] = df[p_supplier].astype(str).str.strip()
    out["Invoice_No"]    = df[p_invoice].apply(clean_invoice)
    out["Taxable_Value"] = safe_float(df[p_taxable])
    out["IGST"]          = safe_float(df[p_igst])
    out["CGST"]          = safe_float(df[p_cgst])
    out["SGST"]          = safe_float(df[p_sgst])
    out["Cess"]          = safe_float(df[p_cess])
    out["Total_Tax"]     = out["IGST"] + out["CGST"] + out["SGST"] + out["Cess"]
    out["Total_Value"]   = out["Taxable_Value"] + out["Total_Tax"]
    if p_itc_avail in df.columns:
        out["ITC_Availability"] = df[p_itc_avail].astype(str).str.strip()
    return out

# ─────────────────────────────────────────────
# Fuzzy matching helpers
# ─────────────────────────────────────────────
def _get_tax(row, books_side):
    """Safely extract total tax from a row regardless of column naming."""
    if books_side:
        return float(row.get("TotalTax_Books", row.get("Total_Tax", 0)) or 0)
    return float(row.get("TotalTax_Portal", row.get("Total_Tax", 0)) or 0)

def _get_taxable(row, books_side):
    return float(row.get("Taxable_Books", row.get("Taxable_Value", 0)) or 0)

def _build_fuzzy_row(match_type, gstin_books, gstin_portal, b_row, p_row, inv_score, amt_tol):
    b_tax  = _get_tax(b_row, True)
    p_tax  = _get_tax(p_row, False)
    b_taxv = _get_taxable(b_row, True)
    p_taxv = _get_taxable(p_row, False)
    amt_ok = abs(b_tax - p_tax) <= amt_tol

    if match_type == "same_gstin":
        status = "Fuzzy Matched ✓" if amt_ok else "Fuzzy Matched ⚠ Amt Diff"
        note   = "Same GSTIN, similar invoice no."
    else:
        status = "Probable Match ✓" if amt_ok else "Probable Match ⚠ Amt Diff"
        note   = "Diff GSTIN, similar invoice no. + amount"

    return {
        "Match_Type":        match_type.replace("_", " ").title(),
        "GSTIN_Books":       gstin_books,
        "GSTIN_Portal":      gstin_portal,
        "Supplier_Books":    b_row.get("Supplier_Name", b_row.get("Supplier_Books", "")),
        "Supplier_Portal":   p_row.get("Supplier_Name", p_row.get("Supplier_Portal", "")),
        "Invoice_Books":     b_row.get("Invoice_No", ""),
        "Invoice_Portal":    p_row.get("Invoice_No", ""),
        "Invoice_Similarity_%": inv_score,
        "Taxable_Books":     round(b_taxv, 2),
        "Taxable_Portal":    round(p_taxv, 2),
        "TotalTax_Books":    round(b_tax, 2),
        "TotalTax_Portal":   round(p_tax, 2),
        "Tax_Difference":    round(b_tax - p_tax, 2),
        "Note":              note,
        "Status":            status,
    }

def apply_fuzzy_matching(only_books_df, only_portal_df, threshold, amt_tol=1.0):
    """
    Two-pass fuzzy matching:

    Pass 1 — Same GSTIN, similar invoice number (threshold%)
              → labelled 'Fuzzy Matched'

    Pass 2 — Different GSTIN, but invoice numbers are similar (threshold%)
              AND tax amounts are close (within amt_tol * 5)
              → labelled 'Probable Match' (needs human review)
    """
    if only_books_df.empty or only_portal_df.empty:
        return only_books_df, only_portal_df, pd.DataFrame()

    fuzzy_rows      = []
    books_used_idx  = set()
    portal_used_idx = set()

    # ── Pass 1: same GSTIN fuzzy match ────────────────────────
    common_gstins = (
        set(only_books_df["GSTIN"].dropna()) &
        set(only_portal_df["GSTIN"].dropna())
    )

    for gstin in common_gstins:
        b_sub = only_books_df[only_books_df["GSTIN"] == gstin]
        p_sub = only_portal_df[only_portal_df["GSTIN"] == gstin]

        for b_idx, b_row in b_sub.iterrows():
            if b_idx in books_used_idx:
                continue
            best_score, best_p_idx = -1, None
            for p_idx, p_row in p_sub.iterrows():
                if p_idx in portal_used_idx:
                    continue
                score = similarity(b_row["Invoice_No"], p_row["Invoice_No"])
                if score >= threshold and score > best_score:
                    best_score, best_p_idx = score, p_idx

            if best_p_idx is not None:
                p_row = only_portal_df.loc[best_p_idx]
                fuzzy_rows.append(_build_fuzzy_row(
                    "same_gstin", gstin, gstin, b_row, p_row, best_score, amt_tol
                ))
                books_used_idx.add(b_idx)
                portal_used_idx.add(best_p_idx)

    # ── Pass 2: cross-GSTIN probable match ────────────────────
    # Criteria (any of the below combinations qualifies):
    #   A) Invoice similar + amounts close
    #   B) Supplier name similar + amounts close (catches name-based mismatches)
    #   C) Invoice similar + supplier name similar (even if small amount diff)
    # In all cases GSTIN must differ from pass-1 matches.

    b_rem = only_books_df.drop(index=list(books_used_idx))
    p_rem = only_portal_df.drop(index=list(portal_used_idx))

    probable_amt_tol    = max(amt_tol * 5, 10.0)   # generous tolerance for cross-GSTIN
    probable_rows       = []
    probable_b_used_idx = set()
    probable_p_used_idx = set()

    for b_idx, b_row in b_rem.iterrows():
        if b_idx in probable_b_used_idx:
            continue
        b_inv  = b_row.get("Invoice_No", "")
        b_name = b_row.get("Supplier_Name", b_row.get("Supplier_Books", ""))
        b_tax  = _get_tax(b_row, True)

        best_combined, best_p_idx = -1, None

        for p_idx, p_row in p_rem.iterrows():
            if p_idx in probable_p_used_idx:
                continue

            p_tax  = _get_tax(p_row, False)
            p_inv  = p_row.get("Invoice_No", "")
            p_name = p_row.get("Supplier_Name", p_row.get("Supplier_Portal", ""))

            amt_close = abs(b_tax - p_tax) <= probable_amt_tol

            inv_score  = similarity(b_inv,  p_inv)
            name_score = similarity(b_name, p_name)

            inv_match  = inv_score  >= threshold
            name_match = name_score >= threshold

            # Must satisfy at least one qualifying combination
            qualifies = (
                (amt_close and inv_match)   or   # A: amount + invoice
                (amt_close and name_match)  or   # B: amount + name
                (inv_match  and name_match)       # C: invoice + name (regardless of amount)
            )

            if not qualifies:
                continue

            # Combined score — weight: invoice 50%, name 30%, amount closeness 20%
            amt_score = max(0, 100 - abs(b_tax - p_tax))   # higher = closer amount
            combined  = inv_score * 0.5 + name_score * 0.3 + min(amt_score, 100) * 0.2

            if combined > best_combined:
                best_combined, best_p_idx = combined, p_idx

        if best_p_idx is not None:
            p_row   = p_rem.loc[best_p_idx]
            gstin_b = str(b_row.get("GSTIN", ""))
            gstin_p = str(p_row.get("GSTIN", ""))
            if gstin_b != gstin_p:
                p_inv  = p_row.get("Invoice_No", "")
                p_name = p_row.get("Supplier_Name", p_row.get("Supplier_Portal", ""))
                inv_s  = similarity(b_inv, p_inv)
                # Store best invoice score for display
                row = _build_fuzzy_row(
                    "cross_gstin", gstin_b, gstin_p, b_row, p_row, inv_s, amt_tol
                )
                row["Name_Similarity_%"] = similarity(b_name, p_name)
                row["Combined_Score_%"]  = round(best_combined, 1)
                probable_rows.append(row)
                probable_b_used_idx.add(b_idx)
                probable_p_used_idx.add(best_p_idx)

    all_fuzzy = pd.DataFrame(fuzzy_rows + probable_rows)

    all_used_b = books_used_idx  | probable_b_used_idx
    all_used_p = portal_used_idx | probable_p_used_idx

    remaining_books  = only_books_df.drop(index=list(all_used_b))
    remaining_portal = only_portal_df.drop(index=list(all_used_p))

    return (
        remaining_books.reset_index(drop=True),
        remaining_portal.reset_index(drop=True),
        all_fuzzy
    )

# ─────────────────────────────────────────────
# GSTIN-Level Reconciliation
# ─────────────────────────────────────────────
def reconcile_gstin_level(books_df, portal_df, tol):
    b_agg = books_df.dropna(subset=["GSTIN"]).groupby("GSTIN", as_index=False).agg(
        Supplier_Name_Books  =("Supplier_Name",  "first"),
        Invoice_Count_Books  =("Invoice_No",     "count"),
        Taxable_Books        =("Taxable_Value",  "sum"),
        IGST_Books           =("IGST",           "sum"),
        CGST_Books           =("CGST",           "sum"),
        SGST_Books           =("SGST",           "sum"),
        Cess_Books           =("Cess",           "sum"),
        TotalTax_Books       =("Total_Tax",      "sum"),
    )
    p_agg = portal_df.dropna(subset=["GSTIN"]).groupby("GSTIN", as_index=False).agg(
        Supplier_Name_Portal =("Supplier_Name",  "first"),
        Invoice_Count_Portal =("Invoice_No",     "count"),
        Taxable_Portal       =("Taxable_Value",  "sum"),
        IGST_Portal          =("IGST",           "sum"),
        CGST_Portal          =("CGST",           "sum"),
        SGST_Portal          =("SGST",           "sum"),
        Cess_Portal          =("Cess",           "sum"),
        TotalTax_Portal      =("Total_Tax",      "sum"),
    )

    merged = pd.merge(b_agg, p_agg, on="GSTIN", how="outer", indicator=True)

    for s in ["Taxable","IGST","CGST","SGST","Cess","TotalTax"]:
        merged[f"{s}_Books"]  = merged[f"{s}_Books"].fillna(0)
        merged[f"{s}_Portal"] = merged[f"{s}_Portal"].fillna(0)
        merged[f"Diff_{s}"]   = merged[f"{s}_Books"] - merged[f"{s}_Portal"]

    merged["Invoice_Count_Books"]  = merged["Invoice_Count_Books"].fillna(0).astype(int)
    merged["Invoice_Count_Portal"] = merged["Invoice_Count_Portal"].fillna(0).astype(int)

    def classify(row):
        if row["_merge"] == "left_only":  return "Only in Books"
        if row["_merge"] == "right_only": return "Only in Portal"
        if abs(row["Diff_TotalTax"]) <= tol and abs(row["Diff_Taxable"]) <= tol:
            return "Matched"
        return "Mismatched"

    merged["Status"] = merged.apply(classify, axis=1)
    merged.drop(columns=["_merge"], inplace=True)
    num_cols = merged.select_dtypes(include=[np.number]).columns
    merged[num_cols] = merged[num_cols].round(2)

    matched     = merged[merged["Status"] == "Matched"].copy()
    mismatched  = merged[merged["Status"] == "Mismatched"].copy()
    only_books  = merged[merged["Status"] == "Only in Books"].copy()
    only_portal = merged[merged["Status"] == "Only in Portal"].copy()
    no_gstin    = books_df[books_df["GSTIN"].isna()].copy()

    return merged, matched, mismatched, only_books, only_portal, no_gstin

# ─────────────────────────────────────────────
# Invoice-Level Reconciliation
# ─────────────────────────────────────────────
def reconcile_invoice_level(books_df, portal_df, tol):
    b = books_df.dropna(subset=["GSTIN"]).copy()
    p = portal_df.dropna(subset=["GSTIN"]).copy()

    b["_key"] = b["GSTIN"] + "|" + b["Invoice_No"]
    p["_key"] = p["GSTIN"] + "|" + p["Invoice_No"]

    b_r = b.rename(columns={"Supplier_Name":"Supplier_Books","Taxable_Value":"Taxable_Books",
                             "IGST":"IGST_Books","CGST":"CGST_Books","SGST":"SGST_Books",
                             "Cess":"Cess_Books","Total_Tax":"TotalTax_Books","Total_Value":"TotalValue_Books"})
    p_r = p.rename(columns={"Supplier_Name":"Supplier_Portal","Taxable_Value":"Taxable_Portal",
                             "IGST":"IGST_Portal","CGST":"CGST_Portal","SGST":"SGST_Portal",
                             "Cess":"Cess_Portal","Total_Tax":"TotalTax_Portal","Total_Value":"TotalValue_Portal"})

    b_cols = ["_key","GSTIN","Invoice_No","Supplier_Books",
              "Taxable_Books","IGST_Books","CGST_Books","SGST_Books","Cess_Books","TotalTax_Books","TotalValue_Books"]
    p_cols = ["_key","Supplier_Portal","Taxable_Portal","IGST_Portal","CGST_Portal",
              "SGST_Portal","Cess_Portal","TotalTax_Portal","TotalValue_Portal"]
    if "ITC_Availability" in p_r.columns:
        p_cols.append("ITC_Availability")

    merged = pd.merge(b_r[b_cols], p_r[p_cols], on="_key", how="outer", indicator=True)

    # Fill GSTIN/Invoice for portal-only rows
    for idx in merged[merged["_merge"] == "right_only"].index:
        key = merged.loc[idx, "_key"]
        if pd.notna(key) and "|" in str(key):
            parts = str(key).split("|", 1)
            merged.loc[idx, "GSTIN"]      = parts[0]
            merged.loc[idx, "Invoice_No"] = parts[1]

    for s in ["Taxable","IGST","CGST","SGST","Cess","TotalTax"]:
        merged[f"{s}_Books"]  = merged[f"{s}_Books"].fillna(0)
        merged[f"{s}_Portal"] = merged[f"{s}_Portal"].fillna(0)
        merged[f"Diff_{s}"]   = merged[f"{s}_Books"] - merged[f"{s}_Portal"]

    def classify(row):
        if row["_merge"] == "left_only":  return "Only in Books"
        if row["_merge"] == "right_only": return "Only in Portal"
        if abs(row["Diff_TotalTax"]) <= tol and abs(row["Diff_Taxable"]) <= tol:
            return "Matched"
        return "Mismatched"

    merged["Status"] = merged.apply(classify, axis=1)
    merged.drop(columns=["_merge","_key"], inplace=True)
    num_cols = merged.select_dtypes(include=[np.number]).columns
    merged[num_cols] = merged[num_cols].round(2)

    matched     = merged[merged["Status"] == "Matched"].copy()
    mismatched  = merged[merged["Status"] == "Mismatched"].copy()
    only_books  = merged[merged["Status"] == "Only in Books"].copy()
    only_portal = merged[merged["Status"] == "Only in Portal"].copy()
    no_gstin    = books_df[books_df["GSTIN"].isna()].copy()

    return merged, matched, mismatched, only_books, only_portal, no_gstin

# ─────────────────────────────────────────────
# Colour-coded Excel export
# ─────────────────────────────────────────────
FILL_GREEN  = PatternFill("solid", fgColor="C6EFCE")   # Matched
FILL_RED    = PatternFill("solid", fgColor="FFC7CE")   # Mismatched
FILL_BLUE   = PatternFill("solid", fgColor="BDD7EE")   # Only in Books
FILL_PURPLE = PatternFill("solid", fgColor="E2CFEF")   # Only in Portal
FILL_YELLOW = PatternFill("solid", fgColor="FFEB9C")   # Fuzzy Matched
FILL_HEADER = PatternFill("solid", fgColor="1F4E79")   # Header row
FONT_HEADER = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
FONT_BOLD   = Font(bold=True, name="Calibri", size=10)
FONT_NORMAL = Font(name="Calibri", size=10)

STATUS_FILL = {
    "Matched":        FILL_GREEN,
    "Mismatched":     FILL_RED,
    "Only in Books":  FILL_BLUE,
    "Only in Portal": FILL_PURPLE,
}

def style_sheet(ws, df, status_col="Status"):
    """Apply colour-coding to a worksheet based on Status column."""
    # Header row
    for cell in ws[1]:
        cell.fill   = FILL_HEADER
        cell.font   = FONT_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Data rows
    status_idx = None
    if status_col in df.columns:
        status_idx = list(df.columns).index(status_col) + 1  # 1-based

    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        status_val = ""
        if status_idx:
            status_val = ws.cell(row=row_idx, column=status_idx).value or ""

        # Determine fill for this row
        fill = None
        for key, f in STATUS_FILL.items():
            if key in str(status_val):
                fill = f
                break
        if fill is None and "Fuzzy" in str(status_val):
            fill = FILL_YELLOW

        for cell in row:
            cell.font      = FONT_NORMAL
            cell.alignment = Alignment(vertical="center")
            if fill:
                cell.fill = fill

    # Auto-width columns
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

    # Freeze header row
    ws.freeze_panes = "A2"

def build_summary_sheet(ws, stats, amt_df):
    """Write a nicely formatted Summary sheet."""
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 20

    # Title
    ws["A1"] = "GST ITC RECONCILIATION REPORT"
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79", name="Calibri")
    ws["A2"] = f"Generated: {datetime.now().strftime('%d %b %Y  %H:%M')}"
    ws["A2"].font = Font(italic=True, size=10, color="595959", name="Calibri")

    # Count stats
    ws["A4"] = "RECORD COUNT SUMMARY"
    ws["A4"].font = FONT_BOLD
    headers = ["Category", "Count"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=5, column=c, value=h)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = Alignment(horizontal="center")

    stat_rows = [
        ("✅ Matched",        stats.get("matched", 0),        FILL_GREEN),
        ("⚠️ Mismatched",     stats.get("mismatched", 0),     FILL_RED),
        ("🔀 Fuzzy Matched",  stats.get("fuzzy", 0),          FILL_YELLOW),
        ("📘 Only in Books",  stats.get("only_books", 0),     FILL_BLUE),
        ("🌐 Only in Portal", stats.get("only_portal", 0),    FILL_PURPLE),
        ("🚫 No GSTIN",       stats.get("no_gstin", 0),       PatternFill("solid", fgColor="EEEEEE")),
    ]
    for i, (label, count, fill) in enumerate(stat_rows, start=6):
        ws.cell(row=i, column=1, value=label).fill = fill
        ws.cell(row=i, column=2, value=count).fill = fill
        ws.cell(row=i, column=1).font = FONT_NORMAL
        ws.cell(row=i, column=2).font = FONT_NORMAL

    # Amount summary
    ws["A13"] = "AMOUNT SUMMARY (₹)"
    ws["A13"].font = FONT_BOLD
    for c, h in enumerate(amt_df.columns, 1):
        cell = ws.cell(row=14, column=c, value=h)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = Alignment(horizontal="center")
    for r, row in enumerate(amt_df.itertuples(index=False), start=15):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = FONT_NORMAL
            if c > 1:
                cell.number_format = "#,##0.00"

    # Colour legend
    ws["A22"] = "COLOUR LEGEND"
    ws["A22"].font = FONT_BOLD
    legend = [
        ("Green",  "Matched — amounts agree within tolerance",       FILL_GREEN),
        ("Red",    "Mismatched — found in both but amounts differ",   FILL_RED),
        ("Yellow", "Fuzzy Matched — invoice numbers nearly match",    FILL_YELLOW),
        ("Blue",   "Only in Books — not found on portal",             FILL_BLUE),
        ("Purple", "Only in Portal — not found in books",             FILL_PURPLE),
    ]
    for i, (colour, desc, fill) in enumerate(legend, start=23):
        ws.cell(row=i, column=1, value=colour).fill = fill
        ws.cell(row=i, column=2, value=desc).fill   = fill
        ws.cell(row=i, column=1).font = FONT_NORMAL
        ws.cell(row=i, column=2).font = FONT_NORMAL

def to_coloured_excel(sheets_dict, stats, amt_df):
    """Build a colour-coded Excel workbook."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # Write all dataframes first
        for name, df in sheets_dict.items():
            if df is None or len(df) == 0:
                df = pd.DataFrame({"Info": ["No records in this category"]})
            df.to_excel(writer, sheet_name=name[:31], index=False)

        wb = writer.book

        # Style each data sheet
        for name, df in sheets_dict.items():
            ws = wb[name[:31]]
            if df is not None and len(df) > 0:
                style_sheet(ws, df)

        # Build summary sheet
        if "Summary" in wb.sheetnames:
            wb.remove(wb["Summary"])
        ws_sum = wb.create_sheet("Summary", 0)
        build_summary_sheet(ws_sum, stats, amt_df)

    buf.seek(0)
    return buf

# ─────────────────────────────────────────────
# Main UI
# ─────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    books_file  = st.file_uploader("📘 Upload **Books** file (.xlsx)", type=["xlsx","xls"])
with col2:
    portal_file = st.file_uploader("🌐 Upload **Portal** file (.xlsx)", type=["xlsx","xls"])

recon_mode = st.radio(
    "🔍 Reconciliation Mode",
    ["GSTIN-Level Summary", "GSTIN + Invoice No (Detailed)"],
    horizontal=True,
    help="Summary mode aggregates per GSTIN. Detailed mode matches each invoice line."
)

run_btn = st.button("🚀 Run Reconciliation", type="primary", use_container_width=True)

# ─────────────────────────────────────────────
if run_btn:
    if not books_file or not portal_file:
        st.error("Please upload both Books and Portal files.")
        st.stop()

    with st.spinner("Reading files..."):
        try:
            raw_books  = pd.read_excel(books_file)
            raw_portal = pd.read_excel(portal_file)
        except Exception as e:
            st.error(f"Error reading files: {e}")
            st.stop()

    missing_b = [c for c in [b_gstin,b_supplier,b_invoice,b_taxable,b_igst,b_cgst,b_sgst,b_cess]
                 if c not in raw_books.columns]
    missing_p = [c for c in [p_gstin,p_supplier,p_invoice,p_taxable,p_igst,p_cgst,p_sgst,p_cess]
                 if c not in raw_portal.columns]
    if missing_b:
        st.error(f"Books file missing columns: {missing_b}. Check sidebar → Books Column Names.")
        st.stop()
    if missing_p:
        st.error(f"Portal file missing columns: {missing_p}. Check sidebar → Portal Column Names.")
        st.stop()

    with st.spinner("Normalising data..."):
        books_df  = load_books(raw_books)
        portal_df = load_portal(raw_portal)

    with st.spinner("Running exact reconciliation..."):
        if recon_mode == "GSTIN-Level Summary":
            full, matched, mismatched, only_books, only_portal, no_gstin = \
                reconcile_gstin_level(books_df, portal_df, tolerance)
        else:
            full, matched, mismatched, only_books, only_portal, no_gstin = \
                reconcile_invoice_level(books_df, portal_df, tolerance)

    with st.spinner("Running fuzzy matching on unmatched invoices..."):
        only_books_rem, only_portal_rem, fuzzy_df = apply_fuzzy_matching(
            only_books.copy(), only_portal.copy(), fuzzy_threshold, amt_tol=tolerance
        )

    # ── Amount summary ──
    bk = books_df.dropna(subset=["GSTIN"])
    amt_rows = []
    for label, bc, pc in [
        ("Taxable Value","Taxable_Value","Taxable_Value"),
        ("IGST","IGST","IGST"), ("CGST","CGST","CGST"), ("SGST","SGST","SGST"),
        ("Cess","Cess","Cess"), ("Total Tax","Total_Tax","Total_Tax"),
    ]:
        b_sum = round(bk[bc].sum(), 2)
        p_sum = round(portal_df[pc].sum(), 2)
        amt_rows.append({"Head": label, "Books (₹)": b_sum, "Portal (₹)": p_sum, "Difference (₹)": round(b_sum - p_sum, 2)})
    amt_df = pd.DataFrame(amt_rows)

    stats = {
        "matched":    len(matched),
        "mismatched": len(mismatched),
        "fuzzy":      len(fuzzy_df),
        "only_books": len(only_books_rem),
        "only_portal":len(only_portal_rem),
        "no_gstin":   len(no_gstin),
    }

    # ── Summary metrics ──
    st.divider()
    st.subheader("📊 Reconciliation Summary")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("✅ Matched",        len(matched))
    c2.metric("⚠️ Mismatched",     len(mismatched))
    c3.metric("🔀 Fuzzy Matched",  len(fuzzy_df))
    c4.metric("📘 Only in Books",  len(only_books_rem))
    c5.metric("🌐 Only in Portal", len(only_portal_rem))
    c6.metric("🚫 No GSTIN",       len(no_gstin))

    if len(fuzzy_df) > 0:
        st.info(f"🔀 **Fuzzy matching** found {len(fuzzy_df)} additional near-matches that exact matching missed. "
                f"Review them in the **Fuzzy Matched** tab.")

    # ── Amount summary ──
    st.subheader("💰 Amount Summary")
    def colour_diff(val):
        if isinstance(val, (int, float)):
            if val > 1:   return "color: red; font-weight: bold"
            if val < -1:  return "color: red; font-weight: bold"
        return ""
    st.dataframe(amt_df.style.applymap(colour_diff, subset=["Difference (₹)"]),
                 use_container_width=True, hide_index=True)

    # ── Detailed tabs ──
    st.divider()
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        f"✅ Matched ({len(matched)})",
        f"⚠️ Mismatched ({len(mismatched)})",
        f"🔀 Fuzzy Matched ({len(fuzzy_df)})",
        f"📘 Only in Books ({len(only_books_rem)})",
        f"🌐 Only in Portal ({len(only_portal_rem)})",
        f"🚫 No GSTIN ({len(no_gstin)})",
    ])

    with tab1:
        st.caption("These entries are found in both Books and Portal with matching amounts.")
        if len(matched) > 0: st.dataframe(matched, use_container_width=True, hide_index=True)
        else: st.success("No exact matches found — try checking fuzzy matched tab!")

    with tab2:
        st.caption("Found in both Books and Portal but the amounts do not match — needs investigation.")
        if len(mismatched) > 0: st.dataframe(mismatched, use_container_width=True, hide_index=True)
        else: st.success("No mismatches — everything balances!")

    with tab3:
        st.caption(
            f"Two types of near-matches are shown here:  "
            f"**Fuzzy Matched** = same GSTIN, invoice numbers are ≥{fuzzy_threshold}% similar.  "
            f"**Probable Match** = different GSTIN but invoice numbers are similar AND tax amounts are close — "
            f"likely a GSTIN mismatch / data entry error. Review all rows manually before accepting."
        )
        if len(fuzzy_df) > 0:
            # Highlight probable matches in a different colour for attention
            same_gstin = fuzzy_df[~fuzzy_df["Match_Type"].str.contains("Cross", na=False)] if "Match_Type" in fuzzy_df.columns else fuzzy_df
            cross_gstin = fuzzy_df[fuzzy_df["Match_Type"].str.contains("Cross", na=False)] if "Match_Type" in fuzzy_df.columns else pd.DataFrame()

            if len(same_gstin) > 0:
                st.markdown("**🔀 Fuzzy Matched — Same GSTIN, similar invoice number:**")
                st.dataframe(same_gstin, use_container_width=True, hide_index=True)
            if len(cross_gstin) > 0:
                st.markdown("**🟡 Probable Match — Different GSTIN, similar invoice + matching amount:**")
                st.warning("⚠️ These have a GSTIN mismatch — could be a typo in books or portal. Verify before using for ITC claim.")
                st.dataframe(cross_gstin, use_container_width=True, hide_index=True)
        else:
            st.info("No fuzzy matches found. Try lowering the sensitivity slider in the sidebar.")

    with tab4:
        st.caption("These entries are in your Books but NOT found on the GST Portal.")
        if len(only_books_rem) > 0: st.dataframe(only_books_rem, use_container_width=True, hide_index=True)
        else: st.info("No unmatched entries in Books.")

    with tab5:
        st.caption("These entries are on the GST Portal but NOT found in your Books.")
        if len(only_portal_rem) > 0: st.dataframe(only_portal_rem, use_container_width=True, hide_index=True)
        else: st.info("No unmatched entries in Portal.")

    with tab6:
        st.caption("Books entries with no GSTIN (e.g. Petty Cash, unregistered vendors).")
        if len(no_gstin) > 0: st.dataframe(no_gstin, use_container_width=True, hide_index=True)
        else: st.info("All books entries have a GSTIN.")

    # ──────────────────────────────────────────
    # SUPPLIER DRILL-DOWN
    # ──────────────────────────────────────────
    st.divider()
    st.subheader("🔎 Supplier Drill-Down")
    st.caption("Pick any GSTIN to see all its invoices from both Books and Portal side by side.")

    all_gstins = sorted(
        set(books_df["GSTIN"].dropna().tolist()) | set(portal_df["GSTIN"].dropna().tolist())
    )

    # Build a label like "27AAJCB2354C1ZZ — SUPPLIER NAME"
    gstin_labels = {}
    for g in all_gstins:
        b_name = books_df[books_df["GSTIN"] == g]["Supplier_Name"].values
        p_name = portal_df[portal_df["GSTIN"] == g]["Supplier_Name"].values
        name = b_name[0] if len(b_name) > 0 else (p_name[0] if len(p_name) > 0 else "")
        gstin_labels[g] = f"{g}  —  {name}"

    selected_gstin = st.selectbox(
        "Select Supplier (GSTIN)",
        options=["— Select a GSTIN —"] + all_gstins,
        format_func=lambda g: gstin_labels.get(g, g) if g != "— Select a GSTIN —" else g
    )

    if selected_gstin != "— Select a GSTIN —":
        b_inv = books_df[books_df["GSTIN"] == selected_gstin].copy()
        p_inv = portal_df[portal_df["GSTIN"] == selected_gstin].copy()

        dc1, dc2 = st.columns(2)

        with dc1:
            st.markdown(f"**📘 Books** — {len(b_inv)} invoice(s)")
            if len(b_inv) > 0:
                show_b = b_inv[["Invoice_No","Taxable_Value","IGST","CGST","SGST","Cess","Total_Tax"]].reset_index(drop=True)
                st.dataframe(show_b, use_container_width=True, hide_index=True)
                st.markdown(f"**Total Tax (Books): ₹{b_inv['Total_Tax'].sum():,.2f}**")
            else:
                st.info("No invoices in Books for this GSTIN.")

        with dc2:
            st.markdown(f"**🌐 Portal** — {len(p_inv)} invoice(s)")
            if len(p_inv) > 0:
                p_show_cols = ["Invoice_No","Taxable_Value","IGST","CGST","SGST","Cess","Total_Tax"]
                if "ITC_Availability" in p_inv.columns:
                    p_show_cols.append("ITC_Availability")
                show_p = p_inv[p_show_cols].reset_index(drop=True)
                st.dataframe(show_p, use_container_width=True, hide_index=True)
                st.markdown(f"**Total Tax (Portal): ₹{p_inv['Total_Tax'].sum():,.2f}**")
            else:
                st.info("No invoices on Portal for this GSTIN.")

        # Difference callout
        b_total = b_inv["Total_Tax"].sum()
        p_total = p_inv["Total_Tax"].sum()
        diff    = round(b_total - p_total, 2)
        if abs(diff) <= tolerance:
            st.success(f"✅ Totals match for this supplier. Difference: ₹{diff:,.2f}")
        else:
            st.error(f"⚠️ Tax difference for this supplier: ₹{diff:,.2f}  "
                     f"(Books ₹{b_total:,.2f} vs Portal ₹{p_total:,.2f})")

    # ──────────────────────────────────────────
    # DOWNLOAD — Colour-coded Excel
    # ──────────────────────────────────────────
    st.divider()
    st.subheader("📥 Download Colour-Coded Report")

    # Add Status column to fuzzy_df for colour logic
    fuzzy_export = fuzzy_df.copy() if len(fuzzy_df) > 0 else pd.DataFrame()

    sheets = {
        "Matched":          matched,
        "Mismatched":       mismatched,
        "Fuzzy Matched":    fuzzy_export,
        "Only in Books":    only_books_rem,
        "Only in Portal":   only_portal_rem,
        "No GSTIN":         no_gstin,
        "Full Recon":       full,
    }

    timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
    mode_tag   = "GSTIN" if recon_mode == "GSTIN-Level Summary" else "Invoice"
    filename   = f"GST_Reconciliation_{mode_tag}_{timestamp}.xlsx"

    excel_bytes = to_coloured_excel(sheets, stats, amt_df)

    st.download_button(
        label="⬇️ Download Full Colour-Coded Report",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

# ── Footer ──
st.divider()
st.caption("GST ITC Reconciliation Tool v2 • Fuzzy Matching • Colour-Coded Export • Supplier Drill-Down • Built for CA Professionals")
