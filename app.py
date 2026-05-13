import streamlit as st
import io
import re
from xml.etree import ElementTree as ET
from parser import TallyParser, EditLogParser
from excel_builder import build_excel

st.set_page_config(
    page_title="Tally Audit Pack Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
h1, h2, h3 { font-family: 'IBM Plex Sans', sans-serif; font-weight: 700; }

.stApp { background: #0f0f0f; color: #e8e8e8; }

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #161616 !important;
    border-right: 1px solid #2d2d2d;
}
section[data-testid="stSidebar"] * { color: #e8e8e8 !important; }

/* Cards */
div[data-testid="metric-container"] {
    background: #1e1e1e;
    border: 1px solid #2d2d2d;
    border-top: 3px solid #f4a100;
    border-radius: 6px;
    padding: 16px;
}

/* Buttons */
.stButton > button {
    background: #f4a100 !important;
    color: #000 !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: 6px !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    width: 100%;
}
.stButton > button:hover { background: #d08c00 !important; }

.stDownloadButton > button {
    background: #22c55e !important;
    color: #000 !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: 6px !important;
    width: 100%;
}
.stDownloadButton > button:hover { background: #16a34a !important; }

/* File uploader */
div[data-testid="stFileUploader"] {
    background: #1e1e1e;
    border: 1.5px dashed #2d2d2d;
    border-radius: 8px;
}
div[data-testid="stFileUploader"]:hover { border-color: #f4a100; }

/* Dataframe */
div[data-testid="stDataFrame"] { border: 1px solid #2d2d2d; border-radius: 6px; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] { gap: 4px; background: #1e1e1e; border-radius: 6px; padding: 4px; }
.stTabs [data-baseweb="tab"] {
    background: transparent;
    border-radius: 4px;
    color: #666;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 12px;
}
.stTabs [aria-selected="true"] { background: #f4a100 !important; color: #000 !important; font-weight: 700 !important; }

/* Expander */
details { background: #1e1e1e; border: 1px solid #2d2d2d; border-radius: 6px; }

/* Info/success boxes */
div[data-testid="stAlert"] { border-radius: 6px; }

/* Header bar */
.header-bar {
    background: #161616;
    border-bottom: 3px solid #f4a100;
    padding: 12px 20px;
    margin: -1rem -1rem 1.5rem -1rem;
    display: flex; align-items: center; gap: 12px;
}
.header-bar h2 { margin: 0; font-size: 1.2rem; }
.badge {
    background: #1e1e1e; border: 1px solid #2d2d2d;
    border-radius: 4px; padding: 2px 8px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px; color: #666;
}
.badge.green { border-color: #22c55e; color: #22c55e; }
.badge.blue  { border-color: #3b82f6; color: #3b82f6; }
.badge.amber { border-color: #f4a100; color: #f4a100; }

/* Log area */
.log-box {
    background: #0d0d0d;
    border: 1px solid #2d2d2d;
    border-radius: 6px;
    padding: 12px 14px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 12px;
    line-height: 1.8;
    max-height: 220px;
    overflow-y: auto;
}
</style>
""", unsafe_allow_html=True)

# ── Header ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-bar">
  <span style="font-size:28px">📊</span>
  <h2 style="color:#f4a100; font-weight:800">Tally<span style="color:#e8e8e8">AuditPack</span> <small style="font-size:14px;color:#666;font-family:'IBM Plex Mono',monospace">v3.0</small></h2>
  <span class="badge blue">✓ XML / ODBC</span>
  <span class="badge amber">✓ .data Backup</span>
  <span class="badge green">✓ Edit Log (MCA)</span>
  <span class="badge" style="border-color:#a855f7;color:#a855f7">✓ ZIP Archive</span>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📂 Step 1 — Upload Tally File")
    main_file = st.file_uploader(
        "Tally Backup / Export",
        type=["xml", "data", "zip", "txt", "xlsx"],
        help="Supports: Tally Prime XML export, .data backup, ZIP archive, tab-separated text export",
        key="main_upload",
    )

    st.markdown("---")
    st.markdown("### 🔍 Step 2 — Edit Log *(Optional)*")
    st.caption("Tally Prime: Gateway → Display More Reports → Audit & Compliance → Edit Log → Ctrl+E → XML")
    edit_log_file = st.file_uploader(
        "Edit Log XML (MCA Audit Trail)",
        type=["xml", "txt"],
        help="Dedicated Edit Log export for MCA compliance sheets",
        key="editlog_upload",
    )

    st.markdown("---")
    st.markdown("### 🏢 Step 3 — Company Info")
    company_name  = st.text_input("Company Name", value=st.session_state.get("company_name", ""), placeholder="Auto-read from file")
    col1, col2    = st.columns(2)
    period_from   = col1.date_input("Period From", value=None, format="DD/MM/YYYY")
    period_to     = col2.date_input("Period To",   value=None, format="DD/MM/YYYY")
    entity_type   = st.selectbox("Entity Type", ["Pvt Ltd", "LLP", "Proprietorship", "Partnership", "Public Ltd", "OPC"])
    currency      = st.text_input("Currency Symbol", value="₹", max_chars=3)

    st.markdown("---")
    st.markdown("### 📋 Step 4 — Select Sheets")

    sheet_groups = {
        "📋 Core Statements": {
            "summary":       ("🏠", "Summary Dashboard"),
            "balance_sheet": ("🏛️", "Balance Sheet"),
            "profit_loss":   ("📈", "P & L Statement"),
            "trial_balance": ("⚖️",  "Trial Balance"),
            "daybook":       ("📋", "Daybook"),
        },
        "💹 Transactions": {
            "sales_reg":     ("🛒", "Sales Register"),
            "sales_summary": ("📊", "Sales Summary"),
            "purchase_reg":  ("📦", "Purchase Register"),
        },
        "👥 Parties": {
            "customer_360":  ("👤", "Customer 360°"),
            "supplier_360":  ("🏭", "Supplier 360°"),
            "debtors":       ("📉", "Debtors Aging"),
        },
        "🧾 Tax & Compliance": {
            "gst_summary":   ("🧾", "GST Summary"),
            "gst_3way":      ("🔄", "GST 3-Way Recon"),
            "tds":           ("💰", "TDS Register"),
        },
        "🔍 Audit Trail (MCA Edit Log)": {
            "editlog_vch":   ("📝", "Edit Log – Vouchers"),
            "editlog_mst":   ("🗂️", "Edit Log – Masters"),
            "editlog_del":   ("🗑️", "Deleted Vouchers"),
            "editlog_stats": ("📊", "Audit Trail Summary"),
        },
        "📈 MIS & Analysis": {
            "monthly_mis":   ("📅", "Monthly MIS"),
            "ledger_wise":   ("📓", "Ledger-wise"),
        },
    }

    selected_sheets = {}
    for group_name, sheets in sheet_groups.items():
        with st.expander(group_name, expanded=True):
            c1, c2 = st.columns(2)
            items = list(sheets.items())
            for i, (sid, (icon, label)) in enumerate(items):
                col = c1 if i % 2 == 0 else c2
                selected_sheets[sid] = col.checkbox(f"{icon} {label}", value=True, key=f"sh_{sid}")

    st.markdown("---")
    generate_btn = st.button("⚡ Generate Audit Pack", use_container_width=True, type="primary")

# ── Main Area ────────────────────────────────────────────────────────────────
if not main_file:
    # Empty state
    st.markdown("""
    <div style="text-align:center;padding:60px 20px;color:#444;">
      <div style="font-size:64px;opacity:0.2;margin-bottom:16px">📊</div>
      <h3 style="color:#555;font-weight:700">Tally Audit Pack Generator</h3>
      <p style="color:#444;max-width:480px;margin:8px auto;font-size:14px">
        Upload a Tally backup file in the sidebar to generate a complete<br>
        Schedule III audit pack with all ledger registers.
      </p>
    </div>
    """, unsafe_allow_html=True)

    with st.expander("📖 How to export from Tally Prime", expanded=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("""
            **📄 XML Export (Recommended)**
            ```
            Gateway of Tally
            → Export
            → Data
            → Format: XML
            → Save file
            ```
            """)
        with col2:
            st.markdown("""
            **💾 .data Backup**
            ```
            Gateway of Tally
            → Company
            → Backup
            → Select company
            → TallyDrive / Local
            ```
            """)
        with col3:
            st.markdown("""
            **🔍 Edit Log (MCA)**
            ```
            Display More Reports
            → Audit & Compliance
            → Edit Log
            → Ctrl+E → XML
            ```
            """)
    st.stop()

# ── Parse file on upload ─────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def parse_main_file(file_bytes: bytes, filename: str):
    return TallyParser.parse(file_bytes, filename)

@st.cache_data(show_spinner=False)
def parse_edit_log(file_bytes: bytes, filename: str):
    return EditLogParser.parse(file_bytes, filename)

with st.spinner("🔍 Parsing Tally file…"):
    try:
        parsed = parse_main_file(main_file.getvalue(), main_file.name)
    except Exception as e:
        st.error(f"❌ Failed to parse file: {e}")
        st.stop()

# Auto-fill company name from parsed data
if parsed.get("company_name") and not st.session_state.get("company_name"):
    st.session_state["company_name"] = parsed["company_name"]
    st.rerun()

company = company_name or parsed.get("company_name", "Company")
from_str = str(period_from) if period_from else parsed.get("from_date", "2025-04-01")
to_str   = str(period_to)   if period_to   else parsed.get("to_date",   "2026-03-31")

# Parse edit log
edit_log_parsed = None
if edit_log_file:
    with st.spinner("🔍 Parsing Edit Log…"):
        try:
            edit_log_parsed = parse_edit_log(edit_log_file.getvalue(), edit_log_file.name)
        except Exception as e:
            st.warning(f"⚠ Edit Log parse error: {e}")
elif parsed.get("edit_log"):
    edit_log_parsed = parsed["edit_log"]

# ── File Info Banner ─────────────────────────────────────────────────────────
v = parsed.get("vouchers", [])
l = parsed.get("ledgers",  {})
el_vch = edit_log_parsed.get("voucher_logs", []) if edit_log_parsed else []
el_mst = edit_log_parsed.get("master_logs",  []) if edit_log_parsed else []

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Vouchers",     f"{len(v):,}")
col2.metric("Ledgers",      f"{len(l):,}")
col3.metric("File Format",  parsed.get("format", "XML").upper())
col4.metric("Edit Log Vch", f"{len(el_vch):,}", delta="MCA" if el_vch else "None", delta_color="normal" if el_vch else "off")
col5.metric("Edit Log Mst", f"{len(el_mst):,}", delta="MCA" if el_mst else "None", delta_color="normal" if el_mst else "off")

st.markdown("---")

# ── Compute trial balance for preview ────────────────────────────────────────
@st.cache_data(show_spinner=False)
def compute_trial_balance(vouchers_tuple, ledgers_str):
    import json
    ledgers = json.loads(ledgers_str)
    vouchers = list(vouchers_tuple)
    tb = {}
    for v in vouchers:
        acc = v["account"]
        if acc not in tb:
            tb[acc] = {"dr": 0.0, "cr": 0.0, "op_dr": 0.0, "op_cr": 0.0, "cls": v.get("cls", "other")}
        tb[acc]["dr"] += v.get("debit", 0)
        tb[acc]["cr"] += v.get("credit", 0)
    for name, ld in ledgers.items():
        if name not in tb:
            tb[name] = {"dr": 0.0, "cr": 0.0, "op_dr": 0.0, "op_cr": 0.0, "cls": "other"}
        op = ld.get("opening_balance", 0)
        tb[name]["op_dr"] = op if op >= 0 else 0
        tb[name]["op_cr"] = abs(op) if op < 0 else 0
    return tb

# ── Generate on button click ─────────────────────────────────────────────────
if generate_btn:
    active_sheets = {k for k, v in selected_sheets.items() if v}
    if not active_sheets:
        st.warning("Please select at least one sheet.")
        st.stop()

    logs = []
    progress_bar = st.progress(0, text="Building audit pack…")

    def log_step(msg, pct):
        logs.append(msg)
        progress_bar.progress(pct, text=msg)

    try:
        log_step("⚙ Computing trial balance…", 10)
        import json
        tb = compute_trial_balance(
            tuple(tuple(sorted(v.items())) for v in parsed["vouchers"]),
            json.dumps(parsed["ledgers"])
        )

        log_step("📊 Building Excel workbook…", 20)
        excel_bytes = build_excel(
            company=company,
            from_date=from_str,
            to_date=to_str,
            entity_type=entity_type,
            currency=currency,
            vouchers=parsed["vouchers"],
            ledgers=parsed["ledgers"],
            groups=parsed.get("groups", {}),
            trial_balance=tb,
            edit_log=edit_log_parsed,
            selected_sheets=active_sheets,
            log_fn=log_step,
        )

        progress_bar.progress(100, text="✅ Done!")
        st.success(f"✅ Audit pack generated with {len(active_sheets)} sheets!")

        fname = f"{company.replace(' ','_')}_AuditPack_{to_str[:4]}.xlsx"
        st.download_button(
            label=f"⬇ Download {fname}",
            data=excel_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        # Show log
        with st.expander("📋 Processing Log", expanded=False):
            st.markdown("<div class='log-box'>" + "<br>".join(f"<span style='color:#22c55e'>{l}</span>" for l in logs) + "</div>", unsafe_allow_html=True)

        st.session_state["last_excel"] = excel_bytes
        st.session_state["last_fname"] = fname
        st.session_state["last_tb"]    = tb

    except Exception as e:
        st.error(f"❌ Generation failed: {e}")
        import traceback; st.code(traceback.format_exc())
        st.stop()

# ── Preview Tabs ─────────────────────────────────────────────────────────────
st.markdown("### 📋 Data Preview")

import pandas as pd

tab_labels = [
    "📋 Daybook", "⚖️ Trial Balance", "🛒 Sales", "📦 Purchases",
    "👤 Customers", "🏭 Suppliers", "🧾 GST", "📝 Edit Log", "📅 MIS"
]
tabs = st.tabs(tab_labels)

vouchers = parsed.get("vouchers", [])
ledgers  = parsed.get("ledgers",  {})

# ── Daybook Tab ──
with tabs[0]:
    if vouchers:
        df = pd.DataFrame(vouchers)[["date","vch_no","vch_type","account","narration","debit","credit","cls"]].copy()
        df.columns = ["Date","Vch No","Vch Type","Account","Narration","Debit","Credit","Class"]
        df["Debit"]  = df["Debit"].apply(lambda x: f"₹{x:,.2f}" if x else "—")
        df["Credit"] = df["Credit"].apply(lambda x: f"₹{x:,.2f}" if x else "—")
        st.dataframe(df, use_container_width=True, height=400)
        st.caption(f"{len(df):,} voucher lines")
    else:
        st.info("No voucher data found.")

# ── Trial Balance Tab ──
with tabs[1]:
    if vouchers:
        import json as _json
        tb_data = compute_trial_balance(
            tuple(tuple(sorted(v.items())) for v in vouchers),
            _json.dumps(ledgers)
        )
        tb_rows = []
        for name, t in tb_data.items():
            cl_raw = t["op_dr"] - t["op_cr"] + t["dr"] - t["cr"]
            tb_rows.append({
                "Account": name,
                "Class": t["cls"],
                "Parent": ledgers.get(name, {}).get("parent", ""),
                "Op Dr": t["op_dr"] or "",
                "Op Cr": t["op_cr"] or "",
                "Trans Dr": t["dr"] or "",
                "Trans Cr": t["cr"] or "",
                "Closing Dr": cl_raw if cl_raw >= 0 else 0,
                "Closing Cr": abs(cl_raw) if cl_raw < 0 else 0,
            })
        df_tb = pd.DataFrame(tb_rows).sort_values("Parent")
        st.dataframe(df_tb, use_container_width=True, height=400)
        st.caption(f"{len(df_tb):,} ledgers")
    else:
        st.info("No ledger data found.")

# ── Sales Tab ──
with tabs[2]:
    sales_vch = [v for v in vouchers if v.get("cls") == "income" and v.get("credit", 0) > 0]
    if sales_vch:
        df_s = pd.DataFrame(sales_vch)[["date","vch_no","vch_type","account","narration","credit"]].copy()
        df_s.columns = ["Date","Vch No","Vch Type","Account","Narration","Amount"]
        total = df_s["Amount"].sum()
        st.dataframe(df_s, use_container_width=True, height=400)
        st.caption(f"{len(df_s):,} sales lines · Total: ₹{total:,.2f}")
    else:
        st.info("No sales vouchers found in this data.")

# ── Purchases Tab ──
with tabs[3]:
    exp_vch = [v for v in vouchers if v.get("cls") in ("purchase","empl_expense","indir_expense") and v.get("debit",0)>0]
    if exp_vch:
        df_e = pd.DataFrame(exp_vch)[["date","vch_no","vch_type","account","narration","debit"]].copy()
        df_e.columns = ["Date","Vch No","Vch Type","Account","Narration","Amount"]
        total = df_e["Amount"].sum()
        st.dataframe(df_e, use_container_width=True, height=400)
        st.caption(f"{len(df_e):,} expense lines · Total: ₹{total:,.2f}")
    else:
        st.info("No purchase/expense vouchers found.")

# ── Customers Tab ──
with tabs[4]:
    debtor_vch = [v for v in vouchers if v.get("cls") == "debtor"]
    if debtor_vch:
        cust_map = {}
        for v in debtor_vch:
            acc = v["account"]
            if acc not in cust_map: cust_map[acc] = {"sales": 0, "receipts": 0}
            cust_map[acc]["sales"]    += v.get("debit", 0)
            cust_map[acc]["receipts"] += v.get("credit", 0)
        rows = []
        for name, mv in cust_map.items():
            op = ledgers.get(name, {}).get("opening_balance", 0)
            cl = op + mv["sales"] - mv["receipts"]
            rows.append({"Customer": name, "Opening": op, "Sales": mv["sales"], "Receipts": mv["receipts"], "Closing": round(cl, 2)})
        df_c = pd.DataFrame(rows).sort_values("Sales", ascending=False)
        st.dataframe(df_c, use_container_width=True, height=400)
        st.caption(f"{len(df_c):,} customers")
    else:
        st.info("No debtor entries found.")

# ── Suppliers Tab ──
with tabs[5]:
    cred_vch = [v for v in vouchers if v.get("cls") == "creditor"]
    if cred_vch:
        sup_map = {}
        for v in cred_vch:
            acc = v["account"]
            if acc not in sup_map: sup_map[acc] = {"purchases": 0, "payments": 0}
            sup_map[acc]["purchases"] += v.get("credit", 0)
            sup_map[acc]["payments"]  += v.get("debit", 0)
        rows = []
        for name, mv in sup_map.items():
            op = ledgers.get(name, {}).get("opening_balance", 0)
            cl = op + mv["purchases"] - mv["payments"]
            rows.append({"Supplier": name, "Opening": op, "Purchases": mv["purchases"], "Payments": mv["payments"], "Closing": round(cl, 2)})
        df_sup = pd.DataFrame(rows).sort_values("Purchases", ascending=False)
        st.dataframe(df_sup, use_container_width=True, height=400)
        st.caption(f"{len(df_sup):,} suppliers")
    else:
        st.info("No creditor entries found.")

# ── GST Tab ──
with tabs[6]:
    gst_vch = [v for v in vouchers if v.get("cls") == "tax_gst"]
    if gst_vch:
        gst_map = {}
        for v in gst_vch:
            acc = v["account"]
            if acc not in gst_map: gst_map[acc] = {"dr": 0, "cr": 0}
            gst_map[acc]["dr"] += v.get("debit", 0)
            gst_map[acc]["cr"] += v.get("credit", 0)
        rows = [{"GST Ledger": name, "Total Dr": mv["dr"], "Total Cr": mv["cr"], "Net": round(mv["cr"]-mv["dr"],2)} for name, mv in gst_map.items()]
        df_gst = pd.DataFrame(rows)
        st.dataframe(df_gst, use_container_width=True, height=400)
    else:
        st.info("No GST ledger entries found.")

# ── Edit Log Tab ──
with tabs[7]:
    if edit_log_parsed:
        vl = edit_log_parsed.get("voucher_logs", [])
        ml = edit_log_parsed.get("master_logs",  [])
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Events", len(vl)+len(ml))
        col2.metric("Deletions",    sum(1 for x in vl if x.get("is_deleted")))
        col3.metric("Cancellations",sum(1 for x in vl if x.get("is_cancelled")))
        col4.metric("Master Changes",len(ml))
        if vl:
            st.markdown("**Voucher Edit Log**")
            df_vl = pd.DataFrame(vl)[["audit_date","audit_time","user","action","vch_type","entity_name","amount","is_deleted","is_cancelled"]].copy()
            df_vl.columns = ["Audit Date","Time","User","Action","Vch Type","Vch No","Amount","Deleted?","Cancelled?"]
            st.dataframe(df_vl, use_container_width=True, height=300)
        if ml:
            st.markdown("**Master Edit Log**")
            df_ml = pd.DataFrame(ml)[["audit_date","audit_time","user","action","master_type","entity_name","parent"]].copy()
            df_ml.columns = ["Audit Date","Time","User","Action","Master Type","Name","Parent"]
            st.dataframe(df_ml, use_container_width=True, height=300)
        if not vl and not ml:
            st.info("Edit log is empty.")
    else:
        st.info("""
        **No Edit Log data loaded.**

        To export Edit Log from Tally Prime:
        1. Gateway → Display More Reports → Audit & Compliance → Edit Log
        2. Select *Voucher Edit Log* or *Master Edit Log*
        3. Press `Ctrl+E` → Format: XML → Save
        4. Upload the XML file using the Edit Log uploader in the sidebar.

        *Note: Edit Log is available from TallyPrime Release 2.1 onwards (MCA Audit Trail compliance).*
        """)

# ── Monthly MIS Tab ──
with tabs[8]:
    if vouchers:
        months = ["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"]
        PL_INCOME  = {"income","other_income"}
        PL_EXPENSE = {"purchase","empl_expense","indir_expense"}

        def fiscal_month(date_str):
            try:
                parts = str(date_str).split("-")
                if len(parts) == 3:
                    m = int(parts[1])
                    return m - 4 if m >= 4 else m + 8
                parts2 = str(date_str).split("/")
                if len(parts2) == 3:
                    d_part, m_part = int(parts2[0]), int(parts2[1])
                    return m_part - 4 if m_part >= 4 else m_part + 8
            except: pass
            return -1

        rows = []
        for i, m in enumerate(months):
            mv = [x for x in vouchers if fiscal_month(x.get("date","")) == i]
            inc  = sum(x.get("credit",0) for x in mv if x.get("cls") in PL_INCOME)
            empl = sum(x.get("debit",0)  for x in mv if x.get("cls") == "empl_expense")
            oth  = sum(x.get("debit",0)  for x in mv if x.get("cls") in ("purchase","indir_expense"))
            net  = inc - empl - oth
            rows.append({"Month":m,"Income":round(inc,2),"Employee Cost":round(empl,2),
                         "Other Expenses":round(oth,2),"Net Profit":round(net,2),
                         "Margin %": f"{net/inc*100:.1f}%" if inc else "—"})
        df_mis = pd.DataFrame(rows)
        st.dataframe(df_mis, use_container_width=True, height=450)

        # Mini bar chart
        df_chart = df_mis.set_index("Month")[["Income","Net Profit"]]
        st.bar_chart(df_chart)
    else:
        st.info("No voucher data found.")

# ── Footer ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<div style='text-align:center;color:#444;font-size:11px;font-family:IBM Plex Mono,monospace'>"
    "TallyAuditPack v3.0 · MCA Audit Trail Compliant · Schedule III · "
    "Supports Tally Prime 2.1+ Edit Log</div>",
    unsafe_allow_html=True
)
