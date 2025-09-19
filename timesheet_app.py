# timesheet_app.py — Cloud-backed Timesheet (Supabase DB) + Original Export Formats
# v28.0 — Restores EXACT exports:
#   (A) Per-job "mm-dd-yyyy - {Job Number} - Daily Time Import.xlsx"
#   (B) Single "mm-dd-yyyy - Daily Time.xlsx" from your template (Daily Time.xlsx)
#
# Dependencies (requirements.txt):
#   streamlit>=1.30
#   pandas>=2.1
#   openpyxl>=3.1
#   xlsxwriter>=3.1
#   supabase>=2.6
#   office365-rest-python-client>=2.5   # only if using SharePoint upload
#
# Secrets (Streamlit Cloud → Settings → Secrets):
#   SUPABASE_URL, SUPABASE_SERVICE_KEY
#   # Optional SharePoint for exports (preferred):
#   TENANT_ID, CLIENT_ID, CLIENT_SECRET, SP_SITE, SP_DRIVE="Documents", SP_EXPORT_FOLDER="Exports"

import os, datetime as dt
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# Supabase DB + optional Storage helper (you already have this file)
from supabase_helpers import (
    add_time_rows,
    fetch_time_entries_for_date,
    delete_by_ids,
    upload_export_bytes,   # used if SharePoint not configured
)

# Original export builder (per-job files)
from exports_legacy import build_per_job_exports

# Prefer SharePoint if secrets present
USE_SHAREPOINT = bool(os.environ.get("SP_SITE")) and bool(os.environ.get("CLIENT_ID"))
if USE_SHAREPOINT:
    try:
        from sharepoint_upload import upload_export_to_sharepoint
    except Exception:
        USE_SHAREPOINT = False

st.set_page_config(page_title="Daily Timesheet", page_icon="🗂️", layout="centered")

# ---- Session init ----
for k, v in {
    "whoami_email": "",
    "entered_app": False,
    "is_admin": True,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

APP_DIR = Path(__file__).parent

# ---- Utils ----
def _clean_headers(df: pd.DataFrame) -> pd.DataFrame:
    try: df.columns = [str(c).strip() for c in df.columns]
    except Exception: pass
    return df

def _first(cols, names):
    s = {str(c) for c in cols}
    for n in names:
        if n in s: return n
    return None

def _pad3(v)->str:
    s = str(v).strip()
    return f"{int(s):03d}" if s.isdigit() else s

# ---- Landing ----
def show_landing():
    jpgs = sorted(APP_DIR.glob("*.jpg"))
    logo_path = str(jpgs[0]) if jpgs else None
    st.markdown("<div style='height:5vh'></div>", unsafe_allow_html=True)
    left, mid, right = st.columns([1,2,1])
    with mid:
        if logo_path: st.image(logo_path, width=300)
        st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
        email = st.text_input("Your work email", st.session_state.get("whoami_email",""), placeholder="name@ptwenergy.com")
        if st.button("Enter"):
            st.session_state["whoami_email"] = (email or "").strip()
            st.session_state["entered_app"] = True
            st.rerun()

if (not st.session_state.get("entered_app")) or (not st.session_state.get("whoami_email","").strip()):
    show_landing(); st.stop()

# ---- Sidebar ----
with st.sidebar:
    st.header("Settings")
    st.text_input("Your work email", key="whoami_email")
    st.caption("Entries → Supabase; Exports → SharePoint (if configured) or Supabase Storage.")

# ---- Lookups from your Excel workbook for lists ----
@st.cache_data(ttl=60)
def load_lookups():
    xlsx_path = os.getenv("STREAMLIT_TIMESHEET_XLSX", str(APP_DIR / "TimeSheet Apps.xlsx"))
    if not os.path.exists(xlsx_path):
        st.warning("Lookup workbook 'TimeSheet Apps.xlsx' not found next to the app.")
        return {
            "employees": pd.DataFrame(columns=["Employee Name","Employee Number","Override Trade Class"]),
            "jobs": pd.DataFrame(columns=["JOB #","AREA #","DESCRIPTION"]),
            "costcodes": pd.DataFrame(columns=["Cost Code","Cost Code Description","Active"]),
            "emp_name_col":"Employee Name",
            "emp_num_col":"Employee Number",
            "emp_trade_col":"Override Trade Class",
            "job_num_col":"JOB #",
            "job_area_col":"AREA #",
            "job_desc_col":"DESCRIPTION",
            "cost_code_col":"Cost Code",
            "paycode_map":{"REG":"211","OT":"212","SUBSISTENCE":"261"},
        }
    employees = pd.read_excel(xlsx_path, sheet_name="Employee List"); _clean_headers(employees)
    jobs      = pd.read_excel(xlsx_path, sheet_name="Job Numbers");   _clean_headers(jobs)
    costcodes = pd.read_excel(xlsx_path, sheet_name="Cost Codes");    _clean_headers(costcodes)
    return {
        "employees": employees, "jobs": jobs, "costcodes": costcodes,
        "emp_name_col": _first(employees.columns, ["Employee Name","Name"]),
        "emp_num_col":  _first(employees.columns, ["Employee Number","Person Number","Emp #"]),
        "emp_trade_col":_first(employees.columns, ["Override Trade Class","Trade Class"]),
        "job_num_col":  _first(jobs.columns, ["JOB #","Job Number","Job #"]),
        "job_area_col": _first(jobs.columns, ["AREA #","Job Area","Area #"]),
        "job_desc_col": _first(jobs.columns, ["DESCRIPTION","Area Description","Description","Area Name"]),
        "cost_code_col":_first(costcodes.columns, ["Cost Code","Class Type"]),
        "paycode_map":{"REG":"211","OT":"212","SUBSISTENCE":"261"},
    }

lk = load_lookups()
employees, jobs, costcodes = lk["employees"], lk["jobs"], lk["costcodes"]
emp_name_col, emp_num_col, emp_trade_col = lk["emp_name_col"], lk["emp_num_col"], lk["emp_trade_col"]
job_num_col, job_area_col, job_desc_col = lk["job_num_col"], lk["job_area_col"], lk["job_desc_col"]
cost_code_col, paycode_map = lk["cost_code_col"], lk["paycode_map"]

# ---- Entry UI ----
st.subheader("Timesheet Entry")
date_val = st.date_input("Date", dt.date.today())

emp_opts = employees[emp_name_col].astype(str).tolist() if emp_name_col else []
sel_emps = st.multiselect("Employees", emp_opts)

job_opts = jobs[job_num_col].astype(str).unique().tolist() if job_num_col else []
sel_job  = st.selectbox("Job Number", [""] + job_opts)

# Areas bound to job
area_labels, area_map = [], {}
if sel_job and job_area_col:
    df = jobs.loc[jobs[job_num_col].astype(str)==str(sel_job)].copy(); _clean_headers(df)
    for _, r in df.iterrows():
        code = _pad3(r.get(job_area_col,""))
        desc = str(r.get(job_desc_col,"") or "").strip()
        lab = f"{code} - {desc}" if desc else code
        if lab not in area_map: area_labels.append(lab); area_map[lab]=code
sel_area_label = st.selectbox("Job Area", [""] + area_labels)
sel_area_code = area_map.get(sel_area_label, "")

# Cost codes → ACTIVE only
def _active_costcodes(df: pd.DataFrame) -> pd.DataFrame:
    df2=df.copy(); _clean_headers(df2)
    active_col = _first(df2.columns, ["Active","Is Active","Enabled","ACTIVE","IS ACTIVE","ENABLED"])
    if active_col:
        def truthy(v):
            if isinstance(v,bool): return v
            s=str(v).strip().lower(); return s in {"true","t","yes","y","1","active","enabled"}
        return df2[df2[active_col].apply(truthy)]
    status_col = _first(df2.columns, ["Status","STATUS"])
    if status_col:
        return df2[df2[status_col].astype(str).str.strip().str.lower()=="active"]
    end_col = _first(df2.columns, ["End Date","Inactive Date","END DATE"])
    if end_col:
        return df2[(df2[end_col].isna()) | (df2[end_col].astype(str).str.strip()=="")]
    return df2

act_codes = _active_costcodes(costcodes)

def _build_code_labels(df, code_col):
    df2=df.copy(); _clean_headers(df2)
    desc_col = _first(df2.columns, ["Cost Code Description","Class Type Description","Description","Cost Code Name","Name"])
    labels, mapping = [], {}
    for _, r in df2.iterrows():
        code = str(r.get(code_col,"") or "").strip()
        if not code: continue
        desc = str(r.get(desc_col,"") or "").strip() if desc_col else ""
        lab = f"{code} - {desc}" if desc else code
        if lab not in mapping: labels.append(lab); mapping[lab]=code
    return labels, mapping

code_labels, code_map = _build_code_labels(act_codes, cost_code_col)
sel_code_label = st.selectbox("Class Type (Cost Code)", [""] + code_labels)
sel_code_code = code_map.get(sel_code_label, "")

rt_hours = st.number_input("RT Hours (per employee)", min_value=0.0, max_value=24.0, step=0.5, value=0.0)
ot_hours = st.number_input("OT Hours (per employee)", min_value=0.0, max_value=24.0, step=0.5, value=0.0)
desc     = st.text_area("Comments (optional)", "", height=100)

# Submit → Supabase insert
if st.button("Submit"):
    if not sel_emps:
        st.warning("Select at least one employee.")
    elif not sel_job or not sel_area_code or not sel_code_code:
        st.warning("Select Job, Area, and Class Type.")
    else:
        rows = []
        for emp_name in sel_emps:
            try:
                emp_row = employees.loc[employees[emp_name_col].astype(str) == str(emp_name)].iloc[0]
            except Exception:
                st.error(f"Employee '{emp_name}' not found."); continue
            rows.append({
                "date": date_val,
                "job_number": str(sel_job),
                "job_area": _pad3(sel_area_code),
                "name": emp_name,
                "class_type": sel_code_code,
                "trade_class": emp_row.get(emp_trade_col,""),
                "employee_number": emp_row.get(emp_num_col,""),
                "rt_hours": float(rt_hours),
                "ot_hours": float(ot_hours),
                "night_shift": False,
                "premium_rate": "",
                "comments": desc,
            })
        n = add_time_rows(rows, created_by=st.session_state.get("whoami_email",""))
        st.success(f"Added {n} row(s)." if n else "No rows were added.")

# ---- What's been added for this day ----
st.markdown("---"); st.subheader("What's been added for this day")
filter_by_job = st.checkbox("Filter by selected Job Number", value=False)
day_rows = fetch_time_entries_for_date(date_str=date_val.strftime("%Y-%m-%d"),
                                       job_number=(sel_job if filter_by_job and sel_job else None))
day_df = pd.DataFrame(day_rows)

if day_df.empty:
    st.caption("empty")
else:
    cols = ["id","job_number","job_area","date","name","class_type","trade_class","employee_number","rt_hours","ot_hours","comments","created_by","created_at"]
    show_cols = [c for c in cols if c in day_df.columns]
    disp = day_df[show_cols].reset_index(drop=True).copy()
    disp.insert(0,"IDX",disp.index)
    st.dataframe(disp.drop(columns=["id"],errors="ignore"), use_container_width=True, hide_index=True)

    opts = ["ALL — Delete all rows shown"] + [f"{i} — {r.get('name','')} ({r.get('job_number','')}/{_pad3(r.get('job_area',''))})" for i, r in disp.iterrows()]
    to_del = st.multiselect("Choose options", opts, placeholder="Choose options")
    if st.button("Delete selected rows"):
        if any(o.startswith("ALL") for o in to_del):
            ids = day_df["id"].tolist()
        else:
            idxs = []
            for o in to_del:
                if o.startswith("ALL"): continue
                try: idxs.append(int(o.split("—")[0].strip()))
                except: pass
            ids = [day_df.iloc[i]["id"] for i in idxs if 0 <= i < len(day_df)]
        if ids:
            d = delete_by_ids(ids); st.success(f"Deleted {d} row(s)."); st.rerun()
        else:
            st.warning("Nothing matched to delete.")

# ---- EXPORTS (original formats) ----
if st.session_state.get("is_admin", True):
    st.markdown("---"); st.subheader("Export Day → TimeEntries + Daily Report")

    with st.form("export_form", clear_on_submit=False):
        export_date = st.date_input("Export Date", dt.date.today())
        do_export = st.form_submit_button("Create Export")

    def _upload(file_bytes: bytes, month_folder: str, filename: str) -> str:
        if USE_SHAREPOINT:
            from sharepoint_upload import upload_export_to_sharepoint
            sp_root = os.environ.get("SP_EXPORT_FOLDER","Exports").strip("/")
            sp_path = f"{sp_root}/{month_folder}/{filename}"
            return upload_export_to_sharepoint(sp_path, file_bytes, link_scope="organization")
        else:
            storage_path = f"{month_folder}/{filename}"
            return upload_export_bytes(file_bytes, storage_path)

    if do_export:
        rows = fetch_time_entries_for_date(export_date.strftime("%Y-%m-%d"), None)
        df_all = pd.DataFrame(rows)
        if df_all.empty:
            st.warning("No matching rows for that date.")
        else:
            month_folder = export_date.strftime("%B")

            # (A) Per-job "Daily Time Import.xlsx" files (exact mapping/order)
            files = build_per_job_exports(df_all, export_date, paycode_map)
            if not files:
                st.info("No per-job files created (no REG/OT hours found).")
            else:
                for fname, data in files.items():
                    link = _upload(data, month_folder, fname)
                    st.success(f"Created: {fname}")
                    st.link_button("Open", link, use_container_width=True)
                    st.download_button(f"Download {fname}", data, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            # (B) Single "mm-dd-yyyy - Daily Time.xlsx" from template
            daily_name = f"{export_date.strftime('%m-%d-%Y')} - Daily Time.xlsx"
            template_path = APP_DIR / "Daily Time.xlsx"
            if not template_path.exists():
                st.warning("Template 'Daily Time.xlsx' not found beside the app.")
            else:
                try:
                    wb = load_workbook(template_path); ws = wb.active
                    try: ws["B1"] = export_date.strftime("%A, %B %d, %Y")
                    except: pass
                    import io
                    out2 = io.BytesIO(); wb.save(out2); out2.seek(0)
                    link2 = _upload(out2.getvalue(), month_folder, daily_name)
                    st.success(f"Created: {daily_name}")
                    st.link_button("Open daily report", link2, use_container_width=True)
                    st.download_button(f"Download {daily_name}", out2.getvalue(), file_name=daily_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                except Exception as e:
                    st.error(f"Failed to build daily report: {e}")

st.caption("Original export formats restored. Rows are stored centrally in Supabase; files saved to SharePoint (if configured) or Supabase Storage.")
