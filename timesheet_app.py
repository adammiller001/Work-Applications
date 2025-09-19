# timesheet_app.py — Cloud-enabled (Supabase DB) + SharePoint/Supabase exports
# v27.0 — Restores exact export filenames:
#   1) "mm-dd-yyyy - {Job Number} - Daily Time Import.xlsx" (one per job with hours)
#   2) "mm-dd-yyyy - Daily Time.xlsx" (single daily report from template)
#
# Requirements in requirements.txt:
#   streamlit>=1.30
#   pandas>=2.1
#   openpyxl>=3.1
#   xlsxwriter>=3.1
#   supabase>=2.6
#   office365-rest-python-client>=2.5   # only needed if using SharePoint upload
#
# Secrets (Streamlit Cloud → Settings → Secrets):
#   SUPABASE_URL = "https://<project>.supabase.co"
#   SUPABASE_SERVICE_KEY = "<service-role-key>"
#   # Optional for SharePoint direct upload (preferred for central folder):
#   TENANT_ID = "..."
#   CLIENT_ID = "..."
#   CLIENT_SECRET = "..."
#   SP_SITE = "https://yourtenant.sharepoint.com/sites/YourSite"
#   SP_DRIVE = "Documents"
#   SP_EXPORT_FOLDER = "Exports"
#
# Place a copy of "Daily Time.xlsx" (your template) next to this file in the repo.

import os, datetime as dt
from datetime import date as _date
from pathlib import Path
from typing import Dict, List

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# --- Supabase helpers (DB + optional Storage) ---
from supabase_helpers import (
    add_time_rows,
    fetch_time_entries_for_date,
    delete_by_ids,
    upload_export_bytes,        # used if SharePoint is not configured
)

# If SharePoint secrets are present, we prefer SharePoint for exports
USE_SHAREPOINT = bool(os.environ.get("SP_SITE")) and bool(os.environ.get("CLIENT_ID"))

if USE_SHAREPOINT:
    try:
        from sharepoint_upload import upload_export_to_sharepoint
    except Exception as _e:
        USE_SHAREPOINT = False  # fallback gracefully if helper is missing

st.set_page_config(page_title="Daily Timesheet", page_icon="🗂️", layout="centered")

# ---------- Initialize session keys (prevents KeyError on first load) ----------
_defaults = {
    "whoami_email": "",
    "entered_app": False,
    "is_admin": True,          # keep admin features on unless you add a roles table later
    "enforce_users": False,
}
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

APP_DIR = Path(__file__).parent

# ---------- Utilities ----------
def _clean_headers(df: pd.DataFrame) -> pd.DataFrame:
    try:
        df.columns = [str(c).strip() for c in df.columns]
    except Exception:
        pass
    return df

def _first(cols, names):
    s = {str(c) for c in cols}
    for n in names:
        if n in s:
            return n
    return None

def _pad_job_area(v) -> str:
    s = str(v).strip()
    return f"{int(s):03d}" if s.isdigit() else s

# ---------- Landing page (logo slightly above center; email + enter) ----------
def show_landing():
    jpgs = sorted(APP_DIR.glob("*.jpg"))
    logo_path = str(jpgs[0]) if jpgs else None

    # Use empty space above to pull content up a bit
    st.markdown("<div style='height:5vh'></div>", unsafe_allow_html=True)
    left, mid, right = st.columns([1, 2, 1])
    with mid:
        if logo_path:
            st.image(logo_path, width=300)
        st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
        email = st.text_input("Your work email", st.session_state.get("whoami_email",""), placeholder="name@ptwenergy.com")
        go = st.button("Enter")
        if go:
            st.session_state["whoami_email"] = (email or "").strip()
            st.session_state["entered_app"] = True
            st.rerun()

if (not st.session_state.get("entered_app", False)) or (not st.session_state.get("whoami_email","").strip()):
    show_landing()
    st.stop()

# ---------- Sidebar ----------
with st.sidebar:
    st.header("Settings")
    st.caption("Cloud-backed: entries → Supabase; exports → SharePoint (if configured) or Supabase Storage.")
    st.text_input("Your work email", key="whoami_email")

# ---------- Lookup loading ----------
@st.cache_data(ttl=60)
def load_local_lookups():
    """
    We keep your Excel lookup tabs for Employees / Jobs / Cost Codes.
    Put 'TimeSheet Apps.xlsx' next to this file in the repo, or set STREAMLIT_TIMESHEET_XLSX.
    """
    xlsx_guess = os.getenv("STREAMLIT_TIMESHEET_XLSX", str(APP_DIR / "TimeSheet Apps.xlsx"))
    if not os.path.exists(xlsx_guess):
        st.warning("Lookup workbook not found. Place 'TimeSheet Apps.xlsx' beside the app or set STREAMLIT_TIMESHEET_XLSX.")
        return {
            "employees": pd.DataFrame(columns=["Employee Name","Employee Number","Override Trade Class"]),
            "jobs": pd.DataFrame(columns=["JOB #","AREA #","DESCRIPTION"]),
            "costcodes": pd.DataFrame(columns=["Cost Code","Cost Code Description","Active"]),
            "emp_name_col": "Employee Name",
            "emp_num_col": "Employee Number",
            "emp_trade_col": "Override Trade Class",
            "job_num_col": "JOB #",
            "job_area_col": "AREA #",
            "job_desc_col": "DESCRIPTION",
            "cost_code_col": "Cost Code",
            "paycode_map": {"REG":"211","OT":"212","SUBSISTENCE":"261"},
        }
    try:
        employees = pd.read_excel(xlsx_guess, sheet_name="Employee List"); _clean_headers(employees)
        jobs      = pd.read_excel(xlsx_guess, sheet_name="Job Numbers");   _clean_headers(jobs)
        costcodes = pd.read_excel(xlsx_guess, sheet_name="Cost Codes");    _clean_headers(costcodes)
    except Exception as e:
        st.error(f"Failed to read lookups from '{xlsx_guess}': {e}")
        employees = pd.DataFrame(columns=["Employee Name","Employee Number","Override Trade Class"])
        jobs      = pd.DataFrame(columns=["JOB #","AREA #","DESCRIPTION"])
        costcodes = pd.DataFrame(columns=["Cost Code","Cost Code Description","Active"])

    return {
        "employees": employees, "jobs": jobs, "costcodes": costcodes,
        "emp_name_col": _first(employees.columns, ["Employee Name","Name"]),
        "emp_num_col":  _first(employees.columns, ["Person Number","Employee Number","Emp #"]),
        "emp_trade_col":_first(employees.columns, ["Override Trade Class","Trade Class"]),
        "job_num_col":  _first(jobs.columns, ["JOB #","Job Number","Job #"]),
        "job_area_col": _first(jobs.columns, ["AREA #","Job Area","Area #"]),
        "job_desc_col": _first(jobs.columns, ["DESCRIPTION","Area Description","Description","Area Name"]),
        "cost_code_col":_first(costcodes.columns, ["Cost Code","Class Type"]),
        "paycode_map": {"REG":"211","OT":"212","SUBSISTENCE":"261"},
    }

look = load_local_lookups()
employees = look["employees"]; jobs = look["jobs"]; costcodes = look["costcodes"]
emp_name_col=look["emp_name_col"]; emp_num_col=look["emp_num_col"]; emp_trade_col=look["emp_trade_col"]
job_num_col=look["job_num_col"]; job_area_col=look["job_area_col"]; job_desc_col=look["job_desc_col"]
cost_code_col=look["cost_code_col"]; paycode_map=look["paycode_map"]

# ---------- Entry UI ----------
st.subheader("Timesheet Entry")
date_val = st.date_input("Date", dt.date.today())

emp_opts = employees[emp_name_col].astype(str).tolist() if emp_name_col else []
sel_emps = st.multiselect("Employees", emp_opts)

job_opts = jobs[job_num_col].astype(str).unique().tolist() if job_num_col else []
sel_job  = st.selectbox("Job Number", [""] + job_opts)

# Job Areas bound to selected Job
area_labels, area_map = [], {}
if sel_job and job_area_col:
    df = jobs.loc[jobs[job_num_col].astype(str)==str(sel_job)].copy(); _clean_headers(df)
    for _, r in df.iterrows():
        code = _pad_job_area(r.get(job_area_col, ""))
        desc = str(r.get(job_desc_col,"") or "").strip()
        lab = f"{code} - {desc}" if desc else code
        if lab not in area_map: area_labels.append(lab); area_map[lab]=code
sel_area_label = st.selectbox("Job Area", [""] + area_labels)
sel_area_code = area_map.get(sel_area_label, "")

# Cost codes -> Active only
def _only_active_costcodes(df: pd.DataFrame) -> pd.DataFrame:
    df2 = df.copy(); _clean_headers(df2)
    active_col = _first(df2.columns, ["Active","Is Active","Enabled","ACTIVE","IS ACTIVE","ENABLED"])
    if active_col:
        def _truthy(x):
            if isinstance(x, bool): return x
            s = str(x).strip().lower()
            return s in {"true","t","yes","y","1","active","enabled"}
        return df2[df2[active_col].apply(_truthy)]
    status_col = _first(df2.columns, ["Status","STATUS"])
    if status_col:
        return df2[df2[status_col].astype(str).str.strip().str.lower() == "active"]
    end_col = _first(df2.columns, ["End Date","Inactive Date","Date End","END DATE"])
    if end_col:
        return df2[(df2[end_col].isna()) | (df2[end_col].astype(str).str.strip() == "")]
    return df2

active_costcodes = _only_active_costcodes(costcodes)

def build_cost_labels(df, code_col):
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

code_labels, code_map = build_cost_labels(active_costcodes, cost_code_col)
sel_code_label = st.selectbox("Class Type (Cost Code)", [""] + code_labels)
sel_code_code = code_map.get(sel_code_label, "")

rt_hours = st.number_input("RT Hours (per employee)", min_value=0.0, max_value=24.0, step=0.5, value=0.0)
ot_hours = st.number_input("OT Hours (per employee)", min_value=0.0, max_value=24.0, step=0.5, value=0.0)
desc     = st.text_area("Comments (optional)", "", height=100)

# --- Submit -> write to Supabase ---
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
                "job_area": _pad_job_area(sel_area_code),
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
        inserted = add_time_rows(rows, created_by=st.session_state.get("whoami_email",""))
        if inserted:
            st.success(f"Added {inserted} row(s).")
        else:
            st.warning("No rows were added.")

# ---------- What's been added for this day (from Supabase) ----------
st.markdown("---")
st.subheader("What's been added for this day")

filter_by_job = st.checkbox("Filter by selected Job Number", value=False)
day_rows = fetch_time_entries_for_date(
    date_str=date_val.strftime("%Y-%m-%d"),
    job_number=(sel_job if filter_by_job and sel_job else None),
)
day_df = pd.DataFrame(day_rows)

if day_df.empty:
    st.caption("empty")
else:
    preferred = ["id","job_number","job_area","date","name","class_type","trade_class",
                 "employee_number","rt_hours","ot_hours","comments","created_by","created_at"]
    show_cols = [c for c in preferred if c in day_df.columns]
    display_df = day_df[show_cols].reset_index(drop=True).copy()
    display_df.insert(0, "IDX", display_df.index)
    st.dataframe(display_df.drop(columns=["id"], errors="ignore"), use_container_width=True, hide_index=True)

    # Delete UI
    base_options = [f"{i} — {r.get('name','')} ({r.get('job_number','')}/{_pad_job_area(r.get('job_area',''))})"
                    for i, r in display_df.iterrows()]
    options = ["ALL — Delete all rows shown"] + base_options
    to_del = st.multiselect("Choose options", options, placeholder="Choose options")

    if st.button("Delete selected rows"):
        if any(opt.startswith("ALL") for opt in to_del):
            ids = day_df["id"].tolist()
        else:
            idxs = []
            for opt in to_del:
                if opt.startswith("ALL"): 
                    continue
                try:
                    idx = int(opt.split("—")[0].strip())
                    idxs.append(idx)
                except Exception:
                    pass
            ids = [day_df.iloc[i]["id"] for i in idxs if 0 <= i < len(day_df)]
        if not ids:
            st.warning("Nothing matched to delete.")
        else:
            deleted = delete_by_ids(ids)
            st.success(f"Deleted {deleted} row(s).")
            st.rerun()

# -------------------- EXPORTS (exact filenames restored) --------------------
if st.session_state.get("is_admin", True):
    st.markdown("---")
    st.subheader("Export Day → TimeEntries + Daily Report")

    with st.form("export_form", clear_on_submit=False):
        export_date = st.date_input("Export Date", dt.date.today())
        export_job_filter = st.selectbox("Filter per-job exports to a single Job (optional)", ["ALL"] + job_opts, index=0)
        do_export = st.form_submit_button("Create Export")

    def _upload_bytes(file_bytes: bytes, month_folder: str, file_name: str) -> str:
        """Uploads to SharePoint (preferred) or Supabase Storage; returns a link."""
        if USE_SHAREPOINT:
            sp_root = os.environ.get("SP_EXPORT_FOLDER", "Exports").strip("/")
            sp_path = f"{sp_root}/{month_folder}/{file_name}"
            return upload_export_to_sharepoint(sp_path, file_bytes, link_scope="organization")
        else:
            storage_path = f"{month_folder}/{file_name}"
            return upload_export_bytes(file_bytes, storage_path)

    if do_export:
        # Pull all rows for that date
        rows_all = fetch_time_entries_for_date(export_date.strftime("%Y-%m-%d"), None)
        df_all = pd.DataFrame(rows_all)

        if df_all.empty:
            st.warning("No matching rows for that date.")
        else:
            month_folder = export_date.strftime("%B")

            # (A) PER-JOB "mm-dd-yyyy - Job Number - Daily Time Import.xlsx"
            jobs_for_day = sorted(df_all["job_number"].astype(str).unique().tolist())
            if export_job_filter != "ALL":
                jobs_for_day = [export_job_filter] if export_job_filter in jobs_for_day else []

            total_files = 0
            for job in jobs_for_day:
                sub = df_all[df_all["job_number"].astype(str) == str(job)].copy()

                # Build TimeEntries rows (REG/OT lines)
                exp_rows = []
                for _, r in sub.iterrows():
                    reg_h = float(r.get("rt_hours", 0) or 0.0)
                    ot_h  = float(r.get("ot_hours", 0) or 0.0)

                    base = {
                        "Date": pd.to_datetime(r.get("date")).strftime("%Y-%m-%d"),
                        "Time Record Type": "",
                        "Person Number": r.get("employee_number", ""),
                        "Employee Name": r.get("name", ""),
                        "Override Trade Class": r.get("trade_class", ""),
                        "Post To Payroll": "Y",
                        "Cost Code / Phase": r.get("class_type", ""),
                        "JobArea": _pad_job_area(r.get("job_area", "")),
                        "Scope Change": "",
                        "Pay Code": "",
                        "Hours": 0.0,
                        "Night Shift": "",
                        "Premium Rate / Subsistence Rate / Travel Rate": r.get("premium_rate", ""),
                        "Comments": "",
                    }
                    if reg_h > 0:
                        t = base.copy(); t["Pay Code"] = paycode_map.get("REG", "211"); t["Hours"] = reg_h; exp_rows.append(t)
                    if ot_h > 0:
                        t = base.copy(); t["Pay Code"] = paycode_map.get("OT", "212");  t["Hours"] = ot_h; exp_rows.append(t)

                if not exp_rows:
                    continue

                out_df = pd.DataFrame(exp_rows, columns=[
                    'Date','Time Record Type','Person Number','Employee Name','Override Trade Class','Post To Payroll',
                    'Cost Code / Phase','JobArea','Scope Change','Pay Code','Hours','Night Shift',
                    'Premium Rate / Subsistence Rate / Travel Rate','Comments'
                ])

                # Exact filename
                file_name = f"{export_date.strftime('%m-%d-%Y')} - {job} - Daily Time Import.xlsx"

                # Save -> upload
                import io
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                    out_df.to_excel(writer, sheet_name="TimeEntries", index=False)
                buf.seek(0)

                link = _upload_bytes(buf.getvalue(), month_folder, file_name)
                total_files += 1
                st.success(f"Created: {file_name}")
                st.link_button("Open file", link, use_container_width=True)
                st.download_button(f"Download {file_name}", buf.getvalue(), file_name=file_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            if total_files == 0:
                st.info("No per-job files were created (no REG/OT hours found).")

            # (B) SINGLE "mm-dd-yyyy - Daily Time.xlsx" from template
            daily_file_name = f"{export_date.strftime('%m-%d-%Y')} - Daily Time.xlsx"
            template_path = APP_DIR / "Daily Time.xlsx"
            if not template_path.exists():
                st.warning("Template 'Daily Time.xlsx' not found beside the app. Add it to the repo for the daily report.")
            else:
                try:
                    wb = load_workbook(template_path)
                    ws = wb.active
                    # Minimal header fill (preserves template layout)
                    try:
                        ws["B1"] = export_date.strftime("%A, %B %d, %Y")
                    except Exception:
                        pass

                    out2 = io.BytesIO()
                    wb.save(out2); out2.seek(0)

                    link2 = _upload_bytes(out2.getvalue(), month_folder, daily_file_name)
                    st.success(f"Created: {daily_file_name}")
                    st.link_button("Open daily report", link2, use_container_width=True)
                    st.download_button(f"Download {daily_file_name}", out2.getvalue(), file_name=daily_file_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                except Exception as e:
                    st.error(f"Failed to build daily report from template: {e}")

st.caption("Entries are stored in Supabase (central DB). Exports are saved to SharePoint if configured; otherwise to Supabase Storage bucket 'exports'.")
