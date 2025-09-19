# timesheet_app.py — Excel workflow with templated formatting + Daily report population + SharePoint upload
# v25.2 (Excel-first; no Supabase)
#
# What’s included
# - Per‑job export: "mm-dd-yyyy - {Job Number} - Daily Time Import.xlsx"
#   EXACT formatting copied from the **TimeEntries** sheet in your "TimeSheet Apps.xlsx":
#   header colors, fonts, column widths, row heights. We use that sheet as a template.
# - Daily report:  "mm-dd-yyyy – Daily Time.xlsx" (en dash) populated like before
#   (date header; simple Indirect/Direct allocation placeholder; work descriptions collated by Job Number).
# - SharePoint upload via Graph (if SP_* secrets exist) AND per‑file download buttons in the UI.
#
# Place in the same folder:
#   - TimeSheet Apps.xlsx      (used for lookups AND as export formatting template)
#   - Daily Time.xlsx          (your daily report template file)
#   - sharepoint_upload.py     (helper for SharePoint upload; optional but recommended)
#
# Optional:
#   - timesheet_default_path.txt : full path to "TimeSheet Apps.xlsx" if it’s somewhere else.
#
# Requirements:
#   streamlit, pandas, openpyxl, xlsxwriter
#   office365-rest-python-client  (only if you use SharePoint upload)

import os, io, json, datetime as dt
from datetime import datetime, date
from pathlib import Path
from typing import Dict, List

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font

# ---- Optional SharePoint helper ----
USE_SHAREPOINT = bool(os.environ.get("SP_SITE")) and bool(os.environ.get("CLIENT_ID"))
if USE_SHAREPOINT:
    try:
        from sharepoint_upload import upload_export_to_sharepoint
    except Exception:
        USE_SHAREPOINT = False

st.set_page_config(page_title="Daily Timesheet", page_icon="🗂️", layout="centered")

APP_DIR = Path(__file__).parent

# ---------- Initialize session keys (avoid KeyError) ----------
for k, v in {
    "whoami_email": "",
    "entered_app": False,
    "is_admin": True,      # keep admin features on unless you wire Users tab gating
    "xlsx_path": os.getenv("STREAMLIT_TIMESHEET_XLSX", "TimeSheet Apps.xlsx"),
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ---------- Landing page ----------
def show_landing():
    jpgs = sorted(APP_DIR.glob("*.jpg"))
    logo_path = str(jpgs[0]) if jpgs else None

    st.markdown("<div style='height:5vh'></div>", unsafe_allow_html=True)
    left, mid, right = st.columns([1,2,1])
    with mid:
        if logo_path: st.image(logo_path, width=300)
        st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
        email = st.text_input("Your work email", st.session_state.get("whoami_email",""), placeholder="name@ptwenergy.com")
        if st.button("Enter"):
            st.session_state["whoami_email"] = (email or "").strip()
            st.session_state["entered_app"] = True
            # detect workbook path
            default_xlsx = os.getenv("STREAMLIT_TIMESHEET_XLSX", "TimeSheet Apps.xlsx")
            sidecar = APP_DIR / "timesheet_default_path.txt"
            if sidecar.exists():
                try:
                    p = sidecar.read_text().strip()
                    if p: default_xlsx = p
                except Exception: pass
            st.session_state["xlsx_path"] = default_xlsx
            st.rerun()

if (not st.session_state.get("entered_app", False)) or (not st.session_state.get("whoami_email","").strip()):
    show_landing()
    st.stop()

xlsx_path = st.session_state.get("xlsx_path", "TimeSheet Apps.xlsx")

# ---------- Sidebar ----------
with st.sidebar:
    st.header("Settings")
    st.caption("Excel‑first. Exports saved locally and uploaded to SharePoint if SP_* secrets are present.")
    st.code(xlsx_path)
    st.text_input("Your work email", key="whoami_email")

# ---------- Helpers ----------
def _clean_headers(df: pd.DataFrame) -> pd.DataFrame:
    try: df.columns = [str(c).strip() for c in df.columns]
    except Exception: pass
    return df

def _first(cols, names):
    s = {str(c) for c in cols}
    for n in names:
        if n in s: return n
    return None

def _pad_job_area(v) -> str:
    s = str(v).strip()
    return f"{int(s):03d}" if s.isdigit() else s

def _read_sheet(path: str, sheet: str, empty_cols: List[str]) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=sheet); _clean_headers(df); return df
    except Exception:
        return pd.DataFrame(columns=empty_cols)

# Lookups + template sheet (TimeEntries) come from TimeSheet Apps.xlsx
if not os.path.exists(xlsx_path):
    st.error("Workbook not found. Place 'TimeSheet Apps.xlsx' next to the app OR create 'timesheet_default_path.txt' with its full path.")
    st.stop()

employees = _read_sheet(xlsx_path, "Employee List", ["Employee Name","Employee Number","Override Trade Class"])
jobs      = _read_sheet(xlsx_path, "Job Numbers",    ["JOB #","AREA #","DESCRIPTION"])
costcodes = _read_sheet(xlsx_path, "Cost Codes",     ["Cost Code","Cost Code Description","Active"])
timedata  = _read_sheet(xlsx_path, "Time Data",      ["Job Number","Job Area","Date","Name","Class Type","Trade Class","Employee Number","RT Hours","OT Hours","Night Shift","Premium Rate / Subsistence Rate / Travel Rate","Comments"])

emp_name_col=_first(employees.columns, ["Employee Name","Name"])
emp_num_col =_first(employees.columns, ["Person Number","Employee Number","Emp #"])
emp_trade_col=_first(employees.columns, ["Override Trade Class","Trade Class"])
job_num_col=_first(jobs.columns, ["JOB #","Job Number","Job #"])
job_area_col=_first(jobs.columns, ["AREA #","Job Area","Area #"])
job_desc_col=_first(jobs.columns, ["DESCRIPTION","Area Description","Description","Area Name"])
cost_code_col=_first(costcodes.columns, ["Cost Code","Class Type"])
paycode_map = {"REG":"211","OT":"212","SUBSISTENCE":"261"}

# ---------- Entry UI ----------
st.subheader("Timesheet Entry")
date_val = st.date_input("Date", dt.date.today())

emp_opts = employees[emp_name_col].astype(str).tolist() if emp_name_col else []
sel_emps = st.multiselect("Employees", emp_opts)

job_opts = jobs[job_num_col].astype(str).unique().tolist() if job_num_col else []
sel_job  = st.selectbox("Job Number", [""] + job_opts)

# Areas tied to chosen job
area_labels, area_map = [], {}
if sel_job and job_area_col:
    df = jobs.loc[jobs[job_num_col].astype(str)==str(sel_job)].copy(); _clean_headers(df)
    for _, r in df.iterrows():
        code = _pad_job_area(r.get(job_area_col,""))
        desc = str(r.get(job_desc_col,"") or "").strip()
        lab = f"{code} - {desc}" if desc else code
        if lab not in area_map: area_labels.append(lab); area_map[lab]=code
sel_area_label = st.selectbox("Job Area", [""] + area_labels)
sel_area_code = area_map.get(sel_area_label, "")

# Active cost codes only
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

def _ensure_time_data_headers(xlsx_file: str):
    wb = load_workbook(xlsx_file)
    if "Time Data" not in wb.sheetnames:
        ws = wb.create_sheet("Time Data")
        ws.append(["Job Number","Job Area","Date","Name","Class Type","Trade Class","Employee Number",
                   "RT Hours","OT Hours","Night Shift","Premium Rate / Subsistence Rate / Travel Rate","Comments"])
        wb.save(xlsx_file); return
    # ensure common headers exist (non-destructive)
    ws = wb["Time Data"]
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    headers = [str(h).strip() if h is not None else "" for h in headers]
    needed = ["RT Hours","OT Hours","Night Shift","Premium Rate / Subsistence Rate / Travel Rate","Comments"]
    changed=False
    for h in needed:
        if h not in headers:
            ws.cell(row=1, column=len(headers)+1, value=h); headers.append(h); changed=True
    if changed: wb.save(xlsx_file)

def _append_row_to_time_data(xlsx_file: str, payload: dict):
    wb = load_workbook(xlsx_file)
    if "Time Data" not in wb.sheetnames:
        ws = wb.create_sheet("Time Data")
        ws.append(["Job Number","Job Area","Date","Name","Class Type","Trade Class","Employee Number",
                   "RT Hours","OT Hours","Night Shift","Premium Rate / Subsistence Rate / Travel Rate","Comments"])
    ws = wb["Time Data"]
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    headers = [str(h).strip() if h is not None else "" for h in headers]
    if not any(headers):
        headers = ["Job Number","Job Area","Date","Name","Class Type","Trade Class","Employee Number",
                   "RT Hours","OT Hours","Night Shift","Premium Rate / Subsistence Rate / Travel Rate","Comments"]
        for idx, h in enumerate(headers, start=1): ws.cell(row=1, column=idx, value=h)
    row_vals = [payload.get(h, "") for h in headers]
    ws.append(row_vals); wb.save(xlsx_file)

if st.button("Submit"):
    if not sel_emps:
        st.warning("Select at least one employee.")
    elif not sel_job or not sel_area_code or not sel_code_code:
        st.warning("Select Job, Area, and Class Type.")
    else:
        try: _ensure_time_data_headers(xlsx_path)
        except Exception: pass
        successes = 0
        for emp_name in sel_emps:
            try:
                emp_row = employees.loc[employees[emp_name_col].astype(str) == str(emp_name)].iloc[0]
            except Exception:
                st.error(f"Employee '{emp_name}' not found."); continue
            payload = {
                "Job Number": str(sel_job),
                "Job Area": _pad_job_area(sel_area_code),
                "Date": date_val.strftime("%Y-%m-%d"),
                "Name": emp_name,
                "Class Type": sel_code_code,
                "Trade Class": emp_row.get(emp_trade_col,""),
                "Employee Number": emp_row.get(emp_num_col,""),
                "RT Hours": float(rt_hours),
                "OT Hours": float(ot_hours),
                "Night Shift": "",
                "Premium Rate / Subsistence Rate / Travel Rate": "",
                "Comments": desc,
            }
            try:
                _append_row_to_time_data(xlsx_path, payload); successes += 1
            except Exception as e:
                st.error(f"Failed to append row for {emp_name}: {e}")
        if successes: st.success(f"Added {successes} row(s) to 'Time Data'.")

# ---------- What's been added for this day ----------
st.markdown("---")
st.subheader("What's been added for this day")

filter_by_job = st.checkbox("Filter by selected Job Number", value=False)

td_all = _read_sheet(xlsx_path, "Time Data", [])
if not td_all.empty:
    td_all["__DateStr"] = td_all["Date"].astype(str).str[:10]
    mask = td_all["__DateStr"] == date_val.strftime("%Y-%m-%d")
    if filter_by_job and sel_job:
        mask = mask & (td_all["Job Number"].astype(str).str.strip() == str(sel_job))
    day_df = td_all[mask].copy()
else:
    day_df = pd.DataFrame(columns=["Job Number","Job Area","Date","Name","Class Type","Trade Class","Employee Number",
                                   "RT Hours","OT Hours","Night Shift","Premium Rate / Subsistence Rate / Travel Rate","Comments"])

if day_df.empty:
    st.caption("empty")
else:
    show_cols = ["Job Number","Job Area","Date","Name","Class Type","Trade Class","Employee Number","RT Hours","OT Hours","Comments"]
    show_cols = [c for c in show_cols if c in day_df.columns]
    display_df = day_df.reset_index(drop=True).copy()
    display_df.insert(0, "IDX", display_df.index)
    st.dataframe(display_df[["IDX"] + show_cols], use_container_width=True, hide_index=True)

# -------------------- EXPORTS --------------------
st.markdown("---")
st.subheader("Export Day → TimeEntries + Daily Report")

with st.form("export_form", clear_on_submit=False):
    export_date = st.date_input("Export Date", dt.date.today())
    do_export = st.form_submit_button("Create Export")

# ---- Formatting helpers: copy styles from template "TimeEntries" ----
def clone_row_styles(src_ws: Worksheet, dst_ws: Worksheet, src_row: int, dst_row: int, max_col: int):
    # copy row height
    if src_row in src_ws.row_dimensions:
        dst_ws.row_dimensions[dst_row].height = src_ws.row_dimensions[src_row].height
    for col in range(1, max_col+1):
        c = get_column_letter(col)
        src_cell = src_ws[f"{c}{src_row}"]
        dst_cell = dst_ws[f"{c}{dst_row}"]
        if src_cell.has_style:
            dst_cell._style = src_cell._style
        # copy number_format explicitly (sometimes style copy misses it)
        dst_cell.number_format = src_cell.number_format

def copy_col_widths(src_ws: Worksheet, dst_ws: Worksheet):
    for col_dim in src_ws.column_dimensions.values():
        key = col_dim.index
        dst_ws.column_dimensions[key].width = col_dim.width

def build_timeentries_df(sub: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, r in sub.iterrows():
        reg_h = float(r.get("RT Hours",0) or 0.0)
        ot_h  = float(r.get("OT Hours",0) or 0.0)
        base = {
            "Date": pd.to_datetime(r.get("Date","")).strftime("%Y-%m-%d"),
            "Time Record Type": "",
            "Person Number": r.get("Employee Number",""),
            "Employee Name": r.get("Name",""),
            "Override Trade Class": r.get("Trade Class",""),
            "Post To Payroll": "Y",
            "Cost Code / Phase": r.get("Class Type",""),
            "JobArea": _pad_job_area(r.get("Job Area","")),
            "Scope Change": "",
            "Pay Code": "",
            "Hours": 0.0,
            "Night Shift": "",
            "Premium Rate / Subsistence Rate / Travel Rate": r.get("Premium Rate / Subsistence Rate / Travel Rate",""),
            "Comments": "",
        }
        if reg_h>0:
            t=base.copy(); t["Pay Code"]=paycode_map.get("REG","211"); t["Hours"]=reg_h; rows.append(t)
        if ot_h>0:
            t=base.copy(); t["Pay Code"]=paycode_map.get("OT","212");  t["Hours"]=ot_h; rows.append(t)
    HEADERS = ['Date','Time Record Type','Person Number','Employee Name','Override Trade Class','Post To Payroll','Cost Code / Phase','JobArea','Scope Change','Pay Code','Hours','Night Shift','Premium Rate / Subsistence Rate / Travel Rate','Comments']
    return pd.DataFrame(rows, columns=HEADERS)

def export_per_job_with_template(time_data_df: pd.DataFrame, job: str, export_date: date, template_book_path: str) -> io.BytesIO:
    # Build data
    subset = time_data_df[time_data_df["Job Number"].astype(str).str.strip() == str(job)].copy()
    out_df = build_timeentries_df(subset)
    # Load template book and sheet
    wb = load_workbook(template_book_path)
    if "TimeEntries" not in wb.sheetnames:
        raise RuntimeError("Template workbook is missing a 'TimeEntries' sheet.")
    ws = wb["TimeEntries"]

    # Find header row (assume row 1) and max columns by header
    headers = [str(c.value).strip() if c.value is not None else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]
    max_col = len(headers)

    # Clear any old data starting row 2 while keeping row heights/styles
    # We'll just overwrite values and blank remaining lines if any exist.
    # Paste values row-by-row, cloning styles from the first data row (row 2).
    data_start = 2
    has_template_data_row = ws.max_row >= 2
    for ridx, row in enumerate(out_df.itertuples(index=False), start=data_start):
        # Ensure the row has styles by cloning from row 2
        if has_template_data_row and ridx != 2:
            clone_row_styles(ws, ws, 2, ridx, max_col)
        for c_idx, val in enumerate(row, start=1):
            ws.cell(row=ridx, column=c_idx, value=val)

    # Blank out any leftover rows below our pasted region (optional)
    last_written = data_start + len(out_df) - 1
    if has_template_data_row and ws.max_row > last_written:
        for r in range(last_written+1, ws.max_row+1):
            for c in range(1, max_col+1):
                ws.cell(row=r, column=c, value=None)

    # Ensure column widths copied to self (no-op here) — template already has them.
    # Save to bytes
    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf

def export_daily_report_populated(xlsx: str, template_path: str, date_str: str, user: str="") -> io.BytesIO:
    # Read day subset
    td = _read_sheet(xlsx, "Time Data", [])
    if td.empty or "Date" not in td.columns:
        return None
    date_col = td["Date"].astype(str).str[:10]
    day = td[date_col == date_str].copy()
    if day.empty:
        return None

    # Load template
    wb = load_workbook(template_path)
    ws = wb.active

    # Header
    try:
        ws["B1"] = pd.to_datetime(date_str).strftime("%A, %B %d, %Y")
        ws["B2"] = "2224138065"  # Job Number (from your earlier spec)
        ws["B3"] = "Pembina"     # Client
    except Exception:
        pass

    # Group entries for comments per Job Number
    descs = {}
    if "Comments" in day.columns and "Job Number" in day.columns:
        for job_num in sorted(day["Job Number"].astype(str).str.strip().unique().tolist()):
            job_comments = day[day["Job Number"].astype(str).str.strip() == job_num]["Comments"].dropna()
            if not job_comments.empty:
                texts = job_comments.astype(str).str.strip().replace("nan","").tolist()
                unique_comments, seen = [], set()
                for t in texts:
                    if t and t not in seen:
                        unique_comments.append(t); seen.add(t)
                if unique_comments:
                    descs[str(job_num)] = unique_comments

    # Clear old descriptions (A264:B400) and repopulate
    bold_underline_font = Font(bold=True, underline="single")
    for r in range(264, 401):
        ws.cell(row=r, column=1, value=None)
        ws.cell(row=r, column=2, value=None)

    row_ptr = 264
    for job_num, comments in descs.items():
        ws.cell(row=row_ptr, column=1, value="Work Description").font = bold_underline_font
        ws.cell(row=row_ptr, column=2, value=job_num).font = bold_underline_font
        row_ptr += 1
        for comment in comments:
            ws.cell(row=row_ptr, column=2, value=comment)
            row_ptr += 1
        row_ptr += 1  # blank spacer

    out = io.BytesIO()
    wb.save(out); out.seek(0)
    return out

def offer_download_and_sharepoint(file_name: str, file_bytes: bytes, month_folder: str):
    # Download button
    st.download_button(f"Download {file_name}", file_bytes, file_name=file_name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
    # SharePoint upload
    if USE_SHAREPOINT:
        try:
            sp_root = os.environ.get("SP_EXPORT_FOLDER","Exports").strip("/")
            sp_path = f"{sp_root}/{month_folder}/{file_name}"
            link = upload_export_to_sharepoint(sp_path, file_bytes, link_scope="organization")
            st.link_button("Open in SharePoint", link, use_container_width=True)
        except Exception as e:
            st.error(f"SharePoint upload failed for {file_name}: {e}")

if do_export:
    # Pull rows for day
    td = _read_sheet(xlsx_path, "Time Data", [])
    if td.empty or "Date" not in td.columns:
        st.warning("No matching rows for that date."); 
    else:
        dmask = td["Date"].astype(str).str[:10] == export_date.strftime("%Y-%m-%d")
        day_df = td[dmask].copy()
        if day_df.empty:
            st.warning("No matching rows for that date.")
        else:
            month_folder = export_date.strftime("%B")

            # (A) Per‑job TimeEntries exports (matching template formatting)
            jobs_for_day = sorted(day_df["Job Number"].astype(str).str.strip().unique().tolist())
            n_files = 0
            for job in jobs_for_day:
                try:
                    buf = export_per_job_with_template(day_df, job, export_date, xlsx_path)
                    file_name = f"{export_date.strftime('%m-%d-%Y')} - {job} - Daily Time Import.xlsx"
                    offer_download_and_sharepoint(file_name, buf.getvalue(), month_folder)
                    n_files += 1
                except Exception as e:
                    st.error(f"Failed to create job export for {job}: {e}")

            if n_files == 0:
                st.info("No per‑job files were created (no REG/OT hours).")

            # (B) Daily Time populated from template
            daily_template = APP_DIR / "Daily Time.xlsx"
            if not daily_template.exists():
                st.warning("Template 'Daily Time.xlsx' not found beside the app.")
            else:
                try:
                    out2 = export_daily_report_populated(xlsx_path, str(daily_template), export_date.strftime("%Y-%m-%d"), st.session_state.get("whoami_email",""))
                    if out2 is not None:
                        daily_name = f"{export_date.strftime('%m-%d-%Y')} – Daily Time.xlsx"
                        offer_download_and_sharepoint(daily_name, out2.getvalue(), month_folder)
                    else:
                        st.info("No Daily Time report created (no rows for that date).")
                except Exception as e:
                    st.error(f"Failed to build daily report: {e}")

st.caption("Exports inherit exact formatting from the TimeEntries sheet in your workbook; the daily report is populated from your Daily Time.xlsx template.")
