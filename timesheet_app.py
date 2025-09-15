
# timesheet_app.py — Streamlit app for Excel-backed daily timesheets (v14)
# Usage:
#   cd "C:\Users\adamm\OneDrive - PTW Energy Services LTD\Apps\Power Apps Spreadsheets\Daily Timesheet App\timesheet_python_kit_full"
#   .\.venv\Scripts\activate
#   streamlit run timesheet_app.py

import os
import datetime as dt
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# Configure the Streamlit page. Use an emoji code for the clock icon and
# collapse the sidebar by default so the home page isn't obscured.
st.set_page_config(
    page_title="Timesheet (Excel)",
    page_icon=":clock3:",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# Determine the default workbook path. If a sidecar file exists with a
# previously saved path, use that; otherwise fall back to the environment
# variable STREAMLIT_TIMESHEET_XLSX or a hard-coded filename.
_DEFAULT_XLSX_ENV = os.getenv("STREAMLIT_TIMESHEET_XLSX", "TimeSheet Apps.xlsx")
_DEFAULT_XLSX_FILE = "timesheet_default_path.txt"
if os.path.exists(_DEFAULT_XLSX_FILE):
    try:
        _file_contents = Path(_DEFAULT_XLSX_FILE).read_text().strip()
        if _file_contents:
            _DEFAULT_XLSX_ENV = _file_contents
    except Exception:
        pass
DEFAULT_XLSX = _DEFAULT_XLSX_ENV

# --------------------------------------------------------------------------
# Landing / Home Page gate
#
# Render a simple front page prompting the user for the workbook path and
# their work email. Only after they click "Enter" will the remainder of the
# application (including the sidebar) be rendered. The sidebar remains
# collapsed on first load.
if not st.session_state.get("entered_app", False) or not st.session_state.get("whoami_email", "").strip():
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        # Display the company logo if present in the current directory.
        _logo_path = None
        for _candidate in ["PTW.jpg", "PTW-Square.jpg", "PTW.png"]:
            if os.path.exists(_candidate):
                _logo_path = _candidate
                break
        if _logo_path:
            # Use the container width for the image to avoid deprecation warnings.
            st.image(_logo_path, use_container_width=True)

        # Pre-fill the workbook path with any value from session_state or the default.
        _prefill = st.session_state.get("xlsx_path", DEFAULT_XLSX)
        path_input = st.text_input(
            "Excel workbook path",
            value=_prefill,
            help="Full path (e.g., C:\\...\\TimeSheet Apps.xlsx). If the workbook lives in SharePoint, use the local OneDrive-synced path.",
            key="home_xlsx_path",
        )

        # Allow the user to save this path as their default.
        if st.button("Set As Default"):
            try:
                with open(_DEFAULT_XLSX_FILE, "w") as f:
                    f.write(path_input.strip())
                st.session_state["xlsx_path"] = path_input.strip()
                st.success("Default path saved.")
            except Exception as e:
                st.error(f"Failed to save default path: {e}")

        # Input for the user's work email. This value will be used to gate
        # access if the 'Users' sheet is enforced later.
        email_input = st.text_input(
            "Your work email",
            value=st.session_state.get("whoami_email", ""),
            placeholder="name@ptwenergy.com",
            key="home_email",
        )

        # Once the user clicks Enter, persist values and rerun.
        if st.button("Enter"):
            st.session_state["entered_app"] = True
            st.session_state["whoami_email"] = email_input.strip()
            st.session_state["xlsx_path"] = path_input.strip()
            # Rerun the app so the main UI renders.
            st.rerun()

    # Stop execution so only the landing page shows.
    st.stop()

# --------------------------------------------------------------------------
# Main title — only displayed once the user has entered the app.
st.title("🕒 Timesheet Entry (Excel backend) — v14")

# Override the DEFAULT_XLSX with any value chosen on the landing page.
DEFAULT_XLSX = st.session_state.get("xlsx_path", DEFAULT_XLSX)

# ---------- utils ----------
def _safe_rerun():
    try:
        if hasattr(st, "rerun"):
            st.rerun()
    except Exception:
        pass

def _clean_headers(df: pd.DataFrame) -> pd.DataFrame:
    try:
        df.columns = [str(c).strip() if c is not None else "" for c in df.columns]
    except Exception:
        pass
    return df

def _first_present(cols, names):
    colset = {str(c) for c in cols}
    for n in names:
        if n in colset:
            return n
    return None

def _pad_job_area(v) -> str:
    s = str(v).strip()
    if s.isdigit():
        try:
            return f"{int(s):03d}"
        except Exception:
            return s
    return s

def _clean_blank(val, upper: bool=False) -> str:
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    s = str(val).strip()
    if s.upper() == "NAN":
        return ""
    return s.upper() if upper else s

def hours_split_by_ot(date_str: str, hours_val: float):
    try:
        d = pd.to_datetime(date_str).date()
    except Exception:
        d = dt.date.today()
    weekday = d.weekday()
    hours = float(hours_val or 0.0)
    if weekday >= 5:
        return 0.0, hours
    if hours <= 8.0:
        return hours, 0.0
    return 8.0, hours - 8.0

# ---------- schema ----------
BASE_TIME_DATA = ["Job Number","Job Area","Date","Name","Class Type","Trade Class","Employee Number"]
OPTIONAL_TIME_DATA = ["RT Hours","OT Hours","Night Shift","Premium Rate / Subsistence Rate / Travel Rate","Comments"]

# ---------- sidebar ----------
with st.sidebar:
    st.subheader("Settings")
    xlsx_path = st.text_input("Excel workbook path", DEFAULT_XLSX,
                              help="Full path (e.g., C:\\...\\TimeSheet Apps.xlsx). If the workbook lives in SharePoint, use the local OneDrive-synced path.")
    filter_emp_cost_active = st.checkbox("Filter Employees & Cost Codes by ACTIVE", value=True)
    filter_jobs_active     = st.checkbox("Filter Job Numbers by ACTIVE", value=True)
    enforce_users          = st.checkbox("Restrict access via 'Users' sheet", value=True,
                                         help="If on, only emails listed as Active in the 'Users' worksheet can use the app.")
    user_email             = st.text_input("Your work email", placeholder="name@ptwenergy.com", key="whoami_email")
    st.caption("Close the Excel file while using this app (Excel locks files).")

# ---------- I/O helpers ----------
def _read_time_data_df(xlsx_file: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(xlsx_file, sheet_name="Time Data")
        _clean_headers(df)
        return df
    except Exception:
        cols = BASE_TIME_DATA + OPTIONAL_TIME_DATA
        return pd.DataFrame(columns=cols)

def _ensure_time_data_headers(xlsx_file: str, headers_to_add: list):
    wb = load_workbook(xlsx_file)
    if "Time Data" not in wb.sheetnames:
        ws = wb.create_sheet("Time Data")
        ws.append(BASE_TIME_DATA + OPTIONAL_TIME_DATA)
        wb.save(xlsx_file)
        return
    ws = wb["Time Data"]
    header_cells = next(ws.iter_rows(min_row=1, max_row=1))
    headers = [str(c.value).strip() if c.value is not None else "" for c in header_cells]
    changed = False
    for h in headers_to_add:
        if h not in headers:
            ws.cell(row=1, column=len(headers)+1, value=h)
            headers.append(h)
            changed = True
    if changed:
        wb.save(xlsx_file)

def _append_dict_row_to_time_data(xlsx_file: str, payload: dict):
    wb = load_workbook(xlsx_file)
    if "Time Data" not in wb.sheetnames:
        ws = wb.create_sheet("Time Data")
        ws.append(BASE_TIME_DATA + OPTIONAL_TIME_DATA)
    ws = wb["Time Data"]
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    headers = [str(h).strip() if h is not None else "" for h in headers]
    if not any(headers):
        headers = BASE_TIME_DATA + OPTIONAL_TIME_DATA
        for idx, h in enumerate(headers, start=1):
            ws.cell(row=1, column=idx, value=h)
    row_vals = [payload.get(h, "") for h in headers]
    ws.append(row_vals)
    wb.save(xlsx_file)

def _replace_time_data_with_df(xlsx_file: str, df: pd.DataFrame):
    from openpyxl.utils.dataframe import dataframe_to_rows
    wb = load_workbook(xlsx_file)
    if "Time Data" in wb.sheetnames:
        idx = wb.sheetnames.index("Time Data")
        wb.remove(wb["Time Data"])
        ws = wb.create_sheet("Time Data", idx)
    else:
        ws = wb.create_sheet("Time Data")
    ws.append([c for c in df.columns])
    for r in dataframe_to_rows(df, index=False, header=False):
        ws.append(r)
    wb.save(xlsx_file)

# ---------- lookups ----------
@st.cache_data(ttl=30)
def load_users(path: str):
    try:
        users = pd.read_excel(path, sheet_name="Users"); _clean_headers(users)
    except Exception:
        users = pd.DataFrame(columns=["Email","Active"])
    email_col = _first_present(users.columns, ["Email","User Email","Work Email","MAIL"])
    active_col = _first_present(users.columns, ["Active","ACTIVE"])
    return users, email_col, active_col

@st.cache_data(ttl=30)
def load_lookups(path: str, filter_emp_cost_active: bool, filter_jobs_active: bool):
    employees = pd.read_excel(path, sheet_name="Employee List"); _clean_headers(employees)
    jobs      = pd.read_excel(path, sheet_name="Job Numbers"); _clean_headers(jobs)
    costcodes = pd.read_excel(path, sheet_name="Cost Codes"); _clean_headers(costcodes)

    # EMPLOYEES
    emp_name_col    = _first_present(employees.columns, ["Employee Name","Name"])
    emp_num_col     = _first_present(employees.columns, ["Person Number","Employee Number","Emp #","Emp Number"])
    emp_trade_col   = _first_present(employees.columns, ["Override Trade Class","Trade Class"])
    emp_trt_col     = _first_present(employees.columns, ["Time Record Type"])
    emp_active_col  = _first_present(employees.columns, ["Active","ACTIVE"])
    emp_night_col   = _first_present(employees.columns, ["Night Shift","NightShift","Night_Shift"])
    emp_prem_col    = _first_present(employees.columns, ["Premium Rate / Subsistence Rate / Travel Rate","Premium Code","Premium","Premium Rate"])

    # JOBS
    job_num_col     = "JOB #" if "JOB #" in jobs.columns else _first_present(jobs.columns, ["JOB #","Job Number","Job #","JOB NUMBER"])
    job_area_col    = "AREA #" if "AREA #" in jobs.columns else _first_present(jobs.columns, ["AREA #","Job Area","Area #","AREA","AREA#"])
    job_desc_col    = "DESCRIPTION" if "DESCRIPTION" in jobs.columns else _first_present(jobs.columns, ["DESCRIPTION","Area Description","Job Area Description","Description","Area Name","Name","Label","Title"])
    job_active_col  = _first_present(jobs.columns, ["ACTIVE","Active"])

    # COST CODES
    cost_code_col   = _first_present(costcodes.columns, ["Cost Code","Class Type","CostCode","Cost-Code","CostCode#"])
    cost_active_col = _first_present(costcodes.columns, ["Active","ACTIVE"])

    # ACTIVE filters
    if filter_emp_cost_active and (emp_active_col in employees.columns):
        try: employees = employees[employees[emp_active_col] == True]
        except Exception: pass
    if filter_emp_cost_active and (cost_active_col in costcodes.columns):
        try: costcodes = costcodes[costcodes[cost_active_col] == True]
        except Exception: pass
    if filter_jobs_active and (job_active_col in jobs.columns):
        try: jobs = jobs[jobs[job_active_col] == True]
        except Exception: pass

    paycode_map: Dict[str,str] = {"REG": "211", "OT": "212", "SUBSISTENCE": "261"}

    return {
        "employees": employees, "jobs": jobs, "costcodes": costcodes, "paycode_map": paycode_map,
        "emp_name_col": emp_name_col, "emp_num_col": emp_num_col, "emp_trade_col": emp_trade_col, "emp_trt_col": emp_trt_col,
        "emp_night_col": emp_night_col, "emp_prem_col": emp_prem_col,
        "job_num_col": job_num_col, "job_area_col": job_area_col, "job_desc_col": job_desc_col,
        "cost_code_col": cost_code_col
    }

# ---------- guard ----------
if not os.path.exists(xlsx_path):
    st.error("Excel file not found. Check the path in the sidebar and try again.")
    st.stop()

# ---------- Users gating ----------
users_df, users_email_col, users_active_col = load_users(xlsx_path)
if 'enforce_users' in st.session_state:
    _enforce = st.session_state['enforce_users']
else:
    _enforce = True
if _enforce:
    allowed = set(
        users_df.loc[users_df.get(users_active_col, True) == True, users_email_col]
        .dropna().astype(str).str.strip().str.lower()
    ) if users_email_col else set()
    if not st.session_state.get('whoami_email', '').strip():
        st.info("Enter your work email in the sidebar to continue (must exist in the 'Users' sheet and be Active).")
        st.stop()
    if users_email_col and st.session_state['whoami_email'].strip().lower() not in allowed:
        st.error("You're not on the 'Users' sheet (Active). Ask an admin to add you.")
        st.stop()

# ---------- load lookups ----------
try:
    look = load_lookups(xlsx_path, st.session_state.get('filter_emp_cost_active', True), st.session_state.get('filter_jobs_active', True))
except Exception as e:
    st.error(f"Couldn't read lookup sheets.\n\n{e}")
    st.stop()

employees = look["employees"]; jobs = look["jobs"]; costcodes = look["costcodes"]
paycode_map = look["paycode_map"]
emp_name_col  = look["emp_name_col"]; emp_num_col = look["emp_num_col"]; emp_trade_col = look["emp_trade_col"]; emp_trt_col=look["emp_trt_col"]
emp_night_col = look["emp_night_col"]; emp_prem_col = look["emp_prem_col"]
job_num_col   = look["job_num_col"];  job_area_col= look["job_area_col"];  job_desc_col = look["job_desc_col"]
cost_code_col = look["cost_code_col"]

# ---------- Diagnostics ----------
with st.expander("Diagnostics"):
    st.write("Workbook:", xlsx_path)
    st.write({"job_num_col": job_num_col, "job_area_col": job_area_col, "job_desc_col": job_desc_col, "cost_code_col": cost_code_col})
    st.write({"emp_name_col": emp_name_col, "emp_num_col": emp_num_col, "emp_trade_col": emp_trade_col, "emp_trt_col": emp_trt_col, "emp_night_col": emp_night_col, "emp_prem_col": emp_prem_col})
    st.write("Rows → Employees:", len(employees), " Jobs:", len(jobs), " Cost Codes:", len(costcodes))

# ---------- After-submit reset ----------
if st.session_state.get("_reset_after_submit", False):
    st.session_state["emp_multiselect"] = []
    st.session_state["job_sel"] = ""
    st.session_state["area_sel"] = ""
    st.session_state["code_sel"] = ""
    st.session_state["rt_hours"] = 0.0
    st.session_state["ot_hours"] = 0.0
    st.session_state["desc_text"] = ""
    st.session_state["_reset_after_submit"] = False

# ---------- Add Time Entry ----------
st.subheader("Add Time Entry")

date_val = st.date_input("Date", dt.date.today(), key="date_val")
if "rt_hours" not in st.session_state: st.session_state["rt_hours"] = 0.0
if "ot_hours" not in st.session_state: st.session_state["ot_hours"] = 0.0

# Employees (multi-select)
emp_opts = employees[emp_name_col].astype(str).tolist() if emp_name_col else []
sel_emps = st.multiselect("Employees (select one or more)", emp_opts, key="emp_multiselect")

# Job Number & Area — labels "code - DESCRIPTION", store 3-digit code
job_opts = jobs[job_num_col].astype(str).unique().tolist() if job_num_col else []
sel_job  = st.selectbox("Job Number", [""] + job_opts, key="job_sel")

area_labels = []; area_map: Dict[str,str] = {}
if sel_job and job_area_col and job_num_col:
    df = jobs.loc[jobs[job_num_col].astype(str) == str(sel_job)].copy(); _clean_headers(df)
    for _, row in df.iterrows():
        code = _pad_job_area(row.get(job_area_col, ""))
        desc = str(row.get(job_desc_col, "")).strip() if job_desc_col else ""
        label = f"{code} - {desc}" if desc else code
        if code and label not in area_map:
            area_labels.append(label); area_map[label] = code

sel_area_label = st.selectbox("Job Area (for selected Job)", [""] + area_labels, key="area_sel")
sel_area_code = area_map.get(sel_area_label, sel_area_label or "")

# Cost Code — labels "code - description", store code only
def build_cost_code_labels(cost_df: pd.DataFrame, cost_code_col: str) -> Tuple[List[str], Dict[str,str]]:
    if not cost_code_col:
        return [], {}
    df = cost_df.copy(); _clean_headers(df)
    desc_col = _first_present(df.columns, ["Cost Code Description","Class Type Description","Description","Cost Code Name","Name","Label","Title"])
    labels, mapping = [], {}
    for _, row in df.iterrows():
        code = str(row.get(cost_code_col, "")).strip()
        if not code:
            continue
        desc = str(row.get(desc_col, "")).strip() if desc_col else ""
        label = f"{code} - {desc}" if desc else code
        if label not in mapping:
            labels.append(label); mapping[label] = code
    return labels, mapping

code_labels, code_map = build_cost_code_labels(costcodes, cost_code_col)
sel_code_label  = st.selectbox("Class Type (Cost Code)", [""] + code_labels, key="code_sel")
sel_code_code = code_map.get(sel_code_label, sel_code_label or "")

# RT/OT & multi-line description
rt_hours = st.number_input("RT Hours (Per Employee)", min_value=0.0, max_value=24.0, step=0.5, key="rt_hours")
ot_hours = st.number_input("OT Hours (Per Employee)", min_value=0.0, max_value=24.0, step=0.5, key="ot_hours")
desc     = st.text_area("Description of work (optional — not exported for now)", "", key="desc_text", height=140)

if st.button("Submit"):
    if not sel_emps:
        st.warning("Please select at least one employee.")
    elif not sel_job or not sel_area_code or not sel_code_code:
        st.warning("Please select a Job Number, Job Area, and Class Type.")
    else:
        try:
            _ensure_time_data_headers(xlsx_path, ["RT Hours","OT Hours","Night Shift","Premium Rate / Subsistence Rate / Travel Rate","Comments"])
        except Exception as e:
            st.warning(f"Could not ensure headers (continuing): {e}")
        successes = 0
        for emp_name in sel_emps:
            try:
                emp_row = employees.loc[employees[emp_name_col].astype(str) == str(emp_name)].iloc[0]
            except Exception:
                st.error(f"Could not find employee '{emp_name}' in Employee List."); continue

            trade_class = emp_row[emp_trade_col] if emp_trade_col and emp_trade_col in emp_row else ""
            emp_num     = emp_row[emp_num_col] if emp_num_col and emp_num_col in emp_row else ""
            night_val   = _clean_blank(emp_row.get(emp_night_col, "") if emp_night_col else "", upper=True)
            prem_val    = _clean_blank(emp_row.get(emp_prem_col, "") if emp_prem_col else "")

            payload = {
                "Job Number": str(sel_job),
                "Job Area": _pad_job_area(sel_area_code),
                "Date": date_val.strftime("%Y-%m-%d"),
                "Name": emp_name,
                "Class Type": sel_code_code,
                "Trade Class": trade_class,
                "Employee Number": emp_num,
                "RT Hours": float(rt_hours),
                "OT Hours": float(ot_hours),
                "Night Shift": night_val,
                "Premium Rate / Subsistence Rate / Travel Rate": prem_val,
                "Comments": desc,
            }
            try:
                _append_dict_row_to_time_data(xlsx_path, payload); successes += 1
            except Exception as e:
                st.error(f"Failed to append row for {emp_name}: {e}")
        if successes:
            st.success(f"Added {successes} row(s) to 'Time Data'.")
            st.session_state["_reset_after_submit"] = True
            _safe_rerun()

# ---------- Day Preview ----------
st.markdown("---"); st.subheader("What’s been added for this day")
preview_job_filter = st.checkbox("Filter by selected Job Number", value=False)
td_all = _read_time_data_df(xlsx_path); _clean_headers(td_all)
if "Date" in td_all.columns:
    mask = td_all["Date"].astype(str).str[:10] == date_val.strftime("%Y-%m-%d")
else:
    mask = pd.Series([False]*len(td_all), index=td_all.index)
if preview_job_filter and sel_job and sel_job != "" and "Job Number" in td_all.columns:
    mask = mask & (td_all["Job Number"].astype(str).str.strip() == str(sel_job).strip())
td_day = td_all.loc[mask].copy()
for c in ["RT Hours","OT Hours","Night Shift","Premium Rate / Subsistence Rate / Travel Rate","Comments"]:
    if c not in td_day.columns:
        td_day[c] = "" if c in ["Night Shift","Premium Rate / Subsistence Rate / Travel Rate","Comments"] else 0.0
if "Job Area" in td_day.columns:
    td_day["Job Area"] = td_day["Job Area"].apply(_pad_job_area)
if "Date" in td_day.columns and len(td_day)>0 and "Hours" in td_day.columns:
    comp = td_day.apply(lambda r: hours_split_by_ot(str(r["Date"]), r.get("Hours", 0)), axis=1)
    comp_df = pd.DataFrame(comp.tolist(), columns=["_rt","_ot"], index=td_day.index)
    zero_mask = (td_day.get("RT Hours", 0).fillna(0)==0) & (td_day.get("OT Hours", 0).fillna(0)==0)
    td_day.loc[zero_mask, "RT Hours"] = comp_df.loc[zero_mask, "_rt"]
    td_day.loc[zero_mask, "OT Hours"] = comp_df.loc[zero_mask, "_ot"]
td_day_display = td_day.reindex(columns=["Job Number","Job Area","Date","Name","Class Type","Trade Class","Employee Number","RT Hours","OT Hours","Night Shift","Premium Rate / Subsistence Rate / Travel Rate","Comments"]).copy()
td_day_display.insert(0, "IDX", td_day.index)
st.dataframe(td_day_display, use_container_width=True, height=360)
del_ids = st.multiselect("Select IDX values to delete from Time Data", td_day.index.tolist())
if st.button("Delete selected rows"):
    if not del_ids:
        st.warning("No rows selected.")
    else:
        new_df = td_all.drop(index=del_ids).copy()
        try:
            _replace_time_data_with_df(xlsx_path, new_df); st.success(f"Deleted {len(del_ids)} row(s) from 'Time Data'.")
        except Exception as e:
            st.error(f"Failed to delete rows: {e}")

# ---------- Export ----------
def build_export_rows(td_subset: pd.DataFrame, employee_list: pd.DataFrame, paycode_map: Dict[str,str]) -> pd.DataFrame:
    td = td_subset.copy(); _clean_headers(td)
    el = employee_list.copy(); _clean_headers(el)
    if "Employee Name" in el.columns:
        merged = td.merge(
            el[["Employee Name","Time Record Type","Person Number","Override Trade Class"]],
            how="left", left_on="Name", right_on="Employee Name"
        )
    else:
        merged = td.copy()
        merged["Time Record Type"] = ""
        merged["Person Number"] = ""
        merged["Override Trade Class"] = ""

    rows = []
    for _, r in merged.iterrows():
        date_s = pd.to_datetime(r.get("Date", "")).strftime("%Y-%m-%d")
        reg_h = float(r.get("RT Hours", 0) or 0.0)
        ot_h  = float(r.get("OT Hours", 0) or 0.0)
        ns_clean = _clean_blank(r.get("Night Shift", ""), upper=True)
        pv_clean = _clean_blank(r.get("Premium Rate / Subsistence Rate / Travel Rate", ""))
        base = {
            "Date": date_s,
            "Time Record Type": r.get("Time Record Type",""),
            "Person Number": r.get("Employee Number", r.get("Person Number","")),
            "Employee Name": r.get("Name",""),
            "Override Trade Class": r.get("Trade Class", r.get("Override Trade Class","")),
            "Post To Payroll": "Y",
            "Cost Code / Phase": r.get("Class Type",""),
            "JobArea": _pad_job_area(r.get("Job Area", r.get("JobArea",""))),
            "Scope Change": "",
            "Pay Code": "",
            "Hours": 0.0,
            "Night Shift": ns_clean,
            "Premium Rate / Subsistence Rate / Travel Rate": pv_clean,
            "Comments": ""
        }
        reg_code = paycode_map.get("REG","211")
        ot_code  = paycode_map.get("OT","212")
        subs_code= paycode_map.get("SUBSISTENCE","261")
        if reg_h > 0:
            t = base.copy(); t["Pay Code"] = reg_code; t["Hours"] = float(reg_h); rows.append(t)
        if ot_h > 0:
            t = base.copy(); t["Pay Code"] = ot_code;  t["Hours"] = float(ot_h); rows.append(t)
        if ("SUBSIST" in pv_clean.upper()) or (pv_clean == "261") or ("261" in pv_clean):
            t = base.copy(); t["Pay Code"] = subs_code; t["Hours"] = 1.0; rows.append(t)
    HEADERS = ['Date','Time Record Type','Person Number','Employee Name','Override Trade Class',
               'Post To Payroll','Cost Code / Phase','JobArea','Scope Change','Pay Code','Hours',
               'Night Shift','Premium Rate / Subsistence Rate / Travel Rate','Comments']
    return pd.DataFrame(rows, columns=HEADERS)

def write_formatted(out_df: pd.DataFrame, out_path: Path):
    NOTE_TEXT = "** Note: Please do not exceed 2000 rows of data."
    HEADERS = ['Date','Time Record Type','Person Number','Employee Name','Override Trade Class',
               'Post To Payroll','Cost Code / Phase','JobArea','Scope Change','Pay Code','Hours',
               'Night Shift','Premium Rate / Subsistence Rate / Travel Rate','Comments']
    out_path.parent.mkdir(parents=True, exist_ok=True)
    import xlsxwriter
    wb = xlsxwriter.Workbook(out_path); ws = wb.add_worksheet("TimeEntries")
    note_fmt = wb.add_format({"bg_color":"#FFF200"})
    base_header   = wb.add_format({"bold":True,"font_color":"#FFFFFF","align":"center","valign":"vcenter","bg_color":"#4F81BD","border":1})
    green_header  = wb.add_format({"bold":True,"font_color":"#FFFFFF","align":"center","valign":"vcenter","bg_color":"#00B050","border":1})
    orange_header = wb.add_format({"bold":True,"font_color":"#FFFFFF","align":"center","valign":"vcenter","bg_color":"#FFC000","border":1})
    text_fmt = wb.add_format({})
    num_fmt  = wb.add_format({"num_format":"0.00"})
    date_fmt = wb.add_format({"num_format":"yyyy-mm-dd"})
    gray_band= wb.add_format({"bg_color":"#F2F2F2"})
    widths = [12,18,14,22,18,14,18,10,14,12,8,12,28,18]
    for col, w in enumerate(widths): ws.set_column(col, col, w)
    ws.write(0, 1, NOTE_TEXT, note_fmt)
    header_formats = [base_header]*len(HEADERS); header_formats[4]=green_header; header_formats[10]=orange_header
    for c, (h, fmt) in enumerate(zip(HEADERS, header_formats)): ws.write(2, c, h, fmt)
    ws.autofilter(2, 0, 2, len(HEADERS)-1); ws.freeze_panes(3, 0)
    start_row = 3
    for r in range(len(out_df)):
        row = out_df.iloc[r]
        for c, h in enumerate(HEADERS):
            val = row[h]; fmt = num_fmt if h=="Hours" else (date_fmt if h=="Date" else text_fmt)
            ws.write(start_row + r, c, val if pd.notna(val) else "", fmt)
    if len(out_df) > 0:
        end_row = start_row + len(out_df) - 1
        ws.conditional_format(start_row, 0, end_row, len(HEADERS)-1, {"type":"formula","criteria":"=MOD(ROW(),2)=0","format": gray_band})
    wb.close()

def export_all_jobs_for_date(xlsx: str, date_str: str, outdir: str, user: str, paycode_map: Dict[str,str]):
    td = _read_time_data_df(xlsx); el = pd.read_excel(xlsx, sheet_name="Employee List")
    _clean_headers(el); _clean_headers(td)
    if "Date" not in td.columns: return []
    date_col = td["Date"].astype(str).str[:10]
    subset = td[date_col == date_str].copy()
    if subset.empty: return []
    job_nums = sorted(subset["Job Number"].astype(str).str.strip().unique().tolist()) if "Job Number" in subset.columns else []
    out_paths = []; wb = load_workbook(xlsx)
    if "Exports Log" not in wb.sheetnames:
        ws_log = wb.create_sheet("Exports Log")
        ws_log.append(["LogID","Date","Job Number","Entries Count","File Name","OneDrive Path","Share Link","Triggered By","Triggered At","Status","Notes"])
    ws_log = wb["Exports Log"]
    dt_obj = pd.to_datetime(date_str).date()
    month_folder = dt_obj.strftime("%Y - %B")
    for job in job_nums:
        sub = subset[subset["Job Number"].astype(str).str.strip() == job].copy()
        out_df = build_export_rows(sub, el, paycode_map)
        file_name = f"{dt_obj.strftime('%d-%m-%Y')} - {job} - Daily Time Import.xlsx"
        out_path = Path(outdir) / month_folder / file_name
        write_formatted(out_df, out_path); out_paths.append(out_path)
        log_id = f"{date_str}-{job}-{datetime.utcnow().strftime('%H%M%S')}"
        status = "Created" if len(out_df)>0 else "Empty"
        notes  = "" if len(out_df)>0 else "No matching rows; created headers only."
        ws_log.append([log_id, date_str, str(job), len(out_df), out_path.name, str(out_path.parent), "", user,
                       datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"), status, notes])
    wb.save(xlsx)
    return out_paths

def export_descriptions_for_date(xlsx: str, date_str: str, outdir: str, user: str):
    td = _read_time_data_df(xlsx); _clean_headers(td)
    if "Date" not in td.columns or "Job Number" not in td.columns or "Comments" not in td.columns: return []
    date_col = td["Date"].astype(str).str[:10]
    subset = td[date_col == date_str].copy()
    if subset.empty: return []
    out_paths = []
    dt_obj = pd.to_datetime(date_str).date()
    month_folder = dt_obj.strftime("%Y - %B")
    for job in sorted(subset["Job Number"].astype(str).str.strip().unique().tolist()):
        descs = (subset.loc[subset["Job Number"].astype(str).str.strip()==job, "Comments"]
                 .dropna().astype(str).map(lambda s: s.strip()).replace({"NAN":""}).tolist())
        seen = set(); uniq = []
        for d in descs:
            if d and d not in seen:
                uniq.append(d); seen.add(d)
        if len(uniq) == 0: continue
        file_name = f"{dt_obj.strftime('%d-%m-%Y')} - {job} - Description.xlsx"
        out_path = Path(outdir) / month_folder / file_name
        out_path.parent.mkdir(parents=True, exist_ok=True)
        import xlsxwriter
        wb = xlsxwriter.Workbook(out_path); ws = wb.add_worksheet("Descriptions")
        ws.set_column(0, 0, 100); ws.write(0, 0, "Description of work"); wrap = wb.add_format({"text_wrap": True})
        for i, text in enumerate(uniq, start=1): ws.write(i, 0, text, wrap)
        wb.close(); out_paths.append(out_path)
    return out_paths

# ---------- Export UI ----------
st.markdown("---"); st.subheader("Export Day → TimeEntries (numeric pay codes, padded JobArea)")
with st.form("export_form"):
    export_date = st.date_input("Export Date", dt.date.today())
    outdir      = st.text_input("Output folder", "Exports")
    user        = st.text_input("Triggered by (optional)", st.session_state.get('whoami_email', ''))
    do_export   = st.form_submit_button("Export ALL Jobs for Date")
    if do_export:
        try:
            paths = export_all_jobs_for_date(xlsx=xlsx_path, date_str=export_date.strftime("%Y-%m-%d"),
                                             outdir=outdir, user=user, paycode_map=paycode_map)
            desc_paths = export_descriptions_for_date(xlsx=xlsx_path, date_str=export_date.strftime("%Y-%m-%d"),
                                                      outdir=outdir, user=user)
            total = len(paths) + len(desc_paths)
            if total == 0:
                st.warning("No matching rows for that date. No files created.")
            else:
                st.success(f"Created {total} file(s): {len(paths)} TimeEntries and {len(desc_paths)} Description file(s).")
        except Exception as e:
            st.error(f"Export failed: {e}")

st.caption("Use OneDrive-synced local paths for the workbook and Output folder so SharePoint stays in sync. 'Users' sheet controls access when enabled.")
