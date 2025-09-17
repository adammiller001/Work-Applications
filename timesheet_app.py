# timesheet_app.py — Streamlit app for Excel-backed daily timesheets (v20)
# Usage:
#   cd "C:\Users\adamm\OneDrive - PTW Energy Services LTD\Apps\Power Apps Spreadsheets\Daily Timesheet App\timesheet_python_kit_full"
#   .\.venv\Scripts\activate
#   streamlit run timesheet_app.py

import os
import datetime as dt
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple
import math
import json
from openpyxl.styles import Font

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell import MergedCell

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
            st.session_state["xlsx_path"] = DEFAULT_XLSX
            # Rerun the app so the main UI renders.
            st.rerun()

    # Stop execution so only the landing page shows.
    st.stop()

# --------------------------------------------------------------------------
# Main title — only displayed once the user has entered the app.
# Remove backend and version from the title for a cleaner display.
st.title("Timesheet Entry")

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
    # Settings are editable only for admin users.  Non-admins can view but not
    # modify these fields.  The user's email field remains editable for all.
    xlsx_path = st.text_input(
        "Excel workbook path",
        DEFAULT_XLSX,
        help="Full path (e.g., C:\\...\\TimeSheet Apps.xlsx). If the workbook lives in SharePoint, use the local OneDrive-synced path.",
        disabled=not st.session_state.get("is_admin", False),
    )
    filter_emp_cost_active = st.checkbox(
        "Filter Employees & Cost Codes by ACTIVE",
        value=True,
        disabled=not st.session_state.get("is_admin", False),
    )
    filter_jobs_active = st.checkbox(
        "Filter Job Numbers by ACTIVE",
        value=True,
        disabled=not st.session_state.get("is_admin", False),
    )
    enforce_users = st.checkbox(
        "Restrict access via 'Users' sheet",
        value=True,
        help="If on, only emails listed as Active in the 'Users' worksheet can use the app.",
        disabled=not st.session_state.get("is_admin", False),
    )
    user_email = st.text_input(
        "Your work email",
        placeholder="name@ptwenergy.com",
        key="whoami_email",
    )
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
    if filter_emp_cost_active and emp_active_col in employees.columns:
        try: employees = employees[employees[emp_active_col] == True].copy()
        except Exception: pass
    if filter_emp_cost_active and cost_active_col in costcodes.columns:
        try: costcodes = costcodes[costcodes[cost_active_col] == True].copy()
        except Exception: pass
    if filter_jobs_active and job_active_col in jobs.columns:
        try: jobs = jobs[jobs[job_active_col] == True].copy()
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

# After user gating, compute whether the current user has admin privileges.  In
# this workbook, the 'Users' sheet contains a 'Type' column (column D) with
# values 'Admin' or 'User'.  Treat rows with 'Admin' (case-insensitive) as
# admins and everyone else as non-admin.  Store the result in session_state
# for use in disabling settings in the sidebar.
def _is_user_admin(df: pd.DataFrame, email_col: str, email: str) -> bool:
    if not email_col:
        return False
    # Identify the column that holds user roles or types.  Prefer 'Type'
    # explicitly; fall back to other common labels.
    admin_candidates = [
        "Type",
        "User Type",
        "Role",
        "ROLE",
        "Admin",
        "Is Admin",
        "IsAdmin",
        "ADMIN",
        "IS_ADMIN",
    ]
    admin_col = next((c for c in admin_candidates if c in df.columns), None)
    if not admin_col:
        return False
    try:
        _lower = str(email or "").strip().lower()
        row = df.loc[df[email_col].astype(str).str.strip().str.lower() == _lower].iloc[0]
        val = str(row.get(admin_col, "")).strip().lower()
        # If using the 'Type' convention, only 'admin' counts as admin.  For
        # other admin columns, treat typical truthy values as admin.
        if admin_col.lower() == "type":
            return val == "admin"
        return val in {"true", "t", "yes", "y", "1", "admin"}
    except Exception:
        return False

st.session_state["is_admin"] = _is_user_admin(users_df, users_email_col, st.session_state.get("whoami_email", ""))

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
# Diagnostics have been removed from the main interface for simplicity.  If you
# need to inspect column mappings or row counts, uncomment the lines below.
# with st.expander("Diagnostics"):
#     st.write("Workbook:", xlsx_path)
#     st.write({"job_num_col": job_num_col, "job_area_col": job_area_col, "job_desc_col": job_desc_col, "cost_code_col": cost_code_col})
#     st.write({"emp_name_col": emp_name_col, "emp_num_col": emp_num_col, "emp_trade_col": emp_trade_col, "emp_trt_col": emp_trt_col, "emp_night_col": emp_night_col, "emp_prem_col": emp_prem_col})
#     st.write("Rows → Employees:", len(employees), " Jobs:", len(jobs), " Cost Codes:", len(costcodes))

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
st.subheader("Timesheet Entry")

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
                "Comments": desc.upper(),
            }
            try:
                _append_dict_row_to_time_data(xlsx_path, payload); successes += 1
            except Exception as e:
                st.error(f"Failed to append row for {emp_name}: {e}")
        if successes:
            st.success(f"Added {successes} row(s) to 'Time Data'.")
            st.session_state["_reset_after_submit"] = True
            _safe_rerun()

def get_user_default_outdir(email):
    try:
        with open("user_defaults.json", "r") as f:
            data = json.load(f)
            return data.get(email, str(Path(xlsx_path).parent))
    except:
        return str(Path(xlsx_path).parent)

def set_user_default_outdir(email, path):
    try:
        with open("user_defaults.json", "r") as f:
            data = json.load(f)
    except:
        data = {}
    data[email] = path
    with open("user_defaults.json", "w") as f:
        json.dump(data, f)

if st.session_state["is_admin"]:
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
    # Provide a deletion multiselect with a "Delete All" option first.  Selecting
    # "Delete All" removes all rows for the current day.  After deletion, the app
    # reruns automatically to refresh the preview without a manual refresh.
    del_options = ["Delete All"] + [str(i) for i in td_day.index.tolist()]
    selected_deletions = st.multiselect("Select rows to delete", del_options)
    if st.button("Delete selected rows"):
        if not selected_deletions:
            st.warning("No rows selected.")
        else:
            if "Delete All" in selected_deletions:
                ids_to_delete = td_day.index.tolist()
            else:
                ids_to_delete = []
                for v in selected_deletions:
                    try:
                        ids_to_delete.append(int(v))
                    except Exception:
                        pass
            if not ids_to_delete:
                st.warning("No valid rows selected.")
            else:
                new_df = td_all.drop(index=ids_to_delete).copy()
                try:
                    _replace_time_data_with_df(xlsx_path, new_df)
                    st.success(f"Deleted {len(ids_to_delete)} row(s) from 'Time Data'.")
                    _safe_rerun()
                except Exception as e:
                    st.error(f"Failed to delete rows: {e}")

# ---------- Daily Time Template Export ----------
def _employee_band(employees_df: pd.DataFrame) -> Dict[str, str]:
    el = employees_df.copy()
    _clean_headers(el)
    col = _first_present(el.columns, ["Indirect / Direct","Indirect/Direct","Band","Category"])
    band_map = {}
    if not col:
        # Default all employees to Direct if the column isn't present.
        for _, r in el.iterrows():
            band_map[str(r.get(emp_name_col, "")).strip()] = "Direct"
        return band_map
    for _, r in el.iterrows():
        name = str(r.get(emp_name_col, "")).strip()
        band = str(r.get(col, "")).strip().title()
        if band not in {"Indirect","Direct"}:
            band = "Direct"
        band_map[name] = band
    return band_map

def _area_label(job_num: str, job_area_code: str, jobs_df: pd.DataFrame) -> str:
    jd = jobs_df.copy(); _clean_headers(jd)
    # Find matching description for the job_num + area code
    if job_num_col and job_area_col:
        try:
            sub = jd[(jd[job_num_col].astype(str).str.strip()==str(job_num).strip()) &
                     (jd[job_area_col].astype(str).map(_pad_job_area).astype(str)==_pad_job_area(job_area_code))]
            desc = ""
            if not sub.empty and job_desc_col in sub.columns:
                desc = str(sub.iloc[0][job_desc_col] or "").strip()
        except Exception:
            desc = ""
    else:
        desc = ""
    area_code = _pad_job_area(job_area_code)
    # Format: Job Number – Area Number – Description
    parts = [str(job_num).strip() if job_num else "", area_code if area_code else "", desc if desc else ""]
    parts = [p for p in parts if p]
    return " – ".join(parts)

def export_daily_time_report(xlsx: str, template_path: str, date_str: str, outdir: str, user: str = "") -> str:
    """
    Build a 'Daily Time.xlsx' from the provided template, filling:
      - Rows 8–30 with Indirect employees
      - Rows 32–261 with Direct employees
      - Two job slots per employee row (Cost Code/Area/RT/OT/Subtotal x2) with 'TOTAL' auto-formula
      - Report date at G5
      - Descriptions at A264/B264 and below grouped by Job Number
    Blank rows in the bands are hidden automatically.
    """
    # Check if template file exists
    if not os.path.exists(template_path):
        st.error(f"Template file 'Daily Time.xlsx' not found in the current directory. Please ensure it exists at {os.getcwd()}")
        return None

    # Load data
    try:
        td = _read_time_data_df(xlsx); _clean_headers(td)
        if "Date" not in td.columns: 
            st.error("No 'Date' column found in 'Time Data' sheet.")
            return None
        el = pd.read_excel(xlsx, sheet_name="Employee List"); _clean_headers(el)
        jd = pd.read_excel(xlsx, sheet_name="Job Numbers"); _clean_headers(jd)
        cc = pd.read_excel(xlsx, sheet_name="Cost Codes"); _clean_headers(cc)
    except Exception as e:
        st.error(f"Failed to load data for Daily Time report: {e}")
        return None

    cost_code_col = _first_present(cc.columns, ["Cost Code","Class Type","CostCode","Cost-Code","CostCode#"])
    cost_desc_col = _first_present(cc.columns, ["Cost Code Description","Class Type Description","Description","Cost Code Name","Name","Label","Title"])
    cost_desc_map = {}
    if cost_code_col and cost_desc_col:
        for _, row in cc.iterrows():
            code = str(row.get(cost_code_col, "")).strip()
            desc = str(row.get(cost_desc_col, "")).strip()
            if code:
                cost_desc_map[code] = desc

    # Filter to date
    date_mask = td["Date"].astype(str).str[:10] == date_str
    day = td.loc[date_mask].copy()
    if day.empty:
        st.warning(f"No data found for date {date_str} in 'Time Data' sheet.")
        return None

    # Build band map and per-employee job allocations
    band_map = _employee_band(el)
    # Normalize numeric hours
    for c in ["RT Hours","OT Hours"]:
        if c not in day.columns: day[c] = 0.0
        day[c] = pd.to_numeric(day[c], errors="coerce").fillna(0.0)
    # Derive trade class from Employee List if present
    trade_map = {}
    if emp_name_col and emp_trade_col and emp_trade_col in el.columns:
        for _, r in el.iterrows():
            trade_map[str(r.get(emp_name_col, "")).strip()] = str(r.get(emp_trade_col, "")).strip()
    # Build per-employee list of job entries (each entry is a dict for one job)
    entries_by_emp: Dict[str, List[dict]] = {}
    for _, r in day.iterrows():
        emp = str(r.get("Name","")).strip()
        if not emp: 
            continue
        job_num = str(r.get("Job Number","")).strip()
        area_code = _pad_job_area(r.get("Job Area",""))
        cost_code = str(r.get("Class Type","")).strip()
        area_text = _area_label(job_num, area_code, jd)
        entry = {
            "job_num": job_num,
            "area": area_text,
            "rt": float(r.get("RT Hours", 0.0) or 0.0),
            "ot": float(r.get("OT Hours", 0.0) or 0.0),
            "cost_desc": cost_desc_map.get(cost_code, ""),
            "cost_code": cost_code,
            "premium": _clean_blank(r.get("Premium Rate / Subsistence Rate / Travel Rate","")),
        }
        entries_by_emp.setdefault(emp, []).append(entry)
    # Sort employees alphabetically for stable output
    all_emps = sorted(entries_by_emp.keys(), key=lambda s: s.upper())
    # Open template and start writing
    try:
        wb = load_workbook(template_path)
        st.write(f"Loaded template sheets: {wb.sheetnames}")  # Debug: List all sheets
        if "Sheet1" not in wb.sheetnames:
            st.error("The 'Sheet1' sheet was not found in the template.")
            return None
        ws = wb["Sheet1"]
        # Set date in G5
        try:
            d = pd.to_datetime(date_str).date()
        except Exception:
            d = dt.date.today()
        ws.cell(row=5, column=7, value=d)  # G5 is row5 column7
        # Locate starting rows
        INDIRECT_START, INDIRECT_END = 8, 30
        DIRECT_START, DIRECT_END = 32, 261
        # Calculate needed rows and extend if necessary
        def calc_needed(band: str) -> int:
            total = 0
            for emp in all_emps:
                if band_map.get(emp, "Direct") != band:
                    continue
                entries = entries_by_emp.get(emp, [])
                chunks = math.ceil(len(entries) / 2) if entries else 0
                total += chunks if entries else 0
            return total
        needed_ind = calc_needed("Indirect")
        avail_ind = INDIRECT_END - INDIRECT_START + 1
        if needed_ind > avail_ind:
            extra = needed_ind - avail_ind
            insert_at = INDIRECT_END + 1
            ws.insert_rows(insert_at, extra)
            INDIRECT_END += extra
            DIRECT_START += extra
            DIRECT_END += extra
        needed_dir = calc_needed("Direct")
        avail_dir = DIRECT_END - DIRECT_START + 1
        if needed_dir > avail_dir:
            extra = needed_dir - avail_dir
            insert_at = DIRECT_END + 1
            ws.insert_rows(insert_at, extra)
            DIRECT_END += extra
        def place_emp_row(row_idx: int, emp_name: str, entries: List[dict]):
            # Columns mapping for job1 and job2 sets
            # A: Name, B: Trade, C: Truck, D: Premium, (E..J) first job, (K..P) second job, Q: TOTAL (already a formula in template)
            ws.cell(row=row_idx, column=1, value=emp_name)  # Name
            ws.cell(row=row_idx, column=2, value=trade_map.get(emp_name, ""))  # Trade Class
            # Premium
            prem = ""
            for e in entries:
                if e.get("premium"):
                    prem = e["premium"]; break
            ws.cell(row=row_idx, column=4, value=prem)
            # Fill first two jobs; if more than two, caller will allocate additional row
            def fill_job(slot: int, job: dict):
                # slot 1 -> starts at col 5; slot 2 -> starts at col 11
                base = 5 if slot == 1 else 11
                ws.cell(row=row_idx, column=base, value=job.get("cost_desc",""))
                ws.cell(row=row_idx, column=base+1, value=job.get("cost_code",""))
                ws.cell(row=row_idx, column=base+2, value=job.get("area",""))
                ws.cell(row=row_idx, column=base+3, value=job.get("rt",0.0))
                ws.cell(row=row_idx, column=base+4, value=job.get("ot",0.0))
                # Subtotal column (base+5) is a formula in the template, leave intact if exists
            if len(entries) >= 1: fill_job(1, entries[0])
            if len(entries) >= 2: fill_job(2, entries[1])
            # TOTAL column typically already has a formula (e.g., =J8+P8). Leave as-is.
        # Split employees by band preserving template ranges
        indirect_rows = list(range(INDIRECT_START, INDIRECT_END+1))
        direct_rows   = list(range(DIRECT_START, DIRECT_END+1))
        # We'll iterate employees and place them in the appropriate band; if a person has >2 jobs, spill to next available row.
        def allocate(rows_range: List[int], band: str):
            row_iter = iter(rows_range)
            used_rows = []
            for emp in [e for e in all_emps if band_map.get(e,"Direct")==band]:
                jobs = entries_by_emp.get(emp, [])
                n = len(jobs)
                if n == 0:
                    continue
                for i in range(0, n, 2):
                    try:
                        r = next(row_iter)
                    except StopIteration:
                        raise RuntimeError(f"Out of rows for {band} employees; extend the template.")
                    place_emp_row(r, emp, jobs[i:i+2])
                    used_rows.append(r)
            # Hide unused rows in this band
            for r in rows_range:
                if r not in used_rows:
                    ws.row_dimensions[r].hidden = True
            return used_rows
        used_indirect = allocate(indirect_rows, "Indirect")
        used_direct   = allocate(direct_rows, "Direct")
        # Descriptions section
        descs = {}
        if "Comments" in day.columns and "Job Number" in day.columns:
            for job in sorted(day["Job Number"].astype(str).str.strip().unique().tolist()):
                texts = (day.loc[day["Job Number"].astype(str).str.strip()==job, "Comments"]
                         .dropna().astype(str).map(lambda s: s.strip()).replace({"NAN":""}).tolist())
                uniq = []
                seen = set()
                for t in texts:
                    if t and t not in seen:
                        uniq.append(t); seen.add(t)
                if uniq:
                    descs[job] = uniq
        # Clear area A264:B400 before writing
        for r in range(264, 401):
            ws.cell(row=r, column=1, value=None)
            ws.cell(row=r, column=2, value=None)
        row_ptr = 264
        bold_underline_font = Font(bold=True, underline="single")
        for job, lines in descs.items():
            desc_cell = ws.cell(row=row_ptr, column=1, value="Work Description")
            desc_cell.font = bold_underline_font
            job_cell = ws.cell(row=row_ptr, column=2, value=job)
            job_cell.font = bold_underline_font
            row_ptr += 1
            for line in lines:
                ws.cell(row=row_ptr, column=2, value=line)
                row_ptr += 1
            row_ptr += 1  # spacer
        # Total hours in row 259 - assume template has formulas, but recalculate if needed
        # For simplicity, leave as template formulas
        # Save file
        out_dir = Path(outdir).as_posix()  # Ensure POSIX-style path for consistency
        out_dir = out_dir.replace("/", "\\")  # Force Windows-style backslashes
        out_dir_path = Path(out_dir)
        out_dir_path.mkdir(parents=True, exist_ok=True)
        dt_obj = pd.to_datetime(date_str).date()
        out_name = f"{dt_obj.strftime('%m-%d-%Y')} – Daily Time.xlsx"
        out_path = out_dir_path / dt_obj.strftime("%B") / out_name
        st.write(f"Attempting to save file to: {out_path}")  # Debug: Show the exact save path
        wb.save(out_path)
        st.write(f"File saved to: {out_path}")  # Debug: Confirm save success
        if os.path.exists(out_path):
            import os
            file_size = os.path.getsize(out_path)
            st.write(f"File verified at: {out_path} (Size: {file_size} bytes)")  # Debug: Confirm file existence and size
        else:
            st.warning(f"File not found at: {out_path} after save attempt")  # Debug: Warn if file is missing
        # Log to Exports Log
        wb_x = load_workbook(xlsx)
        if "Exports Log" not in wb_x.sheetnames:
            ws_log = wb_x.create_sheet("Exports Log")
            ws_log.append(["LogID","Date","Job Number","Entries Count","File Name","OneDrive Path","Share Link","Triggered By","Triggered At","Status","Notes"])
        ws_log = wb_x["Exports Log"]
        log_id = f"{date_str}-ALL-{datetime.utcnow().strftime('%H%M%S')}"
        status = "Created"
        notes = ""
        entries_count = len(day)
        ws_log.append([log_id, date_str, "ALL", entries_count, out_path.name, str(out_path.parent), "", user,
                       datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"), status, notes])
        wb_x.save(xlsx)
        return str(out_path)
    except Exception as e:
        st.error(f"Error generating Daily Time report: {e}")
        return None

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
    HEADERS = [
        'Date',
        'Time Record Type',
        'Person Number',
        'Employee Name',
        'Override Trade Class',
        'Post To Payroll',
        'Cost Code / Phase',
        'JobArea',
        'Scope Change',
        'Pay Code',
        'Hours',
        'Night Shift',
        'Premium Rate / Subsistence Rate / Travel Rate',
        'Comments',
    ]
    out_path.parent.mkdir(parents=True, exist_ok=True)
    import xlsxwriter
    wb = xlsxwriter.Workbook(out_path)
    ws = wb.add_worksheet("TimeEntries")
    note_fmt = wb.add_format({"bg_color": "#FFF200"})
    # Header formats: set wrap text and centre alignment with borders
    base_header = wb.add_format(
        {
            "bold": True,
            "font_color": "#FFFFFF",
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#4F81BD",
            "border": 1,
            "text_wrap": True,
        }
    )
    green_header = wb.add_format(
        {
            "bold": True,
            "font_color": "#FFFFFF",
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#00B050",
            "border": 1,
            "text_wrap": True,
        }
    )
    orange_header = wb.add_format(
        {
            "bold": True,
            "font_color": "#FFFFFF",
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#FFC000",
            "border": 1,
            "text_wrap": True,
        }
    )
    # Data formats: borders on all cells and centre align; apply number/date formats where appropriate
    text_fmt = wb.add_format({"align": "center", "valign": "vcenter", "border": 1})
    num_fmt = wb.add_format({"num_format": "0.00", "align": "center", "valign": "vcenter", "border": 1})
    date_fmt = wb.add_format({"num_format": "yyyy-mm-dd", "align": "center", "valign": "vcenter", "border": 1})
    gray_band = wb.add_format({"bg_color": "#F2F2F2"})
    # Column widths corresponding to each header
    widths = [12, 18, 14, 22, 18, 14, 18, 10, 14, 12, 8, 12, 28, 18]
    for col, w in enumerate(widths):
        ws.set_column(col, col, w)
    # Write note
    ws.write(0, 1, NOTE_TEXT, note_fmt)
    # Build header formats list with custom colours
    header_formats = [base_header] * len(HEADERS)
    header_formats[4] = green_header  # Override Trade Class column gets green header
    header_formats[10] = orange_header  # Hours column gets orange header
    # Write headers on row index 2 (Excel row 3)
    for c, (h, fmt) in enumerate(zip(HEADERS, header_formats)):
        ws.write(2, c, h, fmt)
    # Adjust header row height (row 3) to 38.25 points and enable wrap text
    ws.set_row(2, 38.25)
    # Freeze panes at the beginning of data rows; do not set filters on header row
    ws.freeze_panes(3, 0)
    # Write data
    start_row = 3
    for r in range(len(out_df)):
        row = out_df.iloc[r]
        for c, h in enumerate(HEADERS):
            val = row[h]
            # Choose format based on column
            fmt = num_fmt if h == "Hours" else (date_fmt if h == "Date" else text_fmt)
            ws.write(start_row + r, c, val if pd.notna(val) else "", fmt)
    # Add alternating grey banding on even data rows
    if len(out_df) > 0:
        end_row = start_row + len(out_df) - 1
        ws.conditional_format(
            start_row,
            0,
            end_row,
            len(HEADERS) - 1,
            {
                "type": "formula",
                "criteria": "=MOD(ROW(),2)=0",
                "format": gray_band,
            },
        )
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
    # Create a folder named only by the month (e.g. "September"); year is omitted
    month_folder = dt_obj.strftime("%B")
    for job in job_nums:
        sub = subset[subset["Job Number"].astype(str).str.strip() == job].copy()
        out_df = build_export_rows(sub, el, paycode_map)
        # File names use mm-dd-yyyy format per user request
        file_name = f"{dt_obj.strftime('%m-%d-%Y')} - {job} - Daily Time Import.xlsx"
        out_path = Path(outdir) / month_folder / file_name
        write_formatted(out_df, out_path); out_paths.append(out_path)
        log_id = f"{date_str}-{job}-{datetime.utcnow().strftime('%H%M%S')}"
        status = "Created" if len(out_df)>0 else "Empty"
        notes  = "" if len(out_df)>0 else "No matching rows; created headers only."
        ws_log.append([log_id, date_str, str(job), len(out_df), out_path.name, str(out_path.parent), "", user,
                       datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"), status, notes])
    wb.save(xlsx)
    return out_paths

if st.session_state["is_admin"]:
    # ---------- Export UI ----------
    st.markdown("---")
    st.subheader("Export Day → TimeEntries (numeric pay codes, padded JobArea)")
    # Move "Set to Default" button outside the form
    user = st.text_input("Triggered by (optional)", st.session_state.get('whoami_email', ''))
    default_outdir = get_user_default_outdir(st.session_state["whoami_email"])
    outdir_input = st.text_input("Export Directory", value=default_outdir, key="export_outdir")
    if st.button("Set to Default"):  # Moved outside form
        set_user_default_outdir(st.session_state["whoami_email"], outdir_input)
        st.success("Export directory set as default.")

    # Form with submit button
    with st.form("export_form"):
        export_date = st.date_input("Export Date", dt.date.today())
        do_export = st.form_submit_button("Export ALL Jobs for Date")
        if do_export:
            outdir = st.session_state.get("export_outdir", str(Path(xlsx_path).parent))
            try:
                paths = export_all_jobs_for_date(
                    xlsx=xlsx_path,
                    date_str=export_date.strftime("%Y-%m-%d"),
                    outdir=outdir,
                    user=user,
                    paycode_map=paycode_map,
                )
                daily_path = export_daily_time_report(
                    xlsx=xlsx_path,
                    template_path="Daily Time.xlsx",  # Added back the template_path parameter
                    date_str=export_date.strftime("%Y-%m-%d"),
                    outdir=outdir,
                    user=user,
                )
                n_daily = 1 if daily_path else 0
                if len(paths) + n_daily == 0:
                    st.warning("No matching rows for that date. No files created.")
                else:
                    st.success(
                        f"Created {len(paths)} TimeEntries file(s) and {n_daily} Daily Time report(s)."
                    )
            except Exception as e:
                st.error(f"Export failed: {e}")

    st.caption(
        "Use OneDrive-synced local paths for the workbook. Export files are saved in the workbook's folder in a month-named subfolder (e.g. 'September')."
    )