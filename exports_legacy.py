# exports_legacy.py
# Rebuilds the original per-job "Daily Time Import.xlsx" exports from Supabase rows.
from __future__ import annotations
import io, pandas as pd

EXPORT_COLUMNS = [
    'Date','Time Record Type','Person Number','Employee Name','Override Trade Class',
    'Post To Payroll','Cost Code / Phase','JobArea','Scope Change','Pay Code','Hours',
    'Night Shift','Premium Rate / Subsistence Rate / Travel Rate','Comments'
]

def _pad3(x):
    s = str(x or "").strip()
    return f"{int(s):03d}" if s.isdigit() else s

def _map_hours_rows(df_job: pd.DataFrame, paycode_map: dict) -> pd.DataFrame:
    out = []
    for _, r in df_job.iterrows():
        reg_h = float(r.get("rt_hours", 0) or 0.0)
        ot_h  = float(r.get("ot_hours", 0) or 0.0)
        base = {
            "Date": pd.to_datetime(r.get("date")).strftime("%Y-%m-%d"),
            "Time Record Type": "",
            "Person Number": r.get("employee_number",""),
            "Employee Name": r.get("name",""),
            "Override Trade Class": r.get("trade_class",""),
            "Post To Payroll": "Y",
            "Cost Code / Phase": r.get("class_type",""),
            "JobArea": _pad3(r.get("job_area","")),
            "Scope Change": "",
            "Pay Code": "",
            "Hours": 0.0,
            "Night Shift": "",
            "Premium Rate / Subsistence Rate / Travel Rate": r.get("premium_rate",""),
            "Comments": r.get("comments","") or "",
        }
        if reg_h > 0:
            t = base.copy(); t["Pay Code"] = paycode_map.get("REG","211"); t["Hours"] = reg_h; out.append(t)
        if ot_h > 0:
            t = base.copy(); t["Pay Code"] = paycode_map.get("OT","212");  t["Hours"] = ot_h; out.append(t)
    return pd.DataFrame(out, columns=EXPORT_COLUMNS) if out else pd.DataFrame(columns=EXPORT_COLUMNS)

def build_per_job_exports(df_all: pd.DataFrame, export_date, paycode_map: dict) -> dict[str, bytes]:
    """
    Returns { filename: xlsx_bytes } for each Job Number on that date.
    Filenames EXACTLY: "mm-dd-yyyy - {Job Number} - Daily Time Import.xlsx"
    """
    if df_all.empty:
        return {}
    df_all = df_all.copy()
    df_all["job_number"] = df_all["job_number"].astype(str)

    files: dict[str, bytes] = {}
    for job in sorted(df_all["job_number"].unique().tolist()):
        sub = df_all[df_all["job_number"] == job].copy()
        out_df = _map_hours_rows(sub, paycode_map)
        if out_df.empty:
            continue

        # Write to Excel in memory (sheet name EXACTLY "TimeEntries")
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            out_df.to_excel(writer, sheet_name="TimeEntries", index=False)
        buf.seek(0)

        fname = f"{export_date.strftime('%m-%d-%Y')} - {job} - Daily Time Import.xlsx"
        files[fname] = buf.getvalue()
    return files
