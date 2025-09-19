# supabase_helpers.py
# Utilities for Streamlit + Supabase (DB and Storage) — Python SDK v2
# Usage:
#   from supabase_helpers import sb_client, add_time_rows, fetch_time_entries_for_date, delete_by_ids, upload_export_bytes
#
# Env/Secrets required:
#   SUPABASE_URL, SUPABASE_SERVICE_KEY, SUPABASE_BUCKET (defaults to "exports")

from __future__ import annotations
import os, io, datetime as dt
from typing import List, Dict, Optional
from supabase import create_client, Client

def sb_client() -> Client:
    url = os.environ["SUPABASE_URL"]
    key = os.environ["SUPABASE_SERVICE_KEY"]  # service role for server-side use
    return create_client(url, key)

def _iso_date(v):
    if v is None: return None
    if isinstance(v, (dt.date, dt.datetime)):
        return v.date().isoformat() if isinstance(v, dt.datetime) else v.isoformat()
    return str(v)

def add_time_rows(rows: List[Dict], created_by: str) -> int:
    """Insert one or many time entries; returns number inserted."""
    if not rows: return 0
    rows2 = []
    for r in rows:
        c = dict(r)
        c["created_by"] = created_by
        if "date" in c: c["date"] = _iso_date(c["date"])
        # Normalize numeric hours
        for k in ("rt_hours","ot_hours"):
            if k in c and c[k] is not None:
                try: c[k] = float(c[k])
                except Exception: c[k] = 0.0
        rows2.append(c)
    sb = sb_client()
    res = sb.table("time_entries").insert(rows2).execute()
    return len(res.data or [])

def fetch_time_entries_for_date(date_str: str, job_number: Optional[str]=None) -> list[dict]:
    sb = sb_client()
    q = sb.table("time_entries").select("*").eq("date", date_str)
    if job_number:
        q = q.eq("job_number", str(job_number))
    out = q.order("created_at", desc=False).execute()
    return out.data or []

def delete_by_ids(ids: List[str]) -> int:
    if not ids: return 0
    sb = sb_client()
    res = sb.table("time_entries").delete().in_("id", ids).execute()
    return len(res.data or [])

def upload_export_bytes(content: bytes, path: str, expires_seconds: int = 86400) -> str:
    """Upload bytes to Storage and return a signed URL."""
    bucket = os.environ.get("SUPABASE_BUCKET", "exports")
    sb = sb_client()
    # upsert to allow overwriting same name for same day
    sb.storage.from_(bucket).upload(file=path, file_content=content, file_options={"contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "upsert": True})
    signed = sb.storage.from_(bucket).create_signed_url(path, expires_seconds)
    # SDK returns dict with 'signedURL' in recent versions
    return signed.get("signedURL") or signed.get("signed_url") or ""
