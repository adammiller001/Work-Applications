# supabase_helpers.py (patched to avoid header type error on upload)
from __future__ import annotations
import os, datetime as dt
from typing import List, Dict, Optional
from supabase import create_client, Client

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

def sb_client() -> Client:
    url = os.environ["SUPABASE_URL"].strip()
    key = os.environ["SUPABASE_SERVICE_KEY"].strip()
    return create_client(url, key)

def _iso_date(v):
    if v is None: return None
    if isinstance(v, (dt.date, dt.datetime)):
        return v.date().isoformat() if isinstance(v, dt.datetime) else v.isoformat()
    return str(v)

def add_time_rows(rows: List[Dict], created_by: str) -> int:
    if not rows: return 0
    rows2 = []
    for r in rows:
        c = dict(r)
        c["created_by"] = created_by
        if "date" in c: c["date"] = _iso_date(c["date"])
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
    """
    Upload bytes to Supabase Storage and return a signed URL.
    Uses string values for headers/options to avoid httpx header type errors.
    """
    if not isinstance(content, (bytes, bytearray)):
        raise TypeError("content must be bytes")
    path = str(path).lstrip("/")  # ensure string path
    bucket = os.environ.get("SUPABASE_BUCKET", "exports")

    sb = sb_client()
    # Ensure file options are simple strings
    file_options = {
        "contentType": str(MIME_XLSX),
        "upsert": "true",            # string to avoid header value type errors
        "cacheControl": "3600",      # optional, as string
    }
    sb.storage.from_(bucket).upload(path, content, file_options)
    signed = sb.storage.from_(bucket).create_signed_url(path, int(expires_seconds))
    return signed.get("signedURL") or signed.get("signed_url") or ""
