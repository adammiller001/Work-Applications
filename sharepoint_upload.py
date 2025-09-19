# sharepoint_upload.py
# Upload an in-memory Excel (bytes) to SharePoint via Microsoft Graph
# Requires environment variables (set in Streamlit Secrets):
#   TENANT_ID, CLIENT_ID, CLIENT_SECRET, SP_SITE, SP_DRIVE (e.g. "Documents"), SP_EXPORT_FOLDER (e.g. "Exports")
#
# Usage from your Streamlit app after you build an export into `output` (BytesIO):
#   from sharepoint_upload import upload_export_to_sharepoint
#   link = upload_export_to_sharepoint("Exports/September/file.xlsx", output.getvalue())

import os
from office365.graph_client import GraphClient
from office365.runtime.auth.client_credential import ClientCredential

def _graph() -> GraphClient:
    tenant = os.environ["TENANT_ID"]
    client_id = os.environ["CLIENT_ID"]
    client_secret = os.environ["CLIENT_SECRET"]
    return GraphClient(ClientCredential(client_id, client_secret), tenant=tenant)

def _drive_root(client: GraphClient):
    site_url = os.environ["SP_SITE"]  # e.g. https://yourtenant.sharepoint.com/sites/YourSite
    drive_name = os.environ.get("SP_DRIVE", "Documents")
    site = client.sites.get_by_url(site_url).get().execute_query()
    return site.drives[drive_name].root

def upload_export_to_sharepoint(path_in_library: str, data: bytes, link_scope: str = "organization") -> str:
    """
    path_in_library: e.g. "Exports/September/09-19-2025 - ALL_JOBS - Daily Time Import.xlsx"
    link_scope: "anonymous" for anyone with the link, or "organization" for org-only.
    Returns: a view link URL.
    """
    client = _graph()
    root = _drive_root(client)
    # Upload (creates intermediate folders as needed)
    item = root.upload_file(path_in_library, data).execute_query()
    # Create a view link
    link = item.create_link(type="view", scope=link_scope).execute_query()
    return link.link.web_url
