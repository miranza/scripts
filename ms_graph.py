import logging

import requests
from msal import ConfidentialClientApplication

from settings import (
    APP_AUTHORITY,
    APP_CLIENT_ID,
    APP_CLIENT_SECRET,
    APP_SCOPE,
    BC_AUTHORITY,
    BC_CLIENT_ID,
    BC_CLIENT_SECRET,
    BC_SCOPE,
    SHAREPOINT_BASE_URL,
    SHAREPOINT_DRIVE_NAME,
    SHAREPOINT_FOLDER_PATH,
    SHAREPOINT_HOST,
    SHAREPOINT_SITE_NAME,
)

log = logging.getLogger(__name__)


def get_bc_access_token() -> str:
    log.info("Requesting BC access token")
    app = ConfidentialClientApplication(
        BC_CLIENT_ID, authority=BC_AUTHORITY, client_credential=BC_CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=BC_SCOPE)
    if "access_token" not in result:
        raise ValueError(f"BCAuth failed: {result.get('error_description', result)}")
    log.info("BC token obtained")
    return f"Bearer {result['access_token']}"


def get_app_token() -> str:
    log.info("Requesting Graph app token")
    app = ConfidentialClientApplication(
        APP_CLIENT_ID,
        authority=APP_AUTHORITY,
        client_credential=APP_CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=APP_SCOPE)
    if "access_token" not in result:
        raise ValueError(f"AppAuth failed: {result.get('error_description', result)}")
    log.info("Graph token obtained")
    return result["access_token"]


def graph_get(token, path):
    resp = requests.get(
        f"{SHAREPOINT_BASE_URL}{path}",
        headers={"Authorization": f"Bearer {token}"},
    )
    return resp


def get_site_id(token):
    r = graph_get(token, f"/sites/{SHAREPOINT_HOST}:/sites/{SHAREPOINT_SITE_NAME}")
    r.raise_for_status()
    site_id = r.json()["id"]
    log.info(f"Site ID: {site_id}")
    return site_id


def get_drive_id(token, site_id):
    r = graph_get(token, f"/sites/{site_id}/drives")
    r.raise_for_status()
    drives = r.json().get("value", [])
    for d in drives:
        log.info(f"Drive encontrado: {d['name']} - {d['id']}")
        if d["name"] == SHAREPOINT_DRIVE_NAME:
            return d["id"]

    raise ValueError(f"Drive id not found '{SHAREPOINT_DRIVE_NAME}'")


def sync_sharepoint(filepath):
    filename = filepath.name
    log.info(f"Uploading {filename} to SharePoint")
    token = get_app_token()

    site_id = get_site_id(token)
    drive_id = get_drive_id(token, site_id)

    with open(filepath, "rb") as f:
        data = f.read()

    url = (
        f"{SHAREPOINT_BASE_URL}"
        f"/drives/{drive_id}/root:/{SHAREPOINT_FOLDER_PATH}/{filename}:/content"
    )
    resp = requests.put(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream",
        },
        data=data,
    )
    resp.raise_for_status()
    log.info(f"Uploaded to SharePoint: {filename}")
