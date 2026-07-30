"""
onedrive_uploader.py
====================
Upload video lên OneDrive qua Microsoft Graph API.
Dùng Client Credentials flow (không cần user login).

Setup:
1. Vào https://portal.azure.com → Azure Active Directory → App registrations
2. Tạo app mới → lấy Client ID, Tenant ID
3. Certificates & secrets → tạo Client Secret
4. API permissions → Microsoft Graph → Files.ReadWrite.All (Application)
5. Grant admin consent

Environment variables (hoặc GitHub Secrets):
    ONEDRIVE_CLIENT_ID
    ONEDRIVE_CLIENT_SECRET
    ONEDRIVE_TENANT_ID
    ONEDRIVE_USER_EMAIL  (email tài khoản OneDrive đích)
"""

import os
import math
import time
import msal
import requests


# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
CLIENT_ID = os.getenv("ONEDRIVE_CLIENT_ID", "")
CLIENT_SECRET = os.getenv("ONEDRIVE_CLIENT_SECRET", "")
TENANT_ID = os.getenv("ONEDRIVE_TENANT_ID", "common")  # "common" nếu dùng personal
USER_EMAIL = os.getenv("ONEDRIVE_USER_EMAIL", "me")     # email hoặc "me"

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
ONEDRIVE_FOLDER = "/YouTube-Shorts"   # Thư mục gốc trên OneDrive
CHUNK_SIZE = 5 * 1024 * 1024          # 5 MB per chunk (resumable upload)


# ---------------------------------------------------------------------------
# AUTH
# ---------------------------------------------------------------------------
_token_cache = {"token": None, "expires_at": 0}

def get_access_token() -> str | None:
    """
    Lấy Bearer token qua MSAL Client Credentials.
    Token được cache tự động cho đến khi hết hạn.
    """
    if not CLIENT_ID or not CLIENT_SECRET:
        print("[OneDrive] ❌ Chưa set ONEDRIVE_CLIENT_ID / ONEDRIVE_CLIENT_SECRET")
        return None

    now = time.time()
    if _token_cache["token"] and now < _token_cache["expires_at"] - 60:
        return _token_cache["token"]

    # Personal account → dùng device code flow hoặc delegated
    # Business account → dùng client credentials
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET,
    )
    scopes = ["https://graph.microsoft.com/.default"]
    result = app.acquire_token_for_client(scopes=scopes)

    if "access_token" in result:
        _token_cache["token"] = result["access_token"]
        _token_cache["expires_at"] = now + result.get("expires_in", 3600)
        print("[OneDrive] ✅ Token OK")
        return result["access_token"]
    else:
        print(f"[OneDrive] ❌ Lỗi auth: {result.get('error_description', result)}")
        return None


# ---------------------------------------------------------------------------
# UPLOAD
# ---------------------------------------------------------------------------
def upload_file(local_path: str, remote_filename: str = None,
                folder: str = ONEDRIVE_FOLDER) -> str | None:
    """
    Upload file lên OneDrive.
    - File < 4MB: simple PUT
    - File >= 4MB: resumable upload session
    Trả về URL chia sẻ của file, hoặc None nếu thất bại.
    """
    token = get_access_token()
    if not token:
        return None

    if not os.path.exists(local_path):
        print(f"[OneDrive] ❌ File không tồn tại: {local_path}")
        return None

    file_size = os.path.getsize(local_path)
    if remote_filename is None:
        remote_filename = os.path.basename(local_path)

    headers = {"Authorization": f"Bearer {token}"}

    # Đường dẫn đích trên OneDrive
    if USER_EMAIL == "me":
        drive_path = f"{GRAPH_BASE}/me/drive/root:{folder}/{remote_filename}"
    else:
        drive_path = f"{GRAPH_BASE}/users/{USER_EMAIL}/drive/root:{folder}/{remote_filename}"

    print(f"[OneDrive] Uploading '{remote_filename}' ({file_size / 1024 / 1024:.1f} MB)...")

    if file_size < 4 * 1024 * 1024:
        return _simple_upload(local_path, drive_path, headers)
    else:
        return _resumable_upload(local_path, drive_path, headers, file_size)


def _simple_upload(local_path: str, drive_path: str, headers: dict) -> str | None:
    """Upload file nhỏ (<4MB) bằng PUT request đơn."""
    with open(local_path, "rb") as f:
        data = f.read()
    url = f"{drive_path}:/content"
    resp = requests.put(url, headers={**headers, "Content-Type": "video/mp4"}, data=data, timeout=60)
    if resp.status_code in (200, 201):
        item_id = resp.json().get("id", "")
        return _get_share_link(item_id, headers)
    else:
        print(f"[OneDrive] ❌ Simple upload lỗi {resp.status_code}: {resp.text[:200]}")
        return None


def _resumable_upload(local_path: str, drive_path: str,
                       headers: dict, file_size: int) -> str | None:
    """Upload file lớn (>=4MB) bằng resumable upload session."""
    # Tạo upload session
    session_url = f"{drive_path}:/createUploadSession"
    session_body = {"item": {"@microsoft.graph.conflictBehavior": "replace"}}
    resp = requests.post(session_url, headers=headers, json=session_body, timeout=30)
    if resp.status_code not in (200, 201):
        print(f"[OneDrive] ❌ Tạo upload session lỗi {resp.status_code}")
        return None

    upload_url = resp.json()["uploadUrl"]

    # Upload từng chunk
    with open(local_path, "rb") as f:
        offset = 0
        while offset < file_size:
            chunk = f.read(CHUNK_SIZE)
            chunk_size = len(chunk)
            end = offset + chunk_size - 1

            chunk_headers = {
                "Content-Length": str(chunk_size),
                "Content-Range": f"bytes {offset}-{end}/{file_size}",
                "Content-Type": "video/mp4",
            }
            r = requests.put(upload_url, headers=chunk_headers, data=chunk, timeout=120)
            pct = (offset + chunk_size) / file_size * 100
            print(f"[OneDrive] Upload... {pct:.0f}%")

            if r.status_code == 202:  # Accepted, tiếp tục
                offset += chunk_size
                continue
            elif r.status_code in (200, 201):  # Done
                item_id = r.json().get("id", "")
                print("[OneDrive] ✅ Upload hoàn tất!")
                return _get_share_link(item_id, headers)
            else:
                print(f"[OneDrive] ❌ Chunk upload lỗi {r.status_code}: {r.text[:200]}")
                return None

    return None


def _get_share_link(item_id: str, headers: dict) -> str | None:
    """Tạo share link công khai cho file đã upload."""
    if not item_id:
        return None
    if USER_EMAIL == "me":
        url = f"{GRAPH_BASE}/me/drive/items/{item_id}/createLink"
    else:
        url = f"{GRAPH_BASE}/users/{USER_EMAIL}/drive/items/{item_id}/createLink"

    body = {"type": "view", "scope": "anonymous"}
    resp = requests.post(url, headers=headers, json=body, timeout=30)
    if resp.status_code in (200, 201):
        link = resp.json().get("link", {}).get("webUrl", "")
        print(f"[OneDrive] 🔗 Share link: {link}")
        return link
    return None


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("file", help="Đường dẫn file cần upload")
    parser.add_argument("--folder", default=ONEDRIVE_FOLDER)
    parser.add_argument("--name", help="Tên file trên OneDrive")
    args = parser.parse_args()

    link = upload_file(args.file, remote_filename=args.name, folder=args.folder)
    if link:
        print(f"\n✅ File đã upload: {link}")
    else:
        print("\n❌ Upload thất bại.")
        exit(1)
