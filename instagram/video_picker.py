"""
instagram/video_picker.py
=========================
Chọn ngẫu nhiên 1 video từ kho OneDrive pool để dùng làm nền Shorts.
Đánh dấu video là "used = true" trong pool_index.json sau khi chọn.

Cách dùng trong main.py:
    from instagram.video_picker import pick_pool_video
    local_path = pick_pool_video()  # Tải từ OneDrive về local, trả về path
"""

import os
import sys
import json
import random
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from onedrive_uploader import get_access_token

POOL_INDEX = os.path.join("instagram", "pool_index.json")
DOWNLOAD_DIR = os.path.join("instagram", "downloads")
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def load_index() -> dict:
    if os.path.exists(POOL_INDEX):
        with open(POOL_INDEX, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_index(index: dict):
    with open(POOL_INDEX, "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)


def pick_pool_video(prefer_username: str = None) -> str | None:
    """
    Chọn ngẫu nhiên 1 video chưa dùng từ pool.
    - prefer_username: nếu có, ưu tiên video từ kênh đó
    Trả về local path sau khi tải về, hoặc None nếu không có video khả dụng.
    """
    index = load_index()

    # Lọc video chưa dùng và đã upload OneDrive
    available = [
        (vid_id, info)
        for vid_id, info in index.items()
        if not info.get("used", False)
        and info.get("status") == "done"
        and info.get("onedrive_url")
    ]

    if not available:
        print("[Picker] ❌ Không còn video nào trong pool chưa dùng.")
        return None

    # Ưu tiên username nếu có
    if prefer_username:
        preferred = [(i, v) for i, v in available if v.get("username") == prefer_username]
        if preferred:
            available = preferred

    # Chọn ngẫu nhiên
    vid_id, info = random.choice(available)
    filename = info.get("filename", f"{vid_id}.mp4")
    onedrive_url = info["onedrive_url"]

    print(f"[Picker] Chọn video: {filename} từ @{info.get('username', '?')}")

    # Tải về local
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    local_path = os.path.join(DOWNLOAD_DIR, filename)

    if os.path.exists(local_path):
        print(f"[Picker] File đã có local: {local_path}")
    else:
        print(f"[Picker] Đang tải từ OneDrive...")
        downloaded = _download_from_onedrive_url(onedrive_url, local_path)
        if not downloaded:
            print("[Picker] ❌ Tải thất bại.")
            return None

    # Đánh dấu đã dùng
    index[vid_id]["used"] = True
    save_index(index)
    print(f"[Picker] ✅ Đã đánh dấu '{filename}' là đã dùng.")

    return local_path


def _download_from_onedrive_url(share_url: str, output_path: str) -> bool:
    """
    Tải file từ OneDrive share URL.
    Share URL dạng: https://onedrive.live.com/...
    """
    try:
        # OneDrive share link → direct download bằng cách thêm ?download=1
        download_url = share_url
        if "1drv.ms" in share_url or "onedrive.live.com" in share_url:
            # Convert share link sang direct download
            download_url = share_url.replace("?", "?download=1&") if "?" in share_url \
                else share_url + "?download=1"

        r = requests.get(download_url, stream=True, timeout=120, allow_redirects=True)
        r.raise_for_status()

        with open(output_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)

        size_mb = os.path.getsize(output_path) / 1024 / 1024
        print(f"[Picker] Tải xong: {size_mb:.1f} MB")
        return True
    except Exception as e:
        print(f"[Picker] Lỗi tải: {e}")
        return False


def pool_stats() -> dict:
    """Trả về thống kê pool hiện tại."""
    index = load_index()
    total = len(index)
    unused = sum(1 for v in index.values() if not v.get("used") and v.get("status") == "done")
    used = sum(1 for v in index.values() if v.get("used"))
    failed = sum(1 for v in index.values() if v.get("status") != "done")

    by_channel = {}
    for v in index.values():
        u = v.get("username", "unknown")
        by_channel[u] = by_channel.get(u, 0) + 1

    return {
        "total": total,
        "unused": unused,
        "used": used,
        "failed": failed,
        "by_channel": by_channel,
    }


if __name__ == "__main__":
    stats = pool_stats()
    print("📊 Pool Statistics:")
    print(f"  Tổng:        {stats['total']}")
    print(f"  Chưa dùng:  {stats['unused']}")
    print(f"  Đã dùng:    {stats['used']}")
    print(f"  Thất bại:   {stats['failed']}")
    print(f"  Theo kênh:  {stats['by_channel']}")
