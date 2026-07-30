"""
instagram/downloader.py
=======================
Tải toàn bộ Reels từ 1 hoặc nhiều kênh Instagram bằng yt-dlp.
Video được upload lên OneDrive vào thư mục /Instagram-Pool/<username>/

Cách dùng:
    python instagram/downloader.py --username <ig_username> --max 50
    python instagram/downloader.py --username daquan meme --max 30

Môi trường:
    ONEDRIVE_CLIENT_ID, ONEDRIVE_CLIENT_SECRET, ONEDRIVE_TENANT_ID,
    ONEDRIVE_USER_EMAIL  (xem onedrive_uploader.py)

Lưu ý:
    - Instagram Reels public không cần login với yt-dlp
    - Rate limit: ~50-100 video/giờ để không bị block
    - Chỉ tải video portrait (9:16) hoặc tự convert
"""

import os
import sys
import json
import time
import argparse
import subprocess
import tempfile

# Đảm bảo import được từ thư mục gốc repo
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from onedrive_uploader import upload_file

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
OUTPUT_DIR     = os.path.join("instagram", "downloads")
ONEDRIVE_ROOT  = "/Instagram-Pool"   # Thư mục gốc trên OneDrive
POOL_INDEX     = os.path.join("instagram", "pool_index.json")  # Track đã tải
SLEEP_BETWEEN  = 3       # Giây chờ giữa các lần download (tránh block)
MAX_DEFAULT    = 30      # Số video tối đa mỗi lần chạy


# ---------------------------------------------------------------------------
# POOL INDEX — track video đã tải để không tải lại
# ---------------------------------------------------------------------------
def load_index() -> dict:
    if os.path.exists(POOL_INDEX):
        with open(POOL_INDEX, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}   # {video_id: {"username": ..., "onedrive_url": ..., "used": false}}


def save_index(index: dict):
    os.makedirs(os.path.dirname(POOL_INDEX), exist_ok=True)
    with open(POOL_INDEX, "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# FETCH VIDEO LIST (không download ngay)
# ---------------------------------------------------------------------------
def get_reel_urls(username: str, max_count: int = MAX_DEFAULT) -> list[dict]:
    """
    Lấy danh sách URL Reels từ kênh Instagram bằng yt-dlp --flat-playlist.
    Trả về list [{"id": ..., "url": ..., "title": ...}]
    """
    profile_url = f"https://www.instagram.com/{username}/reels/"
    print(f"\n[Fetch] Lấy danh sách Reels của @{username} (tối đa {max_count})...")

    cmd = [
        "yt-dlp",
        "--flat-playlist",
        "--playlist-end", str(max_count),
        "--print", "%(id)s\t%(webpage_url)s\t%(title)s",
        "--no-warnings",
        "--quiet",
        profile_url,
    ]

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        lines = [l.strip() for l in result.stdout.strip().splitlines() if "\t" in l]
        videos = []
        for line in lines:
            parts = line.split("\t", 2)
            videos.append({
                "id": parts[0],
                "url": parts[1] if len(parts) > 1 else "",
                "title": parts[2] if len(parts) > 2 else "",
                "username": username,
            })
        print(f"[Fetch] Tìm thấy {len(videos)} Reels")
        return videos
    except subprocess.TimeoutExpired:
        print("[Fetch] Timeout khi lấy danh sách.")
        return []
    except Exception as e:
        print(f"[Fetch] Lỗi: {e}")
        return []


# ---------------------------------------------------------------------------
# DOWNLOAD 1 VIDEO
# ---------------------------------------------------------------------------
def download_reel(video_url: str, output_path: str) -> bool:
    """Download 1 Reel bằng yt-dlp, lưu vào output_path."""
    cmd = [
        "yt-dlp",
        "--no-warnings",
        "--quiet",
        "-f", "bestvideo[height<=1920][ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best",
        "--merge-output-format", "mp4",
        "--max-filesize", "150M",
        "--max-downloads", "1",
        "-o", output_path,
        video_url,
    ]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
        if r.returncode == 0 and os.path.exists(output_path):
            size_mb = os.path.getsize(output_path) / 1024 / 1024
            print(f"  ✅ {os.path.basename(output_path)} ({size_mb:.1f} MB)")
            return True
        else:
            print(f"  ❌ yt-dlp lỗi: {r.stderr[:200]}")
            return False
    except subprocess.TimeoutExpired:
        print("  ❌ Timeout download.")
        return False
    except Exception as e:
        print(f"  ❌ Exception: {e}")
        return False


# ---------------------------------------------------------------------------
# CONVERT TO 9:16 PORTRAIT (nếu video là landscape)
# ---------------------------------------------------------------------------
def ensure_portrait(input_path: str, output_path: str, max_seconds: int = 60) -> bool:
    """Dùng FFmpeg crop/scale về 1080x1920, cắt tối đa max_seconds."""
    cmd = [
        "ffmpeg", "-y",
        "-i", input_path,
        "-t", str(max_seconds),
        "-vf", (
            "scale=iw*max(1080/iw\\,1920/ih):ih*max(1080/iw\\,1920/ih),"
            "crop=1080:1920"
        ),
        "-c:v", "libx264", "-preset", "fast", "-crf", "23",
        "-c:a", "aac", "-b:a", "128k",
        "-movflags", "+faststart",
        output_path,
    ]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if r.returncode == 0 and os.path.exists(output_path):
        return True
    print(f"  ⚠️ FFmpeg convert lỗi: {r.stderr[-200:]}")
    return False


# ---------------------------------------------------------------------------
# MAIN PIPELINE
# ---------------------------------------------------------------------------
def process_channel(username: str, max_count: int = MAX_DEFAULT,
                    skip_onedrive: bool = False) -> int:
    """
    Tải Reels từ @username, convert portrait, upload OneDrive, cập nhật index.
    Trả về số video đã xử lý thành công.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    index = load_index()
    videos = get_reel_urls(username, max_count)

    if not videos:
        print(f"[Pipeline] Không lấy được video nào từ @{username}.")
        return 0

    success_count = 0

    for i, v in enumerate(videos, 1):
        vid_id = v["id"]
        vid_url = v["url"]
        title_safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in v["title"])[:40]
        filename = f"{username}_{vid_id}_{title_safe}.mp4"

        # Bỏ qua nếu đã có trong index
        if vid_id in index:
            print(f"  [Skip] #{i} {vid_id} — đã tải trước đó.")
            continue

        print(f"\n[{i}/{len(videos)}] Xử lý: {vid_url}")

        raw_path  = os.path.join(OUTPUT_DIR, f"raw_{vid_id}.mp4")
        port_path = os.path.join(OUTPUT_DIR, filename)

        # 1. Download
        if not download_reel(vid_url, raw_path):
            index[vid_id] = {"username": username, "status": "download_failed", "used": False}
            save_index(index)
            time.sleep(SLEEP_BETWEEN)
            continue

        # 2. Convert portrait
        converted = ensure_portrait(raw_path, port_path)
        if not converted:
            # Dùng raw nếu convert lỗi
            os.rename(raw_path, port_path)
        else:
            try:
                os.remove(raw_path)
            except:
                pass

        # 3. Upload OneDrive
        onedrive_url = None
        if not skip_onedrive:
            remote_folder = f"{ONEDRIVE_ROOT}/{username}"
            onedrive_url = upload_file(port_path, remote_filename=filename, folder=remote_folder)
            if onedrive_url:
                # Xóa local sau khi upload thành công (tiết kiệm GitHub storage)
                try:
                    os.remove(port_path)
                except:
                    pass

        # 4. Cập nhật index
        index[vid_id] = {
            "username": username,
            "url": vid_url,
            "filename": filename,
            "onedrive_url": onedrive_url or "",
            "local_path": "" if onedrive_url else port_path,
            "status": "done",
            "used": False,      # False = chưa dùng làm nền Shorts
        }
        save_index(index)
        success_count += 1

        # Rate limit: ngủ giữa các lần download
        time.sleep(SLEEP_BETWEEN)

    print(f"\n[Pipeline] ✅ Hoàn tất @{username}: {success_count}/{len(videos)} video")
    return success_count


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Instagram Reels Channel Downloader")
    parser.add_argument("--username", nargs="+", required=True,
                        help="Một hoặc nhiều username Instagram (không có @)")
    parser.add_argument("--max", type=int, default=MAX_DEFAULT,
                        help=f"Số video tối đa mỗi kênh (mặc định {MAX_DEFAULT})")
    parser.add_argument("--no-onedrive", action="store_true",
                        help="Chỉ download local, không upload OneDrive")
    args = parser.parse_args()

    total = 0
    for uname in args.username:
        uname = uname.lstrip("@")
        total += process_channel(uname, args.max, skip_onedrive=args.no_onedrive)

    print(f"\n🎉 Tổng cộng đã xử lý: {total} video")
    if total == 0:
        sys.exit(1)
