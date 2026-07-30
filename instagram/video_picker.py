"""
instagram/video_picker.py
=========================
Chọn ngẫu nhiên 1-2 video từ link_pool.json, download về local để dùng làm nền Shorts.
Không cần OneDrive — tải thẳng từ TikTok/Instagram lúc cần.

Cách dùng trong main.py:
    from instagram.video_picker import pick_and_download
    video_path = pick_and_download(prefer_hashtag="#diy")
"""

import os
import sys
import json
import random
import subprocess
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

POOL_FILE    = os.path.join("instagram", "link_pool.json")
DOWNLOAD_DIR = os.path.join("output", "bg_videos")   # Lưu tạm, xóa sau khi dùng
MAX_DURATION = 65   # Giây tối đa cho video nền


# ─────────────────────────────────────────────────────────────────────────────
# POOL I/O
# ─────────────────────────────────────────────────────────────────────────────

def load_pool() -> dict:
    if os.path.exists(POOL_FILE):
        with open(POOL_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_pool(pool: dict):
    with open(POOL_FILE, "w", encoding="utf-8") as f:
        json.dump(pool, f, ensure_ascii=False, separators=(",", ":"))


# ─────────────────────────────────────────────────────────────────────────────
# PICK RANDOM URL
# ─────────────────────────────────────────────────────────────────────────────

def pick_url(prefer_hashtag: str = None, prefer_platform: str = None) -> tuple[str, str] | tuple[None, None]:
    """
    Chọn ngẫu nhiên 1 video chưa dùng và chưa fail từ pool.
    Trả về (video_id, url) hoặc (None, None).
    """
    pool = load_pool()

    candidates = [
        (vid_id, info)
        for vid_id, info in pool.items()
        if not info.get("used") and not info.get("failed")
    ]

    if not candidates:
        print("[Picker] ❌ Pool rỗng hoặc tất cả đã được dùng.")
        return None, None

    # Ưu tiên hashtag / platform nếu có
    filtered = candidates
    if prefer_hashtag:
        tag = prefer_hashtag if prefer_hashtag.startswith("#") else f"#{prefer_hashtag}"
        subset = [(i, v) for i, v in candidates if v.get("hashtag") == tag]
        if subset:
            filtered = subset
    if prefer_platform:
        subset = [(i, v) for i, v in filtered if v.get("platform") == prefer_platform]
        if subset:
            filtered = subset

    vid_id, info = random.choice(filtered)
    print(f"[Picker] Chọn: {info.get('hashtag','?')} | {info.get('platform','?')} | {vid_id}")
    return vid_id, info["url"]


# ─────────────────────────────────────────────────────────────────────────────
# DOWNLOAD VIDEO
# ─────────────────────────────────────────────────────────────────────────────

def download_video(vid_id: str, url: str) -> str | None:
    """
    Download video từ URL, convert về 9:16, trả về local path hoặc None.
    """
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    raw_path    = os.path.join(DOWNLOAD_DIR, f"raw_{vid_id}.mp4")
    output_path = os.path.join(DOWNLOAD_DIR, f"bg_{vid_id}.mp4")

    # Download
    cmd = [
        "yt-dlp",
        "--no-warnings", "--quiet",
        "-f", "bestvideo[height<=1920][ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best",
        "--merge-output-format", "mp4",
        "--max-filesize", "200M",
        "--max-downloads", "1",
        "-o", raw_path,
        url,
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if result.returncode != 0 or not os.path.exists(raw_path):
            print(f"[Picker] ❌ Download lỗi: {result.stderr[:200]}")
            return None
    except subprocess.TimeoutExpired:
        print("[Picker] ❌ Download timeout.")
        return None
    except Exception as e:
        print(f"[Picker] ❌ {e}")
        return None

    # Convert về 9:16 portrait
    if _convert_portrait(raw_path, output_path):
        try: os.remove(raw_path)
        except: pass
        size_mb = os.path.getsize(output_path) / 1024 / 1024
        print(f"[Picker] ✅ Video sẵn sàng: {output_path} ({size_mb:.1f} MB)")
        return output_path
    else:
        # Trả về raw nếu convert lỗi
        print("[Picker] ⚠️ Dùng file raw (convert thất bại).")
        return raw_path


def _convert_portrait(input_path: str, output_path: str) -> bool:
    """FFmpeg: scale + crop về 1080x1920, cắt tối đa MAX_DURATION giây."""
    cmd = [
        "ffmpeg", "-y",
        "-i", input_path,
        "-t", str(MAX_DURATION),
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
    return r.returncode == 0 and os.path.exists(output_path)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN FUNCTION (gọi từ main.py)
# ─────────────────────────────────────────────────────────────────────────────

def pick_and_download(prefer_hashtag: str = None,
                      prefer_platform: str = "tiktok",
                      max_retries: int = 5) -> str | None:
    """
    Chọn ngẫu nhiên + download video từ pool.
    Tự retry nếu download thất bại (đánh dấu failed và chọn cái khác).
    Trả về local path hoặc None.
    """
    pool = load_pool()

    for attempt in range(max_retries):
        vid_id, url = pick_url(prefer_hashtag, prefer_platform)
        if not vid_id:
            return None

        print(f"[Picker] Thử lần {attempt+1}/{max_retries}: {url[:60]}...")
        local_path = download_video(vid_id, url)

        if local_path:
            # Đánh dấu đã dùng
            if vid_id in pool:
                pool[vid_id]["used"] = True
                pool[vid_id]["used_at"] = datetime.now().strftime("%Y-%m-%d")
            save_pool(pool)
            return local_path
        else:
            # Đánh dấu failed, thử cái khác
            if vid_id in pool:
                pool[vid_id]["failed"] = True
            save_pool(pool)
            pool = load_pool()  # Reload để tránh chọn lại

    print(f"[Picker] ❌ Thất bại sau {max_retries} lần thử.")
    return None


def pool_stats() -> dict:
    """Thống kê pool."""
    pool = load_pool()
    total  = len(pool)
    unused = sum(1 for v in pool.values() if not v.get("used") and not v.get("failed"))
    used   = sum(1 for v in pool.values() if v.get("used"))
    failed = sum(1 for v in pool.values() if v.get("failed"))
    by_tag = {}
    for v in pool.values():
        tag = v.get("hashtag", "?")
        by_tag[tag] = by_tag.get(tag, 0) + 1
    return {"total": total, "unused": unused, "used": used, "failed": failed, "by_hashtag": by_tag}


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    stats = pool_stats()
    print("📊 Link Pool Stats:")
    print(f"  Tổng link:     {stats['total']:,}")
    print(f"  Chưa dùng:    {stats['unused']:,}")
    print(f"  Đã dùng:      {stats['used']:,}")
    print(f"  Failed:       {stats['failed']:,}")
    print(f"  Theo hashtag: {stats['by_hashtag']}")

    remaining_days = stats['unused'] // 3 if stats['unused'] > 0 else 0
    print(f"\n  ⏳ Đủ dùng cho ~{remaining_days} ngày (3 video/ngày)")
