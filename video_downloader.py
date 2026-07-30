"""
video_downloader.py
===================
Download video từ TikTok / Instagram / YouTube Shorts bằng yt-dlp.
Fallback: Pexels API (royalty-free, hợp pháp 100%).

Cách dùng:
    python video_downloader.py --url "https://www.tiktok.com/@user/video/123"
    python video_downloader.py --pexels "funny cat" --duration 60
"""

import os
import re
import json
import argparse
import subprocess
import requests
import tempfile

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
OUTPUT_DIR = "output"
PEXELS_API_KEY = os.getenv("PEXELS_API_KEY", "")
TARGET_W, TARGET_H = 1080, 1920   # 9:16 Shorts format


# ---------------------------------------------------------------------------
# YT-DLP DOWNLOADER
# ---------------------------------------------------------------------------
def download_with_ytdlp(url: str, output_path: str, max_duration: int = 65) -> bool:
    """
    Download video từ TikTok/Instagram/YouTube Shorts bằng yt-dlp.
    Chỉ download nếu video <= max_duration giây.
    Trả về True nếu thành công.
    """
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    # Bước 1: Check duration trước khi download
    try:
        result = subprocess.run(
            ["yt-dlp", "--print", "duration", "--no-warnings", url],
            capture_output=True, text=True, timeout=30
        )
        duration_str = result.stdout.strip()
        if duration_str and duration_str.replace(".", "").isdigit():
            duration = float(duration_str)
            if duration > max_duration:
                print(f"[yt-dlp] Video quá dài ({duration:.0f}s > {max_duration}s), bỏ qua.")
                return False
    except Exception as e:
        print(f"[yt-dlp] Không check được duration: {e}, tiếp tục download...")

    # Bước 2: Download với format tốt nhất (ưu tiên 9:16 portrait)
    cmd = [
        "yt-dlp",
        "--no-warnings",
        "--quiet",
        "-f", "bestvideo[height<=1920][ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best",
        "--merge-output-format", "mp4",
        "--max-filesize", "100M",
        "-o", output_path,
        "--no-playlist",
        url,
    ]

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if result.returncode == 0 and os.path.exists(output_path):
            size_mb = os.path.getsize(output_path) / (1024 * 1024)
            print(f"[yt-dlp] ✅ Download thành công: {output_path} ({size_mb:.1f} MB)")
            return True
        else:
            print(f"[yt-dlp] ❌ Lỗi: {result.stderr[:300]}")
            return False
    except subprocess.TimeoutExpired:
        print("[yt-dlp] ❌ Timeout khi download.")
        return False
    except FileNotFoundError:
        print("[yt-dlp] ❌ yt-dlp không được cài đặt. Chạy: pip install yt-dlp")
        return False


# ---------------------------------------------------------------------------
# FFMPEG: Convert & crop về 9:16
# ---------------------------------------------------------------------------
def convert_to_portrait(input_path: str, output_path: str,
                         max_duration: int = 58) -> bool:
    """
    Convert video thành format 9:16 (1080x1920), cắt tối đa max_duration giây.
    Dùng FFmpeg crop + scale thông minh.
    """
    cmd = [
        "ffmpeg", "-y",
        "-i", input_path,
        "-t", str(max_duration),
        # Scale để chiều nhỏ hơn bằng 1080 hoặc 1920, rồi crop center
        "-vf", (
            "scale=iw*max(1080/iw\\,1920/ih):ih*max(1080/iw\\,1920/ih),"
            "crop=1080:1920"
        ),
        "-c:v", "libx264",
        "-preset", "fast",
        "-crf", "23",
        "-c:a", "aac",
        "-b:a", "128k",
        "-movflags", "+faststart",
        output_path,
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if result.returncode == 0 and os.path.exists(output_path):
            print(f"[FFmpeg] ✅ Convert thành công: {output_path}")
            return True
        else:
            print(f"[FFmpeg] ❌ Lỗi: {result.stderr[-300:]}")
            return False
    except Exception as e:
        print(f"[FFmpeg] ❌ Exception: {e}")
        return False


# ---------------------------------------------------------------------------
# PEXELS FALLBACK
# ---------------------------------------------------------------------------
def download_from_pexels(query: str, output_path: str,
                          min_duration: int = 15) -> bool:
    """
    Tìm và download video royalty-free từ Pexels API.
    Cần PEXELS_API_KEY trong environment.
    """
    if not PEXELS_API_KEY:
        print("[Pexels] ❌ PEXELS_API_KEY chưa được set.")
        return False

    headers = {"Authorization": PEXELS_API_KEY}
    search_url = f"https://api.pexels.com/videos/search?query={query}&per_page=15&orientation=portrait"

    try:
        resp = requests.get(search_url, headers=headers, timeout=15)
        resp.raise_for_status()
        videos = resp.json().get("videos", [])
    except Exception as e:
        print(f"[Pexels] ❌ Lỗi search: {e}")
        return False

    # Lọc video đủ dài, ưu tiên portrait
    candidates = []
    for v in videos:
        dur = v.get("duration", 0)
        if dur < min_duration:
            continue
        for file in v.get("video_files", []):
            if file.get("quality") in ["hd", "sd"] and file.get("width", 0) < file.get("height", 1):
                candidates.append({
                    "url": file["link"],
                    "w": file["width"],
                    "h": file["height"],
                    "duration": dur,
                })
                break

    if not candidates:
        # Thử landscape nếu không có portrait
        for v in videos:
            if v.get("duration", 0) >= min_duration:
                for file in v.get("video_files", []):
                    if file.get("quality") == "hd":
                        candidates.append({
                            "url": file["link"],
                            "w": file["width"],
                            "h": file["height"],
                            "duration": v["duration"],
                        })
                        break

    if not candidates:
        print(f"[Pexels] ❌ Không tìm thấy video cho query: '{query}'")
        return False

    # Download video đầu tiên tìm được
    video_url = candidates[0]["url"]
    print(f"[Pexels] Downloading: {video_url[:80]}...")
    try:
        with requests.get(video_url, stream=True, timeout=60) as r:
            r.raise_for_status()
            with open(output_path, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
        print(f"[Pexels] ✅ Download thành công: {output_path}")
        return True
    except Exception as e:
        print(f"[Pexels] ❌ Lỗi download: {e}")
        return False


# ---------------------------------------------------------------------------
# MAIN FUNCTION
# ---------------------------------------------------------------------------
def get_source_video(source_url: str = None, pexels_query: str = "funny moments",
                     output_path: str = None) -> str | None:
    """
    Hàm chính: thử download từ URL nguồn, nếu fail thì fallback Pexels.
    Trả về path file video đã convert về 9:16, hoặc None nếu thất bại.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if output_path is None:
        output_path = os.path.join(OUTPUT_DIR, "source_video.mp4")

    portrait_path = output_path.replace(".mp4", "_portrait.mp4")

    # --- Thử yt-dlp ---
    if source_url:
        raw_path = output_path.replace(".mp4", "_raw.mp4")
        if download_with_ytdlp(source_url, raw_path):
            if convert_to_portrait(raw_path, portrait_path):
                try:
                    os.remove(raw_path)
                except:
                    pass
                return portrait_path
            else:
                # Dùng raw nếu convert lỗi
                return raw_path

    # --- Fallback Pexels ---
    print("[Downloader] Fallback sang Pexels...")
    raw_path = output_path.replace(".mp4", "_pexels_raw.mp4")
    if download_from_pexels(pexels_query, raw_path):
        if convert_to_portrait(raw_path, portrait_path):
            try:
                os.remove(raw_path)
            except:
                pass
            return portrait_path
        return raw_path

    print("[Downloader] ❌ Tất cả nguồn đều thất bại.")
    return None


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Video Downloader")
    parser.add_argument("--url", help="URL TikTok/Instagram/YouTube")
    parser.add_argument("--pexels", help="Pexels search query (fallback)")
    parser.add_argument("--output", default=os.path.join(OUTPUT_DIR, "source_video.mp4"))
    args = parser.parse_args()

    result = get_source_video(
        source_url=args.url,
        pexels_query=args.pexels or "funny moments",
        output_path=args.output,
    )

    if result:
        print(f"\n✅ Video sẵn sàng tại: {result}")
    else:
        print("\n❌ Không download được video.")
        exit(1)
