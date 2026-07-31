import os
import json
import random
import subprocess
from concurrent.futures import ThreadPoolExecutor

POOL_FILE = os.path.join("instagram", "link_pool.json")
OUTPUT_DIR = "local_videos"

def load_pool():
    if os.path.exists(POOL_FILE):
        with open(POOL_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def download_video(vid_id, item):
    url = item.get("url")
    title = item.get("title", "video").replace("/", "_").replace("\\", "_")[:50]
    out_path = os.path.join(OUTPUT_DIR, f"{vid_id}_{title}.mp4")
    
    if os.path.exists(out_path):
        return True, "Đã tồn tại"

    print(f"⏳ Đang tải: {vid_id}...")
    cmd = [
        "yt-dlp",
        "-f", "b[ext=mp4]/mp4",
        "--no-warnings",
        "--quiet",
        "-o", out_path,
        url
    ]
    
    try:
        res = subprocess.run(cmd, capture_output=True, timeout=60)
        if res.returncode == 0 and os.path.exists(out_path):
            print(f"✅ Đã tải xong: {vid_id}")
            return True, "Thành công"
        else:
            print(f"❌ Lỗi tải {vid_id}: {res.stderr.decode('utf-8', errors='ignore')}")
            return False, "Lỗi yt-dlp"
    except Exception as e:
        print(f"❌ Lỗi mạng {vid_id}: {e}")
        return False, str(e)

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    pool = load_pool()
    
    # Lọc ra các video chưa được sử dụng và dài 15-45s
    unused_videos = {k: v for k, v in pool.items() if not v.get("used") and not v.get("failed")}
    
    if not unused_videos:
        print("⚠️ Không có video nào khả dụng trong pool!")
        return

    # Sếp yêu cầu tải test, mình sẽ tải 20 video làm mẫu (tránh treo máy)
    LIMIT = 20
    items_to_download = list(unused_videos.items())
    random.shuffle(items_to_download)
    items_to_download = items_to_download[:LIMIT]
    
    print(f"🚀 Bắt đầu tiến trình tải {len(items_to_download)} video siêu tốc...")
    
    success_count = 0
    # Tải đa luồng 4 video cùng lúc
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {executor.submit(download_video, k, v): k for k, v in items_to_download}
        for future in futures:
            is_success, msg = future.result()
            if is_success:
                success_count += 1
                
    print(f"🎉 HOÀN TẤT! Đã tải thành công {success_count}/{len(items_to_download)} video.")
    print(f"👉 Hãy kiểm tra thư mục: {os.path.abspath(OUTPUT_DIR)}")

if __name__ == "__main__":
    main()
