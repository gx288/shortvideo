import os
import json
import random
import shutil
import subprocess
from concurrent.futures import ThreadPoolExecutor

POOL_FILE = os.path.join("instagram", "link_pool.json")
OUTPUT_DIR_VERT = os.path.join("local_videos", "vertical")
OUTPUT_DIR_HORZ = os.path.join("local_videos", "horizontal")
TEMP_DIR = os.path.join("local_videos", "temp")

def load_pool():
    if os.path.exists(POOL_FILE):
        with open(POOL_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def check_orientation(video_path):
    """Sử dụng ffprobe để kiểm tra kích thước video."""
    cmd = [
        "ffprobe", "-v", "error", "-select_streams", "v:0",
        "-show_entries", "stream=width,height", "-of", "csv=p=0", video_path
    ]
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
        w, h = map(int, res.stdout.strip().split(','))
        return "vertical" if h > w else "horizontal"
    except Exception:
        # Fallback: mặc định coi là dọc nếu không thể phân tích
        return "vertical"

def download_video(vid_id, item):
    url = item.get("url")
    title = item.get("title", "video").replace("/", "_").replace("\\", "_")[:50]
    filename = f"{vid_id}_{title}.mp4"
    temp_path = os.path.join(TEMP_DIR, filename)
    
    # Kiểm tra xem đã tồn tại ở 2 thư mục đích chưa
    vert_path = os.path.join(OUTPUT_DIR_VERT, filename)
    horz_path = os.path.join(OUTPUT_DIR_HORZ, filename)
    if os.path.exists(vert_path) or os.path.exists(horz_path):
        return True, "Đã tồn tại"

    print(f"⏳ Đang tải: {vid_id}...")
    cmd = [
        "yt-dlp",
        "-f", "b[ext=mp4]/mp4",
        "--no-warnings",
        "--quiet",
        "-o", temp_path,
        url
    ]
    
    try:
        res = subprocess.run(cmd, capture_output=True, timeout=90)
        if res.returncode == 0 and os.path.exists(temp_path):
            # Phân loại dọc/ngang
            orientation = check_orientation(temp_path)
            if orientation == "vertical":
                shutil.move(temp_path, vert_path)
                final_path = vert_path
            else:
                shutil.move(temp_path, horz_path)
                final_path = horz_path
                
            print(f"✅ Đã tải xong: {vid_id} -> {orientation}")
            return True, "Thành công"
        else:
            print(f"❌ Lỗi tải {vid_id}")
            return False, "Lỗi yt-dlp"
    except Exception as e:
        print(f"❌ Lỗi mạng {vid_id}: {e}")
        return False, str(e)

def main():
    os.makedirs(OUTPUT_DIR_VERT, exist_ok=True)
    os.makedirs(OUTPUT_DIR_HORZ, exist_ok=True)
    os.makedirs(TEMP_DIR, exist_ok=True)
    
    pool = load_pool()
    unused_videos = {k: v for k, v in pool.items() if not v.get("used") and not v.get("failed")}
    
    if not unused_videos:
        print("⚠️ Không có video nào khả dụng trong pool!")
        return

    LIMIT = 1000
    items_to_download = list(unused_videos.items())
    random.shuffle(items_to_download)
    items_to_download = items_to_download[:LIMIT]
    
    print(f"🚀 Bắt đầu tiến trình tải {len(items_to_download)} video siêu tốc...")
    
    success_count = 0
    # Tải đa luồng 16 video cùng lúc để max băng thông
    with ThreadPoolExecutor(max_workers=16) as executor:
        futures = {executor.submit(download_video, k, v): k for k, v in items_to_download}
        for future in futures:
            is_success, msg = future.result()
            if is_success:
                success_count += 1
                
    # Dọn dẹp thư mục temp
    try:
        shutil.rmtree(TEMP_DIR)
    except:
        pass
        
    print(f"🎉 HOÀN TẤT! Đã tải thành công {success_count}/{len(items_to_download)} video.")

if __name__ == "__main__":
    main()
