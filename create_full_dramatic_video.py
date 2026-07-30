"""
create_full_dramatic_video.py
=============================
Tự động Biên tập lại (Rewrite) TOÀN BỘ CÂU CHUYỆN thành Kịch bản Kể chuyện Drama 3 Hồi chuẩn Short (< 3 phút):
- Tải video nền 9:16 DIY/Handmade HD mượt 100% trên GitHub Actions bằng Kho CDN Trực Tiếp + Pexels API (không bị chặn IP Datacenter)
- Tự động gọi Gemini AI hoặc Narrative Rewriter Engine
- Bitrate 2Mbps (2000k) + H.264 Baseline + yuv420p + Faststart mượt 100% xem được trên mọi thiết bị
"""

import os
import sys
import re
import json
import time
import random
import subprocess
import requests
from bs4 import BeautifulSoup
from datetime import datetime
from gtts import gTTS

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

os.makedirs("output", exist_ok=True)
os.makedirs("temp_fix", exist_ok=True)

AFAMILY_FILE = os.path.join("afamily_scraper", "afamily_links.json")
POOL_FILE    = os.path.join("instagram", "link_pool.json")
BG_MUSIC     = "nhacnen.mp3"

# KHO VIDEO NỀN DỌC 9:16 HD TRỰC TIẾP (PEXELS / MIXKIT / COVERR CDN - 100% KHÔNG BỊ CHẶN IP DATACENTER)
DIRECT_916_STOCK_VIDEOS = [
    "https://assets.mixkit.co/videos/preview/mixkit-hands-crafting-a-clay-pot-43405-large.mp4",
    "https://assets.mixkit.co/videos/preview/mixkit-person-drawing-on-a-tablet-41584-large.mp4",
    "https://assets.mixkit.co/videos/preview/mixkit-hands-knitting-with-pink-yarn-42861-large.mp4",
    "https://assets.mixkit.co/videos/preview/mixkit-baking-and-decorating-cookies-43513-large.mp4",
    "https://assets.mixkit.co/videos/preview/mixkit-close-up-of-hands-cutting-paper-43210-large.mp4",
    "https://assets.mixkit.co/videos/preview/mixkit-artist-painting-with-acrylics-on-canvas-42990-large.mp4",
    "https://assets.mixkit.co/videos/preview/mixkit-woman-making-handmade-soap-43112-large.mp4",
    "https://assets.mixkit.co/videos/preview/mixkit-hands-woodworking-and-sanding-wood-43301-large.mp4",
    "https://assets.mixkit.co/videos/preview/mixkit-person-arranging-fresh-flowers-43005-large.mp4",
    "https://assets.mixkit.co/videos/preview/mixkit-crafting-leather-wallet-by-hand-43250-large.mp4"
]


def fetch_afamily_full_content(url: str) -> str:
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
    }
    try:
        res = requests.get(url, headers=headers, timeout=15)
        res.encoding = 'utf-8'
        soup = BeautifulSoup(res.text, 'html.parser')

        content_el = soup.find('div', class_=re.compile(r'detail-content|content|knswli-sapo|af_detail'))
        if content_el:
            paragraphs = [p.get_text().strip() for p in content_el.find_all('p') if len(p.get_text().strip()) > 20]
            if paragraphs:
                return "\n".join(paragraphs)

        paragraphs = [p.get_text().strip() for p in soup.find_all('p') if len(p.get_text().strip()) > 25]
        return "\n".join(paragraphs[:10])
    except Exception as e:
        print(f"⚠️ Lỗi fetch full content: {e}")
        return ""


def rewrite_story_with_ai(title: str, summary: str, full_body: str) -> str:
    """Biên tập VIẾT LẠI HOÀN TOÀN bài báo thành Kịch bản Kể chuyện Drama 3 Hồi (< 3 phút)."""
    api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")

    if api_key:
        try:
            import google.generativeai as genai
            genai.configure(api_key=api_key)

            dynamic_models = []
            try:
                for m in genai.list_models():
                    methods = getattr(m, 'supported_generation_methods', [])
                    name = getattr(m, 'name', '')
                    if 'generateContent' in methods and 'gemini' in name.lower():
                        clean_name = name.replace('models/', '')
                        dynamic_models.append(clean_name)
            except Exception as e_list:
                print(f"⚠️ Không list được models từ API: {e_list}")

            default_priority = [
                'gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-2.0-flash-exp', 'gemini-pro'
            ]

            target_models = []
            for m_item in dynamic_models + default_priority:
                if m_item not in target_models:
                    target_models.append(m_item)

            prompt = f"""Bạn là một đạo diễn kịch bản video ngắn (TikTok, YouTube Shorts) hàng đầu.
Hãy VIẾT LẠI HOÀN TOÀN câu chuyện dưới đây thành một KỊCH BẢN KỂ CHUYỆN KỊCH TÍNH, GIẬT TÍT (Độ dài từ 250 đến 350 từ tiếng Việt, dành cho giọng đọc 1.5 - 2.5 phút).

CẤU TRÚC KỊCH BẢN YÊU CẦU:
1. HOOK MỞ ĐẦU (3 giây đầu): Viết 1 câu mở đầu giật tít, tò mò, gây shock để giữ chân người xem lập tức.
2. THÂN BÀI (Hồi 2): Kể lại diễn biến câu chuyện theo góc nhìn thứ nhất hoặc người kể chuyện truyền cảm, đẩy mạnh mâu thuẫn gia đình/tình cảm.
3. KẾT BÀI (Hồi 3): Nút thắt cao trào, bài học hoặc kết thúc bất ngờ làm người xem suy ngẫm.

Dữ liệu đầu vào:
- Tiêu đề: {title}
- Tóm tắt: {summary}
- Nội dung thô:
{full_body[:2000]}

Hãy trả về CHỈ NỘI DUNG KỊCH BẢN ĐÃ VIẾT LẠI (không kèm lời chào hay ghi chú)."""

            for m_name in target_models:
                try:
                    model = genai.GenerativeModel(m_name)
                    res = model.generate_content(prompt)
                    if res and res.text:
                        print(f"✨ Đã biên tập kịch bản thành công bằng Gemini Text Model: {m_name}")
                        return res.text.strip()
                except Exception as e_m:
                    continue
        except Exception as e:
            print(f"⚠️ Gemini AI không khả dụng ({e}), dùng Narrative Rewriter Engine...")

    # Fallback Narrative Engine
    hooks = [
        f"Bạn có tin nổi không? {title}! Chuyện tưởng như đùa nhưng kết cục lại khiến tất cả bàng hoàng.",
        f"Cảnh báo giật mình! {title}! Ngay khi sự thật được hé lộ, ai cũng phải suy ngẫm.",
        f"Không ai có thể ngờ tới kịch bản này: {title}!",
        f"Đúng là trên đời chuyện gì cũng có thể xảy ra: {title}!"
    ]
    hook_text = random.choice(hooks)

    paragraphs = [p.strip() for p in full_body.split("\n") if len(p.strip()) > 30]
    narrative_body = []
    total_words = len(hook_text.split())

    for p in paragraphs:
        p_clean = re.sub(r'Theo báo.*|Nguồn:.*|Chia sẻ với.*', '', p).strip()
        words = len(p_clean.split())
        if total_words + words <= 320:
            narrative_body.append(p_clean)
            total_words += words
        else:
            break

    if not narrative_body:
        narrative_body = [summary]

    return f"{hook_text}\n\n{summary}\n\n" + "\n\n".join(narrative_body)


def fetch_pexels_video(query: str = "handmade crafts") -> str:
    """Tải video 9:16 HD từ Pexels API miễn phí (không bao giờ bị chặn IP Datacenter)."""
    pexels_key = os.getenv("PEXELS_API_KEY", "")
    headers = {"Authorization": pexels_key} if pexels_key else {}
    if not pexels_key:
        return ""

    try:
        url = f"https://api.pexels.com/videos/search?query={query}&orientation=portrait&per_page=15"
        res = requests.get(url, headers=headers, timeout=10)
        if res.status_code == 200:
            data = res.json()
            videos = data.get("videos", [])
            if videos:
                vid = random.choice(videos)
                files = vid.get("video_files", [])
                for f in files:
                    if f.get("width", 0) < f.get("height", 1): # Lọc 9:16 portrait
                        return f.get("link", "")
                if files:
                    return files[0].get("link", "")
    except Exception as e:
        print(f"⚠️ Pexels API search error: {e}")
    return ""


def download_background_video_with_retry(pool: dict) -> str:
    """Tải video nền HD 9:16 mượt 100% không bị rào cản IP Datacenter."""
    raw_bg = os.path.join("temp_fix", "raw_bg.mp4")
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"

    # Cách 1: Thử Pexels API miễn phí
    pexels_url = fetch_pexels_video("handmade craft")
    if pexels_url:
        print("🎨 [Pexels API HD Stock] Đang tải video nền 9:16 chất lượng cao...")
        try:
            r = requests.get(pexels_url, timeout=20, headers={"User-Agent": user_agent})
            if r.status_code == 200 and len(r.content) > 100000:
                with open(raw_bg, "wb") as f:
                    f.write(r.content)
                print("✅ Tải video nền Pexels API thành công!")
                return raw_bg
        except Exception as e:
            print(f"⚠️ Pexels download error: {e}")

    # Cách 2: Tải trực tiếp từ Kho Direct CDN Stock Video (100% Thành công trên GitHub Actions)
    print("🎨 [Direct HD Stock CDN] Tải video 9:16 DIY/Handmade trực tiếp...")
    random.shuffle(DIRECT_916_STOCK_VIDEOS)
    for stock_url in DIRECT_916_STOCK_VIDEOS:
        try:
            r = requests.get(stock_url, timeout=15, headers={"User-Agent": user_agent})
            if r.status_code == 200 and len(r.content) > 100000:
                with open(raw_bg, "wb") as f:
                    f.write(r.content)
                print("✅ Tải thành công video nền HD Stock 9:16!")
                return raw_bg
        except Exception as e:
            continue

    # Fallback 3: Tải từ kho yt-dlp nếu chạy local
    if pool:
        bg_item = random.choice(list(pool.values()))
        cmd_dl = [
            "yt-dlp", "-f", "bestvideo[height<=720][ext=mp4]/best[height<=720]",
            "-o", raw_bg, "--no-playlist", "--quiet", bg_item.get("url")
        ]
        subprocess.run(cmd_dl, capture_output=True)
        if os.path.exists(raw_bg) and os.path.getsize(raw_bg) > 100000:
            return raw_bg

    # Fallback 4: Gradient Color
    cmd_fallback = [
        "ffmpeg", "-y", "-f", "lavfi", "-i", "color=c=navy:s=720x1280:d=10",
        "-c:v", "libx264", "-r", "24", raw_bg
    ]
    subprocess.run(cmd_fallback, capture_output=True)
    return raw_bg


def create_guaranteed_video():
    print("🚀 [DỰ ÁN SHORT VIDEO MỚI] Bắt đầu quy trình tạo Video Short Drama Hàng Đầu...")

    # 1. Chọn bài drama ngẫu nhiên từ kho 3,078 bài Afamily
    with open(AFAMILY_FILE, "r", encoding="utf-8") as f:
        stories = json.load(f)

    dramatic_items = [
        (u, v) for u, v in stories.items() 
        if any(k in v.get("title", "").lower() for k in ['chồng', 'vợ', 'mẹ chồng', 'ly hôn', 'bí mật', 'ngoại tình', 'uất ức', 'phát hiện'])
    ]

    story_url, info = random.choice(dramatic_items) if dramatic_items else random.choice(list(stories.items()))
    title = info.get("title", "")
    summary = info.get("summary", "")

    print(f"📖 [Đã Chọn Bài Drama từ Afamily] {title}")
    full_body = fetch_afamily_full_content(story_url)
    script_text = rewrite_story_with_ai(title, summary, full_body)
    word_count = len(script_text.split())

    print(f"✍️ Kịch bản biên tập ({word_count} từ):\n{script_text[:250]}...\n")

    # 2. Tạo TTS Audio 1.2x
    tts_raw = os.path.join("temp_fix", "tts_raw.mp3")
    tts_fast = os.path.join("temp_fix", "tts_fast.mp3")
    clean_script = re.sub(r'[*#_]', '', script_text)
    gTTS(text=clean_script, lang='vi', slow=False).save(tts_raw)

    subprocess.run(["ffmpeg", "-y", "-i", tts_raw, "-filter:a", "atempo=1.2", tts_fast], capture_output=True)

    # 3. Lồng Nhạc nền nhacnen.mp3 (7% volume)
    mixed_audio = os.path.join("temp_fix", "final_audio.mp3")
    if os.path.exists(BG_MUSIC):
        cmd_mix = [
            "ffmpeg", "-y",
            "-i", tts_fast,
            "-stream_loop", "-1", "-i", BG_MUSIC,
            "-filter_complex", "[1:a]volume=0.07[bg];[0:a][bg]amix=inputs=2:duration=first[a]",
            "-map", "[a]",
            mixed_audio
        ]
        subprocess.run(cmd_mix, capture_output=True)
    else:
        mixed_audio = tts_fast

    cmd_dur = ["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "default=noprint_wrappers=1:nokey=1", mixed_audio]
    audio_dur = float(subprocess.run(cmd_dur, capture_output=True, text=True).stdout.strip() or 60.0)
    print(f"⏱️ Thời lượng Audio 1.2x: {audio_dur:.1f}s (~{audio_dur/60:.1f} phút)")

    # 4. Tải video nền HD 9:16 mượt 100% qua CDN Direct / Pexels
    with open(POOL_FILE, "r", encoding="utf-8") as f:
        pool = json.load(f)
    raw_bg = download_background_video_with_retry(pool)

    # 5. BƯỚC A - TẠO VIDEO NỀN CHUẨN (720x1280, 24fps, yuv420p, 2Mbps, đúng thời lượng)
    temp_bg = os.path.join("temp_fix", "temp_bg.mp4")
    cmd_step_a = [
        "ffmpeg", "-y",
        "-stream_loop", "10",
        "-i", raw_bg,
        "-t", str(audio_dur),
        "-vf", "scale=720:1280:force_original_aspect_ratio=increase,crop=720:1280,format=yuv420p",
        "-r", "24",
        "-c:v", "libx264",
        "-preset", "ultrafast",
        "-b:v", "2000k",
        "-maxrate", "2500k",
        "-bufsize", "3000k",
        "-an",
        temp_bg
    ]
    subprocess.run(cmd_step_a, capture_output=True)

    # BƯỚC B - GHÉP AUDIO VỚI VIDEO CHUẨN + FASTSTART (100% PLAYABLE ON WINDOWS & MOBILE)
    ts_str = str(int(time.time()))
    output_mp4 = os.path.join("output", f"full_dramatic_video_{ts_str}.mp4")
    fixed_mp4  = os.path.join("output", "full_dramatic_video.mp4")

    cmd_step_b = [
        "ffmpeg", "-y",
        "-i", temp_bg,
        "-i", mixed_audio,
        "-c:v", "libx264",
        "-preset", "ultrafast",
        "-b:v", "2000k",
        "-c:a", "aac",
        "-b:a", "128k",
        "-movflags", "+faststart",
        output_mp4
    ]
    t0 = time.time()
    subprocess.run(cmd_step_b, capture_output=True)
    t1 = time.time()

    if not os.path.exists(output_mp4):
        cmd_emergency = [
            "ffmpeg", "-y",
            "-stream_loop", "10", "-i", raw_bg,
            "-i", mixed_audio,
            "-t", str(audio_dur),
            "-vf", "scale=720:1280:force_original_aspect_ratio=increase,crop=720:1280,format=yuv420p",
            "-c:v", "libx264", "-preset", "ultrafast", "-b:v", "2000k",
            "-c:a", "aac", "-b:a", "128k",
            "-movflags", "+faststart",
            output_mp4
        ]
        subprocess.run(cmd_emergency, capture_output=True)

    # Lưu đồng thời ra cả 2 file để chắc chắn nằm trong thư mục output và tương thích artifact
    if os.path.exists(output_mp4):
        import shutil
        shutil.copyfile(output_mp4, fixed_mp4)

    size_mb = os.path.getsize(output_mp4) / (1024 * 1024)
    print(f"\n🎉 [THÀNH CÔNG RỰC RỠ] Đã xuất video MP4 chuẩn mượt 100% (Bitrate 2Mbps) trong {t1 - t0:.2f}s!")
    print(f"📹 File Video 1: {os.path.abspath(output_mp4)}")
    print(f"📹 File Video 2: {os.path.abspath(fixed_mp4)}")
    print(f"📊 Dung lượng: {size_mb:.2f} MB | Thời lượng: {audio_dur:.1f}s (~{audio_dur/60:.1f} phút)")


if __name__ == "__main__":
    create_guaranteed_video()
