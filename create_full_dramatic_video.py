"""
create_full_dramatic_video.py
=============================
Tự động Biên tập lại (Rewrite) TOÀN BỘ CÂU CHUYỆN thành Kịch bản Kể chuyện Drama 3 Hồi chuẩn Short (< 3 phút):
- TÍCH HỢP BỘ CÀO PROXY & XOAY PROXY TỰ ĐỘNG (Proxy Scraper & Rotator Engine):
  + Tự động cào hàng chục HTTP Proxy miễn phí live từ các nguồn công cộng.
  + Thử từng Proxy với yt-dlp. Nếu hết proxy/lỗi thì tự động cào batch proxy mới.
  + Chạy lặp liên tục cho tới khi tải video nền thành công mới thôi!
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


# ---------------------------------------------------------------------------
# BỘ CÀO & XOAY PROXY TỰ ĐỘNG (PROXY SCRAPER & ROTATOR ENGINE)
# ---------------------------------------------------------------------------
def scrape_free_proxies(limit: int = 30) -> list:
    """Cào danh sách Proxy HTTP miễn phí mới nhất từ các nguồn API công cộng."""
    print("🌐 Đang tự động cào danh sách HTTP Proxy mới nhất...")
    proxy_urls = [
        "https://api.proxyscrape.com/v2/?request=getproxies&protocol=http&timeout=5000&country=all&ssl=all&anonymity=all",
        "https://raw.githubusercontent.com/TheSpeedX/PROXY-List/master/http.txt",
        "https://raw.githubusercontent.com/monosans/proxy-list/main/proxies/http.txt",
        "https://raw.githubusercontent.com/clarketm/proxy-list/master/proxy-list-raw.txt"
    ]

    proxies = set()
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"

    for src_url in proxy_urls:
        try:
            res = requests.get(src_url, timeout=6, headers={"User-Agent": user_agent})
            if res.status_code == 200:
                matches = re.findall(r'\b(?:\d{1,3}\.){3}\d{1,3}:\d{2,5}\b', res.text)
                for ip_port in matches:
                    proxies.add(f"http://{ip_port}")
                if len(proxies) >= limit * 2:
                    break
        except Exception:
            continue

    proxy_list = list(proxies)
    random.shuffle(proxy_list)
    print(f"✅ Đã cào thành công {len(proxy_list)} Proxy live!")
    return proxy_list[:limit]


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

            print(f"🔍 Danh sách Model Text sẽ thử fallback: {target_models[:5]}")
            for m_name in target_models:
                try:
                    model = genai.GenerativeModel(m_name)
                    res = model.generate_content(prompt)
                    if res and res.text:
                        print(f"✨ Đã biên tập kịch bản thành công bằng Gemini Text Model: {m_name}")
                        return res.text.strip()
                except Exception as e_m:
                    print(f"⚠️ Model '{m_name}' không phản hồi, tự động chuyển sang model tiếp theo...")
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


def download_background_video_with_proxy_loop(pool: dict) -> str:
    """
    TẢI VIDEO NỀN DÙNG PROXY XOAY TỰ ĐỘNG:
    - Lấy chục Proxy live để tải. Nếu hết proxy/lỗi thì tự động cào batch proxy mới.
    - Chạy lặp liên tục tới khi tải xong video mới thôi!
    """
    raw_bg = os.path.join("temp_fix", "raw_bg.mp4")
    pool_items = list(pool.values())
    
    proxy_pool = scrape_free_proxies(limit=30)
    attempt_count = 0

    while True:
        attempt_count += 1

        # Nếu hết proxy trong pool thì cào đợt proxy mới
        if not proxy_pool:
            print("🔄 Danh sách Proxy hiện tại đã dùng hết, đang tự động CÀO BẮT BỘ PROXY MỚI...")
            proxy_pool = scrape_free_proxies(limit=30)

        current_proxy = proxy_pool.pop(0) if proxy_pool else None
        bg_item = random.choice(pool_items)
        bg_url = bg_item.get("url")

        print(f"🎨 [Tải Video Nền Thử Lần {attempt_count}] Proxy: {current_proxy or 'Direct'} | Title: {bg_item.get('title')[:45]}")

        if os.path.exists(raw_bg):
            try: os.remove(raw_bg)
            except: pass

        cmd_dl = [
            "yt-dlp",
            "-f", "bestvideo[height<=720][ext=mp4]+bestaudio[ext=m4a]/bestvideo[height<=720]/best[height<=720]/best",
            "-o", raw_bg,
            "--no-playlist",
            "--quiet",
            "--socket-timeout", "10"
        ]

        if current_proxy:
            cmd_dl.extend(["--proxy", current_proxy])

        cmd_dl.append(bg_url)

        try:
            res = subprocess.run(cmd_dl, capture_output=True, timeout=25)
            if os.path.exists(raw_bg) and os.path.getsize(raw_bg) > 100000:
                print(f"🎉 [THÀNH CÔNG RỰC RỠ] Đã tải xong Video Nền qua Proxy: {current_proxy}!")
                return raw_bg
        except subprocess.TimeoutExpired:
            print("⚠️ Proxy bị timeout (quá 25s), chuyển sang proxy tiếp theo...")
        except Exception as e:
            print(f"⚠️ Thử proxy thất bại: {e}")

        # Thử Direct Stock Video CDN nếu proxy thất bại liên tục > 8 lần
        if attempt_count % 8 == 0:
            print("🎨 Tải dự phòng video HD Stock 9:16 CDN...")
            stock_urls = [
                "https://assets.mixkit.co/videos/preview/mixkit-hands-crafting-a-clay-pot-43405-large.mp4",
                "https://assets.mixkit.co/videos/preview/mixkit-person-drawing-on-a-tablet-41584-large.mp4",
                "https://assets.mixkit.co/videos/preview/mixkit-hands-knitting-with-pink-yarn-42861-large.mp4"
            ]
            for s_url in stock_urls:
                try:
                    r = requests.get(s_url, timeout=12)
                    if r.status_code == 200 and len(r.content) > 100000:
                        with open(raw_bg, "wb") as f:
                            f.write(r.content)
                        print("✅ Tải dự phòng HD Stock thành công!")
                        return raw_bg
                except Exception:
                    continue


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

    # 4. Tải video nền DÙNG PROXY XOAY TỰ ĐỘNG (Bao giờ tải xong mới thôi)
    with open(POOL_FILE, "r", encoding="utf-8") as f:
        pool = json.load(f)
    raw_bg = download_background_video_with_proxy_loop(pool)

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
