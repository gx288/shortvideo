"""
eva_scraper/script_rewriter.py
===============================
Sử dụng Gemini AI (hoặc rule-based fallback) để viết lại nội dung bài viết 
thành kịch bản kể chuyện video YouTube Shorts / TikTok cuốn hút:

Tiêu chuẩn kịch bản Gemini AI:
- Thời lượng: 1.5 đến 3 phút (khoảng 250 - 400 từ tiếng Việt)
- Cấu trúc: 
    1. Câu Hook mở đầu gây tò mò trong 3 giây đầu (giữ chân người xem)
    2. Hành văn tự nhiên, truyền cảm, giàu cảm xúc dành cho giọng đọc Google TTS
    3. Trọng tâm là diễn biến kịch tính & cao trào của câu chuyện

Environment variables:
    GEMINI_API_KEY hoặc GOOGLE_API_KEY (hoặc secrets.GEMINI_API_KEY trên GitHub Actions)

Cách dùng:
    python eva_scraper/script_rewriter.py --batch 20
"""

import os
import sys
import re
import json
import time
import random
import argparse
from datetime import datetime

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

DATA_DIR    = os.path.join("eva_scraper", "data")
SCRIPTS_DIR = os.path.join("eva_scraper", "scripts")
MIN_TARGET_WORDS = 220   # ~1.5 phút
MAX_TARGET_WORDS = 420   # ~3 phút

# Khởi tạo Gemini AI nếu có thư viện & API key
GEMINI_AVAILABLE = False
try:
    import google.generativeai as genai
    api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if api_key:
        genai.configure(api_key=api_key)
        GEMINI_AVAILABLE = True
except Exception:
    GEMINI_AVAILABLE = False


# ─────────────────────────────────────────────────────────────────────────────
# GEMINI AI REWRITER WITH FALLBACK
# ─────────────────────────────────────────────────────────────────────────────

GEMINI_MODELS = [
    'gemini-2.5-flash',
    'gemini-1.5-flash',
    'gemini-flash-latest'
]

PROMPT_TEMPLATE = """Bạn là một chuyên gia biên tập kịch bản video ngắn (YouTube Shorts, TikTok, Reels) chuyên nghiệp.
Hãy viết lại bài viết tâm sự sau thành một KỊCH BẢN KỂ CHUYỆN HẤP DẪN để đọc thành video ngắn.

YÊU CẦU BẮT BUỘC:
1. MỞ ĐẦU (HOOK 3 GIÂY ĐẦU): Thêm 1-2 câu mở đầu cực kỳ kịch tính, gây tò mò hoặc bất ngờ để giữ chân người xem ngay lập tức.
2. NỘI DUNG CHÍNH: Tóm tắt và viết lại cốt chuyện diễn biến mạch lạc, tập trung vào mâu thuẫn, cảm xúc và cái kết đáng suy ngẫm.
3. HÀNH VĂN: Tự nhiên, văn thoại gần gũi, truyền cảm, tối ưu cho giọng đọc AI (Google TTS) phát âm mượt mà.
4. ĐỘ DÀI: BẮT BUỘC trong khoảng 250 đến 380 từ tiếng Việt (thời lượng video đúng 2 phút đến 2.5 phút). Không quá ngắn và không vượt quá 400 từ.
5. CHỈ TRẢ VỀ NỘI DUNG KỊCH BẢN ĐỌC. Không thêm lời chào, không thêm chú thích hay ký tự đặc biệt như (*), (#).

Tiêu đề: {title}
Nội dung bài viết:
{content}
"""


def rewrite_with_gemini(title: str, content: str) -> str | None:
    """Gọi Gemini AI API để chuyển bài viết thành kịch bản video."""
    if not GEMINI_AVAILABLE:
        return None

    prompt = PROMPT_TEMPLATE.format(title=title, content=content[:3000])

    for model_name in GEMINI_MODELS:
        try:
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            if response and response.text:
                script_text = _clean_ai_output(response.text)
                print(f"      [Gemini AI - {model_name}] Thành công!")
                return script_text
        except Exception as e:
            print(f"      [!] Gemini model {model_name} lỗi: {str(e)[:60]}...")
            time.sleep(1)

    return None


def _clean_ai_output(text: str) -> str:
    """Loại bỏ ký tự markdown, chú thích thừa của AI."""
    text = re.sub(r'\*+', '', text)
    text = re.sub(r'#{1,6}\s', '', text)
    text = re.sub(r'\[.*?\]', '', text)
    text = re.sub(r'^(Kịch bản|Lời dẫn|Host|MC):\s*', '', text, flags=re.I)
    text = re.sub(r'\n+', ' ', text)
    text = re.sub(r'\s{2,}', ' ', text)
    return text.strip()


# ─────────────────────────────────────────────────────────────────────────────
# RULE-BASED FALLBACK REWRITER
# ─────────────────────────────────────────────────────────────────────────────

def rewrite_rule_based(title: str, content: str) -> str:
    """Fallback viết lại kịch bản bằng thuật toán nếu không có Gemini API key."""
    clean_title = re.sub(r'^\s*(Tâm sự|Chuyện cũ|Hot):\s*', '', title, flags=re.I).strip()
    hooks = [
        f"Có những câu chuyện khiến người ta không thể nào quên. {clean_title}.",
        f"Đừng bao giờ vội đánh giá một người qua vẻ bề ngoài. {clean_title}.",
        f"Câu chuyện ngày hôm nay có lẽ sẽ khiến nhiều người phải suy ngẫm. {clean_title}.",
    ]
    hook = random.choice(hooks)

    paragraphs = [p.strip() for p in content.split("\n") if p.strip()] or [content]
    full_text = " ".join(paragraphs)

    sentences = re.split(r'(?<=[.!?:])\s+', full_text)
    rewritten = [hook]

    for s in sentences:
        s = s.strip()
        if len(s) > 10:
            rewritten.append(s)

    final_script = " ".join(rewritten)
    words = final_script.split()

    if len(words) > MAX_TARGET_WORDS:
        truncated = " ".join(words[:MAX_TARGET_WORDS])
        last_punct = max(truncated.rfind('.'), truncated.rfind('!'), truncated.rfind('?'))
        if last_punct > len(truncated) * 0.7:
            final_script = truncated[:last_punct + 1]
        else:
            final_script = truncated + "..."

    return final_script


def rewrite_story_text(title: str, content: str) -> str:
    """Hàm chính: Ưu tiên Gemini AI, fallback sang Rule-based."""
    if GEMINI_AVAILABLE:
        script = rewrite_with_gemini(title, content)
        if script:
            return script
    return rewrite_rule_based(title, content)


# ─────────────────────────────────────────────────────────────────────────────
# BATCH REWRITER PROCESSOR
# ─────────────────────────────────────────────────────────────────────────────

def process_batch_rewrite(batch_size: int = 20) -> int:
    os.makedirs(SCRIPTS_DIR, exist_ok=True)

    if not os.path.exists(DATA_DIR):
        print(f"[Rewriter] ❌ Chưa có dữ liệu trong {DATA_DIR}")
        return 0

    json_files = [os.path.join(DATA_DIR, f) for f in os.listdir(DATA_DIR) if f.endswith(".json")]
    if not json_files:
        print("[Rewriter] ❌ Không tìm thấy file bài viết JSON nào.")
        return 0

    processed = 0
    engine_used = "Gemini AI" if GEMINI_AVAILABLE else "Rule-based Engine"
    print(f"[Rewriter] 🚀 Bắt đầu tạo kịch bản (Engine: {engine_used})...")

    for filepath in json_files:
        if processed >= batch_size:
            break

        basename = os.path.basename(filepath)
        script_filename = basename.replace(".json", "_script.json")
        script_filepath = os.path.join(SCRIPTS_DIR, script_filename)

        if os.path.exists(script_filepath):
            continue

        try:
            with open(filepath, "r", encoding="utf-8") as f:
                article = json.load(f)

            title   = article.get("title", "")
            content = article.get("content", "")

            if not content:
                continue

            script_text = rewrite_story_text(title, content)
            word_count  = len(script_text.split())
            est_seconds = int((word_count / 180) * 60)   # ~180 từ/phút ở 1.25x

            script_data = {
                "article_url":  article.get("url", ""),
                "title":        title,
                "script_text":  script_text,
                "word_count":   word_count,
                "est_seconds":  est_seconds,
                "engine":       "Gemini AI" if GEMINI_AVAILABLE else "Rule-based",
                "created_at":   datetime.now().strftime("%Y-%m-%d %H:%M"),
            }

            with open(script_filepath, "w", encoding="utf-8") as f:
                json.dump(script_data, f, ensure_ascii=False, indent=2)

            print(f"  ✅ [{processed+1}/{batch_size}] {title[:35]}... → {word_count} từ (~{est_seconds}s)")
            processed += 1

            if GEMINI_AVAILABLE:
                time.sleep(1)

        except Exception as e:
            print(f"  ❌ Lỗi xử lý {basename}: {e}")

    print(f"\n[Rewriter] ✅ Đã hoàn thành kịch bản cho {processed} bài viết.")
    return processed


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Script Rewriter for Story Videos with Gemini AI")
    parser.add_argument("--batch", type=int, default=20, help="Số bài viết cần viết lại kịch bản")
    args = parser.parse_args()

    process_batch_rewrite(args.batch)
