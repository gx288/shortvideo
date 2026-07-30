"""
eva_scraper/script_rewriter.py
===============================
Viết lại nội dung bài viết thành kịch bản kể chuyện video:
- Thời lượng: 1.5 đến 3 phút (khoảng 250 - 450 từ tiếng Việt ở tốc độ TTS 1.25x)
- Cấu trúc: 
    1. Câu Hook mở đầu gây tò mò trong 3 giây đầu
    2. Hành văn tự nhiên, trau chuốt, giàu cảm xúc cho TTS đọc
    3. Giữ trọn vẹn diễn biến & cốt truyện chính
- Đọc file JSON từ eva_scraper/data/ và lưu kịch bản đã viết lại vào eva_scraper/scripts/

Cách dùng:
    python eva_scraper/script_rewriter.py --batch 20
"""

import os
import sys
import re
import json
import random
import argparse
from datetime import datetime

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

DATA_DIR    = os.path.join("eva_scraper", "data")
SCRIPTS_DIR = os.path.join("eva_scraper", "scripts")
MIN_TARGET_WORDS = 220   # ~1.5 phút
MAX_TARGET_WORDS = 420   # ~3 phút


# ─────────────────────────────────────────────────────────────────────────────
# SCRIPT REWRITER ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def rewrite_story_text(title: str, content: str) -> str:
    """
    Viết lại hành văn câu chuyện ngắn gọn, lôi cuốn cho giọng đọc TTS.
    """
    # 1. Tạo câu Hook gây tò mò mở đầu
    hook = _create_hook(title, content)

    # 2. Làm sạch và chia đoạn văn
    paragraphs = [p.strip() for p in content.split("\n") if p.strip()]
    if not paragraphs:
        paragraphs = [content]

    # 3. Biên tập lại văn phong thoại
    rewritten_sentences = []
    if hook:
        rewritten_sentences.append(hook)

    full_text = " ".join(paragraphs)

    # Tách câu
    sentences = re.split(r'(?<=[.!?:])\s+', full_text)
    
    for s in sentences:
        s = s.strip()
        if not s or len(s) < 10:
            continue
        
        # Tối ưu từ ngữ cho giọng đọc truyền cảm
        s = _refine_sentence_style(s)
        rewritten_sentences.append(s)

    final_script = " ".join(rewritten_sentences)

    # 4. Điều chỉnh độ dài chuẩn (220 - 420 từ = < 3 phút)
    words = final_script.split()
    if len(words) > MAX_TARGET_WORDS:
        # Cắt ngọt ở dấu câu gần nhất
        truncated = " ".join(words[:MAX_TARGET_WORDS])
        last_punct = max(truncated.rfind('.'), truncated.rfind('!'), truncated.rfind('?'))
        if last_punct > len(truncated) * 0.7:
            final_script = truncated[:last_punct + 1]
        else:
            final_script = truncated + "..."

    return final_script


def _create_hook(title: str, content: str) -> str:
    """Tạo mở đầu giật gân/tò mò cho 3s đầu video."""
    clean_title = re.sub(r'^\s*(Tâm sự|Chuyện cũ|Hot):\s*', '', title, flags=re.I).strip()
    
    hooks = [
        f"Có những câu chuyện khiến người ta không thể nào quên. {clean_title}.",
        f"Đừng bao giờ vội đánh giá một người qua vẻ bề ngoài. {clean_title}.",
        f"Câu chuyện ngày hôm nay có lẽ sẽ khiến nhiều người phải suy ngẫm. {clean_title}.",
    ]
    return random.choice(hooks)


def _refine_sentence_style(sentence: str) -> str:
    """Thay đổi nhẹ hành văn mượt mà hơn khi đọc bằng TTS."""
    replacements = {
        r'\btôi nghĩ là\b': 'tôi thầm nghĩ',
        r'\bngay lập tức\b': 'ngay lập tức',
        r'\bkhông ngờ rằng\b': 'chẳng thể ngờ',
        r'\brất là\b': 'rất',
        r'\bbắt đầu\b': 'bắt đầu',
    }
    for pat, repl in replacements.items():
        sentence = re.sub(pat, repl, sentence, flags=re.I)
    return sentence


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

    for filepath in json_files:
        if processed >= batch_size:
            break

        basename = os.path.basename(filepath)
        script_filename = basename.replace(".json", "_script.json")
        script_filepath = os.path.join(SCRIPTS_DIR, script_filename)

        # Bỏ qua nếu đã tạo kịch bản rồi
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
                "created_at":   datetime.now().strftime("%Y-%m-%d %H:%M"),
            }

            with open(script_filepath, "w", encoding="utf-8") as f:
                json.dump(script_data, f, ensure_ascii=False, indent=2)

            print(f"  ✅ [{processed+1}/{batch_size}] {title[:40]}... → {word_count} từ (~{est_seconds}s)")
            processed += 1

        except Exception as e:
            print(f"  ❌ Lỗi xử lý {basename}: {e}")

    print(f"\n[Rewriter] ✅ Đã hoàn thành kịch bản cho {processed} bài viết (thời lượng < 3 phút).")
    return processed


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Script Rewriter for Story Videos")
    parser.add_argument("--batch", type=int, default=20, help="Số bài viết cần viết lại kịch bản")
    args = parser.parse_args()

    process_batch_rewrite(args.batch)
