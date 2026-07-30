"""
eva_scraper/full_batch_scraper.py
==================================
Thu thập TOÀN BỘ bài viết và kịch bản video trong 1 LẦN CHẠY DUY NHẤT bằng Multithreading.
Tốc độ: 500 - 1,000 bài viết trong vòng ~1 phút (Nhanh gấp 20x so với Selenium).

Quy trình 1 lần duy nhất:
1. Quét song song 500-1000 ID bài viết từ Eva.vn.
2. Tải toàn bộ nội dung chi tiết bài viết thô vào eva_scraper/data/.
3. Tự động viết lại kịch bản video ngắn < 3 phút vào eva_scraper/scripts/.

Chạy:
    python eva_scraper/full_batch_scraper.py --count 300
    python eva_scraper/full_batch_scraper.py --count 1000
"""

import os
import sys
import re
import json
import time
import hashlib
import argparse
import requests
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

OUTPUT_DIR  = os.path.join("eva_scraper", "data")
SCRIPTS_DIR = os.path.join("eva_scraper", "scripts")
LINKS_FILE  = os.path.join("eva_scraper", "links_master.json")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/125.0.0.0 Safari/537.36",
    "Accept-Language": "vi-VN,vi;q=0.9",
}

STORY_CATEGORIES = ['/tam-su/', '/tinh-yeu-gioi-tinh/', '/gia-dinh/', '/chuyen-eva/', '/goc-tam-su/']


# ─────────────────────────────────────────────────────────────────────────────
# MULTITHREADED ID SCANNER & CONTENT SCRAPER
# ─────────────────────────────────────────────────────────────────────────────

def process_single_id(art_id: int) -> dict | None:
    """Tải và bóc tách 1 bài viết theo ID trong 1 request duy nhất."""
    url = f"https://eva.vn/tam-su/x-c391a{art_id}.html"
    try:
        r = requests.get(url, headers=HEADERS, allow_redirects=True, timeout=10)
        final_url = r.url

        if r.status_code != 200 or not any(cat in final_url for cat in STORY_CATEGORIES):
            return None

        soup = BeautifulSoup(r.text, "lxml")

        # Tiêu đề
        h1 = soup.select_one("h1.fw-bold, h1.color-main, h1")
        title = h1.get_text(strip=True) if h1 else ""

        # Nội dung
        content_el = (
            soup.select_one("div#baiviet-container") or
            soup.select_one("div.eva-cont-art__info") or
            soup.select_one("div.article-content") or
            soup.select_one("div.content-detail") or
            soup.select_one("article .content") or
            soup.select_one("main")
        )

        paragraphs = []
        if content_el:
            for p in content_el.find_all("p"):
                t = p.get_text(strip=True)
                if len(t) > 20:
                    paragraphs.append(t)

        content = " ".join(paragraphs)
        content = re.sub(r"\s{2,}", " ", content).strip()
        words = len(content.split())

        if words < 100:  # Bỏ bài quá ngắn
            return None

        url_hash = hashlib.md5(final_url.encode()).hexdigest()[:8]
        slug = re.sub(r"[^\w]", "_", title.lower())[:40].strip("_")
        filename = f"{url_hash}_{slug}.json"
        filepath = os.path.join(OUTPUT_DIR, filename)

        article_data = {
            "url": final_url,
            "title": title,
            "content": content,
            "words": words,
            "scraped_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }

        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(article_data, f, ensure_ascii=False, indent=2)

        return {"id": art_id, "url": final_url, "title": title, "file": filepath, "words": words}

    except Exception:
        return None


def run_full_batch_scrape(start_id: int = 678450, count: int = 300, max_workers: int = 20) -> int:
    """
    Chạy cào 1 THỂ hàng loạt hàng trăm bài viết bằng Multithreading.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(SCRIPTS_DIR, exist_ok=True)

    print(f"🚀 [FULL BATCH] Bắt đầu cào 1 THỂ {count} bài viết từ ID {start_id} lùi dần...")
    print(f"⚡ Số luồng xử lý đồng thời (Multithreading): {max_workers} threads\n")

    id_list = list(range(start_id, start_id - count, -1))
    success_count = 0

    start_time = time.time()

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(process_single_id, art_id): art_id for art_id in id_list}

        for i, future in enumerate(as_completed(futures), 1):
            res = future.result()
            if res:
                success_count += 1
                title_preview = res["title"][:45]
                print(f"  [{success_count}/{count}] ✅ ID {res['id']}: {title_preview}... ({res['words']} từ)")
            
            if i % 50 == 0:
                elapsed = time.time() - start_time
                print(f"   ⏱️ Đã xử lý {i}/{count} ID ({elapsed:.1f}s)...")

    elapsed_total = time.time() - start_time
    print(f"\n🎉 [HOÀN THÀNH CÀO 1 THỂ] Thu thập {success_count} bài viết thành công trong {elapsed_total:.1f} giây!")

    # Tự động viết lại kịch bản cho toàn bộ bài viết mới cào
    print("\n📝 [FULL BATCH] Đang chuyển toàn bộ bài viết thành KỊCH BẢN VIDEO < 3 phút...")
    from eva_scraper.script_rewriter import process_batch_rewrite
    process_batch_rewrite(batch_size=success_count)

    return success_count


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Full Batch Story Scraper (Multithreaded)")
    parser.add_argument("--start-id", type=int, default=678450, help="ID bắt đầu quét")
    parser.add_argument("--count", type=int, default=300, help="Tổng số ID cần quét 1 thể (ví dụ 300, 500, 1000)")
    parser.add_argument("--threads", type=int, default=20, help="Số luồng xử lý đồng thời (mặc định 20)")
    args = parser.parse_args()

    run_full_batch_scrape(start_id=args.start_id, count=args.count, max_workers=args.threads)
