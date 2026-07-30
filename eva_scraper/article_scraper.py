"""
eva_scraper/article_scraper.py
==============================
Lấy nội dung từng bài viết từ links_master.json.
- Đọc các URL chưa được scrape (scraped: false)
- Fetch HTML từng bài
- Lưu file riêng: eva_scraper/data/<article_id>.json
- Update links_master.json (scraped: true, file: "path")

Chạy:
    python eva_scraper/article_scraper.py --batch 50   # Lấy 50 bài
    python eva_scraper/article_scraper.py --batch 500  # Lấy 500 bài
    python eva_scraper/article_scraper.py --stats      # Xem thống kê
"""

import os
import sys
import re
import json
import time
import random
import hashlib
import argparse
import requests
from bs4 import BeautifulSoup
from datetime import datetime

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

OUTPUT_DIR  = os.path.join("eva_scraper", "data")
LINKS_FILE  = os.path.join("eva_scraper", "links_master.json")
MIN_WORDS   = 100    # Bỏ bài quá ngắn
MAX_WORDS   = 800    # Cắt bài quá dài (~80s TTS)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/125.0.0.0 Safari/537.36",
    "Accept-Language": "vi-VN,vi;q=0.9",
}


# ─────────────────────────────────────────────────────────────────────────────
# MASTER FILE I/O
# ─────────────────────────────────────────────────────────────────────────────

def load_master() -> dict:
    if os.path.exists(LINKS_FILE):
        with open(LINKS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_master(master: dict):
    with open(LINKS_FILE, "w", encoding="utf-8") as f:
        json.dump(master, f, ensure_ascii=False, indent=2)


# ─────────────────────────────────────────────────────────────────────────────
# SCRAPE 1 BÀI VIẾT
# ─────────────────────────────────────────────────────────────────────────────

def scrape_article(url: str) -> dict | None:
    """
    Fetch và parse 1 bài viết.
    Trả về dict hoặc None nếu thất bại / quá ngắn.
    """
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        r.raise_for_status()
        r.encoding = "utf-8"
    except Exception as e:
        print(f"  ❌ Fetch lỗi: {e}")
        return None

    soup = BeautifulSoup(r.text, "lxml")

    # ── Tiêu đề ────────────────────────────────────────────────────────────────
    title = ""
    h1 = soup.select_one("h1.fw-bold, h1.color-main, h1")
    if h1:
        title = h1.get_text(strip=True)

    # ── Nội dung chính ─────────────────────────────────────────────────────────
    content_el = (
        soup.select_one("div#baiviet-container") or
        soup.select_one("div.eva-cont-art__info") or
        soup.select_one("div.article-content") or
        soup.select_one("div.content-detail") or
        soup.select_one("article .content")
    )

    if not content_el:
        # Fallback: lấy tất cả <p> trong main
        content_el = soup.select_one("main") or soup.select_one("body")

    # Lấy tất cả đoạn văn
    paragraphs = []
    if content_el:
        for p in content_el.find_all("p"):
            text = p.get_text(strip=True)
            if len(text) > 20:  # Bỏ qua dòng quá ngắn
                paragraphs.append(text)

    content = " ".join(paragraphs)

    # Làm sạch
    content = _clean_text(content)

    if len(content.split()) < MIN_WORDS:
        print(f"  ⏭️ Bài quá ngắn ({len(content.split())} từ), bỏ qua.")
        return None

    # Cắt nếu quá dài
    content = _truncate(content, MAX_WORDS)

    # ── Tác giả ────────────────────────────────────────────────────────────────
    author = ""
    author_el = (
        soup.select_one(".authorName") or
        soup.select_one(".eva-author-time-art") or
        soup.select_one("p.author") or
        soup.select_one("[class*='author']")
    )
    if author_el:
        author = author_el.get_text(strip=True)

    # ── Ngày đăng ──────────────────────────────────────────────────────────────
    date = ""
    date_el = (
        soup.select_one("time") or
        soup.select_one("[class*='date']") or
        soup.select_one("[class*='time']")
    )
    if date_el:
        date = date_el.get("datetime", "") or date_el.get_text(strip=True)

    return {
        "url":     url,
        "title":   title,
        "content": content,
        "author":  author,
        "date":    date,
        "words":   len(content.split()),
        "scraped_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


def _clean_text(text: str) -> str:
    """Loại bỏ ký tự thừa, emoji, HTML entities."""
    text = re.sub(r"&[a-z]+;", " ", text)          # HTML entities
    text = re.sub(r"<[^>]+>", " ", text)            # HTML tags còn sót
    text = re.sub(r"https?://\S+", "", text)         # URLs
    text = re.sub(r"[^\w\s.,!?;:\"'()\-–—]", " ", text)  # Ký tự lạ
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()


def _truncate(text: str, max_words: int) -> str:
    words = text.split()
    if len(words) <= max_words:
        return text
    truncated = " ".join(words[:max_words])
    # Tìm dấu câu gần nhất để không cắt giữa câu
    for sep in [".", "!", "?"]:
        idx = truncated.rfind(sep)
        if idx > len(truncated) * 0.7:
            return truncated[:idx + 1]
    return truncated + "..."


# ─────────────────────────────────────────────────────────────────────────────
# SAVE ARTICLE FILE
# ─────────────────────────────────────────────────────────────────────────────

def save_article(article: dict) -> str:
    """
    Lưu file JSON cho 1 bài viết.
    Tên file: <md5_url[:8]>_<slug_title>.json
    Trả về filepath.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Tạo tên file từ URL
    url_hash = hashlib.md5(article["url"].encode()).hexdigest()[:8]
    slug = re.sub(r"[^\w]", "_", article["title"].lower())[:40].strip("_")
    filename = f"{url_hash}_{slug}.json"
    filepath = os.path.join(OUTPUT_DIR, filename)

    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(article, f, ensure_ascii=False, indent=2)

    return filepath


# ─────────────────────────────────────────────────────────────────────────────
# MAIN BATCH SCRAPER
# ─────────────────────────────────────────────────────────────────────────────

def scrape_batch(batch_size: int = 50, prefer_unscraped: bool = True) -> int:
    """
    Lấy nội dung batch_size bài chưa được scrape.
    Trả về số bài đã lấy thành công.
    """
    master = load_master()

    # Lấy danh sách chưa scrape
    pending = [
        (url, info)
        for url, info in master.items()
        if not info.get("scraped") and not info.get("failed")
    ]

    if not pending:
        print("✅ Tất cả bài đã được scrape!")
        return 0

    targets = pending[:batch_size]
    print(f"[Scraper] Sẽ lấy {len(targets)}/{len(pending)} bài chưa scrape...")

    success = 0
    fail    = 0

    for i, (url, info) in enumerate(targets, 1):
        title_preview = info.get("title", url)[:50]
        print(f"\n[{i}/{len(targets)}] {title_preview}")
        print(f"  URL: {url}")

        article = scrape_article(url)

        if article:
            filepath = save_article(article)
            master[url]["scraped"] = True
            master[url]["file"]    = filepath
            master[url]["words"]   = article["words"]
            master[url]["title"]   = article["title"] or info.get("title", "")
            print(f"  ✅ {article['words']} từ → {os.path.basename(filepath)}")
            success += 1
        else:
            master[url]["failed"]  = True
            master[url]["scraped"] = False
            fail += 1

        # Lưu định kỳ mỗi 10 bài
        if i % 10 == 0:
            save_master(master)
            print(f"  💾 Auto-saved master ({i}/{len(targets)})")

        # Rate limit
        time.sleep(random.uniform(0.8, 2.0))

    save_master(master)

    total = len(master)
    scraped_total = sum(1 for v in master.values() if v.get("scraped"))
    print(f"\n✅ Batch hoàn tất: {success} thành công, {fail} thất bại")
    print(f"📊 Tổng: {total} links | {scraped_total} đã có nội dung | {total - scraped_total} chưa lấy")

    return success


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Eva.vn Article Content Scraper")
    parser.add_argument("--batch", type=int, default=50,
                        help="Số bài cần lấy nội dung mỗi lần chạy (mặc định 50)")
    parser.add_argument("--stats", action="store_true", help="Xem thống kê")
    args = parser.parse_args()

    if args.stats:
        master = load_master()
        scraped  = sum(1 for v in master.values() if v.get("scraped"))
        failed   = sum(1 for v in master.values() if v.get("failed"))
        pending  = len(master) - scraped - failed
        print(f"📊 Eva.vn Tâm sự Master Stats:")
        print(f"   Tổng links:      {len(master):,}")
        print(f"   Đã có nội dung:  {scraped:,}")
        print(f"   Chưa lấy:        {pending:,}")
        print(f"   Thất bại:        {failed:,}")
        print(f"   Đủ dùng (3/ngày): ~{scraped//3} ngày")
    else:
        scrape_batch(args.batch)
