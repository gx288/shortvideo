"""
eva_scraper/link_crawler.py
===========================
Thu thập link bài viết từ eva.vn/tam-su bằng:
1. Crawl trang danh sách chính & các chuyên mục con
2. Quét dải ID bài viết trực tiếp (ID Router: /tam-su/x-c391a{ID}.html)

Chạy:
    python eva_scraper/link_crawler.py --max-pages 20
    python eva_scraper/link_crawler.py --scan-ids --start-id 678450 --count 200
"""

import os
import sys
import re
import json
import time
import random
import argparse
import requests
from bs4 import BeautifulSoup
from datetime import datetime

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# ─────────────────────────────────────────────────────────────────────────────
BASE_URL    = "https://eva.vn/tam-su-p1638c3.html"
OUTPUT_DIR  = os.path.join("eva_scraper", "data")
LINKS_FILE  = os.path.join("eva_scraper", "links_master.json")

CATEGORIES = [
    ("Tâm sự", "https://eva.vn/tam-su-c3.html"),
    ("Chuyện eva", "https://eva.vn/chuyen-eva-c391.html"),
    ("Góc tâm sự", "https://eva.vn/goc-tam-su-c392.html"),
    ("Tình yêu giới tính", "https://eva.vn/tinh-yeu-gioi-tinh-c3a1638.html"),
    ("Gia đình", "https://eva.vn/gia-dinh-c390.html"),
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/125.0.0.0 Safari/537.36",
    "Accept-Language": "vi-VN,vi;q=0.9",
    "Referer": "https://eva.vn/",
}


# ─────────────────────────────────────────────────────────────────────────────
# LOAD / SAVE MASTER
# ─────────────────────────────────────────────────────────────────────────────

def load_master() -> dict:
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    if os.path.exists(LINKS_FILE):
        with open(LINKS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_master(master: dict):
    with open(LINKS_FILE, "w", encoding="utf-8") as f:
        json.dump(master, f, ensure_ascii=False, indent=2)


# ─────────────────────────────────────────────────────────────────────────────
# METHOD 1: CATEGORY PAGE CRAWLER
# ─────────────────────────────────────────────────────────────────────────────

def extract_links(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "lxml")
    results = []
    articles = soup.select("article") or soup.select("div.eva-cm-spec-h-sma, div.eva-cm-spec-h-big")

    for art in articles:
        links = art.select("a[href]")
        title_el = art.select_one("h2 a, h3 a, h4 a, .title a")
        url   = title_el.get("href", "") if title_el else (links[0].get("href", "") if links else "")
        title = title_el.get_text(strip=True) if title_el else ""

        if not url or ".html" not in url:
            continue
        if url.startswith("/"):
            url = "https://eva.vn" + url
        if not url.startswith("https://eva.vn"):
            continue

        results.append({"url": url, "title": title})

    return results


def crawl_categories() -> int:
    master = load_master()
    existing = set(master.keys())
    new_count = 0

    print("[Crawler] Thu thập link từ các chuyên mục Tâm sự...")
    for name, cat_url in CATEGORIES:
        try:
            r = requests.get(cat_url, headers=HEADERS, timeout=15)
            r.raise_for_status()
            r.encoding = "utf-8"
            items = extract_links(r.text)
            added = 0
            for item in items:
                if item["url"] not in existing:
                    master[item["url"]] = {
                        "title": item["title"],
                        "scraped": False,
                        "file": "",
                        "added_at": datetime.now().strftime("%Y-%m-%d"),
                    }
                    existing.add(item["url"])
                    new_count += 1
                    added += 1
            print(f"  Chuyên mục '{name}': {len(items)} bài tìm thấy ({added} mới)")
        except Exception as e:
            print(f"  ❌ Lỗi '{name}': {e}")

    save_master(master)
    print(f"\n✅ Tổng link hiện có trong links_master.json: {len(master)}")
    return new_count


# ─────────────────────────────────────────────────────────────────────────────
# METHOD 2: DIRECT ID SCANNER (Fast & Complete)
# ─────────────────────────────────────────────────────────────────────────────

def scan_ids_range(start_id: int = 678450, count: int = 200) -> int:
    """
    Quét trực tiếp dải ID bài viết trên eva.vn.
    Nhận diện bài thuộc nhóm Tâm sự và tự động lưu vào master.
    """
    master = load_master()
    existing = set(master.keys())
    new_count = 0

    story_categories = ['/tam-su/', '/tinh-yeu-gioi-tinh/', '/gia-dinh/', '/chuyen-eva/', '/goc-tam-su/']

    print(f"\n[ID Scanner] Quét {count} ID bài viết từ {start_id} lùi dần...")

    for i, art_id in enumerate(range(start_id, start_id - count, -1), 1):
        url = f"https://eva.vn/tam-su/x-c391a{art_id}.html"
        try:
            r = requests.get(url, headers=HEADERS, allow_redirects=True, timeout=5)
            final_url = r.url

            # Kiểm tra xem có phải bài viết thuộc chuyên mục tâm sự không
            if any(cat in final_url for cat in story_categories) and ".html" in final_url:
                if final_url not in existing:
                    # Lấy tiêu đề từ bài viết
                    soup = BeautifulSoup(r.text, "lxml")
                    h1 = soup.select_one("h1.fw-bold, h1.color-main, h1")
                    title = h1.get_text(strip=True) if h1 else ""

                    master[final_url] = {
                        "title": title,
                        "scraped": False,
                        "file": "",
                        "added_at": datetime.now().strftime("%Y-%m-%d"),
                    }
                    existing.add(final_url)
                    new_count += 1
                    print(f"  [{i}/{count}] ✅ Story (ID {art_id}): {title[:50]}")
                else:
                    print(f"  [{i}/{count}] ⏭️ ID {art_id} đã có trong master")

        except Exception as e:
            print(f"  [{i}/{count}] ❌ ID {art_id} lỗi: {e}")

        # Lưu định kỳ mỗi 20 bài
        if i % 20 == 0:
            save_master(master)

        time.sleep(random.uniform(0.1, 0.3))

    save_master(master)
    print(f"\n✅ Hoàn tất quét ID: Thêm {new_count} link mới. Tổng master: {len(master)}")
    return new_count


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Eva.vn Tâm sự Link Crawler")
    parser.add_argument("--scan-ids", action="store_true", help="Dùng phương pháp quét ID trực tiếp")
    parser.add_argument("--start-id", type=int, default=678450, help="ID bắt đầu quét")
    parser.add_argument("--count", type=int, default=100, help="Số lượng ID quét")
    parser.add_argument("--stats", action="store_true", help="Xem thống kê master file")
    args = parser.parse_args()

    if args.stats:
        master = load_master()
        scraped = sum(1 for v in master.values() if v.get("scraped"))
        pending = len(master) - scraped
        print(f"📊 Master: {len(master)} links | {scraped} đã lấy nội dung | {pending} chưa lấy")
    elif args.scan_ids:
        scan_ids_range(args.start_id, args.count)
    else:
        crawl_categories()
