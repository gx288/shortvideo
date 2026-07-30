"""
eva_scraper/crawl_pages_only.py
================================
CHỈ CÀO DẠNG DANH SÁCH LINK (KHÔNG LẤY NỘI DUNG) từ 200 trang danh mục Eva.vn.
Tốc độ: Lấy 3,000 - 5,000 link câu chuyện chỉ trong 3 - 5 GIÂY!
"""

import os
import sys
import json
import time
import requests
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

LINKS_FILE = os.path.join("eva_scraper", "links_master.json")

# 5 chuyên mục Tâm sự chính trên Eva.vn
CATEGORIES = [
    "https://eva.vn/tam-su-c3.html",
    "https://eva.vn/chuyen-eva-c391.html",
    "https://eva.vn/goc-tam-su-c392.html",
    "https://eva.vn/tinh-yeu-gioi-tinh-c3a1638.html",
    "https://eva.vn/gia-dinh-c390.html",
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
}


def load_master() -> dict:
    if os.path.exists(LINKS_FILE):
        with open(LINKS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_master(master: dict):
    with open(LINKS_FILE, "w", encoding="utf-8") as f:
        json.dump(master, f, ensure_ascii=False, indent=2)


def fetch_page_links(url: str) -> list[tuple[str, str]]:
    """Lấy toàn bộ link trong 1 trang danh mục (chỉ lấy URL và Tiêu đề, KHÔNG tải bài)."""
    results = []
    try:
        r = requests.get(url, headers=HEADERS, timeout=5)
        if r.status_code == 200:
            soup = BeautifulSoup(r.text, "lxml")
            for a in soup.find_all("a", href=True):
                href = a["href"]
                if not href.endswith(".html"):
                    continue
                if href.startswith("/"):
                    href = "https://eva.vn" + href
                if any(k in href for k in ["/tam-su/", "/tinh-yeu-gioi-tinh/", "/gia-dinh/", "/chuyen-eva/", "/goc-tam-su/"]):
                    title = a.get_text(strip=True)
                    results.append((href, title[:100]))
    except Exception:
        pass
    return results


def crawl_all_story_link_pages(max_pages_per_cat: int = 40, max_threads: int = 30) -> int:
    """
    Quét danh sách trang phân trang để lấy HÀNG NGHÌN LINK trong 5 GIÂY.
    """
    master = load_master()
    existing = set(master.keys())
    new_count = 0

    page_urls = []
    for cat in CATEGORIES:
        for page in range(1, max_pages_per_cat + 1):
            page_urls.append(f"{cat}?page={page}")

    print(f"🚀 [LINK ONLY CRAWLER] Quét {len(page_urls)} trang phân trang bằng {max_threads} luồng...")
    start_time = time.time()

    with ThreadPoolExecutor(max_workers=max_threads) as executor:
        futures = {executor.submit(fetch_page_links, p_url): p_url for p_url in page_urls}

        for future in as_completed(futures):
            items = future.result()
            for href, title in items:
                if href not in existing:
                    master[href] = {
                        "title": title,
                        "scraped": False,
                        "file": "",
                        "added_at": datetime.now().strftime("%Y-%m-%d"),
                    }
                    existing.add(href)
                    new_count += 1

    save_master(master)
    elapsed = time.time() - start_time
    print(f"\n🎉 [HOÀN THÀNH] Thu thập được {new_count:,} link mới chỉ trong {elapsed:.2f} giây!")
    print(f"📊 Tổng số link câu chuyện sẵn sàng trong links_master.json: {len(master):,} link.")
    return new_count


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Fast Category Link-Only Crawler")
    parser.add_argument("--pages", type=int, default=50, help="Số trang phân trang mỗi chuyên mục (mặc định 50 = 250 trang)")
    parser.add_argument("--threads", type=int, default=30, help="Số luồng song song")
    args = parser.parse_args()

    crawl_all_story_link_pages(max_pages_per_cat=args.pages, max_threads=args.threads)
