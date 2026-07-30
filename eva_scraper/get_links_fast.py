"""
eva_scraper/get_links_fast.py
==============================
Chỉ cào duy nhất DANH SÁCH LINK (không tải nội dung, không chạy ngầm gây lag).
Chạy nhẹ nhàng trong 2-3 giây.
"""

import os
import sys
import json
import requests
from bs4 import BeautifulSoup
from datetime import datetime

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

LINKS_FILE = os.path.join("eva_scraper", "links_master.json")

CATEGORIES = [
    "https://eva.vn/tam-su-c3.html",
    "https://eva.vn/chuyen-eva-c391.html",
    "https://eva.vn/goc-tam-su-c392.html",
    "https://eva.vn/tinh-yeu-gioi-tinh-c3a1638.html",
    "https://eva.vn/gia-dinh-c390.html",
    "https://eva.vn/tam-su-p1638c3.html"
]

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}


def load_master() -> dict:
    if os.path.exists(LINKS_FILE):
        with open(LINKS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_master(master: dict):
    with open(LINKS_FILE, "w", encoding="utf-8") as f:
        json.dump(master, f, ensure_ascii=False, indent=2)


def get_links_only():
    master = load_master()
    existing = set(master.keys())
    new_count = 0

    print("🚀 [Fast Link Collector] Bắt đầu quét danh sách link (Chỉ lấy URL, nhẹ 100%)...")

    # Quét các trang danh mục chính + các trang phân trang 1->10
    urls_to_scrape = []
    for cat in CATEGORIES:
        urls_to_scrape.append(cat)
        for page in range(1, 10):
            urls_to_scrape.append(f"{cat}?page={page}")

    for url in urls_to_scrape:
        try:
            r = requests.get(url, headers=HEADERS, timeout=5)
            if r.status_code != 200:
                continue
            soup = BeautifulSoup(r.text, "lxml")
            for a in soup.find_all("a", href=True):
                href = a["href"]
                if not href.endswith(".html"):
                    continue
                if href.startswith("/"):
                    href = "https://eva.vn" + href
                if any(k in href for k in ["/tam-su/", "/tinh-yeu-gioi-tinh/", "/gia-dinh/", "/chuyen-eva/", "/goc-tam-su/"]):
                    if href not in existing:
                        title = a.get_text(strip=True)
                        master[href] = {
                            "title": title[:100],
                            "scraped": False,
                            "file": "",
                            "added_at": datetime.now().strftime("%Y-%m-%d"),
                        }
                        existing.add(href)
                        new_count += 1
        except Exception:
            pass

    save_master(master)
    print(f"✅ Hoàn tất! Thêm {new_count} link mới. Tổng link trong links_master.json: {len(master)}")
    return new_count


if __name__ == "__main__":
    get_links_only()
