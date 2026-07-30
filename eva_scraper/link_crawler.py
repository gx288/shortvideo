"""
eva_scraper/link_crawler.py
===========================
Thu thập link bài viết từ eva.vn/tam-su bằng cách:
1. Lấy trang đầu tiên
2. Parse AJAX endpoint từ nút "Xem thêm"
3. Loop gọi AJAX POST để lấy thêm trang
4. Lưu toàn bộ link vào links_master.json

Chạy:
    python eva_scraper/link_crawler.py --max-pages 50
    python eva_scraper/link_crawler.py --max-pages 200  # ~10k bài
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
AJAX_BASE   = "https://eva.vn/ajax"
OUTPUT_DIR  = os.path.join("eva_scraper", "data")
LINKS_FILE  = os.path.join("eva_scraper", "links_master.json")

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
    """
    Master file structure:
    {
      "url": {
        "title": "...",
        "scraped": false,
        "file": "",
        "added_at": "2026-07-30"
      }
    }
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    if os.path.exists(LINKS_FILE):
        with open(LINKS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_master(master: dict):
    with open(LINKS_FILE, "w", encoding="utf-8") as f:
        json.dump(master, f, ensure_ascii=False, indent=2)


# ─────────────────────────────────────────────────────────────────────────────
# PARSE ARTICLE LINKS FROM HTML
# ─────────────────────────────────────────────────────────────────────────────

def extract_links(html: str) -> list[dict]:
    """Trích xuất link bài viết từ HTML trang danh sách."""
    soup = BeautifulSoup(html, "lxml")
    results = []

    # Lấy tất cả article tags
    articles = soup.select("article")
    if not articles:
        # Fallback: tìm link trong div danh sách
        articles = soup.select("div.eva-cm-spec-h-sma, div.eva-cm-spec-h-big")

    for art in articles:
        # Lấy link đầu tiên trong article (link ảnh hoặc tiêu đề)
        links = art.select("a[href]")
        title_el = art.select_one("h2 a, h3 a, h4 a, .title a")

        url   = ""
        title = ""

        if title_el:
            url   = title_el.get("href", "")
            title = title_el.get_text(strip=True)
        elif links:
            url = links[0].get("href", "")

        if not url or "eva.vn" not in url and not url.startswith("/"):
            continue

        # Normalize URL
        if url.startswith("/"):
            url = "https://eva.vn" + url
        if not url.startswith("https://eva.vn"):
            continue

        results.append({"url": url, "title": title})

    return results


# ─────────────────────────────────────────────────────────────────────────────
# PARSE AJAX INFO FROM PAGE
# ─────────────────────────────────────────────────────────────────────────────

def parse_ajax_info(html: str) -> dict | None:
    """
    Tìm thông tin AJAX từ nút 'Xem thêm'.
    Kết quả: {"div_id": "...", "ajax_url": "...", "post_data": {...}, "current_page": 1}
    """
    soup = BeautifulSoup(html, "lxml")

    # Tìm nút Xem thêm
    load_more = (
        soup.find("a", class_="btn-load-h") or
        soup.find("a", string=re.compile(r"Xem\s+thêm", re.I)) or
        soup.find("a", attrs={"href": re.compile(r"AjaxAction", re.I)})
    )

    if not load_more:
        print("[AJAX] Không tìm thấy nút 'Xem thêm'. Có thể chỉ có 1 trang.")
        return None

    href = load_more.get("href", "")
    print(f"[AJAX] Tìm thấy nút: {href[:120]}")

    # Parse: AjaxAction('div_id', 'url', true, 'POST', '{json_data}')
    match = re.search(
        r"AjaxAction\(['\"]([^'\"]+)['\"],\s*['\"]([^'\"]+)['\"].*?['\"]POST['\"].*?'(\{[^']+\})'",
        href, re.DOTALL
    )
    if not match:
        # Thử pattern khác
        match = re.search(
            r"AjaxAction\(['\"]([^'\"]+)['\"],\s*['\"]([^'\"]+)['\"]",
            href
        )

    if not match:
        print(f"[AJAX] Không parse được AjaxAction: {href[:200]}")
        return None

    div_id   = match.group(1)
    ajax_url = match.group(2)
    post_str = match.group(3) if match.lastindex >= 3 else "{}"

    # Parse JSON post data
    try:
        # Chuẩn hóa: thay key không có quote
        post_str = re.sub(r'(\w+):', r'"\1":', post_str)
        post_data = json.loads(post_str)
    except Exception:
        post_data = {}

    return {
        "div_id":   div_id,
        "ajax_url": ajax_url if ajax_url.startswith("http") else f"https://eva.vn{ajax_url}",
        "post_data": post_data,
        "current_page": post_data.get("v_page", post_data.get("page", 1)),
    }


# ─────────────────────────────────────────────────────────────────────────────
# FETCH AJAX PAGE
# ─────────────────────────────────────────────────────────────────────────────

def fetch_ajax_page(ajax_url: str, post_data: dict, page: int) -> str | None:
    """Gọi AJAX POST để lấy HTML trang tiếp theo."""
    data = {**post_data, "v_page": page}
    try:
        r = requests.post(
            ajax_url,
            data=data,
            headers={**HEADERS, "X-Requested-With": "XMLHttpRequest"},
            timeout=15
        )
        r.raise_for_status()
        r.encoding = "utf-8"
        return r.text
    except Exception as e:
        print(f"[AJAX] Lỗi page {page}: {e}")
        return None


# ─────────────────────────────────────────────────────────────────────────────
# MAIN CRAWLER
# ─────────────────────────────────────────────────────────────────────────────

def crawl_links(max_pages: int = 50) -> int:
    """
    Thu thập link từ tất cả các trang.
    Trả về số link MỚI thêm vào master.
    """
    master = load_master()
    existing = set(master.keys())
    new_count = 0

    # ── Trang đầu tiên ────────────────────────────────────────────────────────
    print(f"[Crawler] Fetching trang đầu: {BASE_URL}")
    try:
        r = requests.get(BASE_URL, headers=HEADERS, timeout=15)
        r.raise_for_status()
        r.encoding = "utf-8"
        first_html = r.text
    except Exception as e:
        print(f"[Crawler] Lỗi trang đầu: {e}")
        return 0

    # Parse link trang đầu
    links = extract_links(first_html)
    print(f"[Crawler] Trang 1: {len(links)} bài")
    for item in links:
        if item["url"] not in existing:
            master[item["url"]] = {
                "title": item["title"],
                "scraped": False,
                "file": "",
                "added_at": datetime.now().strftime("%Y-%m-%d"),
            }
            existing.add(item["url"])
            new_count += 1

    # Parse AJAX info
    ajax_info = parse_ajax_info(first_html)
    if not ajax_info:
        print("[Crawler] Chỉ có 1 trang hoặc không tìm được AJAX endpoint.")
        save_master(master)
        return new_count

    ajax_url  = ajax_info["ajax_url"]
    post_data = ajax_info["post_data"]
    print(f"[Crawler] AJAX URL: {ajax_url}")
    print(f"[Crawler] POST data mẫu: {post_data}")

    # ── Loop các trang tiếp theo ───────────────────────────────────────────────
    for page in range(2, max_pages + 1):
        print(f"[Crawler] Fetching trang {page}/{max_pages}...", end=" ")
        html = fetch_ajax_page(ajax_url, post_data, page)

        if not html or len(html.strip()) < 100:
            print("→ Hết dữ liệu, dừng.")
            break

        links = extract_links(html)
        if not links:
            print("→ Không có bài mới, dừng.")
            break

        added = 0
        for item in links:
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

        print(f"→ {len(links)} bài, {added} mới. Tổng: {len(master)}")

        # Lưu định kỳ mỗi 10 trang
        if page % 10 == 0:
            save_master(master)

        # Rate limit
        time.sleep(random.uniform(0.5, 1.5))

    save_master(master)
    print(f"\n✅ Hoàn tất! Thêm {new_count} link mới. Tổng trong master: {len(master)}")
    return new_count


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Eva.vn Tâm sự Link Crawler")
    parser.add_argument("--max-pages", type=int, default=50,
                        help="Số trang tối đa (mỗi trang ~15 bài, 700 trang ≈ 10k bài)")
    parser.add_argument("--stats", action="store_true", help="Xem thống kê master file")
    args = parser.parse_args()

    if args.stats:
        master = load_master()
        scraped = sum(1 for v in master.values() if v.get("scraped"))
        not_scraped = len(master) - scraped
        print(f"📊 Master: {len(master)} links | {scraped} đã lấy nội dung | {not_scraped} chưa lấy")
    else:
        crawl_links(args.max_pages)
