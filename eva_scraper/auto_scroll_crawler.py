"""
eva_scraper/auto_scroll_crawler.py
==================================
Sử dụng Playwright (Headless Browser) tự động:
1. Mở trang https://eva.vn/tam-su-p1638c3.html
2. Cuộn xuống cuối trang
3. Tự động click nút "Xem thêm" liên tục đến khi hết bài mới dừng
4. Trích xuất TOÀN BỘ link câu chuyện (chỉ lấy URL, KHÔNG tải nội dung bài)
5. Lưu vào eva_scraper/links_master.json

Chạy:
    python eva_scraper/auto_scroll_crawler.py --max-clicks 100
    python eva_scraper/auto_scroll_crawler.py --max-clicks 500  # Crawl hàng nghìn bài
"""

import os
import sys
import json
import time
import argparse
from datetime import datetime
from playwright.sync_api import sync_playwright

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

BASE_URL   = "https://eva.vn/tam-su-p1638c3.html"
LINKS_FILE = os.path.join("eva_scraper", "links_master.json")
STORY_CATEGORIES = ['/tam-su/', '/tinh-yeu-gioi-tinh/', '/gia-dinh/', '/chuyen-eva/', '/goc-tam-su/']


def load_master() -> dict:
    if os.path.exists(LINKS_FILE):
        with open(LINKS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_master(master: dict):
    with open(LINKS_FILE, "w", encoding="utf-8") as f:
        json.dump(master, f, ensure_ascii=False, indent=2)


def crawl_with_auto_scroll(max_clicks: int = 200):
    master = load_master()
    existing = set(master.keys())
    initial_count = len(existing)

    print(f"🚀 [Playwright] Mở trình duyệt ẩn (Headless) cào link từ {BASE_URL}...")
    print(f"🔄 Sẽ tự động cuộn trang & click nút 'Xem thêm' tối đa {max_clicks} lần...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        try:
            page.goto(BASE_URL, wait_until="domcontentloaded", timeout=30000)
            time.sleep(2)

            no_new_count = 0
            click_num = 0

            for click_num in range(1, max_clicks + 1):
                # Cuộn xuống cuối trang
                page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                time.sleep(1)

                # Tìm nút Xem thêm bằng nhiều selector
                load_more_btn = (
                    page.query_selector("a.btn-load-h") or
                    page.query_selector("a.view-more") or
                    page.query_selector("a:has-text('Xem thêm')") or
                    page.query_selector(".btn-view-more") or
                    page.query_selector("a[href*='javascript']")
                )

                if not load_more_btn:
                    # Thử cuộn chuột 500px xuống dưới rồi tìm lại
                    page.evaluate("window.scrollBy(0, 500)")
                    time.sleep(1)
                    load_more_btn = page.query_selector("a:has-text('Xem thêm')")

                if not load_more_btn:
                    print(f"  [Click #{click_num}] 🛑 Không tìm thấy thêm nút 'Xem thêm'. Thử cuộn tiếp...")
                    no_new_count += 1
                    if no_new_count >= 15:
                        print("  🛑 Hết nút Xem thêm hoàn toàn.")
                        break
                    continue

                try:
                    load_more_btn.scroll_into_view_if_needed()
                    load_more_btn.click(force=True)
                    time.sleep(1.8)  # Chờ AJAX tải bài
                except Exception as e:
                    print(f"  [Click #{click_num}] ⚠️ Lỗi click nút Xem thêm: {e}")
                    time.sleep(1)

                # Trích xuất toàn bộ link hiện có trong trang
                hrefs = page.eval_on_selector_all(
                    "a[href]",
                    "elements => elements.map(e => ({ href: e.href, title: e.innerText.trim() }))"
                )

                added_this_click = 0
                for item in hrefs:
                    href  = item.get("href", "")
                    title = item.get("title", "")

                    if not href.endswith(".html"):
                        continue

                    if any(cat in href for cat in STORY_CATEGORIES):
                        if href not in existing:
                            master[href] = {
                                "title": title[:100],
                                "scraped": False,
                                "file": "",
                                "added_at": datetime.now().strftime("%Y-%m-%d"),
                            }
                            existing.add(href)
                            added_this_click += 1

                print(f"  [Click #{click_num}/{max_clicks}] ⚡ Thêm {added_this_click} link mới. Tổng kho: {len(master)} link")

                if added_this_click > 0:
                    no_new_count = 0
                    save_master(master)

        except Exception as e:
            print(f"❌ Lỗi trong quá trình chạy Playwright: {e}")

        finally:
            browser.close()

    save_master(master)
    new_added = len(master) - initial_count
    print(f"\n🎉 [HOÀN THÀNH PLAYWRIGHT] Thêm {new_added} link mới. Tổng link câu chuyện trong links_master.json: {len(master)}")
    return new_added


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Auto Scroll & Load More Button Link Crawler for Eva.vn")
    parser.add_argument("--max-clicks", type=int, default=100, help="Số lần click nút Xem thêm tối đa")
    args = parser.parse_args()

    crawl_with_auto_scroll(max_clicks=args.max_clicks)
