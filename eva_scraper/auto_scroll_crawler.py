"""
eva_scraper/auto_scroll_crawler.py
==================================
Sử dụng Playwright (Headless Browser) tự động:
1. Mở trang https://eva.vn/tam-su-p1638c3.html
2. Cuộn xuống cuối trang
3. Tự động click nút "Xem thêm" liên tục đến khi hết bài mới dừng (Tối đa 1,000 lần)
4. Trích xuất TOÀN BỘ (URL + Tiêu đề + Mô tả Sapo) trực tiếp từ các thẻ trên trang danh mục
5. Lưu trực tiếp vào eva_scraper/links_master.json

Chạy:
    python eva_scraper/auto_scroll_crawler.py --max-clicks 1000
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


def crawl_with_auto_scroll(max_clicks: int = 1000):
    master = load_master()
    existing = set(master.keys())
    initial_count = len(existing)

    print(f"🚀 [Playwright] Mở trình duyệt ẩn (Headless) cào link + tiêu đề + mô tả từ {BASE_URL}...")
    print(f"🔄 Sẽ tự động cuộn trang & click nút 'Xem thêm' tối đa {max_clicks} lần...\n")

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

            for click_num in range(1, max_clicks + 1):
                # Cuộn xuống cuối trang
                page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                time.sleep(0.8)

                # Tìm nút Xem thêm
                load_more_btn = (
                    page.query_selector("a.btn-load-h") or
                    page.query_selector("a.view-more") or
                    page.query_selector("a:has-text('Xem thêm')") or
                    page.query_selector(".btn-view-more")
                )

                if not load_more_btn:
                    page.evaluate("window.scrollBy(0, 400)")
                    time.sleep(0.8)
                    load_more_btn = page.query_selector("a:has-text('Xem thêm')")

                if not load_more_btn:
                    no_new_count += 1
                    if no_new_count >= 10:
                        print("  🛑 Hết nút Xem thêm hoàn toàn.")
                        break
                    continue

                try:
                    load_more_btn.scroll_into_view_if_needed()
                    load_more_btn.click(force=True)
                    time.sleep(1.5)  # Chờ AJAX tải thêm bài
                except Exception:
                    time.sleep(1)

                # Trích xuất URL + Tiêu đề + Mô tả Sapo trực tiếp trên trang danh mục
                articles = page.evaluate("""
                    () => {
                        const results = [];
                        const links = document.querySelectorAll('a[href*=".html"]');

                        links.forEach(a => {
                            const href = a.href;
                            if (!href || (!href.includes('/tam-su/') && !href.includes('/chuyen-eva/') && !href.includes('/goc-tam-su/'))) return;

                            let parent = a.closest('.eva-cont-kat__item, .kat-art-item, article, div.news-item, div[class*="item"]') || a.parentElement;
                            let title = a.innerText.trim();
                            let sapo = '';

                            if (parent) {
                                const titleEl = parent.querySelector('h2, h3, h4, .title, a.title');
                                if (titleEl) title = titleEl.innerText.trim();

                                const sapoEl = parent.querySelector('.eva-cont-kat__info, .sapo, .kat-sapo, .desc, .summary, p');
                                if (sapoEl) sapo = sapoEl.innerText.trim();
                            }

                            if (!sapo || sapo === title) {
                                sapo = title;
                            }

                            if (href && title && title.length > 5) {
                                results.push({ href: href, title: title, summary: sapo });
                            }
                        });
                        return results;
                    }
                """)

                added_this_click = 0
                for item in articles:
                    href    = item.get("href", "")
                    title   = item.get("title", "")
                    summary = item.get("summary", "")

                    if not href.endswith(".html"):
                        continue

                    if any(cat in href for cat in STORY_CATEGORIES):
                        if href not in existing or not master.get(href, {}).get("summary"):
                            master[href] = {
                                "title": title[:150],
                                "summary": summary[:300],
                                "scraped": False,
                                "file": "",
                                "added_at": datetime.now().strftime("%Y-%m-%d"),
                            }
                            existing.add(href)
                            added_this_click += 1

                print(f"  [Click #{click_num}/{max_clicks}] ⚡ Thêm {added_this_click} link mới (có Tiêu đề & Mô tả). Tổng kho: {len(master)} link")

                if added_this_click > 0:
                    no_new_count = 0
                    save_master(master)

        except Exception as e:
            print(f"❌ Lỗi Playwright: {e}")

        finally:
            browser.close()

    save_master(master)
    new_added = len(master) - initial_count
    print(f"\n🎉 [HOÀN THÀNH PLAYWRIGHT] Thêm {new_added} link mới. Tổng link câu chuyện trong links_master.json: {len(master)}")
    return new_added


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Auto Scroll & Load More Button Link Crawler for Eva.vn")
    parser.add_argument("--max-clicks", type=int, default=1000, help="Số lần click nút Xem thêm tối đa")
    args = parser.parse_args()

    crawl_with_auto_scroll(max_clicks=args.max_clicks)
