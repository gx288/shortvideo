"""
eva_scraper/auto_scroll_crawler.py
==================================
Tự động cuộn trang & click nút "XEM THÊM" chuẩn xác trên Eva.vn:
- Trang mục tiêu: https://eva.vn/tinh-yeu-gioi-tinh-c3.html
- Chỉ lấy các bài viết trong khối #centerContent
- Trích xuất đủ 3 thông tin: URL + Tiêu đề + Mô tả Sapo (KHÔNG vào trang con)
- Lưu liên tục vào eva_scraper/links_master.json

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

TARGET_URL = "https://eva.vn/tinh-yeu-gioi-tinh-c3.html"
LINKS_FILE = os.path.join("eva_scraper", "links_master.json")
STORY_CATEGORIES = ['/tam-su/', '/tinh-yeu-gioi-tinh/', '/gia-dinh/', '/chuyen-eva/', '/goc-tam-su/', '/chuyen-tinh-yeu/']


def load_master() -> dict:
    if os.path.exists(LINKS_FILE):
        with open(LINKS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_master(master: dict):
    with open(LINKS_FILE, "w", encoding="utf-8") as f:
        json.dump(master, f, ensure_ascii=False, indent=2)


def crawl_center_content_stories(max_clicks: int = 1000):
    master = load_master()
    existing = set(master.keys())
    initial_count = len(existing)

    print(f"🚀 [Playwright] Mở trang {TARGET_URL} (Lọc duy nhất trong khối #centerContent)...")
    print(f"🔄 Sẽ cuộn & click nút 'XEM THÊM' tối đa {max_clicks} lần...\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        try:
            page.goto(TARGET_URL, wait_until="domcontentloaded", timeout=30000)
            time.sleep(2)

            no_new_count = 0

            for click_num in range(1, max_clicks + 1):
                # 1. Trích xuất toàn bộ bài viết trong #centerContent hiện tại
                articles = page.evaluate("""
                    () => {
                        const center = document.querySelector('#centerContent') || document.querySelector('.centerContent');
                        if (!center) return [];

                        const results = [];
                        const links = center.querySelectorAll('a[href*=".html"]');

                        links.forEach(a => {
                            const href = a.href;
                            if (!href || href.includes('javascript') || href.includes('#')) return;

                            // Tìm thẻ chứa bao ngoài
                            let parent = a.closest('div[class*="item"], article, li, div.eva-news-trend-h-list') || a.parentElement;
                            let title = a.innerText.trim();
                            let sapo = '';

                            if (parent) {
                                const titleEl = parent.querySelector('h2, h3, h4, .title, a.title, strong');
                                if (titleEl && titleEl.innerText.trim()) title = titleEl.innerText.trim();

                                const sapoEl = parent.querySelector('.sapo, .kat-sapo, p, .desc, .summary, .txt-sapo');
                                if (sapoEl) sapo = sapoEl.innerText.trim();
                            }

                            if (!sapo || sapo === title) {
                                // Lấy đoạn văn bản bổ sung trong thẻ nếu không có class .sapo
                                if (parent) {
                                    const allText = parent.innerText.trim().split('\\n');
                                    for (let t of allText) {
                                        t = t.strip ? t.strip() : t.trim();
                                        if (t && t !== title && t.length > 20) {
                                            sapo = t;
                                            break;
                                        }
                                    }
                                }
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

                print(f"  [Click #{click_num}/{max_clicks}] ⚡ Thu thập {added_this_click} bài mới. Tổng kho: {len(master)} bài (có Tiêu đề + Mô tả)")

                if added_this_click > 0:
                    no_new_count = 0
                    save_master(master)
                else:
                    no_new_count += 1

                # 2. Tìm nút XEM THÊM trong #centerContent
                page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                time.sleep(0.5)

                load_more_btn = (
                    page.query_selector("#centerContent a[href*='javascript:AjaxAction']") or
                    page.query_selector("a[href*='javascript:AjaxAction']") or
                    page.query_selector("#centerContent a:has-text('XEM THÊM')") or
                    page.query_selector("a:has-text('XEM THÊM')")
                )

                if not load_more_btn:
                    page.evaluate("window.scrollBy(0, -600)")
                    time.sleep(0.5)
                    load_more_btn = page.query_selector("a[href*='javascript:AjaxAction']") or page.query_selector("a:has-text('XEM THÊM')")

                if not load_more_btn:
                    if no_new_count >= 5:
                        print("  🛑 Không tìm thấy nút 'XEM THÊM' trong #centerContent nữa. Đã load hết bài!")
                        break
                    continue

                try:
                    load_more_btn.scroll_into_view_if_needed()
                    load_more_btn.click(force=True)
                    time.sleep(1.8)  # Chờ AJAX nạp thêm 10 bài mới vào DOM
                except Exception as e:
                    print(f"  [Click #{click_num}] ⚠️ Lỗi click XEM THÊM: {e}")
                    time.sleep(1)

        except Exception as e:
            print(f"❌ Lỗi Playwright: {e}")

        finally:
            browser.close()

    save_master(master)
    new_added = len(master) - initial_count
    print(f"\n🎉 [HOÀN THÀNH CÀO #centerContent] Thu thập thêm {new_added} bài viết có Tiêu đề & Mô tả Sapo. Tổng kho: {len(master)} bài")
    return new_added


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Scrape #centerContent stories from Eva.vn/tinh-yeu-gioi-tinh-c3.html")
    parser.add_argument("--max-clicks", type=int, default=1000, help="Số lần click XEM THÊM tối đa")
    args = parser.parse_args()

    crawl_center_content_stories(max_clicks=args.max_clicks)
