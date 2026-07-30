"""
eva_scraper/crawl_all_subcategories.py
======================================
Cào HÀNG NGHÌN bài viết (URL + Tiêu đề + Mô tả Sapo trong #centerContent)
bằng cách quét lần lượt TẤT CẢ các chuyên mục con thuộc Tâm sự - Tình yêu - Gia đình trên Eva.vn.

Danh sách chuyên mục con:
1. https://eva.vn/tam-su-c391.html (Tâm sự)
2. https://eva.vn/nghe-thuat-lam-vo-c408.html (Nghệ thuật làm vợ)
3. https://eva.vn/me-chong-nang-dau-c210.html (Mẹ chồng nàng dâu)
4. https://eva.vn/bi-mat-phong-the-c54.html (Bí mật phòng the)
5. https://eva.vn/chuyen-tinh-yeu-c4.html (Chuyện tình yêu)
6. https://eva.vn/chuyen-eva-c391.html (Chuyện Eva)
7. https://eva.vn/goc-tam-su-c392.html (Góc tâm sự)
8. https://eva.vn/gia-dinh-c390.html (Gia đình)
9. https://eva.vn/day-con-c14.html (Dạy con)

Chạy:
    python eva_scraper/crawl_all_subcategories.py
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

LINKS_FILE = os.path.join("eva_scraper", "links_master.json")

# Danh sách tất cả chuyên mục con của Eva.vn
SUBCATEGORIES = [
    "https://eva.vn/tam-su-c391.html",
    "https://eva.vn/nghe-thuat-lam-vo-c408.html",
    "https://eva.vn/me-chong-nang-dau-c210.html",
    "https://eva.vn/bi-mat-phong-the-c54.html",
    "https://eva.vn/chuyen-tinh-yeu-c4.html",
    "https://eva.vn/chuyen-eva-c391.html",
    "https://eva.vn/goc-tam-su-c392.html",
    "https://eva.vn/gia-dinh-c390.html",
    "https://eva.vn/day-con-c14.html",
    "https://eva.vn/tinh-yeu-gioi-tinh-c3.html",
]

STORY_CATEGORIES = [
    '/tam-su/', '/tinh-yeu-gioi-tinh/', '/gia-dinh/', '/chuyen-eva/',
    '/goc-tam-su/', '/chuyen-tinh-yeu/', '/nghe-thuat-lam-vo/',
    '/me-chong-nang-dau/', '/bi-mat-phong-the/', '/day-con/'
]


def load_master() -> dict:
    if os.path.exists(LINKS_FILE):
        with open(LINKS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_master(master: dict):
    with open(LINKS_FILE, "w", encoding="utf-8") as f:
        json.dump(master, f, ensure_ascii=False, indent=2)


def crawl_subcategory(page, cat_url: str, master: dict, existing: set, max_clicks_per_cat: int = 100) -> int:
    print(f"\n📂 [CÀO CHUYÊN MỤC] {cat_url} ...")
    added_in_cat = 0

    try:
        page.goto(cat_url, wait_until="domcontentloaded", timeout=30000)
        time.sleep(1.5)

        no_new_count = 0

        for click_num in range(1, max_clicks_per_cat + 1):
            # 1. Trích xuất bài viết trong #centerContent
            articles = page.evaluate("""
                () => {
                    const center = document.querySelector('#centerContent') || document.querySelector('.centerContent') || document.body;
                    const results = [];
                    const links = center.querySelectorAll('a[href*=".html"]');

                    links.forEach(a => {
                        const href = a.href;
                        if (!href || href.includes('javascript') || href.includes('#')) return;

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
                        added_in_cat += 1

            if added_this_click > 0:
                print(f"  [Click #{click_num}] ⚡ +{added_this_click} bài mới. Tổng kho: {len(master)} bài")
                no_new_count = 0
                save_master(master)
            else:
                no_new_count += 1

            # 2. Tìm nút Xem thêm
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(0.5)

            load_more_btn = (
                page.query_selector("a[href*='javascript:AjaxAction']") or
                page.query_selector("a:has-text('XEM THÊM')") or
                page.query_selector("a:has-text('Xem thêm')")
            )

            if not load_more_btn:
                page.evaluate("window.scrollBy(0, -500)")
                time.sleep(0.5)
                load_more_btn = page.query_selector("a[href*='javascript:AjaxAction']") or page.query_selector("a:has-text('XEM THÊM')")

            if not load_more_btn:
                if no_new_count >= 3:
                    print(f"  🛑 Đã load hết bài trong chuyên mục {cat_url}")
                    break
                continue

            try:
                load_more_btn.scroll_into_view_if_needed()
                load_more_btn.click(force=True)
                time.sleep(1.5)
            except Exception:
                time.sleep(1)

    except Exception as e:
        print(f"  ❌ Lỗi khi cào chuyên mục {cat_url}: {e}")

    return added_in_cat


def crawl_all_categories(max_clicks_per_cat: int = 100):
    master = load_master()
    existing = set(master.keys())
    initial_count = len(existing)

    print(f"🚀 [Playwright] Bắt đầu cào HÀNG NGHÌN BÀI VIẾT qua {len(SUBCATEGORIES)} chuyên mục con...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        for cat_url in SUBCATEGORIES:
            crawl_subcategory(page, cat_url, master, existing, max_clicks_per_cat)

        browser.close()

    save_master(master)
    total_added = len(master) - initial_count
    print(f"\n🎉 [HOÀN THÀNH TOÀN BỘ CHUYÊN MỤC] Đã thu thập thêm {total_added} bài viết mới. TỔNG KHO MASTER: {len(master)} BÀI VIẾT!")
    return total_added


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Crawl thousands of stories across all Eva.vn subcategories")
    parser.add_argument("--max-clicks", type=int, default=100, help="Số lần click XEM THÊM tối đa mỗi chuyên mục")
    args = parser.parse_args()

    crawl_all_categories(max_clicks_per_cat=args.max_clicks)
