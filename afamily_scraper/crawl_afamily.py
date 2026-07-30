"""
afamily_scraper/crawl_afamily.py
================================
Cào CÂU CHUYỆN TÂM SỰ GIA ĐÌNH 100% từ Afamily.vn (https://afamily.vn/tam-su-gia-dinh.chn).
Sử dụng Playwright (Headless):
- Cuộn trang & tự động click nút "Xem thêm" (a.load-more-btn) liên tục.
- Trích xuất đủ 3 thông tin: URL + Tiêu đề + Mô tả Sapo ngắn (Không vào trang con).
- Tự động tích hợp vào eva_scraper/links_master.json & giao diện Web Excel.

Chạy:
    python afamily_scraper/crawl_afamily.py --max-clicks 500
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

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

TARGET_URL = "https://afamily.vn/tam-su-gia-dinh.chn"
LINKS_FILE = os.path.join("eva_scraper", "links_master.json")


def load_master() -> dict:
    if os.path.exists(LINKS_FILE):
        with open(LINKS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_master(master: dict):
    with open(LINKS_FILE, "w", encoding="utf-8") as f:
        json.dump(master, f, ensure_ascii=False, indent=2)


def crawl_afamily_stories(max_clicks: int = 500):
    master = load_master()
    existing = set(master.keys())
    initial_count = len(existing)

    print(f"🚀 [Afamily Playwright] Mở trang {TARGET_URL}...")
    print(f"🔄 Sẽ cuộn & click nút 'Xem thêm' tối đa {max_clicks} lần...\n")

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
                # 1. Trích xuất bài viết hiện có trên trang Afamily
                articles = page.evaluate("""
                    () => {
                        const results = [];
                        const links = document.querySelectorAll('a[href*=".chn"]');

                        links.forEach(a => {
                            const href = a.href;
                            if (!href || href.includes('javascript') || href.includes('#')) return;

                            let parent = a.closest('li, article, div[class*="item"], div[class*="news"], div.knswli') || a.parentElement;
                            let title = a.innerText.trim();
                            let sapo = '';

                            if (parent) {
                                const titleEl = parent.querySelector('h2, h3, h4, .title, a.title, .knswli-title, strong');
                                if (titleEl && titleEl.innerText.trim()) title = titleEl.innerText.trim();

                                const sapoEl = parent.querySelector('.sapo, .knswli-sapo, p, .desc, .summary');
                                if (sapoEl) sapo = sapoEl.innerText.trim();
                            }

                            if (!sapo || sapo === title) {
                                if (parent) {
                                    const allText = parent.innerText.trim().split('\\n');
                                    for (let t of allText) {
                                        t = t.trim();
                                        if (t && t !== title && t.length > 20) {
                                            sapo = t;
                                            break;
                                        }
                                    }
                                }
                            }

                            if (href && title && title.length > 10 && sapo && sapo.length > 15 && sapo !== title) {
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

                    if not href.endswith(".chn"):
                        continue

                    # Loại trừ tin showbiz / quảng cáo trên Afamily nếu có
                    item_str = (title + " " + summary).lower()
                    if any(kw in item_str for kw in ['quảng cáo', 'mua sắm', 'khuyến mãi', 'trắc nghiệm']):
                        continue

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

                print(f"  [Afamily Click #{click_num}/{max_clicks}] ⚡ +{added_this_click} câu chuyện mới. Tổng kho: {len(master)} bài")

                if added_this_click > 0:
                    no_new_count = 0
                    save_master(master)
                else:
                    no_new_count += 1

                # 2. Tìm nút Xem thêm của Afamily
                page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                time.sleep(0.5)

                load_more_btn = (
                    page.query_selector("a.load-more-btn") or
                    page.query_selector("a.btn-readmore") or
                    page.query_selector(".load-more-cell a") or
                    page.query_selector("a:has-text('Xem thêm')")
                )

                if not load_more_btn:
                    page.evaluate("window.scrollBy(0, -500)")
                    time.sleep(0.5)
                    load_more_btn = page.query_selector("a.load-more-btn") or page.query_selector("a:has-text('Xem thêm')")

                if not load_more_btn:
                    if no_new_count >= 5:
                        print("  🛑 Không tìm thấy nút 'Xem thêm' trên Afamily nữa. Đã load hết bài!")
                        break
                    continue

                try:
                    load_more_btn.scroll_into_view_if_needed()
                    load_more_btn.click(force=True)
                    time.sleep(1.5)  # Chờ AJAX nạp thêm bài
                except Exception as e:
                    print(f"  [Click #{click_num}] ⚠️ Lỗi click Xem thêm: {e}")
                    time.sleep(1)

        except Exception as e:
            print(f"❌ Lỗi Playwright: {e}")

        finally:
            browser.close()

    save_master(master)
    new_added = len(master) - initial_count
    print(f"\n🎉 [HOÀN THÀNH AFAMILY] Thu thập thêm {new_added} câu chuyện Tâm sự Gia đình. Tổng kho: {len(master)} bài")

    # Rebuild Bảng Excel Web
    from eva_scraper.build_viewer import generate_html_viewer
    generate_html_viewer()

    return new_added


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Crawl Afamily.vn Tâm Sự Gia Đình Stories")
    parser.add_argument("--max-clicks", type=int, default=300, help="Số lần click Xem thêm tối đa")
    args = parser.parse_args()

    crawl_afamily_stories(max_clicks=args.max_clicks)
