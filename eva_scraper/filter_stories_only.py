"""
eva_scraper/filter_stories_only.py
===================================
LOẠI BỎ TOÀN BỘ CÁC CHUYÊN MỤC DẠNG TIN TỨC / SHOWBIZ / DẠY CON / DINH DƯỠNG.
Chỉ giữ lại duy nhất 100% CÂU CHUYỆN THỰC TẾ (Tâm sự, Tình yêu, Nghệ thuật làm vợ, Mẹ chồng nàng dâu, Bí mật phòng thế).

Các chuyên mục GIỮ LẠI (Pure Story Categories):
- Tâm sự (/tam-su/, /goc-tam-su/)
- Tình yêu - Giới tính (/tinh-yeu-gioi-tinh/)
- Chuyện tình yêu (/chuyen-tinh-yeu/)
- Nghệ thuật làm vợ (/nghe-thuat-lam-vo/)
- Mẹ chồng nàng dâu (/me-chong-nang-dau/)
- Bí mật phòng thế (/bi-mat-phong-the/)

Các chuyên mục BỊ LOẠI BỎ HOÀN TOÀN (News / Celebrity Categories):
- Dạy con (/day-con/)
- Gia đình / Nuôi con (/gia-dinh/, /nuoi-con/)
- Tin tức (/tin-tuc/)
- Dinh dưỡng (/dinh-duong/)
- Chuyện Eva (/chuyen-eva/)
- Làm mẹ (/lam-me/)

Chạy:
    python eva_scraper/filter_stories_only.py
"""

import os
import sys
import json

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

LINKS_FILE = os.path.join("eva_scraper", "links_master.json")

# Danh sách duy nhất các đường dẫn Chuyên mục CHUẨN CÂU CHUYỆN
ALLOWED_CATEGORIES = [
    '/tam-su/', '/tinh-yeu-gioi-tinh/', '/chuyen-tinh-yeu/',
    '/nghe-thuat-lam-vo/', '/me-chong-nang-dau/', '/bi-mat-phong-the/', '/goc-tam-su/'
]

# Chuyên mục Tin tức / Showbiz / Mẹ & Bé cần LOẠI BỎ HOÀN TOÀN
EXCLUDED_CATEGORIES = [
    '/day-con/', '/gia-dinh/', '/tin-tuc/', '/dinh-duong/',
    '/chuyen-eva/', '/nuoi-con/', '/lam-me/', '/day-con'
]

# Các từ khóa bài rác / trắc nghiệm / bói toán cần LOẠI BỎ
EXCLUDE_KEYWORDS = [
    'trắc nghiệm', 'trac-nghiem', 'con giáp', 'con giap', '12 con giáp',
    'tử vi', 'tu-vi', 'bói', 'phong thủy', 'phong thuy', 'cung hoàng đạo',
    'tướng số', 'tướng mặt', 'bức tranh', 'quạt', 'bàn tay', 'nốt ruồi',
    'hoa hậu', 'vtv', 'showbiz', 'sao việt', 'mỹ nhân', 'diễn viên'
]


def is_pure_story(url: str, info: dict) -> bool:
    url_lower = url.lower()
    title = info.get("title", "").lower()
    summary = info.get("summary", "").lower()

    # 1. Loại bỏ nếu thuộc Chuyên mục Tin tức / Showbiz
    if any(ex in url_lower for ex in EXCLUDED_CATEGORIES):
        return False

    # 2. Phải thuộc đúng Chuyên mục Câu chuyện
    if not any(allow in url_lower for allow in ALLOWED_CATEGORIES):
        return False

    # 3. Loại bỏ nếu chứa từ khóa trắc nghiệm / tin tức showbiz
    item_str = title + " " + summary + " " + url_lower
    for kw in EXCLUDE_KEYWORDS:
        if kw in item_str:
            return False

    return True


def filter_links_master():
    if not os.path.exists(LINKS_FILE):
        print("❌ File eva_scraper/links_master.json không tồn tại!")
        return

    with open(LINKS_FILE, "r", encoding="utf-8") as f:
        master = json.load(f)

    initial_total = len(master)
    clean_master = {}
    removed_count = 0

    for url, info in master.items():
        if is_pure_story(url, info):
            clean_master[url] = info
        else:
            removed_count += 1

    with open(LINKS_FILE, "w", encoding="utf-8") as f:
        json.dump(clean_master, f, ensure_ascii=False, indent=2)

    print(f"🧹 [LOẠI BỎ CHUYÊN MỤC TIN TỨC] Đã xóa {removed_count} bài viết dạng Tin tức / Showbiz / Dạy con!")
    print(f"✅ Kho CÂU CHUYỆN THẦN THỦY 100% còn lại: {len(clean_master)} / {initial_total} bài.")

    # Rebuild lại file giao diện Web stories_viewer.html
    from eva_scraper.build_viewer import generate_html_viewer
    generate_html_viewer()


if __name__ == "__main__":
    filter_links_master()
