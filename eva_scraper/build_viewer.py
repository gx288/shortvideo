"""
eva_scraper/build_viewer.py
============================
Tạo trang Web HTML dạng bảng Excel tương tác cao (Interactive Table Viewer)
hiển thị TOÀN BỘ bài viết cào được từ links_master.json.

Tính năng:
- Bảng kiểu Excel siêu gọn gàng (Compact Density), không tốn diện tích
- Tìm kiếm tức thì (Search Filter) theo tiêu đề, mô tả, link
- Lọc theo từng Chuyên mục (Filter by Category)
- Thống kê tổng số bài, chuyên mục
- Mở link gốc Eva.vn trong tab mới khi click

Chạy:
    python eva_scraper/build_viewer.py
"""

import os
import sys
import json
from datetime import datetime

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

LINKS_FILE   = os.path.join("eva_scraper", "links_master.json")
VIEWER_FILE  = os.path.join("eva_scraper", "stories_viewer.html")


def detect_category(url: str) -> str:
    if '/tam-su/' in url or '/tam-su-c' in url:
        return 'Tâm sự'
    elif '/chuyen-tinh-yeu/' in url or '/chuyen-tinh-yeu-c' in url:
        return 'Chuyện tình yêu'
    elif '/tinh-yeu-gioi-tinh/' in url or '/tinh-yeu-gioi-tinh-c' in url:
        return 'Tình yêu - Giới tính'
    elif '/me-chong-nang-dau/' in url or '/me-chong-nang-dau-c' in url:
        return 'Mẹ chồng nàng dâu'
    elif '/nghe-thuat-lam-vo/' in url or '/nghe-thuat-lam-vo-c' in url:
        return 'Nghệ thuật làm vợ'
    elif '/bi-mat-phong-the/' in url or '/bi-mat-phong-the-c' in url:
        return 'Bí mật phòng thế'
    elif '/gia-dinh/' in url or '/gia-dinh-c' in url:
        return 'Gia đình'
    elif '/day-con/' in url or '/day-con-c' in url:
        return 'Dạy con'
    elif '/goc-tam-su/' in url or '/goc-tam-su-c' in url:
        return 'Góc tâm sự'
    elif '/chuyen-eva/' in url or '/chuyen-eva-c' in url:
        return 'Chuyện Eva'
    return 'Tâm sự'


def generate_html_viewer():
    if not os.path.exists(LINKS_FILE):
        print("❌ Chưa tìm thấy file eva_scraper/links_master.json")
        return

    with open(LINKS_FILE, "r", encoding="utf-8") as f:
        master = json.load(f)

    # Chuyển master dict thành danh sách bài viết
    data_list = []
    idx = 1
    for url, info in master.items():
        title = info.get("title", "").strip() or "Untitled Story"
        summary = info.get("summary", "").strip() or title
        category = detect_category(url)

        data_list.append({
            "stt": idx,
            "url": url,
            "title": title,
            "summary": summary,
            "category": category,
            "added_at": info.get("added_at", datetime.now().strftime("%Y-%m-%d")),
        })
        idx += 1

    json_data = json.dumps(data_list, ensure_ascii=False)

    html_content = f"""<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Kho Bài Viết Eva.vn - ShortVideo Engine</title>
    <!-- Google Fonts: Inter -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {{
            --bg-body: #0f172a;
            --bg-card: #1e293b;
            --bg-table-stripe: #182234;
            --bg-hover: #334155;
            --border-color: #334155;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --accent-blue: #38bdf8;
            --accent-green: #4ade80;
            --accent-purple: #c084fc;
            --accent-amber: #fbbf24;
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }}

        body {{
            font-family: 'Inter', system-ui, -apple-system, sans-serif;
            background-color: var(--bg-body);
            color: var(--text-main);
            padding: 16px;
            font-size: 13px;
            line-height: 1.4;
        }}

        /* Header Bar */
        .header-bar {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 12px;
            padding: 12px 16px;
            background: var(--bg-card);
            border-radius: 8px;
            border: 1px solid var(--border-color);
        }}

        .title-area h1 {{
            font-size: 18px;
            font-weight: 700;
            color: var(--accent-blue);
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .title-area p {{
            font-size: 12px;
            color: var(--text-muted);
            margin-top: 2px;
        }}

        /* Stat Badges */
        .stats-group {{
            display: flex;
            gap: 12px;
        }}

        .stat-badge {{
            background: rgba(56, 189, 248, 0.1);
            border: 1px solid rgba(56, 189, 248, 0.3);
            padding: 6px 12px;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            color: var(--accent-blue);
        }}

        .stat-badge span {{
            color: #fff;
            font-size: 14px;
            font-weight: 700;
        }}

        /* Controls Bar (Filter & Search) */
        .controls-bar {{
            display: flex;
            gap: 12px;
            margin-bottom: 12px;
            flex-wrap: wrap;
        }}

        .search-box {{
            flex: 1;
            min-width: 250px;
            position: relative;
        }}

        .search-box input {{
            width: 100%;
            padding: 8px 12px;
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 6px;
            color: var(--text-main);
            font-size: 13px;
            outline: none;
        }}

        .search-box input:focus {{
            border-color: var(--accent-blue);
        }}

        .filter-select {{
            padding: 8px 12px;
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 6px;
            color: var(--text-main);
            font-size: 13px;
            outline: none;
            cursor: pointer;
        }}

        /* Excel-Style Table Wrapper */
        .table-container {{
            background: var(--bg-card);
            border-radius: 8px;
            border: 1px solid var(--border-color);
            overflow-x: auto;
            max-height: calc(100vh - 150px);
            position: relative;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            text-align: left;
            table-layout: fixed;
        }}

        thead {{
            position: sticky;
            top: 0;
            z-index: 10;
            background-color: #1e293b;
            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
        }}

        th {{
            padding: 10px 12px;
            font-weight: 600;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: var(--text-muted);
            border-bottom: 2px solid var(--border-color);
            border-right: 1px solid var(--border-color);
            background-color: #1e293b;
        }}

        td {{
            padding: 8px 12px;
            border-bottom: 1px solid var(--border-color);
            border-right: 1px solid var(--border-color);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }}

        tr:nth-child(even) {{
            background-color: var(--bg-table-stripe);
        }}

        tr:hover {{
            background-color: var(--bg-hover);
        }}

        /* Column Widths (Optimal Space Saving) */
        .col-stt {{ width: 50px; text-align: center; }}
        .col-cat {{ width: 150px; }}
        .col-title {{ width: 380px; }}
        .col-summary {{ width: 550px; }}
        .col-link {{ width: 90px; text-align: center; }}

        /* Category Tag Styling */
        .tag {{
            display: inline-block;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 600;
        }}
        .tag-tam-su {{ background: rgba(56, 189, 248, 0.15); color: #38bdf8; }}
        .tag-tinh-yeu {{ background: rgba(244, 63, 94, 0.15); color: #fb7185; }}
        .tag-gia-dinh {{ background: rgba(74, 222, 128, 0.15); color: #4ade80; }}
        .tag-me-chong {{ background: rgba(251, 191, 36, 0.15); color: #fbbf24; }}
        .tag-other {{ background: rgba(192, 132, 252, 0.15); color: #c084fc; }}

        .link-btn {{
            color: var(--accent-blue);
            text-decoration: none;
            font-weight: 600;
        }}
        .link-btn:hover {{
            text-decoration: underline;
        }}

        /* Custom Scrollbar */
        ::-webkit-scrollbar {{
            width: 8px;
            height: 8px;
        }}
        ::-webkit-scrollbar-track {{
            background: var(--bg-body);
        }}
        ::-webkit-scrollbar-thumb {{
            background: var(--border-color);
            border-radius: 4px;
        }}
    </style>
</head>
<body>

    <div class="header-bar">
        <div class="title-area">
            <h1>📊 Bảng Kho Bài Viết Eva.vn</h1>
            <p>Hiển thị danh sách câu chuyện, tiêu đề và mô tả sapo thu thập tự động</p>
        </div>
        <div class="stats-group">
            <div class="stat-badge">Tổng bài viết: <span id="stat-total">0</span></div>
            <div class="stat-badge" style="border-color: rgba(74, 222, 128, 0.3); background: rgba(74, 222, 128, 0.1); color: var(--accent-green);">
                Hiển thị: <span id="stat-showing" style="color: var(--accent-green);">0</span>
            </div>
        </div>
    </div>

    <div class="controls-bar">
        <div class="search-box">
            <input type="text" id="searchInput" placeholder="🔍 Tìm kiếm theo tiêu đề, mô tả hoặc từ khóa..." oninput="filterTable()">
        </div>
        <select class="filter-select" id="catFilter" onchange="filterTable()">
            <option value="ALL">All Categories (Tất cả chuyên mục)</option>
            <option value="Tâm sự">Tâm sự</option>
            <option value="Tình yêu - Giới tính">Tình yêu - Giới tính</option>
            <option value="Chuyện tình yêu">Chuyện tình yêu</option>
            <option value="Mẹ chồng nàng dâu">Mẹ chồng nàng dâu</option>
            <option value="Nghệ thuật làm vợ">Nghệ thuật làm vợ</option>
            <option value="Gia đình">Gia đình</option>
            <option value="Dạy con">Dạy con</option>
        </select>
    </div>

    <div class="table-container">
        <table>
            <thead>
                <tr>
                    <th class="col-stt">STT</th>
                    <th class="col-cat">Chuyên Mục</th>
                    <th class="col-title">Tiêu Đề Bài Viết</th>
                    <th class="col-summary">Mô Tả / Sapo Ngắn</th>
                    <th class="col-link">Link Gốc</th>
                </tr>
            </thead>
            <tbody id="tableBody">
                <!-- Dynamic Rows -->
            </tbody>
        </table>
    </div>

    <script>
        const rawData = {json_data};

        function getTagClass(cat) {{
            if (cat.includes('Tâm sự')) return 'tag-tam-su';
            if (cat.includes('Tình yêu')) return 'tag-tinh-yeu';
            if (cat.includes('Gia đình')) return 'tag-gia-dinh';
            if (cat.includes('Mẹ chồng')) return 'tag-me-chong';
            return 'tag-other';
        }}

        function renderTable(data) {{
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';

            data.forEach((item, index) => {{
                const tr = document.createElement('tr');
                const tagClass = getTagClass(item.category);

                tr.innerHTML = `
                    <td class="col-stt">${{index + 1}}</td>
                    <td class="col-cat"><span class="tag ${{tagClass}}">${{item.category}}</span></td>
                    <td class="col-title" title="${{escapeHtml(item.title)}}"><strong>${{escapeHtml(item.title)}}</strong></td>
                    <td class="col-summary" title="${{escapeHtml(item.summary)}}">${{escapeHtml(item.summary)}}</td>
                    <td class="col-link"><a href="${{item.url}}" target="_blank" class="link-btn">Xem Link 🔗</a></td>
                `;
                tbody.appendChild(tr);
            }});

            document.getElementById('stat-showing').innerText = data.length.toLocaleString();
        }}

        function escapeHtml(text) {{
            if (!text) return '';
            return text.replace(/&/g, "&amp;")
                       .replace(/</g, "&lt;")
                       .replace(/>/g, "&gt;")
                       .replace(/"/g, "&quot;")
                       .replace(/'/g, "&#039;");
        }}

        function filterTable() {{
            const query = document.getElementById('searchInput').value.toLowerCase().trim();
            const cat = document.getElementById('catFilter').value;

            const filtered = rawData.filter(item => {{
                const matchCat = (cat === 'ALL' || item.category === cat);
                const matchQuery = !query || 
                                   item.title.toLowerCase().includes(query) || 
                                   item.summary.toLowerCase().includes(query) ||
                                   item.url.toLowerCase().includes(query);
                return matchCat && matchQuery;
            }});

            renderTable(filtered);
        }}

        // Init
        document.getElementById('stat-total').innerText = rawData.length.toLocaleString();
        renderTable(rawData);
    </script>
</body>
</html>
"""

    with open(VIEWER_FILE, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"🎉 [BUILD HOÀN TẤT] Đã tạo file Web Bảng Excel: {VIEWER_FILE}")
    print(f"📊 Tổng số bài viết hiển thị: {len(data_list)} bài.")


if __name__ == "__main__":
    generate_html_viewer()
