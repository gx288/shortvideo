"""
eva_scraper/build_viewer.py
============================
Tạo trang Web HTML dạng bảng Excel tương tác cao (Interactive Table Viewer)
phân biệt RÕ RÀNG NGUỒN TRANG (Eva.vn vs Afamily.vn) không bị lộn xộn.

Tính năng:
- Bảng kiểu Excel siêu gọn gàng (Compact Density)
- Phân loại rõ từng Trang nguồn: Eva.vn | Afamily.vn
- Bộ lọc theo Nguồn trang (Filter by Source) & Chuyên mục
- Tìm kiếm tức thì theo tiêu đề, mô tả, link

Chạy:
    python eva_scraper/build_viewer.py
"""

import os
import sys
import json
from datetime import datetime

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

LINKS_FILE      = os.path.join("eva_scraper", "links_master.json")
EVA_LINKS_FILE  = os.path.join("eva_scraper", "eva_links.json")
AFAMILY_FILE    = os.path.join("afamily_scraper", "afamily_links.json")
VIEWER_FILE     = os.path.join("eva_scraper", "stories_viewer.html")


def detect_source(url: str) -> str:
    if 'afamily.vn' in url:
        return 'Afamily.vn'
    return 'Eva.vn'


def detect_category(url: str) -> str:
    if 'afamily.vn' in url:
        return 'Tâm sự Gia đình'
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
    elif '/goc-tam-su/' in url or '/goc-tam-su-c' in url:
        return 'Góc tâm sự'
    return 'Tâm sự'


def split_source_files(master: dict):
    """Tách biệt master dict ra 2 file json riêng biệt theo từng trang."""
    eva_dict = {}
    afamily_dict = {}

    for url, info in master.items():
        src = detect_source(url)
        info["source"] = src
        if src == "Afamily.vn":
            afamily_dict[url] = info
        else:
            eva_dict[url] = info

    os.makedirs("afamily_scraper", exist_ok=True)
    with open(EVA_LINKS_FILE, "w", encoding="utf-8") as f:
        json.dump(eva_dict, f, ensure_ascii=False, indent=2)

    with open(AFAMILY_FILE, "w", encoding="utf-8") as f:
        json.dump(afamily_dict, f, ensure_ascii=False, indent=2)

    print(f"📁 [TÁCH NGUỒN FILE] Eva.vn: {len(eva_dict)} bài → {EVA_LINKS_FILE}")
    print(f"📁 [TÁCH NGUỒN FILE] Afamily.vn: {len(afamily_dict)} bài → {AFAMILY_FILE}")


def generate_html_viewer():
    if not os.path.exists(LINKS_FILE):
        print("❌ Chưa tìm thấy file eva_scraper/links_master.json")
        return

    with open(LINKS_FILE, "r", encoding="utf-8") as f:
        master = json.load(f)

    # Tách file riêng
    split_source_files(master)

    # Chuyển master dict thành danh sách bài viết
    data_list = []
    idx = 1
    for url, info in master.items():
        title = info.get("title", "").strip() or "Untitled Story"
        summary = info.get("summary", "").strip() or title
        source = detect_source(url)
        category = info.get("category") or detect_category(url)

        data_list.append({
            "stt": idx,
            "source": source,
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
    <title>Kho Bài Viết Tâm Sự (Eva.vn & Afamily.vn)</title>
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
            --accent-pink: #f43f5e;
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
            padding: 10px 12px;
            border-bottom: 1px solid var(--border-color);
            border-right: 1px solid var(--border-color);
            vertical-align: top;
        }}

        tr:nth-child(even) {{
            background-color: var(--bg-table-stripe);
        }}

        tr:hover {{
            background-color: var(--bg-hover);
        }}

        /* Column Widths (Tự động co giãn vừa khít 100% màn hình) */
        .col-stt {{ width: 50px; text-align: center; white-space: nowrap; }}
        .col-source {{ width: 110px; text-align: center; white-space: nowrap; }}
        .col-cat {{ width: 140px; white-space: nowrap; }}
        .col-title {{ 
            width: 35%; 
            white-space: normal !important; 
            word-break: break-word; 
            line-height: 1.45;
            color: #f1f5f9;
        }}
        .col-summary {{ 
            width: auto; 
            white-space: normal !important; 
            word-break: break-word; 
            line-height: 1.5;
            color: #cbd5e1;
        }}
        .col-link {{ width: 90px; text-align: center; white-space: nowrap; }}

        /* Tag Styling */
        .tag {{
            display: inline-block;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 600;
        }}
        .src-eva {{ background: rgba(56, 189, 248, 0.2); color: #38bdf8; border: 1px solid rgba(56, 189, 248, 0.4); }}
        .src-afamily {{ background: rgba(244, 63, 94, 0.2); color: #fb7185; border: 1px solid rgba(244, 63, 94, 0.4); }}

        .tag-tam-su {{ background: rgba(56, 189, 248, 0.15); color: #38bdf8; }}
        .tag-gia-dinh {{ background: rgba(74, 222, 128, 0.15); color: #4ade80; }}
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
            <h1>📊 Bảng Kho Bài Viết Tâm Sự (Tách Nguồn)</h1>
            <p>Phân biệt rõ ràng giữa Eva.vn và Afamily.vn (Có Tìm kiếm & Bộ lọc nguồn)</p>
        </div>
        <div class="stats-group">
            <div class="stat-badge">Tổng bài viết: <span id="stat-total">0</span></div>
            <div class="stat-badge" style="border-color: rgba(244, 63, 94, 0.4); background: rgba(244, 63, 94, 0.1); color: #fb7185;">
                Afamily: <span id="stat-afamily" style="color: #fb7185;">0</span>
            </div>
            <div class="stat-badge" style="border-color: rgba(56, 189, 248, 0.4); background: rgba(56, 189, 248, 0.1); color: #38bdf8;">
                Eva.vn: <span id="stat-eva" style="color: #38bdf8;">0</span>
            </div>
            <div class="stat-badge" style="border-color: rgba(74, 222, 128, 0.4); background: rgba(74, 222, 128, 0.1); color: var(--accent-green);">
                Hiển thị: <span id="stat-showing" style="color: var(--accent-green);">0</span>
            </div>
        </div>
    </div>

    <div class="controls-bar">
        <div class="search-box">
            <input type="text" id="searchInput" placeholder="🔍 Tìm kiếm theo tiêu đề, mô tả hoặc từ khóa..." oninput="filterTable()">
        </div>
        <select class="filter-select" id="sourceFilter" onchange="filterTable()">
            <option value="ALL">🌐 Tất Cả Trang Nguồn (Eva.vn & Afamily.vn)</option>
            <option value="Afamily.vn">🔥 Afamily.vn (Tâm Sự Gia Đình)</option>
            <option value="Eva.vn">💖 Eva.vn (Tâm Sự & Gia Đình)</option>
        </select>
    </div>

    <div class="table-container">
        <table>
            <thead>
                <tr>
                    <th class="col-stt">STT</th>
                    <th class="col-source">Trang Nguồn</th>
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

        function renderTable(data) {{
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';

            data.forEach((item, index) => {{
                const tr = document.createElement('tr');
                const srcClass = item.source === 'Afamily.vn' ? 'src-afamily' : 'src-eva';

                tr.innerHTML = `
                    <td class="col-stt">${{index + 1}}</td>
                    <td class="col-source"><span class="tag ${{srcClass}}">${{item.source}}</span></td>
                    <td class="col-cat"><span class="tag tag-tam-su">${{item.category}}</span></td>
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
            const source = document.getElementById('sourceFilter').value;

            const filtered = rawData.filter(item => {{
                const matchSrc = (source === 'ALL' || item.source === source);
                const matchQuery = !query || 
                                   item.title.toLowerCase().includes(query) || 
                                   item.summary.toLowerCase().includes(query) ||
                                   item.url.toLowerCase().includes(query);
                return matchSrc && matchQuery;
            }});

            renderTable(filtered);
        }}

        // Init stats
        const afamilyCount = rawData.filter(i => i.source === 'Afamily.vn').length;
        const evaCount = rawData.filter(i => i.source === 'Eva.vn').length;

        document.getElementById('stat-total').innerText = rawData.length.toLocaleString();
        document.getElementById('stat-afamily').innerText = afamilyCount.toLocaleString();
        document.getElementById('stat-eva').innerText = evaCount.toLocaleString();

        renderTable(rawData);
    </script>
</body>
</html>
"""

    with open(VIEWER_FILE, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"🎉 [BUILD HOÀN TẤT] Đã tạo file Web Bảng Excel tách nguồn: {VIEWER_FILE}")
    print(f"📊 Tổng số bài viết hiển thị: {len(data_list)} bài (Afamily: {len([d for d in data_list if d['source']=='Afamily.vn'])}, Eva: {len([d for d in data_list if d['source']=='Eva.vn'])}).")


if __name__ == "__main__":
    generate_html_viewer()
