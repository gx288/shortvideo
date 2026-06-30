import os
import glob
import gspread
from google.oauth2.service_account import Credentials
from urllib.parse import urlparse, unquote

output_dir = "output"

print("Đang quét dọn các video dư thừa và video đã đăng...")

# Kết nối Google Sheets
scopes = ['https://www.googleapis.com/auth/spreadsheets']
try:
    creds = Credentials.from_service_account_file('google_sheets_key.json', scopes=scopes)
    gc = gspread.authorize(creds)
except Exception as e:
    print(f"LỖI KHỞI TẠO: {e}")
    exit(1)
    
SHEET_ID = '14tqKftTqlesnb0NqJZU-_f1EsWWywYqO36NiuDdmaTo'

try:
    spreadsheet = gc.open_by_key(SHEET_ID)
except Exception as e:
    print(f"LỖI: Không thể mở spreadsheet: {e}")
    exit(1)

# Danh sách tên các file CẦN GIỮ LẠI (Chưa đăng)
keep_filenames = set()

for worksheet in spreadsheet.worksheets():
    print(f"\nKiểm tra sheet: {worksheet.title}")
    header = worksheet.row_values(1)
    link_col = None
    status_col = None
    for idx, val in enumerate(header, 1):
        c = str(val).strip().lower()
        if c == "link video":
            link_col = idx
        elif c == "đã đăng video?":
            status_col = idx
            
    if not link_col or not status_col:
        print("-> Bỏ qua sheet này vì không có đủ 2 cột cần thiết.")
        continue

    rows = worksheet.get_all_values()
    for i, row in enumerate(rows):
        if i == 0:  # header
            continue

        status_value = row[status_col - 1].strip().lower() if len(row) >= status_col else ""
        link_value = row[link_col - 1].strip() if len(row) >= link_col else ""
        
        if not link_value:
            continue
            
        # Kiểm tra xem video này CÓ BỊ COI LÀ ĐÃ ĐĂNG KHÔNG?
        is_published = "đăng" in status_value or status_value in ["ok", "xong", "true", "yes", "x", "v"]
        
        # Nếu chưa đăng -> GIỮ LẠI
        if not is_published:
            parsed = urlparse(link_value)
            filename = unquote(os.path.basename(parsed.path))
            keep_filenames.add(filename)

print(f"\nTổng số file video CẦN GIỮ (chưa đăng): {len(keep_filenames)}")

# Lấy danh sách toàn bộ các file .mp4 hiện có trong thư mục output
all_mp4_files = glob.glob(os.path.join(output_dir, "*.mp4"))

import subprocess

deleted_count = 0
for filepath in all_mp4_files:
    filename = os.path.basename(filepath)
    if filename not in keep_filenames:
        # Xóa file bằng git rm để xóa cả trên máy và trên repo
        try:
            subprocess.run(["git", "rm", "-f", filepath], check=True, capture_output=True)
            print(f"✅ Đã xóa file dư thừa/đã đăng: {filename}")
            deleted_count += 1
        except Exception as e:
            # Nếu git rm lỗi (file chưa được git theo dõi), xóa thủ công
            try:
                os.remove(filepath)
                print(f"✅ Đã xóa file local: {filename}")
                deleted_count += 1
            except Exception as ex:
                print(f"❌ Lỗi xóa file {filename}: {ex}")

if deleted_count == 0:
    print("\nKhông có video nào dư thừa cần xóa.")
else:
    print(f"\nĐang đồng bộ việc xóa {deleted_count} video lên GitHub...")
    try:
        subprocess.run(["git", "config", "--global", "user.name", "github-actions[bot]"], check=True)
        subprocess.run(["git", "config", "--global", "user.email", "github-actions[bot]@users.noreply.github.com"], check=True)
        subprocess.run(["git", "commit", "-m", f"Auto delete {deleted_count} used/extra videos"], check=True, capture_output=True)
        subprocess.run(["git", "push"], check=True, capture_output=True)
        print("🎉 ĐÃ XÓA TRÊN GITHUB THÀNH CÔNG! Thư mục repo đã sạch sẽ.")
    except Exception as e:
        print(f"⚠️ Lỗi khi đồng bộ lên GitHub (Có thể do không có quyền push): {e}")

