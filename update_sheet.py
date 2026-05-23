import gspread
import os
from google.oauth2.service_account import Credentials

output_dir = "output"
SHEET_ID = '14tqKftTqlesnb0NqJZU-_f1EsWWywYqO36NiuDdmaTo'

def update_sheet():
    try:
        # ====================== ĐỌC THÔNG TIN TỪ MAIN.PY ======================
        with open(os.path.join(output_dir, "current_sheet.txt"), "r", encoding="utf-8") as f:
            current_sheet_name = f.read().strip()
        
        with open(os.path.join(output_dir, "current_row.txt"), "r", encoding="utf-8") as f:
            target_row = int(f.read().strip())
        
        with open(os.path.join(output_dir, "clean_title.txt"), "r", encoding="utf-8") as f:
            clean_title = f.read().strip()

        print(f"Đang cập nhật sheet: {current_sheet_name}")
        print(f"Dòng cần cập nhật: {target_row}")

        # ====================== TẠO LINK VIDEO ======================
        video_url = f"https://raw.githubusercontent.com/gx288/shortvideo/main/output/output_video_{clean_title}.mp4"

        # Kiểm tra video tồn tại
        video_path = os.path.join(output_dir, f"output_video_{clean_title}.mp4")
        if not os.path.exists(video_path):
            print(f"❌ Không tìm thấy video: {video_path}")
            return

        file_size_mb = os.path.getsize(video_path) / (1024 * 1024)
        print(f"Video size: {file_size_mb:.2f} MB")

        # ====================== KẾT NỐI GOOGLE SHEET ======================
        scopes = ['https://www.googleapis.com/auth/spreadsheets']
        creds = Credentials.from_service_account_file('google_sheets_key.json', scopes=scopes)
        gc = gspread.authorize(creds)
        worksheet = gc.open_by_key(SHEET_ID).worksheet(current_sheet_name)

        # Cột Link video và Đã đăng video? (theo log của bạn)
        link_col = 11
        status_col = 12

        # ====================== CẬP NHẬT ======================
        worksheet.update_cell(target_row, link_col, video_url)
        print(f"✅ Đã ghi link vào dòng {target_row}, cột {link_col}")

        if file_size_mb > 5:
            worksheet.update_cell(target_row, status_col, ">5MB")
            print(f"Đánh dấu >5MB tại cột {status_col}")
        else:
            worksheet.update_cell(target_row, status_col, "")
            print("Video ≤ 5MB → Đánh dấu 'Đã tạo'")

        print("CẬP NHẬT GOOGLE SHEET THÀNH CÔNG!")

    except FileNotFoundError as e:
        print(f"❌ Thiếu file: {e}")
        print("→ main.py phải chạy trước update_sheet.py")
    except Exception as e:
        print(f"❌ Lỗi khi cập nhật sheet: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    update_sheet()
