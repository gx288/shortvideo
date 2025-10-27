import os
import re
import unicodedata
import gspread
from google.oauth2.service_account import Credentials

# Hàm xử lý tên file để loại bỏ dấu và ký tự đặc biệt
def clean_filename(text, max_length=50):
    # Chuẩn hóa Unicode: chuyển các ký tự có dấu thành không dấu
    text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII')
    # Thay khoảng trắng bằng dấu gạch dưới
    text = text.replace(' ', '_')
    # Chỉ giữ chữ cái, số, dấu gạch dưới, và dấu gạch ngang
    text = re.sub(r'[^\w-]', '', text)
    # Loại bỏ các dấu gạch dưới liên tiếp
    text = re.sub(r'_+', '_', text)
    # Cắt ngắn tên file nếu quá dài
    text = text[:max_length].strip('_')
    # Nếu tên rỗng hoặc chỉ chứa ký tự không hợp lệ, trả về tên mặc định
    if not text or text == '':
        text = f"video_{random.randint(1000, 9999)}"
    return text.lower()

# Read clean_title from file
try:
    with open(os.path.join("output", "clean_title.txt"), "r") as f:
        clean_title = f.read().strip()
except FileNotFoundError:
    print("Error: clean_title.txt not found. Cannot update sheet.")
    exit(1)

# Construct raw GitHub URL
video_url = f"https://raw.githubusercontent.com/gx288/shortvideo/main/output/output_video_{clean_title}.mp4"

# Check video file size
video_path = os.path.join("output", f"output_video_{clean_title}.mp4")
if not os.path.exists(video_path):
    print(f"Error: Video file {video_path} not found")
    exit(1)

file_size_mb = os.path.getsize(video_path) / (1024 * 1024)  # Convert to MB
print(f"Video size: {file_size_mb:.2f} MB")

# Update sheet
scopes = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file('google_sheets_key.json', scopes=scopes)
gc = gspread.authorize(creds)
SHEET_ID = '14tqKftTqlesnb0NqJZU-_f1EsWWywYqO36NiuDdmaTo'
worksheet = gc.open_by_key(SHEET_ID).worksheet('Phòng mạch')

# Find row with empty column H (to handle row changes)
rows = worksheet.get_all_values()
selected_row_num = None
for i, row in enumerate(rows):
    if i == 0:  # Skip header
        continue
    if len(row) > 7 and (not row[7] or row[7].strip() == ''):
        selected_row_num = i + 1
        break

if not selected_row_num:
    print("Error: No row with empty column H found. Cannot update sheet.")
    exit(1)

try:
    # Update column H with video URL
    worksheet.update_cell(selected_row_num, 8, video_url)
    print(f"Updated row {selected_row_num}, column H with {video_url}")
    
    # Update column I if file size > 5MB
    if file_size_mb > 5:
        worksheet.update_cell(selected_row_num, 9, ">5MB")
        print(f"Updated row {selected_row_num}, column I with '>5MB'")
except Exception as e:
    print(f"Error updating sheet: {e}")
    exit(1)
