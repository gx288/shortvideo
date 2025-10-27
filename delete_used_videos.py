import os
import re
import gspread
from google.oauth2.service_account import Credentials
import unicodedata
import subprocess

# Hàm xử lý tên file để loại bỏ dấu và ký tự đặc biệt (reused from main.py)
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

# Directory setup
output_dir = "output"

# Read from Google Sheets
print("Reading from Google Sheets to find used videos...")
scopes = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file('google_sheets_key.json', scopes=scopes)
gc = gspread.authorize(creds)

SHEET_ID = '14tqKftTqlesnb0NqJZU-_f1EsWWywYqO36NiuDdmaTo'
worksheet = gc.open_by_key(SHEET_ID).worksheet('Phòng mạch')

rows = worksheet.get_all_values()
files_to_delete = []

# Process rows to find those with non-empty column I
for i, row in enumerate(rows):
    if i == 0:  # Skip header
        continue
    if len(row) > 8 and row[8] and row[8].strip() != '':  # Check if column I is non-empty
        raw_content = row[1] if len(row) > 1 else ''  # Column B (title)
        raw_content = re.sub(r'\*+', '', raw_content)  # Remove asterisks
        raw_content = re.sub(r'[\U0001F600-\U0001F64F\U0001F300-\U0001F5FF\U0001F680-\U0001F6FF\U0001F1E0-\U0001F1FF\U00002702-\U000027B0\U000024C2-\U0001F251]+', '', raw_content)  # Remove emojis
        raw_content = re.sub(r'#\w+\s*', '', raw_content)  # Remove hashtags
        lines = [line.strip() for line in raw_content.split('\n') if line.strip()]
        title_text = lines[0].replace('Tiêu đề:', '').strip() if lines else 'Untitled'
        clean_title = clean_filename(title_text)
        video_file = os.path.join(output_dir, f"output_video_{clean_title}.mp4")
        if os.path.exists(video_file):
            files_to_delete.append(video_file)
            print(f"  Found video to delete: {video_file}")
        else:
            print(f"  Video not found: {video_file}")

# Delete files and commit changes
if files_to_delete:
    print("Deleting used video files...")
    for file in files_to_delete:
        try:
            subprocess.run(["git", "rm", file], check=True)
            print(f"  Removed {file} from git")
        except subprocess.CalledProcessError as e:
            print(f"  Warning: Failed to remove {file}: {e}")
    
    # Commit the deletions
    try:
        subprocess.run(["git", "commit", "-m", "Delete used videos based on Google Sheet column I"], check=True)
        subprocess.run(["git", "push"], check=True)
        print("  Committed and pushed deletions")
    except subprocess.CalledProcessError as e:
        print(f"  Warning: Failed to commit/push deletions: {e}")
else:
    print("No used videos to delete.")
