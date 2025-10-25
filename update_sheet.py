import os
import gspread
from google.oauth2.service_account import Credentials

# Read selected_row_num from file
try:
    with open(os.path.join("output", "selected_row_num.txt"), "r") as f:
        selected_row_num = int(f.read().strip())
except FileNotFoundError:
    print("Error: selected_row_num.txt not found. Cannot update sheet.")
    exit(1)

# Update sheet
scopes = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file('google_sheets_key.json', scopes=scopes)
gc = gspread.authorize(creds)
SHEET_ID = '14tqKftTqlesnb0NqJZU-_f1EsWWywYqO36NiuDdmaTo'
worksheet = gc.open_by_key(SHEET_ID).worksheet('Phòng mạch')

video_url = os.environ.get('VIDEO_URL', 'output/output_video.mp4')
try:
    worksheet.update_cell(selected_row_num, 8, video_url)
    print(f"Updated row {selected_row_num} with {video_url}")
except Exception as e:
    print(f"Error updating sheet: {e}")
    exit(1)
