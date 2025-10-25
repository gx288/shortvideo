import os
import gspread
from google.oauth2.service_account import Credentials

# Read clean_title from file
try:
    with open(os.path.join("output", "clean_title.txt"), "r") as f:
        clean_title = f.read().strip()
except FileNotFoundError:
    print("Error: clean_title.txt not found. Cannot update sheet.")
    exit(1)

# Construct raw GitHub URL
video_url = f"https://raw.githubusercontent.com/gx288/shortvideo/main/output/output_video_{clean_title}.mp4"

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
    worksheet.update_cell(selected_row_num, 8, video_url)
    print(f"Updated row {selected_row_num} with {video_url}")
except Exception as e:
    print(f"Error updating sheet: {e}")
    exit(1)
