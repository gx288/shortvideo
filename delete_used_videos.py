import os
import gspread
from google.oauth2.service_account import Credentials
import subprocess
from urllib.parse import urlparse, unquote

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
        video_url = row[7] if len(row) > 7 else ''  # Column H (URL)
        if not video_url:
            print(f"  Row {i + 1}: No URL found in column H. Skipping.")
            continue
        # Extract filename from URL
        parsed_url = urlparse(video_url)
        filename = unquote(os.path.basename(parsed_url.path))
        video_file = os.path.join(output_dir, filename)
        if os.path.exists(video_file):
            files_to_delete.append(video_file)
            print(f"  Found video to delete: {video_file} (from URL: {video_url})")
        else:
            print(f"  Video not found: {video_file} (from URL: {video_url})")

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
