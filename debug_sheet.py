import gspread
from google.oauth2.service_account import Credentials

print("Checking Google Sheets rows 418-422...")
scopes = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file('google_sheets_key.json', scopes=scopes)
gc = gspread.authorize(creds)
SHEET_ID = '14tqKftTqlesnb0NqJZU-_f1EsWWywYqO36NiuDdmaTo'
worksheet = gc.open_by_key(SHEET_ID).worksheet("Khoahocyhoc")

header = worksheet.row_values(1)
link_col = None
for idx, cell in enumerate(header, 1):
    if cell.strip() == "Link video":
        link_col = idx
print(f"Link video column is: {link_col}")

rows = worksheet.get_all_values()
for i in range(417, 422):
    row = rows[i]
    val = row[link_col - 1] if len(row) >= link_col else "<EMPTY-LIST>"
    print(f"Row {i+1} Column {link_col}: '{val}' (len={len(val)})")
