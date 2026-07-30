# 🎬 YouTube Shorts Auto Generator

Hệ thống tự động tạo YouTube Shorts: lấy nội dung từ web, dùng Google TTS đọc bằng giọng Việt, ghép với video nguồn, lưu kết quả lên OneDrive.

## 🔧 Pipeline

```
Google Sheets (queue)
    ↓  Make.com trigger
GitHub Actions
    ├── story_scraper.py  → lấy câu chuyện từ web / Reddit
    ├── Google TTS        → tạo narration.mp3 (vi-VN)
    ├── video_downloader.py → download video nền (yt-dlp / Pexels)
    ├── main.py           → ghép video + audio, hiệu ứng panning
    └── onedrive_uploader.py → upload .mp4 lên OneDrive
```

## 📁 Cấu trúc

```
├── .github/workflows/video-generation.yml   # GitHub Actions
├── main.py               # Tạo video chính (TTS + ghép ảnh/video)
├── story_scraper.py      # Scrape câu chuyện từ web / Reddit
├── video_downloader.py   # Download video TikTok/Instagram/Pexels
├── onedrive_uploader.py  # Upload lên OneDrive via Graph API
├── delete_used_videos.py # Dọn dẹp video đã đăng
├── update_sheet.py       # Cập nhật Google Sheets
└── requirements.txt
```

## ⚙️ GitHub Secrets cần thiết

| Secret | Mô tả |
|--------|-------|
| `GOOGLE_TTS_KEY` | Google Cloud Service Account JSON (base64 hoặc raw) |
| `GOOGLE_SHEETS_KEY` | Google Sheets Service Account JSON |
| `PEXELS_API_KEY` | Pexels API key (miễn phí) |
| `ONEDRIVE_CLIENT_ID` | Azure App Client ID |
| `ONEDRIVE_CLIENT_SECRET` | Azure App Client Secret |
| `ONEDRIVE_TENANT_ID` | Azure Tenant ID |
| `ONEDRIVE_USER_EMAIL` | Email OneDrive đích |

## 🚀 Chạy thủ công

Vào **Actions** tab → **Generate Video** → **Run workflow** → điền:
- `Số video`: số lượng cần tạo
- `Link TikTok/Instagram` (tuỳ chọn): video nền
- `Link câu chuyện` (tuỳ chọn): URL trang web cần scrape
- `Upload OneDrive`: true/false

## 💰 Chi phí

| Service | Free tier |
|---------|-----------|
| GitHub Actions | 2,000 phút/tháng (public repo = miễn phí) |
| Google TTS | 1M ký tự/tháng |
| Pexels API | Không giới hạn |
| OneDrive | 5GB free |
| Make.com | 1,000 operations/tháng |
