# 📋 DANH SÁCH VIỆC CẦN USER THỰC HIỆN (USER_ACTIONS.md)

Tài liệu này tổng hợp các thao tác thủ công bạn (User) có thể thực hiện trên GitHub UI hoặc Make.com.

---

## 🚀 1. KÍCH HOẠT TẠO VIDEO TRÊN GITHUB ACTIONS (Cloud Render)

Mọi quá trình render video nặng (FFmpeg, MoviePy, ghép TTS, cập nhật Google Sheets) được chạy 100% trên đám mây của **GitHub Actions** để không tốn RAM/CPU máy local.

### 📌 Các bước kích hoạt trên GitHub:
1. Vào link: [GitHub Actions - Generate Video Workflow](https://github.com/gx288/shortvideo/actions/workflows/video-generation.yml)
2. Click nút **Run workflow** ở góc phải.
3. Tùy chọn (để trống nếu muốn chạy tự động):
   * **run_count**: Số video muốn tạo (ví dụ `1`, `2`, `3`).
   * **video_url**: (Tùy chọn) Nhập link video TikTok/Reels cụ thể nếu muốn chỉ định video nền.
4. Click **Run workflow** màu xanh.

---

## ⚡ 2. KÍCH HOẠT QUA MAKE.COM (Tự động hóa 100%)
- Bạn có thể tạo 1 HTTP Webhook trên Make.com để gọi GitHub API Trigger workflow tự động mà không cần vào trang GitHub:
  - **URL**: `https://api.github.com/repos/gx288/shortvideo/actions/workflows/video-generation.yml/dispatches`
  - **Method**: `POST`
  - **Header**: `Authorization: token <GH_TOKEN>`
  - **Body**: `{"ref": "main"}`

---

## 🔑 3. CÁC SECRETS ĐÃ CÓ TRÊN GITHUB REPO
- `GH_TOKEN` (Đã có)
- `GOOGLE_TTS_KEY` (Đã có)
- `GOOGLE_SHEETS_KEY` (Đã có)
- `GEMINI_API_KEY` (Tùy chọn - Thêm nếu muốn dùng Gemini AI viết kịch bản)
