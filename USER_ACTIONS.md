# 📋 DANH SÁCH VIỆC CẦN LÀM (USER_ACTIONS.md)

Tài liệu hướng dẫn các bước tiếp theo để vận hành hệ thống tạo Video Short tự động.

---

## 📊 BƯỚC 1: XEM BẢNG KHO BÀI VIẾT (LOCAL)
- Mở file **[`eva_scraper/stories_viewer.html`](file:///d:/AT/github/shortvideo/eva_scraper/stories_viewer.html)** trên trình duyệt.
- Bạn có thể xem, tìm kiếm và lọc qua **3,150 câu chuyện chuẩn** (Afamily: 3,078 bài, Eva: 72 bài).

---

## 📝 BƯỚC 2: CHẠY LẤY FULL KỊCH BẢN VIDEO (< 3 PHÚT)
Nếu bạn muốn cào nội dung chi tiết bài viết và biên tập thành kịch bản video ngắn ngay tại local, chạy lệnh:
```bash
python eva_scraper/full_batch_scraper.py --count 50
```
*(Lệnh này chạy 20 luồng siêu tốc, tự động tạo kịch bản video có câu Hook 3s vào folder `eva_scraper/scripts/`)*

---

## 📤 BƯỚC 3: ĐỒNG BỘ DỮ LIỆU LÊN GITHUB
Khi bạn sẵn sàng đẩy dữ liệu bài viết mới lên GitHub để cloud render video, chạy:
```bash
git add eva_scraper/ afamily_scraper/
git commit -m "Update 3150 story links and viewer"
git push origin main
```

---

## 🚀 BƯỚC 4: KÍCH HOẠT RENDER VIDEO TRÊN GITHUB ACTIONS (Cloud Render)
Render video nặng (MoviePy, TTS, background video) chạy 100% trên đám mây GitHub Actions không tốn RAM/CPU máy bạn:
1. Vào link: [GitHub Actions - Generate Video Workflow](https://github.com/gx288/shortvideo/actions/workflows/video-generation.yml)
2. Click nút **Run workflow** góc phải.
3. Nhập số video muốn tạo (ví dụ `1`, `3`, `5`).
4. Click **Run workflow** màu xanh.

---

## ⚡ BƯỚC 5: TỰ ĐỘNG HÓA QUA MAKE.COM (Webhook Trigger)
- Bạn có thể cài Webhook trên Make.com để gọi tự động không cần bấm tay:
  - **URL**: `https://api.github.com/repos/gx288/shortvideo/actions/workflows/video-generation.yml/dispatches`
  - **Method**: `POST`
  - **Header**: `Authorization: token <GH_TOKEN>`
  - **Body**: `{"ref": "main"}`
