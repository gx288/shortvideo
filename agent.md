# 🤖 AGENT GUIDELINES & HYBRID AUTOMATION WORKFLOW

Tài liệu này hướng dẫn quy trình phối hợp: **Tạo dữ liệu tại Local → Push Git → Chạy Render Video tự động trên GitHub Actions (Cloud)**.

---

## 🎯 MỤC TIÊU TỔNG THỂ

1. **Local**: Crawl truyện, bóc tách nội dung, viết lại kịch bản (< 3 phút) & chuẩn bị kho video nền → Push dữ liệu dạng `.json` lên GitHub.
2. **GitHub Actions (Cloud)**: Tự động chạy `main.py` trên Ubuntu Runner của GitHub để ghép TTS + video nền + render video `.mp4` + cập nhật Google Sheets.

---

## 🔄 QUY TRÌNH TASK CHI TIẾT

### 📌 GIAI ĐOẠN 1: Chuẩn bị dữ liệu tại Local (hoặc Scheduled Actions)

#### 🔹 Task 1: Quét Hàng NGHÌN Link truyện Eva.vn siêu tốc
```bash
python eva_scraper/scan_all_links.py --start-id 678500 --count 5000 --threads 50
```
- Kết quả: Thu thập hàng NGHÌN link câu chuyện khả dụng lưu vào `eva_scraper/links_master.json`.

#### 🔹 Task 2: Viết lại Kịch bản Video ngắn (< 3 phút)
```bash
python eva_scraper/script_rewriter.py --batch 20
```
- Kết quả: Tự động thêm câu Hook 3s đầu, tối ưu 220 - 400 từ (~2-2.5 phút) lưu vào `eva_scraper/scripts/`.

#### 🔹 Task 3: Thu thập Link video nền
```bash
python instagram/hashtag_crawler.py --hashtag diy funny lifehack --limit 50
```
- Kết quả: Lưu danh sách link video nền vào `instagram/link_pool.json`.

#### 🔹 Task 4: Push dữ liệu lên GitHub
```bash
git add eva_scraper/ instagram/
git commit -m "Update story scripts & video pool"
git push origin main
```

---

### 📌 GIAI ĐOẠN 2: Render Video trên GitHub Actions (Cloud)

#### 🔹 Task 5: Chạy Workflow Render Video trên GitHub
- **Cách 1 (Thủ công / Make.com)**: Vào GitHub Actions → chọn **"Generate Video"** → Click **Run workflow**.
- **Cách 2 (Tự động theo lịch)**: Workflow `.github/workflows/video-generation.yml` tự động chạy trên GitHub Cloud để render video.

---

## ⚙️ HƯỚNG DẪN DÀNH CHO AGENT

Mỗi khi được đánh thức bởi hẹn giờ (`schedule`):
1. **Subtask 1**: Xem file [agent.md](file:///d:/AT/github/shortvideo/agent.md).
2. **Subtask 2**: Thực hiện Task dữ liệu hiện tại (Task 1, 2, 3 hoặc 4).
3. **Subtask 3**: Push dữ liệu JSON mới lên `origin main`.
4. **Subtask 4**: Hẹn giờ ngắn 120s bằng tool `schedule`.
5. **Subtask 5**: Cập nhật file [USER_ACTIONS.md](file:///d:/AT/github/shortvideo/USER_ACTIONS.md) nếu cần User bấm nút Trigger workflow trên GitHub.
