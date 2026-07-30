# 🤖 AGENT GUIDELINES & AUTOMATION TASK WORKFLOW

Tài liệu này hướng dẫn Agent tự động từng bước thực hiện quy trình tạo YouTube Shorts từ truyện Tâm sự Eva.vn.

---

## 🎯 MỤC TIÊU TỔNG THỂ

Tự động hóa hoàn toàn 100% pipeline tạo video ngắn (YouTube Shorts 9:16, thời lượng < 3 phút):
1. **Crawl link truyện**: Lấy link câu chuyện từ `eva.vn/tam-su` vào `eva_scraper/links_master.json`.
2. **Crawl nội dung truyện**: Bóc tách nội dung chi tiết bài viết lưu thành từng file JSON trong `eva_scraper/data/`.
3. **Biên tập kịch bản video**: Viết lại câu chuyện với câu Hook mở đầu hấp dẫn, thời lượng 2 - 3 phút (~250-420 từ) lưu vào `eva_scraper/scripts/`.
4. **Crawl video nền**: Thu thập link video ngắn từ YouTube Shorts / TikTok vào `instagram/link_pool.json`.
5. **Tạo video Shorts**: Ghép kịch bản + TTS giọng Việt + video nền + hiệu ứng Lissajous panning + nhạc nền.
6. **Đồng bộ & Hẹn giờ**: Cập nhật trạng thái và tự động lên lịch (`schedule`) cho task tiếp theo.

---

## 🔄 QUY TRÌNH THỰC HIỆN TỰ ĐỘNG CỦA AGENT

Mỗi khi Agent bắt đầu hoặc được đánh thức bởi hẹn giờ (`schedule`), Agent **BẮT BUỘC** làm theo 3 bước:

### 📌 BƯỚC 1: Đọc tài liệu hướng dẫn `agent.md`
- Luôn xem nội dung [agent.md](file:///d:/AT/github/shortvideo/agent.md) để nắm mục tiêu, tiêu chuẩn và kiểm tra tiến độ hiện tại.

### 📌 BƯỚC 2: Thực hiện Task được giao theo thứ tự
Kiểm tra xem task nào cần làm tiếp theo:

#### 🔹 Task 1: Thu thập Link câu chuyện Eva.vn
```bash
python eva_scraper/link_crawler.py --scan-ids --start-id 678450 --count 100
```
- Kiểm tra file `eva_scraper/links_master.json`. Nếu chưa đủ >50 link, chạy lệnh trên.

#### 🔹 Task 2: Cào nội dung chi tiết câu chuyện
```bash
python eva_scraper/article_scraper.py --batch 30
```
- Đọc các link chưa scrape trong `eva_scraper/links_master.json` và lưu các file `.json` bài viết vào `eva_scraper/data/`.

#### 🔹 Task 3: Viết lại Kịch bản Video ngắn (< 3 phút)
```bash
python eva_scraper/script_rewriter.py --batch 20
```
- Đọc file JSON từ `eva_scraper/data/`, tạo câu Hook gây tò mò trong 3 giây đầu, biên tập hành văn mượt mà và tối ưu độ dài **220 - 420 từ (thời lượng 1.5 - 3 phút)** vào `eva_scraper/scripts/`.

#### 🔹 Task 4: Thu thập Link video nền (YouTube Shorts / TikTok)
```bash
python instagram/hashtag_crawler.py --hashtag diy funny lifehack --limit 50
```
- Kiểm tra `instagram/link_pool.json`. Đảm bảo có ít nhất >20 link video nền khả dụng (`used: false`).

#### 🔹 Task 5: Tạo Video Shorts hoàn chỉnh
```bash
python main.py
```
- Tạo video `.mp4` trong thư mục `output/`, tự động ghi link và cập nhật Google Sheets.

### 📌 BƯỚC 3: Tự động Hẹn giờ ngắn (`schedule`) cho Task tiếp theo
- **QUAN TRỌNG**: Sau khi kết thúc bất kỳ task nào, Agent **BẮT BUỘC** gọi tool `schedule` với thời gian ngắn:
- Thời gian hẹn giờ: `DurationSeconds: 120` (2 phút).
- Prompt hẹn giờ mẫu: `Đọc agent.md và tiếp tục chạy Task [N] tiếp theo trong pipeline.`
- Nếu cần User làm gì thủ công, ghi ngay vào file [USER_ACTIONS.md](file:///d:/AT/github/shortvideo/USER_ACTIONS.md).

---

## ⚙️ BÁO CÁO & LƯU TRỮ

- Mọi thay đổi code / dữ liệu thành công đều phải commit & push lên git repository `origin main`.
- Không commit các file tạm, file media nặng, credentials JSON keys (đã cài đặt trong `.gitignore`).
