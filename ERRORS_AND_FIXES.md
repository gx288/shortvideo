# 🛠️ TỔNG HỢP LỖI THƯỜNG GẶP & CÁCH FIX NGẮN GỌN (ERRORS_AND_FIXES.md)

Tài liệu ghi chú nhanh các sự cố kỹ thuật trong quá trình làm video Short tự động và cách khắc phục tối ưu.

---

### ❌ Lỗi 1: File MP4 xuất ra nặng gần 1GB và KHÔNG mở xem được
* **Nguyên nhân**: Sử dụng tham số ghép luồng trực tiếp `FFmpeg -c:v copy` từ video 4K/HD gốc khiến chỉ số `moov` atom nằm sai vị trí và lệch timestamp.
* **Cách Fix Ngắn Gọn**: 
  * Mã hóa nén H.264 Baseline 720p siêu nhẹ (~10MB) và thêm `-movflags +faststart`:
  ```bash
  ffmpeg -y -i bg.mp4 -i audio.mp3 -vf "scale=720:1280:force_original_aspect_ratio=increase,crop=720:1280,format=yuv420p,fps=24" -c:v libx264 -preset superfast -crf 28 -movflags +faststart -c:a aac -b:a 96k -shortest output.mp4
  ```

---

### ❌ Lỗi 2: Nhầm lẫn bài viết Showbiz / Trắc nghiệm / Thiếu Mô tả
* **Nguyên nhân**: Các trang Eva.vn / Afamily.vn chứa các bài bói toán, trắc nghiệm 12 con giáp, tin tức showbiz không phải câu chuyện thực tế.
* **Cách Fix Ngắn Gọn**: 
  * Chạy bộ lọc `eva_scraper/filter_stories_only.py`:
  * Bắt buộc bài viết phải có cả **Tiêu đề (>=10 từ)** + **Mô tả Sapo (>=15 từ)**.
  * Loại bỏ hoàn toàn các url thuộc `/day-con/`, `/gia-dinh/`, `/tin-tuc/`, `/dinh-duong/`.

---

### ❌ Lỗi 3: Nhạc nền quá to đè mất giọng đọc TTS
* **Nguyên nhân**: Âm lượng nhạc nền nguyên bản lồng ghép át mất tiếng đọc câu chuyện.
* **Cách Fix Ngắn Gọn**: 
  * Set âm lượng `nhacnen.mp3` ở mức **5% - 8%** (`volume=0.07`), giữ giọng TTS ở 100%:
  ```python
  bg_music = bg_music.volumex(0.07)  # hoặc FFmpeg: volume=0.07
  ```

---

### ❌ Lỗi 4: Giọng đọc TTS quá chậm so với nhịp lướt TikTok/Shorts
* **Nguyên nhân**: Giọng gTTS chuẩn đọc ~120 từ/phút nghe bị chậm chạp và thiếu kịch tính.
* **Cách Fix Ngắn Gọn**: 
  * Tăng tốc giọng đọc TTS lên **1.2x** bằng FFmpeg `atempo=1.2` hoặc MoviePy:
  ```python
  import moviepy.video.fx as vfx
  tts_audio = vfx.MultiplySpeed(1.2).apply(tts_audio_raw)
  ```

---

### ❌ Lỗi 5: Video nền bị lệch chuẩn size / bitrate khi nối nhiều clip
* **Nguyên nhân**: Các video nền tải từ YouTube Shorts / TikTok có độ phân giải, bitrate và âm thanh gốc khác nhau.
* **Cách Fix Ngắn Gọn**: 
  * Ngay khi tải clip về, tự động chuẩn hóa về 1 bộ khung chung **720x1280, 24fps, yuv420p, No Audio (`-an`)**:
  ```bash
  yt-dlp -f "bestvideo[height<=720][ext=mp4]/best" -o raw.mp4 <url>
  ffmpeg -i raw.mp4 -vf "scale=720:1280:force_original_aspect_ratio=increase,crop=720:1280,format=yuv420p,fps=24" -c:v libx264 -crf 26 -an std_clip.mp4
  ```
