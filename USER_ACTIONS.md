# 📋 DANH SÁCH VIỆC CẦN USER THỰC HIỆN (USER_ACTIONS.md)

Tài liệu này tổng hợp các công việc cần bạn (User) thao tác thủ công khi cần thiết. 
Mỗi khi có việc mới, Agent sẽ cập nhật vào file này.

---

## 🟢 CÁC VIỆC CẦN LÀM HIỆN TẠI (Tùy chọn nâng cao)

### 1. Thêm Secret `GEMINI_API_KEY` trên GitHub (Để dùng Gemini AI viết lại kịch bản)
- **Mục đích**: Giúp Gemini AI tự động viết lại câu chuyện hay hơn, giữ chân người xem video Shorts.
- **Thực hiện**:
  1. Vào [GitHub Repo Settings → Secrets and variables → Actions](https://github.com/gx288/shortvideo/settings/secrets/actions)
  2. Click **New repository secret**
  3. **Name**: `GEMINI_API_KEY`
  4. **Value**: Điền API key Gemini của bạn (lấy free tại [Google AI Studio](https://aistudio.google.com/))
  5. Click **Add secret**

---

### 2. Tùy chỉnh Hashtag video nền (Nếu muốn đổi chủ đề video)
- **Mục đích**: Đổi danh sách hashtag để thu thập video nền theo sở thích (thay vì mặc định `#diy`, `#funny`, `#lifehack`).
- **Thực hiện**:
  - Mở file [instagram/hashtags.txt](file:///d:/AT/github/shortvideo/instagram/hashtags.txt) và điền các hashtag bạn muốn (mỗi hashtag 1 dòng).

---

## ✅ TRẠNG THÁI HỆ THỐNG HIỆN TẠI
- Pipeline tự động 100% không cần bạn làm gì thêm nếu dùng cấu hình mặc định.
- Mọi hẹn giờ tự động chạy ngắn **120 giây (2 phút)** cho từng task.
