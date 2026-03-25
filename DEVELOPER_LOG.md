# 🤖 Joint Agent Mission Control

## 🎯 Current Global Objective
Xây dựng và duy trì hệ thống PCT Report Generator tự động hóa báo cáo học tập.

## 📝 Current Task List
- [x] **Fix KeyError**: Đã xử lý lỗi `Average_time_per_lesson` khi gặp DataFrame rỗng.
- [>] **Environment Setup**: Hướng dẫn người dùng chạy app bằng `python` thay vì `npm`.
- [ ] **Deployment**: Kiểm tra tính tương thích trên Vercel sau khi sửa lỗi.

## 🤝 Handover Protocol (MUST READ)
*Mỗi khi chuyển đổi Agent (Gemini <-> Claude), Agent hiện tại PHẢI cập nhật phần "Last State" bên dưới.*

### 📍 Last State (Updated: 2026-03-25 22:34)
- **Agent vừa làm:** Antigravity (Gemini)
- **Công việc vừa đạt được:** 
    - Phân tích và sửa lỗi `KeyError: 'Average_time_per_lesson'` trong `report_grade_10.py` and `report_grade_11.py`.
    - Khởi tạo `GEMINI.md` và `DEVELOPER_LOG.md` theo quy tắc Global.
- **Vấn đề đang gặp:** Người dùng đang cố chạy project bằng lệnh `npm run dev` trong khi đây là project Python/Flask.
- **Chỉ dẫn cho Agent tiếp theo (Claude):** 
    - Đảm bảo người dùng biết cách khởi động app bằng `run.bat` hoặc lệnh `python app.py`.
    - Hỗ trợ thêm nếu có lỗi về mismatch tên cột trong các file Excel mới.
- **File quan trọng cần đọc:** `report_grade_10.py`, `app.py`.
