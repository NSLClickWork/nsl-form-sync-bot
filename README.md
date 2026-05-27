# NSL Form Sync Bot 🤖

Bot tự động hóa việc thu thập và gộp dữ liệu từ Microsoft Forms cho dự án NSL.

## 🎯 Chức năng chính
- Tự động quét toàn bộ thư mục trong Group `NSL | Assessment Forms` trên SharePoint.
- Tìm kiếm các file Excel kết quả (chứa từ khóa `NSL Assessment Cent`).
- Tự động phân tích, đếm cột và chọn ra những cột dữ liệu xuất hiện ở hầu hết các Form (kèm theo cột "Số điện thoại").
- Trộn (Merge) tất cả dữ liệu từ các form khác nhau thành một file duy nhất.
- Tự động đẩy file tổng `Master_Assessment_Responses.xlsx` vào thư mục OneDrive cá nhân (`MANAGEMENT/Assessment`) của tài khoản quản trị.

## ⚙️ Cơ chế hoạt động
Dự án này sử dụng **GitHub Actions** để chạy tự động theo lịch (Cron Job).
- **Tần suất:** Mặc định chạy mỗi 10 phút một lần (`*/10 * * * *`).
- **Nền tảng:** Node.js với thư viện `xlsx` để xử lý dữ liệu Excel và `@microsoft/microsoft-graph-client` để tương tác với Microsoft Graph API.

## 🔑 Bảo mật
- Thông tin xác thực (Client ID, Tenant ID) được mã hóa trong source code.
- `CLIENT_SECRET` được bảo mật bằng tính năng **Secrets** của GitHub Actions.

---
*Developed for NSL Click & Work UG.*
