# ☕ 23 COFFEE — Hệ Thống Quản Lý Cost & Tự Động Hóa Menu F&B

Ứng dụng web nội bộ giúp chủ quán **23 COFFEE** tự động hóa toàn bộ quy trình tính cost, bóc tách công thức SOP, định giá bán theo chuẩn 30% Food Cost, lưu trữ dữ liệu vĩnh viễn và **tự động đồng bộ hai chiều trực tiếp vào Google Sheets online**.

---

## 🔒 1. Chính Sách Bảo Mật (Giới Hạn Chỉ Đúng 2 Người Truy Cập)

Hệ thống được khóa chặt bằng cơ chế **Session Authentication** (bảo vệ mọi route và API). Khi bạn public web lên internet, chỉ có 2 tài khoản sau mới có quyền đăng nhập và xem dữ liệu giá vốn nội bộ:

| STT | Vai Trò | Tên Đăng Nhập | Mật Khẩu Mặc Định | Ghi Chú |
| :-: | :--- | :---: | :---: | :--- |
| **1** | **Quản Trị Viên (Admin)** | `admin` | `admin` | Toàn quyền quản trị |
| **2** | **Chủ Quán (Tấn Đạt)** | `tandat` | `admin` | Toàn quyền quản trị |

> 💡 **Đổi mật khẩu:** Sau khi đăng nhập, bạn có thể bấm vào biểu tượng 🔑 (chìa khóa) ở góc trên bên phải để đổi sang mật khẩu riêng bất kỳ lúc nào.

---

## 📊 2. Tự Động Lưu Dữ Liệu & Đồng Bộ Trực Tiếp Vào Google Sheets

Hệ thống giải quyết triệt để bài toán **Lưu trữ dữ liệu vĩnh viễn trên Cloud** và **Tự động cập nhật trực tiếp vào Google Sheets**:
* **Link Google Sheets Kết Nối:** [Bảng Cost & Chiến Lược Định Giá 23 Coffee](https://docs.google.com/spreadsheets/d/1uGRxJEPZljoVKlbBXRatjPiIpk2NBnXldwWSfVYVy_Y/edit?usp=sharing)
* **Cơ chế lưu trữ:**
  1. Khi bạn thêm món mới, sửa giá nguyên liệu hay đổi công thức trên Web, hệ thống tự động lưu vào **Cơ sở dữ liệu SQLite** và bộ nhớ đệm.
  2. Ngay lập tức, một tiến trình chạy ngầm tự động đẩy toàn bộ dữ liệu sang **Google Sheets** chỉ sau 1-2 giây.
  3. Bất cứ ai mở link Google Sheets ra đều thấy các bảng tính nhảy số và đổi màu theo thời gian thực (real-time).
  4. Nếu bạn chỉnh sửa dữ liệu trên Google Sheets, trên web có nút **"Kéo Từ Google Về Web"** để cập nhật lại hệ thống ngay lập tức!

### ⚡ Hướng dẫn kích hoạt Webhook trên Google Sheets (Chỉ làm 1 lần - mất 1 phút):
1. Mở file Google Sheets: `https://docs.google.com/spreadsheets/d/1uGRxJEPZljoVKlbBXRatjPiIpk2NBnXldwWSfVYVy_Y/edit?usp=sharing`
2. Vào menu **Tiện ích mở rộng** (*Extensions*) $\rightarrow$ chọn **Apps Script**.
3. Dán toàn bộ nội dung trong tệp [`google_sheets_script.gs`](google_sheets_script.gs) vào và bấm biểu tượng Lưu (Ctrl+S).
4. Bấm nút xanh **Triển khai** (*Deploy*) ở góc phải trên $\rightarrow$ chọn **Triển khai mới** (*New deployment*).
5. Chọn loại **Ứng dụng web** (*Web app*), mục **Ai có quyền truy cập** chọn **Bất kỳ ai** (*Anyone*) $\rightarrow$ Bấm **Triển khai**.
6. Sao chép **URL ứng dụng web** (Web App URL) và dán vào nút **Google Sheets** trên thanh điều hướng của Web App!

---

## 🌟 3. Các Tính Năng Nổi Bật

1. **Quản Lý Món & Công Thức Barista (Recipe SOP):**
   * Thêm món mới, chọn nguyên liệu từ dropdown, nhập định lượng (g/ml).
   * **Live Cost Preview:** Tự động tính tổng cost, hao hụt (5%), bao bì (1.100đ), giá gợi ý 30% và lãi gộp ngay khi đang gõ.
   * Huy hiệu đánh giá tự động: `✅ Chuẩn Ngành (25-35%)`, `💎 Lãi Dày (<25%)`, `⚠️ Cost Cao (>35%)`.

2. **Kho Nguyên Liệu Tự Động Quy Đổi:**
   * Thêm nguyên liệu với giá mua sỉ và quy cách $\rightarrow$ Hệ thống tự chia ra đơn giá cơ sở (đ/g, đ/ml).
   * Khi nhà cung cấp tăng/giảm giá: Chỉ cần sửa giá ở 1 ô trong Kho, toàn bộ các món trong menu tự động cập nhật lại giá vốn!

3. **Tự Động Đồng Bộ & Căn Chỉnh Vào Excel Chuẩn 5 Sheet:**
   * Mỗi thao tác thêm/sửa/xóa đều tự động ghi vào tệp: `Bang_Cost_Va_Chien_Luoc_Dinh_Gia_23_Coffee.xlsx`.
   * **100% Công thức động** (`SUM`, `ROUND`, `CEILING`, `AVERAGE`, `ROUNDUP`), kẻ viền bảng, tô màu và định dạng tiền tệ VNĐ chuyên nghiệp.

4. **Dự Toán Điểm Hòa Vốn Thực Tế:**
   * Quản lý chi phí cố định (Mặt bằng, lương, điện, nước, internet/POS, khấu hao...).
   * Tự động tính: Giá bán trung bình (AOV), Lãi gộp trung bình (CM), **Số ly hòa vốn / ngày** và **Số ly cần bán để đạt lãi ròng 15.000.000 đ/tháng**.

---

## 💻 4. Hướng Dẫn Chạy Trên Máy Tính Cá Nhân (Local)

1. Bấm đúp chuột vào tệp: **`Chay_App_Quan_Ly_Cost.bat`**.
2. Trình duyệt sẽ tự động mở trang web tại địa chỉ: **`http://localhost:5000`**.
3. Đăng nhập bằng tài khoản `admin` / mật khẩu `admin`.

---

## 🌐 5. Hướng Dẫn Đưa Lên Mạng Công Khai (Public Web Miễn Phí 100%)

Để bạn và cổ đông có thể truy cập bằng điện thoại/iPad từ bất cứ đâu:

### Cách triển khai qua Render.com (Khuyên dùng - 2 phút):
1. Đăng ký tài khoản miễn phí tại **[render.com](https://render.com)**.
2. Chọn **New +** $\rightarrow$ **Web Service**.
3. Chọn **Build and deploy from a Git repository** $\rightarrow$ Kết nối với repository GitHub: `https://github.com/tandat1512/23coffee`.
4. Render sẽ tự động nhận diện tệp `render.yaml` và `requirements.txt`:
   * **Name:** `23coffee`
   * **Runtime:** `Python`
   * **Build Command:** `pip install -r requirements.txt`
   * **Start Command:** `gunicorn app:app`
5. Bấm **Create Web Service**.
6. Sau 1-2 phút, bạn sẽ nhận được đường link web HTTPS miễn phí (ví dụ: `https://23coffee.onrender.com`).
7. Mở link trên iPhone/iPad và đăng nhập bằng tài khoản `admin` hoặc `tandat` (mật khẩu `admin`)!

---

## 📁 6. Cấu Trúc Thư Mục Dự Án

```
23coffee/
├── app.py                      # Backend Flask & Xác thực phân quyền 2 người
├── db_service.py               # Cơ sở dữ liệu SQLite lưu trữ dữ liệu vĩnh viễn
├── google_sheets_service.py    # Dịch vụ tự động đồng bộ hai chiều Google Sheets
├── google_sheets_script.gs     # Mã nguồn Google Apps Script Webhook
├── google_config.json          # Cấu hình liên kết Google Sheets
├── excel_service.py            # Động cơ tự động sinh và căn chỉnh file Excel 5 sheet
├── data_cost.json              # Bộ nhớ đệm dữ liệu món, nguyên liệu, chi phí
├── requirements.txt            # Danh sách thư viện Python (Flask, openpyxl, requests...)
├── Procfile                    # Cấu hình máy chủ web cho Render/Railway
├── render.yaml                 # Cấu hình 1-click deploy lên Render
├── Chay_App_Quan_Ly_Cost.bat   # Phím tắt chạy 1-click trên máy tính Windows
├── templates/
│   ├── index.html              # Giao diện chính hệ thống quản trị (hỗ trợ iPhone/iPad)
│   └── login.html              # Màn hình đăng nhập bảo mật
└── README.md                   # Hướng dẫn chi tiết
```
