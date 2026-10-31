# ☕ 23 COFFEE — Hệ Thống Quản Lý Cost & Tự Động Hóa Menu F&B

Ứng dụng web nội bộ giúp chủ quán **23 COFFEE** tự động hóa toàn bộ quy trình tính cost, bóc tách công thức SOP, định giá bán theo chuẩn 30% Food Cost và đồng bộ trực tiếp vào file Excel chuẩn 5 Sheet.

---

## 🔒 1. Chính Sách Bảo Mật (Giới Hạn Chỉ Đúng 2 Người Truy Cập)

Hệ thống được khóa chặt bằng cơ chế **Session Authentication** (bảo vệ mọi route và API). Khi bạn public web lên internet, chỉ có 2 tài khoản sau mới có quyền đăng nhập và xem dữ liệu giá vốn nội bộ:

| STT | Vai Trò | Tên Đăng Nhập | Mật Khẩu Mặc Định | Ghi Chú |
| :-: | :--- | :---: | :---: | :--- |
| **1** | **Chủ Quán (Tấn Đạt)** | `tandat` | `tandat@23coffee` | Toàn quyền quản trị |
| **2** | **Cổ Đông / Quản Lý** | `partner` | `partner@23coffee` | Toàn quyền quản trị |

> 💡 **Đổi mật khẩu:** Sau khi đăng nhập, bạn có thể bấm vào biểu tượng 🔑 (chìa khóa) ở góc trên bên phải để đổi sang mật khẩu riêng bất kỳ lúc nào.

---

## 🌟 2. Các Tính Năng Nổi Bật

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

## 💻 3. Hướng Dẫn Chạy Trên Máy Tính Cá Nhân (Local)

1. Bấm đúp chuột vào tệp: **`Chay_App_Quan_Ly_Cost.bat`**.
2. Trình duyệt sẽ tự động mở trang web tại địa chỉ: **`http://localhost:5000`**.
3. Đăng nhập bằng tài khoản `tandat` hoặc `partner`.

---

## 🌐 4. Hướng Dẫn Đưa Lên Mạng Công Khai (Public Web Miễn Phí 100%)

Để bạn và cổ đông có thể truy cập bằng điện thoại từ bất cứ đâu:

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
7. Mở link trên điện thoại và đăng nhập bằng tài khoản `tandat` hoặc `partner`!

---

## 📁 5. Cấu Trúc Thư Mục Dự Án

```
23coffee/
├── app.py                      # Backend Flask & Xác thực phân quyền 2 người
├── excel_service.py            # Động cơ tự động sinh và căn chỉnh file Excel 5 sheet
├── data_cost.json              # Cơ sở dữ liệu món, nguyên liệu, chi phí
├── requirements.txt            # Danh sách thư viện Python (Flask, openpyxl, gunicorn...)
├── Procfile                    # Cấu hình máy chủ web cho Render/Railway
├── render.yaml                 # Cấu hình 1-click deploy lên Render
├── Chay_App_Quan_Ly_Cost.bat   # Phím tắt chạy 1-click trên máy tính Windows
├── templates/
│   ├── index.html              # Giao diện chính hệ thống quản trị
│   └── login.html              # Màn hình đăng nhập bảo mật
└── README.md                   # Hướng dẫn chi tiết
```
