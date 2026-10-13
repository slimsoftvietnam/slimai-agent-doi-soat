# Phong cách dashboard thống kê

Agent xuất **hai dashboard cùng số liệu**, cùng một phong cách:

| Dashboard | Ở đâu | Dùng khi |
|---|---|---|
| **Excel** | Trang `Tóm tắt` trong file kết quả | Gửi kế toán, lưu sổ, in |
| **HTML** | File `.html` cùng tên, cạnh file Excel | Chủ shop, quản lý xem nhanh trên máy tính hoặc điện thoại |

File này dành cho AI dựng agent: đọc rồi làm dashboard đúng phong cách. Bản dựng sẵn để so: `4-ket-qua-dung-san/1-doi-soat-chuyen-khoan.html` và trang `Tóm tắt` trong `1-doi-soat-chuyen-khoan.xlsx`.

## 1. Cảm giác chung

Phần mềm quản trị cao cấp: nền xám xanh rất nhạt, các **thẻ trắng bo góc lớn**, gần như không đổ bóng, nhiều khoảng trắng, chữ đen rõ. Con số là thứ nổi nhất trên trang.

## 2. Màu

| Vai trò | Mã màu | Dùng cho |
|---|---|---|
| Nền trang | `#EEF1F5` | Toàn trang |
| Thẻ | `#FFFFFF` | Mọi khối nội dung |
| Nền phụ | `#F4F6F9` | Dòng danh sách, ô số phụ |
| Chữ chính | `#0D0F12` | Số lớn, tiêu đề |
| Chữ phụ | `#5B6170` / `#9097A3` | Nhãn, mô tả |
| **Xanh chanh** | `#C8F26D` (chữ `#1F2A06`) | Mục đang chọn, nhãn % tăng |
| **Cam** | `#F4772E` | Cột nổi bật (giờ tiền về nhiều nhất) |
| **Tím** | `#9B87F5` | Đường và vùng cộng dồn |
| Tăng / giảm | `#16A34A` ▲ / `#E5484D` ▼ | Dòng chênh lệch dưới số lớn |
| Trạng thái | xanh lá đủ · cam thiếu · đỏ chưa trả · vàng cần kiểm tra | Luôn kèm chữ, không dùng màu một mình |

Chế độ tối: đổi nền, thẻ, chữ; giữ xanh chanh, cam, tím.

## 3. Bố cục (từ trên xuống)

1. **Thanh trên cùng:** logo trái · menu dạng viên thuốc ở giữa (Tổng quan · Việc cần làm · Đơn hàng · Hóa đơn), mục đang xem tô xanh chanh · nút giao diện, chuông, thẻ "Báo cáo cuối ngày + ngày" bên phải.
2. **Lời chào:** "Đối soát cuối ngày 👋" + một dòng tóm tắt; bên phải nút "Xuất báo cáo".
3. **4 ô số chính:** Tổng giá trị đơn · Còn phải thu · Đơn sẵn sàng giao · Khách cần nhắc. Mỗi ô: nhãn, biểu tượng trong vòng tròn, **số lớn**, dòng ▲/▼ có màu, "Xem chi tiết ↗".
4. **Dòng nhắc:** số tiền lấy từ ảnh khách gửi, chưa phải xác nhận tiền đã vào tài khoản.
5. **Tổng quan thu tiền** (rộng 2/3): số đã thu + nhãn xanh chanh `84%`, hai số phụ (phải thu, còn phải thu), **biểu đồ cột gạch chéo xám**, cột giờ cao điểm **màu cam** có thẻ ghi chú.
6. **Ý chính hôm nay** (rộng 1/3): 3–4 dòng, mỗi dòng một ô nền nhạt: biểu tượng · câu kết luận in đậm · một câu giải thích.
7. **Giao dịch trong ngày:** số giao dịch, nhãn "Cao điểm 11h", **đường tím cộng dồn** có vùng mờ phía dưới.
8. **Việc cần làm:** lọc dạng viên thuốc; mỗi dòng: ô trạng thái màu · tên khách · mã đơn, SĐT · việc cần làm · số tiền · nút "Chép tin nhắn".
9. Giá trị theo trạng thái · Thu theo kênh bán · Ảnh khách gửi · Đơn lớn cần xem sao kê · Hóa đơn điện tử · Bảng tất cả đơn có lọc và tìm kiếm.

## 4. Số liệu

- Mỗi con số xuất hiện **một lần** ở chỗ nổi nhất; nơi khác chỉ nhắc khi giải thích.
- Tiền viết `16.970.000đ`; HTML có nút đổi sang rút gọn `16,97 tr`.
- Số trên Excel và HTML **phải khớp nhau** và khớp trang `Đơn hàng`.
- "Đã trả đủ" luôn hiểu là **đủ theo ảnh**; đơn lớn phải xem sao kê.

## 5. Thiết bị và hiệu ứng (HTML)

- Rộng tối đa 1240px, căn giữa. Máy tính bảng: 2 cột. Điện thoại: 1 cột, 4 ô số chính thành lưới 2×2, bảng đơn thành thẻ, menu vuốt ngang.
- Hiệu ứng nhẹ: khối hiện dần, số chạy từ 0, cột mọc từ đáy, đường tím tự vẽ. Tắt khi máy bật "giảm chuyển động" và khi in.
- Một file HTML tự chứa, mở offline, không tải gì từ mạng.

## 6. Excel

- Tắt đường lưới; tô nền `#EEF1F5`, các khối ô trắng làm thẻ; font Arial.
- Biểu đồ cột dùng hoa văn gạch chéo, cột giờ cao điểm màu cam; biểu đồ đường tím cho tiền cộng dồn; biểu đồ thanh màu theo trạng thái. Số liệu biểu đồ để ở trang ẩn `_bieu_do`.
- In ngang, vừa một trang chiều rộng.
