# Đáp án bộ dữ liệu mẫu · Ngày 12/10/2026

Dùng để **nghiệm thu**: chạy agent xong, so kết quả với bảng dưới. Agent đưa thêm một đơn vào *Cần kiểm tra* thì chấp nhận được. Agent báo **"Đã trả đủ" cho đơn chưa trả đủ** thì là lỗi.

## A. Đơn hàng × ảnh chuyển khoản (16 đơn, 17 ảnh)

| Đơn | Khách | Cần trả | Ảnh (đuôi tên tệp) | Chuyện gì xảy ra | Kết quả đúng |
|---|---|---:|---|---|---|
| DH1001 | Nguyễn Văn An | 450.000 | `…353390` | Ghi mã đơn, đủ tiền | **Đã trả đủ** |
| DH1002 | Trần Thị Bình | 1.200.000 | `…411205` | Chuyển trước 700.000 | **Trả thiếu** 500.000 |
| DH1003 | Lê Minh Châu | 320.000 | `…502117`, `…002388` | Nội dung là SĐT; khách gửi lại đúng ảnh cũ | **Đã trả đủ**, ảnh thứ hai là **trùng** |
| DH1004 | Phạm Quốc Dũng | 890.000 | — | Chưa gửi ảnh | **Chưa trả** |
| DH1005 | Hoàng Thu Hà | 560.000 | `…618842`, `…733015` | Trả 2 lần: 300.000 + 260.000 | **Đã trả đủ** |
| DH1006 | Võ Thanh Em | 250.000 | `…840331` | Người nhà (Võ Văn Phúc) chuyển hộ, nội dung "chuyen tien" | **Cần kiểm tra** (chỉ khớp số tiền) |
| DH1007 | Đỗ Gia Huy | 1.500.000 | `…921456` | Không ghi mã, có tên + đúng tiền | **Đã trả đủ** |
| DH1008 | Cty Nội thất Gỗ Việt | 5.400.000 | `…034470` | Doanh nghiệp, chụp internet banking | **Đã trả đủ** (đơn lớn: xem sao kê) |
| DH1009 | Bùi Ngọc Lan | 780.000 | `…115902` | Chuyển 800.000 | **Cần kiểm tra** (dư 20.000) |
| DH1010 | Đặng Hữu Phước | 350.000 | `…207719` | Ảnh nghiêng, mờ, bị nén | **Đã trả đủ** |
| DH1011 | Ngô Thị Kim | 620.000 | `…350288` | Khách cắt ảnh, mất phần số tiền | **Cần kiểm tra** (không đọc được số tiền) |
| DH1012 | Cty Sự kiện Ánh Dương | 3.240.000 | `…462531` | Doanh nghiệp, internet banking | **Đã trả đủ** (đơn lớn: xem sao kê) |
| DH1013 | Trịnh Văn Long | 1.150.000 | `…573004` | Chuyển **1 lần 1.550.000 cho 2 đơn** | **Đã trả đủ** |
| DH1014 | Trịnh Văn Long | 400.000 | `…573004` | (chung ảnh với DH1013) | **Đã trả đủ** |
| DH1015 | Lý Thanh Tâm | 2.100.000 | `…680147` | Ví điện tử, mã đơn ở "Lời nhắn" | **Đã trả đủ** (đơn lớn: xem sao kê) |
| DH1016 | Mai Phương Thảo | 990.000 | — | Chưa gửi ảnh | **Chưa trả** |

**Ảnh không thuộc đơn nào hôm nay**

| Ảnh | Chuyện gì xảy ra | Kết quả đúng |
|---|---|---|
| `…791260` | Phan Văn Quang trả đơn **DH0987 của hôm qua**, không có trong danh sách | **Chưa ghép được**, giữ lại để tra đơn cũ |
| `…115630` | Khách gửi nhầm ảnh bó hoa | **Không phải biên nhận** |

**Tổng:** 10 đã trả đủ · 1 trả thiếu · 2 chưa trả · 3 cần kiểm tra.

## B. Hóa đơn điện tử × bảng kê phần mềm (8 tệp, 8 số hóa đơn)

| Số HĐ | Đơn | Tệp | Chuyện gì xảy ra | Kết quả đúng |
|---|---|---|---|---|
| 00000121 | DH1001 | PDF | Khớp | **Khớp** |
| 00000122 | DH1002 | **Ảnh chụp** hóa đơn in | Khớp, nhưng ảnh không đọc được dòng thuế | **Khớp**; dòng tô đỏ "chưa kiểm được phép cộng" |
| 00000123 | DH1005 | PDF | Phần mềm gõ sai tổng tiền: 650.000, hóa đơn 560.000 | **Sai khác: Tổng tiền** |
| 00000124 | DH1007 | PDF | Hóa đơn ngày 11/10, phần mềm ghi 12/10 | **Sai khác: Ngày** |
| 00000125 | DH1009 | — | Phần mềm có, không thấy tệp hóa đơn | **Chỉ có ở phần mềm** |
| 00000126 | DH1008 | PDF ×2 | Tải trùng một hóa đơn hai lần | **Khóa trùng** |
| 00000127 | DH1012 | PDF | MST người mua trên hóa đơn sai 2 số cuối (…201, đúng là …210) | **Sai khác: MST người mua** |
| 00000128 | DH1015 | PDF | Có hóa đơn, phần mềm chưa ghi nhận | **Chỉ có ở tệp hóa đơn** |

**Câu hỏi gợi mở cho lớp:** DH1002 mới trả 700.000 nhưng đã xuất hóa đơn 1.200.000. Agent có nên tự báo chuyện này không? Nếu muốn, thêm một dòng vào **P** và **S** rồi bảo AI sửa.
