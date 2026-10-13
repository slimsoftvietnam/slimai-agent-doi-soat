# THẺ NHIỆM VỤ · AGENT ĐỐI SOÁT THANH TOÁN ZALO

**Người dùng:** sales, chủ shop, kế toán  
**Đầu vào:** ảnh chuyển khoản tải từ Zalo và Excel danh sách đơn hàng  
**Đầu ra:** Excel đối soát (trang `Tóm tắt` là dashboard) và dashboard HTML cùng số liệu, theo [phong cách dashboard](DASHBOARD_STYLE.md)

## Tình huống thực tế

Khách chuyển khoản rồi gửi ảnh qua Zalo. Cuối ngày, kế toán phải mở từng ảnh và dò với danh sách đơn để biết:

- Đơn nào đã đủ tiền theo ảnh.
- Khách nào trả thiếu.
- Đơn nào chưa có ảnh phù hợp.
- Ảnh nào bị trùng, bị cắt, chuyển hộ hoặc cần kiểm tra.

Agent đọc ảnh, chuẩn hóa dữ liệu, ghép ảnh với đơn, chặn giao dịch trùng và tính số tiền còn thiếu hoặc nhận dư.

## Kết quả

Agent phân loại đơn thành:

1. **Đã trả đủ:** tổng tiền trên ảnh ghép chắc bằng số phải thu.
2. **Trả thiếu:** số tiền trên ảnh thấp hơn số phải thu.
3. **Chưa có ảnh khớp:** chưa tìm thấy ảnh đủ căn cứ.
4. **Cần kiểm tra:** ảnh mờ, thiếu dữ liệu, chuyển hộ, nhận dư hoặc có nhiều đơn cùng khớp.

Excel lưu chi tiết từng đơn và từng ảnh; trang `Tóm tắt` là **dashboard Excel**. File `.html` cùng tên là **dashboard HTML**. Cả hai cùng phong cách ([DASHBOARD_STYLE.md](DASHBOARD_STYLE.md)): 4 số chính (tổng giá trị đơn, còn phải thu, đơn sẵn sàng giao, khách cần nhắc), tổng quan thu tiền theo giờ, ý chính hôm nay, việc cần làm, giá trị theo trạng thái, ảnh cần xem lại, hóa đơn. Dashboard HTML có thêm bộ lọc đơn, nút chép tin nhắn, chế độ tối và in PDF; xem tốt trên điện thoại.

## Mô hình PEAS

| Thành phần | Nội dung |
|---|---|
| **P · Performance** | Kết quả phải cho biết rõ: **đơn nào đủ tiền theo ảnh; đơn nào trả thiếu và thiếu bao nhiêu; đơn nào chưa có ảnh khớp; ảnh nào cần người xem lại; tổng phải thu, đã nhận và còn thiếu**. 100% đơn và ảnh phải xuất hiện trong báo cáo; ảnh trùng không được cộng hai lần; trường hợp mơ hồ không được tự ghép; số liệu dashboard Excel và dashboard HTML phải khớp nhau và khớp trang `Đơn hàng`; 4 số chính nhìn thấy ngay khi mở, không cần cuộn; mỗi ngoại lệ phải có lý do và việc cần làm. |
| **E · Environment** | Ảnh từ nhiều ngân hàng, chất lượng không đồng đều; Excel xuất từ phần mềm bán hàng; khách có thể chuyển thiếu, chuyển hộ, trả nhiều lần hoặc gửi trùng ảnh. Agent chỉ thấy ảnh và danh sách đơn, không thấy sao kê ngân hàng. |
| **A · Actuators** | OCR ảnh; chuẩn hóa dữ liệu; phát hiện trùng; ghép ảnh với đơn; tính tiền; phân loại; tạo việc cần làm; xuất Excel có trang dashboard và dashboard HTML cùng phong cách. Agent không gửi Zalo, xác nhận giao hàng, hoàn tiền hoặc sửa dữ liệu nguồn. |
| **S · Sensors** | Từ ảnh: số tiền, nội dung, người chuyển, mã giao dịch, thời gian và tên tệp. Từ Excel: mã đơn, khách hàng, số điện thoại và số phải thu. Từ người dùng: đường dẫn, tên trang và ánh xạ cột. |

> `Đã trả đủ` chỉ có nghĩa là đủ theo ảnh đã cung cấp. Sao kê ngân hàng là nguồn xác nhận cuối cùng.
