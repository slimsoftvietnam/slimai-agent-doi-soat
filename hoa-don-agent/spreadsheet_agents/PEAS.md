# Thiết kế tác tử xử lý hóa đơn và bảng tính theo PEAS

PEAS gồm **Performance measure** (thước đo hiệu quả), **Environment** (môi trường), **Actuators** (cơ cấu hành động) và **Sensors** (cảm biến/đầu vào). Chương trình có một tác tử điều phối và ba tác tử chuyên trách. Việc kiểm tra dữ liệu và ghép bản ghi dùng quy tắc có thể giải thích; OCR hỗ trợ đọc hóa đơn và ảnh chuyển khoản.

## 1. Tác tử điều phối

| Thành phần | Thiết kế |
|---|---|
| **P — Thước đo hiệu quả** | Chọn đúng luồng theo lệnh; không ghi đè dữ liệu nguồn; lưu rõ nguồn và kết quả; dừng với thông báo cụ thể nếu thiếu tệp, thư viện hoặc khóa API. |
| **E — Môi trường** | Python 3.9+, hệ thống tệp cục bộ, môi trường ảo của dự án, Excel `.xlsx`; Gemini API tùy chọn khi đọc hóa đơn. |
| **A — Hành động** | Nhận lệnh `hoa-don`, `lam-sach`, `doi-chieu` hoặc `doi-soat-thanh-toan`; kiểm tra đầu vào; gọi tác tử chuyên trách; ghi bảng tính kết quả. |
| **S — Cảm biến** | Đường dẫn và phần mở rộng tệp, cấu trúc bảng tính, kết quả OCR/API, trường dữ liệu và lỗi đọc/ghi. |

## 2. Tác tử đọc hóa đơn

| Thành phần | Thiết kế |
|---|---|
| **P — Thước đo hiệu quả** | Trích được thông tin bên bán/mua, MST, số hóa đơn, ngày, phương thức thanh toán, số tiền và dòng hàng; giữ trường không đọc được là trống; hiển thị phần đọc được nhưng cảnh báo chữ viết tay/OCR yếu; phát hiện sai lệch số học để người dùng kiểm tra; không tự kết luận hóa đơn hợp lệ về thuế. |
| **E — Môi trường** | PDF có văn bản hoặc bản quét, ảnh JPG/JPEG/PNG, Python, PDFium/PyMuPDF, pdfplumber, Tesseract, Excel; Gemini API tùy chọn. Chất lượng ảnh, bố cục, chữ viết tay và dữ liệu ngôn ngữ OCR có thể ảnh hưởng độ chính xác. |
| **A — Hành động** | Phân loại trang, trích văn bản PDF, OCR toàn trang và OCR riêng vùng đầu trang/bảng hàng/tổng tiền; lấy trường bằng biểu thức/bảng hoặc Gemini; chuẩn hóa ngày/số; kiểm tra MST 10/13 chữ số, trường bắt buộc, dòng hàng và phép cộng; xuất trang tổng hợp cùng trang chi tiết mặt hàng; tô đỏ dòng cần rà soát. |
| **S — Cảm biến** | Văn bản PDF, ảnh từng trang/vùng ảnh, tên tệp, điểm tin cậy OCR, trường và dòng hàng đã trích, định dạng ngày và phép cộng tiền trước thuế + thuế GTGT. |

## 3. Tác tử làm sạch và đối chiếu bảng tính

| Thành phần | Thiết kế |
|---|---|
| **P — Thước đo hiệu quả** | Chuẩn hóa nhất quán khoảng trắng, MST, ngày và số tiền; không làm mất hàng/cột nguồn; ghép bản ghi theo khóa người dùng chọn; đánh dấu thiếu, trùng, chỉ có ở một phía hoặc sai khác. |
| **E — Môi trường** | Một hoặc hai tệp `.xlsx`, trang tính nguồn, tên cột, openpyxl và hệ thống tệp cục bộ. |
| **A — Hành động** | `lam-sach`: tạo bản sao đã chuẩn hóa, thêm trạng thái/lý do và tô đỏ dòng trùng hoặc cần rà soát. `doi-chieu`: ghép hai bảng theo khóa, so sánh các cột đã chọn và tạo trang kết quả; tô đỏ sai khác/thiếu, tô vàng dòng chỉ xuất hiện một phía. |
| **S — Cảm biến** | Tên trang tính/cột, giá trị từng ô, kiểu dữ liệu, khóa ghép, giá trị chuẩn hóa và số dòng/khóa trùng. |

## 4. Tác tử đối soát ảnh chuyển khoản với đơn hàng

Tình huống, mục tiêu, kết quả cần đạt, PEAS dịch từng chữ (P = thước đo hiệu quả / KPI, E = môi trường / bàn làm việc, A = cơ cấu chấp hành / đôi tay, S = cảm biến / đôi mắt), luật ghép và nghiệm thu nằm tại [THE_NHIEM_VU.md](THE_NHIEM_VU.md). Dữ liệu thử: [bo_du_lieu_mau_doi_soat](../../bo_du_lieu_mau_doi_soat/README.md).

| Thành phần | Thiết kế |
|---|---|
| **P — Thước đo hiệu quả** | 100% đơn có trạng thái, 100% ảnh có mặt trong kết quả; ≥ 95% ảnh rõ nét đọc đúng số tiền; 0 đơn bị báo "Đã trả đủ" sai; 0 giao dịch cộng hai lần; mỗi ảnh đã ghép có căn cứ. 50 ảnh xử lý và rà xong dưới 10 phút. |
| **E — Môi trường** | Thư mục ảnh JPG/PNG tải từ Zalo, nhiều ngân hàng/ví; Excel danh sách đơn với tên cột khác nhau. Khó: khách không ghi mã đơn, người nhà chuyển hộ, ảnh cắt mất số tiền, ảnh gửi trùng, một đơn trả nhiều lần, ảnh không phải biên nhận. Tesseract, Pillow, openpyxl chạy tại máy. |
| **A — Hành động** | Đọc chữ trên ảnh; ghép ảnh với đơn; cộng tiền theo đơn; xếp Đã trả đủ / Trả thiếu / Chưa trả / Cần kiểm tra; tạo Excel mới có Tóm tắt, tô màu, việc cần làm, tin nhắn nhắc soạn sẵn. **Không được:** gửi tin Zalo, xác nhận đơn hay giao hàng, sửa/xóa ảnh hoặc file đơn, ghi đè kết quả cũ, ghép chỉ vì trùng số tiền, đoán số tiền, cộng ảnh trùng, kết luận tiền đã vào tài khoản. |
| **S — Cảm biến** | Trên ảnh: số tiền, nội dung chuyển khoản, người chuyển, mã giao dịch, thời gian; dấu vân tay tệp ảnh. Trong file đơn: mã đơn, tên khách, số điện thoại, số phải thu. |

## Ranh giới quyết định

- Tác tử không sửa tệp nguồn và không xóa dòng trùng; nó chỉ đánh dấu để người dùng quyết định.
- Nếu OCR thấp, trường bắt buộc thiếu, MST sai định dạng hoặc phép cộng tiền không khớp thì dòng hóa đơn được tô đỏ và ghi lý do. Các trường yếu vẫn hiện khi có nội dung đọc được; ghi chú phân biệt chúng với dữ liệu đã xác nhận.
- So khớp chỉ dựa trên cột khóa được chỉ định. Tác tử không tự đoán hai dòng là cùng một giao dịch.
- Đối soát chuyển khoản ghép chắc bằng mã đơn, số điện thoại, hoặc tên khách kèm đúng số tiền. Chỉ có tên hoặc chỉ trùng số tiền thì chỉ gợi ý; đơn đó vào `Cần kiểm tra` chứ không bị báo `Chưa trả`.
- Ảnh giống hệt hoặc lặp mã giao dịch được đánh dấu trùng và không cộng tiền lần hai. Ảnh chưa ghép được luôn được giữ để người dùng kiểm tra.
- Tác tử xử lý ảnh đã tải về; không tự đọc hội thoại, gửi tin Zalo, xác nhận giao hàng hoặc thực hiện giao dịch tài chính.
- Không có cột thuế/tổng tiền thì bỏ qua phép kiểm tra tương ứng; kết quả không được coi là tư vấn thuế hay kiểm toán.

## Mã nguồn nền được tái sử dụng

Tác tử hóa đơn kế thừa quy trình từ [avi-hue/smart-ocr](https://github.com/avi-hue/smart-ocr), giấy phép MIT. Mã cần thiết và giấy phép nằm trong `smart_ocr_base/`; phần làm sạch/đối chiếu được giữ riêng để dùng khóa và quy tắc do người chạy chọn.
