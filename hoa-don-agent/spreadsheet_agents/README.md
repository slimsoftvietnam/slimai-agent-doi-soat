# Agent đối soát hóa đơn, thanh toán và bảng tính

Gồm ba tác tử chuyên trách:

1. **Đọc hóa đơn**: nhận dạng MST, số hóa đơn, ngày và số tiền từ PDF/ảnh; xuất Excel và tô đỏ dòng cần rà soát.
2. **Làm sạch, đối chiếu bảng tính**: chuẩn hóa khoảng trắng, MST, ngày, số tiền; đánh dấu dòng trùng; ghép hai bảng theo khóa và nêu sai khác.
3. **Đối soát chuyển khoản** (luồng chính của tình huống): dò ảnh chuyển khoản tải từ Zalo với danh sách đơn, trả lời ai đã trả đủ, ai trả thiếu, ai chưa trả.

Thiết kế theo khung PEAS tại [PEAS.md](PEAS.md). Tình huống sử dụng, quy tắc quyết định và điều kiện hoàn thành nằm trong [THE_NHIEM_VU.md](THE_NHIEM_VU.md).

## Cài đặt

Xem [INSTALL.md](../INSTALL.md). Các lệnh bên dưới chạy từ thư mục skill `hoa-don-agent/` (trên Windows dùng `spreadsheet_agents\.venv\Scripts\python.exe`).

Tác tử hóa đơn kế thừa pipeline từ repo Smart-OCR có giấy phép MIT: trích văn bản PDF, nhận dạng chữ trên trang quét/ảnh bằng Tesseract, chuẩn hóa dữ liệu rồi xuất Excel. Mã nguồn cần thiết và giấy phép được giữ trong [smart_ocr_base](smart_ocr_base/). Phần nhận dạng đã được điều chỉnh để hiểu nhãn MST và hóa đơn tiếng Việt.

Cần cài Tesseract cùng dữ liệu ngôn ngữ tiếng Việt (`vie`) để nhận dạng tiếng Việt tốt hơn. Có thể cấu hình `TESSERACT_CMD` và `TESSERACT_LANG` trong `.env`.

Mặc định hóa đơn được xử lý tại máy, không gửi lên dịch vụ ngoài. Để dùng Gemini nhận dạng ngữ nghĩa, đặt `GEMINI_API_KEY` trong `.env` rồi thêm `--che-do gemini`. Có thể đặt `GEMINI_MODEL`; mặc định là `gemini-3.5-flash`. Tác tử bảng tính chạy cục bộ.

## Đọc hóa đơn

```sh
spreadsheet_agents/.venv/bin/python spreadsheet_agents/agent.py hoa-don ./hoa-don/ --dau-ra ./output/hoa-don.xlsx
spreadsheet_agents/.venv/bin/python spreadsheet_agents/agent.py hoa-don ./hoa-don/ --dau-ra ./output/hoa-don.xlsx --che-do gemini
```

Hỗ trợ PDF, JPG, JPEG và PNG. Mỗi hóa đơn tạo một dòng trong trang `Hóa đơn`; các mặt hàng được đưa vào trang `Chi tiết mặt hàng`. Ảnh được OCR toàn trang và quét thêm ba vùng chồng lấn: thông tin đầu trang, bảng hàng hóa và tổng tiền. Excel có thêm địa chỉ hai bên, MST người mua, phương thức thanh toán và số tiền bằng chữ.

Nếu OCR thấp, trường quan trọng bị thiếu hoặc phép cộng tiền không khớp, cột `Cần rà soát` ghi lý do và hàng được tô đỏ. Kết quả đọc được vẫn được giữ để đối chiếu, kèm ghi chú khi độ tin cậy thấp. Điểm OCR không xác nhận tính hợp pháp của hóa đơn; chữ viết tay mờ vẫn cần người dùng xác nhận.

## Làm sạch bảng tính

```sh
spreadsheet_agents/.venv/bin/python spreadsheet_agents/agent.py lam-sach ./nguon.xlsx --dau-ra ./output/nguon-sach.xlsx --trang-tinh Sheet1 --khoa "Mã hóa đơn"
```

Chuẩn hóa văn bản, MST, ngày và số tiền; thêm cột `Trạng thái dữ liệu` và `Lý do rà soát`. Dòng có ô trống, khóa trùng hoặc khóa rỗng được tô đỏ nhưng không bị xóa. Bỏ `--khoa` nếu không cần dò trùng.

## Đối chiếu hai bảng tính

```sh
spreadsheet_agents/.venv/bin/python spreadsheet_agents/agent.py doi-chieu ./ke-toan.xlsx ./nha-cung-cap.xlsx \
  --trang-trai Sheet1 --trang-phai Sheet1 --khoa "Mã hóa đơn" \
  --cot "MST,Ngày hóa đơn,Tổng tiền" --dau-ra ./output/doi-chieu.xlsx
```

Bảng kết quả có trang `Đối chiếu`, giá trị hai bên được đặt cạnh nhau theo khóa. Sai khác, khóa trùng, thiếu khóa và thiếu giá trị ở cột được chọn đều được giữ lại và tô đỏ; khóa chỉ có ở một bên được tô vàng.

## Đối soát ảnh chuyển khoản với đơn hàng

Tình huống, mục tiêu và PEAS: [THE_NHIEM_VU.md](THE_NHIEM_VU.md). Tải ảnh biên nhận từ Zalo vào một thư mục, xuất danh sách đơn ra Excel, rồi chạy:

```sh
spreadsheet_agents/.venv/bin/python spreadsheet_agents/agent.py doi-soat-thanh-toan ./anh-zalo/ ./danh-sach-don.xlsx
```

Agent tự nhận cột mã đơn, tên khách, số phải thu và SĐT theo các tên thường gặp (`Mã đơn`, `Mã ĐH`, `Khách hàng`, `Họ tên`, `Phải thu`, `Tổng tiền`, `SĐT`...). Tên cột khác thì chỉ định bằng `--cot-ma-don`, `--cot-khach-hang`, `--cot-so-tien`, `--cot-dien-thoai`; trang tính khác trang đầu thì thêm `--trang-tinh`. Kết quả mặc định ở `output/doi-soat-<ngày>.xlsx`; tên đã có thì tự thêm số, không ghi đè.

Excel kết quả có ba trang:

- `Tóm tắt`: số đơn và tên khách theo bốn nhóm **Đã trả đủ · Trả thiếu · Chưa trả · Cần kiểm tra**, tổng phải thu và tổng đã nhận theo ảnh.
- `Đơn hàng`: mỗi đơn một dòng, có số còn thiếu, ảnh đã ghép, căn cứ ghép, việc cần làm và tin nhắn nhắc soạn sẵn (agent không gửi).
- `Ảnh chuyển khoản`: mỗi ảnh một dòng, có số tiền, nội dung, người chuyển, đơn được ghép và kết quả.

Ghép chắc khi có mã đơn, số điện thoại, hoặc tên khách kèm đúng số tiền. Chỉ có tên hoặc chỉ trùng số tiền thì agent gợi ý và đưa đơn vào `Cần kiểm tra`, không báo `Chưa trả`. Ảnh giống hệt hoặc trùng mã giao dịch không được cộng lần hai. Nhận dư cũng vào `Cần kiểm tra`. Số tiền lấy từ ảnh khách gửi; đơn giá trị lớn cần xem sao kê trước khi giao.

Bộ dữ liệu mẫu: [`bo_du_lieu_mau_doi_soat/`](../../bo_du_lieu_mau_doi_soat/README.md), gồm 17 ảnh chuyển khoản (5 kiểu màn hình), 8 hóa đơn điện tử, file xuất từ phần mềm bán hàng, kết quả dựng sẵn và đáp án. Tạo lại bằng `spreadsheet_agents/tao_bo_du_lieu_mau.py`.

### Đối chiếu hóa đơn điện tử với bảng kê phần mềm

Đọc hóa đơn bằng `hoa-don`, rồi `doi-chieu` với trang bảng kê, khóa `Số hóa đơn`:

```sh
spreadsheet_agents/.venv/bin/python spreadsheet_agents/agent.py doi-chieu output/hoa-don.xlsx ./phan-mem.xlsx \
  --trang-trai "Hóa đơn" --trang-phai "Hóa đơn điện tử" --khoa "Số hóa đơn" \
  --cot "Ngày hóa đơn,MST người mua,Tổng tiền" --dau-ra output/doi-chieu-hoa-don.xlsx
```

Hai bên cùng trống (khách lẻ không có MST) được coi là khớp. Ảnh chụp hóa đơn in cũng được đọc theo nhãn của hóa đơn điện tử; thiếu dòng thuế thì tô đỏ vì chưa kiểm được phép cộng.

Tên trang tính và tên cột phải khớp với tệp nguồn. Tác tử giữ nguyên tệp gốc, không tự xóa dòng trùng hoặc sửa giá trị nghiệp vụ.
