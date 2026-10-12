# Bộ dữ liệu mẫu · Đối soát chuyển khoản và hóa đơn điện tử

Một ngày bán hàng (12/10/2026) của shop hoa giả lập **Công ty TNHH Hoa Mộc Demo**: 16 đơn, 17 ảnh khách gửi qua Zalo, 8 tệp hóa đơn điện tử và một file Excel xuất từ phần mềm bán hàng.

> **Toàn bộ là dữ liệu giả lập.** Tên ngân hàng (Sen Vàng Bank, Biển Xanh Bank, Đồng Tiến Bank, Ví Mây), doanh nghiệp, MST, số tài khoản và khách hàng đều không có thật. Hóa đơn không có giá trị pháp lý. Không dùng ảnh chuyển khoản thật của khách để dạy nếu chưa được đồng ý.

## Có gì trong bộ này

| Thư mục | Nội dung | Mô phỏng |
|---|---|---|
| `1-anh-chuyen-khoan-zalo/` | 17 ảnh JPG/PNG, tên tệp kiểu Zalo | 5 kiểu màn hình: app ngân hàng kiểu cũ, app ngân hàng mới, ví điện tử, biên lai, internet banking trên máy tính |
| `2-hoa-don-dien-tu/` | 7 PDF + 1 ảnh chụp | Hóa đơn GTGT điện tử có mã CQT, ký hiệu `1C26THM`, thuế suất 8% |
| `3-xuat-tu-phan-mem/phan-mem-ban-hang-12-10-2026.xlsx` | Trang `Đơn hàng` (16 đơn) và `Hóa đơn điện tử` (bảng kê 7 hóa đơn) | File xuất từ phần mềm bán hàng, tên cột kiểu `Mã ĐH`, `SĐT`, `Khách cần trả` |
| `4-ket-qua-dung-san/` | 3 file Excel kết quả và dashboard HTML | Dùng khi máy học viên lỗi, để vẫn làm được phần nghiệm thu |
| `DAP_AN.md` | Kết quả đúng cho từng đơn, ảnh, hóa đơn | Dùng để nghiệm thu |

## Các tình huống khó được cài sẵn

**Ảnh chuyển khoản:** trả đủ · trả thiếu · trả dư · trả hai lần · một lần trả cho hai đơn · người nhà chuyển hộ · không ghi mã đơn · nội dung là SĐT · ảnh gửi trùng · ảnh mờ, nghiêng · ảnh cắt mất số tiền · ảnh trả đơn hôm qua · ảnh không phải biên nhận · khách chưa gửi ảnh.

**Hóa đơn điện tử:** khớp · ảnh chụp hóa đơn in · phần mềm gõ sai tổng tiền · lệch ngày · sai MST người mua · tải trùng tệp · có trong phần mềm nhưng thiếu tệp · có tệp nhưng phần mềm chưa ghi nhận.

## Chạy agent trên bộ này

Từ thư mục gốc repo, sau khi cài theo [INSTALL.md](../hoa-don-agent/INSTALL.md) (đặt `PY=hoa-don-agent/spreadsheet_agents/.venv/bin/python`, `A=hoa-don-agent/spreadsheet_agents/agent.py`):

1. Đọc hóa đơn điện tử:

```bash
$PY $A hoa-don bo_du_lieu_mau_doi_soat/2-hoa-don-dien-tu --dau-ra output/hoa-don-doc-duoc.xlsx
```

2. Đối chiếu hóa đơn với bảng kê trong phần mềm:

```bash
$PY $A doi-chieu output/hoa-don-doc-duoc.xlsx bo_du_lieu_mau_doi_soat/3-xuat-tu-phan-mem/phan-mem-ban-hang-12-10-2026.xlsx --trang-trai "Hóa đơn" --trang-phai "Hóa đơn điện tử" --khoa "Số hóa đơn" --cot "Ngày hóa đơn,MST người mua,Tổng tiền" --dau-ra output/doi-chieu-hoa-don.xlsx
```

3. Đối soát ảnh chuyển khoản, đưa kết quả hóa đơn vào dashboard:

```bash
$PY $A doi-soat-thanh-toan bo_du_lieu_mau_doi_soat/1-anh-chuyen-khoan-zalo bo_du_lieu_mau_doi_soat/3-xuat-tu-phan-mem/phan-mem-ban-hang-12-10-2026.xlsx --doi-chieu-hoa-don output/doi-chieu-hoa-don.xlsx --dau-ra output/doi-soat-chuyen-khoan.xlsx
```

Bước 3 tạo thêm `output/doi-soat-chuyen-khoan.html`: dashboard mở bằng trình duyệt: số chính, ý chính hôm nay, việc cần làm, đơn hàng và hóa đơn trên một màn hình. Mở file này để xem kết quả.

Hoặc trong Claude Code: `/hoa-don-agent đối soát bộ mẫu bo_du_lieu_mau_doi_soat`.

## Tạo lại hoặc đổi số liệu

Sửa danh sách `ORDERS`, `TRANSFERS`, `INVOICES` trong [`tao_bo_du_lieu_mau.py`](../hoa-don-agent/spreadsheet_agents/tao_bo_du_lieu_mau.py), rồi chạy:

```bash
$PY hoa-don-agent/spreadsheet_agents/tao_bo_du_lieu_mau.py bo_du_lieu_mau_doi_soat
```

Nhớ cập nhật `DAP_AN.md` và chạy lại `4-ket-qua-dung-san/`.
