# Agent đối soát chuyển khoản và hóa đơn · SlimAI

Khách chuyển khoản xong chụp màn hình gửi qua Zalo. Cuối ngày phải dò từng ảnh với danh sách đơn: **ai trả rồi, ai trả thiếu, ai chưa trả**. Agent này làm phần dò đó, và đối chiếu thêm hóa đơn điện tử với bảng kê trong phần mềm.

Tài liệu thực hành cho tình huống 12, [Khóa học AI Agent cho người đi làm](https://ai.slim.vn) của SlimAI.

## Bắt đầu bằng một prompt

Mở Claude Code hoặc Codex tại một thư mục trống rồi dán prompt trong [PROMPT_CAI_VA_CHAY_THU.md](PROMPT_CAI_VA_CHAY_THU.md). AI tự tải repo, cài skill, chạy thử trên bộ dữ liệu mẫu và so với đáp án.

## Trong repo có gì

| Thư mục | Nội dung |
|---|---|
| [`hoa-don-agent/`](hoa-don-agent/SKILL.md) | Skill cho Claude Code / Codex: `SKILL.md`, [cài đặt](hoa-don-agent/INSTALL.md), mã chạy trong `spreadsheet_agents/` |
| [`hoa-don-agent/spreadsheet_agents/THE_NHIEM_VU.md`](hoa-don-agent/spreadsheet_agents/THE_NHIEM_VU.md) | Thẻ nhiệm vụ: tình huống, mục tiêu, kết quả cần đạt, PEAS, luật ghép, nghiệm thu |
| [`bo_du_lieu_mau_doi_soat/`](bo_du_lieu_mau_doi_soat/README.md) | 17 ảnh chuyển khoản, 8 hóa đơn điện tử, file xuất từ phần mềm bán hàng, kết quả dựng sẵn và [đáp án](bo_du_lieu_mau_doi_soat/DAP_AN.md) |

## PEAS trong một bảng

| Chữ | Dịch | Nhớ nhanh | Agent này |
|---|---|---|---|
| **P** · Performance measure | Thước đo hiệu quả | KPI | 100% đơn có trạng thái; **0 đơn báo "Đã trả đủ" sai**; 0 ảnh cộng hai lần |
| **E** · Environment | Môi trường | Bàn làm việc | Ảnh Zalo nhiều ngân hàng, hóa đơn PDF/ảnh, file Excel từ phần mềm |
| **A** · Actuators | Cơ cấu chấp hành | Đôi tay | Đọc, ghép, cộng, tô màu, soạn tin nhắc. **Không** nhắn khách, không ghép chỉ vì trùng số tiền, không sửa file gốc |
| **S** · Sensors | Cảm biến | Đôi mắt | Số tiền, nội dung chuyển khoản, người chuyển, mã giao dịch; mã đơn, SĐT, số phải thu |

## Lưu ý

- Bộ dữ liệu mẫu **hoàn toàn giả lập**: ngân hàng, doanh nghiệp, MST, số tài khoản và khách hàng đều không có thật. Hóa đơn mẫu không có giá trị pháp lý.
- Agent chỉ thấy ảnh khách gửi, không thấy tài khoản ngân hàng. Đơn giá trị lớn vẫn phải xem sao kê trước khi giao.
- Phần đọc hóa đơn kế thừa [avi-hue/smart-ocr](https://github.com/avi-hue/smart-ocr) (giấy phép MIT, xem `hoa-don-agent/spreadsheet_agents/smart_ocr_base/LICENSE`).
