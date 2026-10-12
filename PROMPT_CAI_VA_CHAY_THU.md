# Một prompt: tải skill + bộ dữ liệu mẫu, rồi chạy thử

Mở Claude Code hoặc Codex tại một thư mục trống, dán nguyên prompt bên dưới.

```text
Hãy cài và chạy thử skill đối soát chuyển khoản và hóa đơn từ repo công khai:
https://github.com/slimsoftvietnam/slimai-agent-doi-soat

Làm tuần tự trong thư mục làm việc hiện tại:
1. Tải repo vào thư mục con `slimai-agent-doi-soat` bằng git clone. Nếu máy không có git, tải ZIP nhánh main từ GitHub và giải nén. Đọc README.md và hoa-don-agent/INSTALL.md trước khi cài.
2. Cài skill bằng cách chép nguyên thư mục `hoa-don-agent` vào `.claude/skills/hoa-don-agent/` nếu đây là Claude Code, hoặc `.agents/skills/hoa-don-agent/` nếu đây là Codex. Tạo thư mục cha nếu cần. Nếu skill đích đã tồn tại, cập nhật thư mục đó theo bản trong repo.
3. Cần Python 3.10+ và Tesseract có ngôn ngữ tiếng Việt (`vie`). Kiểm tra trước khi cài. Nếu thiếu phần mềm hệ thống, hỏi tôi trước khi cài; nếu đã có đủ, tạo môi trường `.venv` và cài `hoa-don-agent/spreadsheet_agents/requirements.txt` theo INSTALL.md.
4. Đọc `hoa-don-agent/SKILL.md`, `hoa-don-agent/spreadsheet_agents/THE_NHIEM_VU.md` và README của bộ dữ liệu mẫu. Chạy thử đủ ba bước bằng chế độ OCR cục bộ, không dùng API ngoài:
   - Đối soát ảnh trong `bo_du_lieu_mau_doi_soat/1-anh-chuyen-khoan-zalo` với file `bo_du_lieu_mau_doi_soat/3-xuat-tu-phan-mem/phan-mem-ban-hang-12-10-2026.xlsx`, lưu `output/doi-soat-chuyen-khoan.xlsx`.
   - Đọc hóa đơn trong `bo_du_lieu_mau_doi_soat/2-hoa-don-dien-tu`, lưu `output/hoa-don-doc-duoc.xlsx`.
   - Đối chiếu kết quả đọc hóa đơn với trang `Hóa đơn điện tử` trong file bán hàng, khóa `Số hóa đơn`, so các cột `Ngày hóa đơn,MST người mua,Tổng tiền`, lưu `output/doi-chieu-hoa-don.xlsx`.
5. Không sửa/xóa dữ liệu mẫu hoặc ghi đè file nguồn. So kết quả với `bo_du_lieu_mau_doi_soat/DAP_AN.md`; báo bốn nhóm đơn hàng (đã trả đủ, trả thiếu, chưa trả, cần kiểm tra), các hóa đơn khớp/lệch/trùng/chỉ có ở một bên, và mọi khác biệt so với đáp án. Nếu OCR không chắc, giữ nguyên trạng thái cần rà soát. Mở các file kết quả khi hoàn tất.
```

**Kết quả đúng cho đơn hàng:** 10 đã trả đủ · 1 trả thiếu (DH1002, thiếu 500.000đ) · 2 chưa trả (DH1004, DH1016) · 3 cần kiểm tra (DH1006, DH1009, DH1011).

**Kết quả đúng cho hóa đơn:** 00000123 lệch tổng tiền · 00000124 lệch ngày · 00000127 lệch MST người mua · 00000126 trùng · 00000125 chỉ có trong phần mềm · 00000128 chỉ có trong tệp hóa đơn. 00000122 khớp nhưng chưa kiểm được phép cộng thuế từ ảnh.

Toàn bộ bộ dữ liệu là giả lập, không có dữ liệu khách hàng thật.
