# Thông tin về mã nguồn Smart-OCR

Repo gốc: [avi-hue/smart-ocr](https://github.com/avi-hue/smart-ocr), lấy ngày 09/10/2026.

Thư mục này chứa các phần cần thiết của pipeline PDF/OCR/trích xuất trường, được điều chỉnh cho dự án. Giấy phép MIT gốc được giữ nguyên tại `LICENSE`.

Các điều chỉnh gồm nhãn hóa đơn và mẫu MST Việt Nam, mặc định Tesseract tiếng Việt/Anh kèm phương án lùi về tiếng Anh, cùng đường dẫn lưu tệp tạm. Tác tử ở `../agent.py` gọi pipeline này và tạo bảng tính tiếng Việt. Phần làm sạch và đối chiếu bảng tính là mã riêng của dự án.
