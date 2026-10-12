# Cài đặt Agent đối soát chuyển khoản và hóa đơn

## 1. Đặt skill

Chép thư mục `hoa-don-agent` (giữ nguyên cấu trúc) vào:

| Công cụ | Theo dự án | Dùng chung mọi dự án |
|---|---|---|
| Claude Code | `<dự-án>/.claude/skills/hoa-don-agent` | `~/.claude/skills/hoa-don-agent` |
| Codex | `<dự-án>/.agents/skills/hoa-don-agent` | `~/.codex/skills/hoa-don-agent` |

## 2. Cài Python 3.10+ và Tesseract (có tiếng Việt)

- macOS: `brew install python@3.12 tesseract tesseract-lang`
- Ubuntu/Debian: `sudo apt-get install python3 python3-venv tesseract-ocr tesseract-ocr-vie`
- Windows: cài Python từ python.org, cài Tesseract bản UB Mannheim (`winget install UB-Mannheim.TesseractOCR`), chọn thêm ngôn ngữ Vietnamese khi cài. Nếu `tesseract` không có trong PATH, tạo tệp `hoa-don-agent/.env` với dòng `TESSERACT_CMD=C:\Program Files\Tesseract-OCR\tesseract.exe`.

Kiểm tra: `tesseract --list-langs` phải có `vie`.

## 3. Tạo môi trường chạy

Trong thư mục `hoa-don-agent`:

```sh
python3 -m venv spreadsheet_agents/.venv
spreadsheet_agents/.venv/bin/python -m pip install -r spreadsheet_agents/requirements.txt
```

Windows PowerShell:

```powershell
py -m venv spreadsheet_agents\.venv
spreadsheet_agents\.venv\Scripts\python.exe -m pip install -r spreadsheet_agents\requirements.txt
```

## 4. Kiểm tra nhanh

```sh
spreadsheet_agents/.venv/bin/python spreadsheet_agents/agent.py doi-soat-thanh-toan --help
```

Mở lại Claude Code hoặc Codex để skill được nhận, rồi gõ `/hoa-don-agent` (Codex: `$hoa-don-agent`).

## Quyền riêng tư

Mặc định xử lý tại máy bằng Tesseract, không gửi ảnh đi đâu. Chỉ khi chọn `--che-do gemini` cho việc đọc hóa đơn mới cần `GEMINI_API_KEY` trong `.env`. Không chia sẻ `.env`, ảnh chuyển khoản thật hay hóa đơn thật cùng gói skill.
