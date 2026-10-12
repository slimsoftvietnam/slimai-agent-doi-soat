"""Dashboard cho kết quả đối soát: trang Tóm tắt trong Excel và tệp HTML tự chứa.

Nhận dữ liệu đã đối soát từ agent.py; không đọc ảnh, không tính lại trạng thái.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from datetime import date, datetime
import json
from pathlib import Path
import re

STATUS_ORDER = ("Đã trả đủ", "Trả thiếu", "Chưa trả", "Cần kiểm tra")
# Bảng màu trạng thái cố định: tốt / nghiêm trọng vừa / nghiêm trọng / cảnh báo.
STATUS_HEX = {"Đã trả đủ": "0CA30C", "Trả thiếu": "EC835A", "Chưa trả": "D03B3B", "Cần kiểm tra": "FAB219"}
STATUS_TINT = {"Đã trả đủ": "E3F4E3", "Trả thiếu": "FCEAE2", "Chưa trả": "F8E1E1", "Cần kiểm tra": "FEF3DB"}
PROOF_KINDS = ("Đã ghép", "Gợi ý", "Chưa ghép", "Không đọc được", "Trùng")
TEMPLATE = Path(__file__).with_name("dashboard.html")


def _hour(text) -> int | None:
    match = re.search(r"(?<!\d)([01]?\d|2[0-3]):([0-5]\d)", str(text or ""))
    return int(match.group(1)) if match else None


def _day(text) -> date | None:
    for d, m, y in re.findall(r"(?<!\d)(\d{1,2})/(\d{1,2})/(\d{4})", str(text or "")):
        try:
            return date(int(y), int(m), int(d))
        except ValueError:
            continue
    return None


def build_report(orders: list[dict], proofs: list[dict], *, large_order: float, source_images: str,
                 source_orders: str, invoices: list[dict] | None = None) -> dict:
    """orders: rows already carrying status/action/received/missing. proofs: rows carrying kind."""
    days = Counter(d for d in (_day(p.get("transferred_at")) for p in proofs) if d)
    report_day = days.most_common(1)[0][0] if days else date.today()

    totals = {
        "due": sum(o["due"] or 0 for o in orders),
        "received": sum(o["received"] for o in orders),
        "outstanding": sum(o["missing"] or 0 for o in orders),
        "surplus": sum(max(o["received"] - (o["due"] or 0), 0) for o in orders if o["due"] is not None),
        "surplus_count": sum(1 for o in orders if o["due"] is not None and o["received"] > o["due"] + 1),
    }
    by_status = {s: {"count": 0, "due": 0.0, "received": 0.0, "missing": 0.0} for s in STATUS_ORDER}
    for o in orders:
        bucket = by_status[o["status"]]
        bucket["count"] += 1
        bucket["due"] += o["due"] or 0
        bucket["received"] += o["received"]
        bucket["missing"] += o["missing"] or 0

    kinds = Counter(p["kind"] for p in proofs)
    usable = len(proofs) - kinds["Trùng"]
    hourly = defaultdict(lambda: {"amount": 0.0, "count": 0})
    for p in proofs:
        if p["kind"] == "Đã ghép" and p.get("amount") and _hour(p.get("transferred_at")) is not None:
            slot = hourly[_hour(p.get("transferred_at"))]
            slot["amount"] += p["amount"]
            slot["count"] += 1

    no_hour = [p for p in proofs if p["kind"] == "Đã ghép" and p.get("amount") and _hour(p.get("transferred_at")) is None]

    channels = defaultdict(lambda: {"due": 0.0, "received": 0.0, "count": 0})
    for o in orders:
        if o.get("channel"):
            c = channels[str(o["channel"])]
            c["due"] += o["due"] or 0
            c["received"] += min(o["received"], o["due"] or 0)
            c["count"] += 1

    return {
        "title": "Đối soát chuyển khoản",
        "date": report_day.strftime("%d/%m/%Y"),
        "generated": datetime.now().strftime("%H:%M %d/%m/%Y"),
        "sources": {"images": source_images, "orders": source_orders},
        "large_order": large_order,
        "totals": totals,
        "status": [{"name": s, **by_status[s]} for s in STATUS_ORDER],
        "ready": by_status["Đã trả đủ"]["count"],
        "todo": len(orders) - by_status["Đã trả đủ"]["count"],
        "proofs": {
            "total": len(proofs),
            "kinds": {k: kinds[k] for k in PROOF_KINDS},
            "auto_rate": kinds["Đã ghép"] / usable if usable else 0,
            "duplicate_amount": sum(p.get("amount") or 0 for p in proofs if p["kind"] == "Trùng"),
            "unassigned_amount": sum(p.get("amount") or 0 for p in proofs if p["kind"] in ("Gợi ý", "Chưa ghép")),
        },
        "hourly": [{"hour": h, **hourly[h]} for h in sorted(hourly)],
        "hourly_missing": {"count": len(no_hour), "amount": sum(p["amount"] for p in no_hour)},
        "channels": [{"name": k, **v} for k, v in sorted(channels.items(), key=lambda kv: -kv[1]["due"])],
        "large": [o for o in orders if o["status"] == "Đã trả đủ" and (o["due"] or 0) >= large_order],
        "orders": orders,
        "images": [{k: p.get(k) for k in ("source_file", "amount", "transferred_at", "matched_order", "result", "kind")} for p in proofs],
        "invoices": invoices,
    }


def read_invoice_reconciliation(path: Path) -> list[dict]:
    """Read the 'Đối chiếu' sheet written by the doi-chieu command."""
    from openpyxl import load_workbook
    ws = load_workbook(path, data_only=True)["Đối chiếu"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    items = []
    for row in rows[1:]:
        status = str(row[-1] or "")
        kind = ("Khớp" if status == "Khớp" else "Trùng" if status.startswith("Khóa trùng")
                else "Một bên" if status.startswith("Chỉ có") else "Lệch")
        # doi-chieu ghi "bảng trái" = tệp hóa đơn đã đọc, "bảng phải" = bảng kê phần mềm; đổi sang lời dễ hiểu.
        readable = (status.replace("Khóa trùng, không ghép duy nhất", "Tệp hóa đơn bị tải trùng")
                    .replace("Thiếu dữ liệu: ", "Thiếu "))
        # Hạ chữ đầu của tên cột ("Tổng tiền" → "tổng tiền") nhưng giữ nguyên viết tắt như "MST".
        readable = re.sub(r"Sai khác: (\w)(\w)", lambda m: "Lệch " + (m.group(1).lower() if m.group(2).islower() else m.group(1)) + m.group(2), readable)
        readable = re.sub(r"Chỉ có ở bảng trái \([^)]*\)", "Có tệp hóa đơn, phần mềm chưa ghi nhận", readable)
        readable = re.sub(r"Chỉ có ở bảng phải \([^)]*\)", "Có trong phần mềm, thiếu tệp hóa đơn", readable)
        items.append({"number": row[0], "status": readable, "kind": kind})
    return items


# ---------------- Excel ----------------

def excel_dashboard(ws, report: dict) -> None:
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.series import DataPoint
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    ink, muted, line = "1B1F24", "6B7280", "E5E7EB"
    ws.sheet_view.showGridLines = False
    for col, width in zip("ABCDEFGHIJKLM", (2, 15, 15, 15, 2, 15, 15, 15, 2, 15, 15, 15, 2)):
        ws.column_dimensions[col].width = width

    def fill(hex_):
        return PatternFill("solid", fgColor="FF" + hex_)

    def put(cell, value, size=10, bold=False, color=ink, align="left", fmt=None):
        c = ws[cell]
        c.value = value
        c.font = Font(name="Arial", size=size, bold=bold, color="FF" + color)
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=False)
        if fmt:
            c.number_format = fmt
        return c

    # Dải tiêu đề
    for row in (1, 2, 3):
        for col in "ABCDEFGHIJKLM":
            ws[f"{col}{row}"].fill = fill("111827")
    ws.row_dimensions[2].height = 30
    put("B2", f"{report['title'].upper()} · {report['date']}", 18, True, "FFFFFF")
    put("B3", f"{len(report['orders'])} đơn · {report['proofs']['total']} ảnh · tạo lúc {report['generated']}", 10, False, "9CA3AF")

    # Bốn ô số liệu
    t = report["totals"]
    pct = t["received"] / t["due"] if t["due"] else 0
    tiles = [
        ("B", "D", "ĐÃ THU THEO ẢNH", t["received"], "#,##0\"đ\"", f"{pct:.0%} của {t['due']:,.0f}đ phải thu".replace(",", "."), "2A78D6"),
        ("F", "H", "CÒN PHẢI THU", t["outstanding"], "#,##0\"đ\"", f"{report['todo']} đơn chưa xong", "D03B3B"),
        ("J", "L", "ĐƠN SẴN SÀNG GIAO", report["ready"], "0", f"trên {len(report['orders'])} đơn", "0CA30C"),
    ]
    for first, last, label, value, fmt, sub, accent in tiles:
        for r in range(5, 9):
            for col in range(ord(first), ord(last) + 1):
                ws[f"{chr(col)}{r}"].fill = fill("F9FAFB")
        for col in range(ord(first), ord(last) + 1):
            ws[f"{chr(col)}5"].border = Border(top=Side(style="thick", color="FF" + accent))
        put(f"{first}5", label, 9, True, muted)
        ws.merge_cells(f"{first}6:{last}7")
        put(f"{first}6", value, 24, True, ink, fmt=fmt)
        put(f"{first}8", sub, 9, False, muted)
    ws.row_dimensions[6].height = 22
    ws.row_dimensions[7].height = 22

    # Bảng trạng thái + biểu đồ
    r0 = 10
    put(f"B{r0}", "THEO TRẠNG THÁI", 11, True)
    headers = ("Trạng thái", "Số đơn", "Giá trị đơn", "Đã nhận", "Còn thiếu")
    for i, h in enumerate(headers):
        c = put(f"{'BCDFG'[i]}{r0 + 1}", h, 9, True, muted, "left" if i == 0 else "right")
        c.border = Border(bottom=Side(style="thin", color="FF" + line))
    for i, s in enumerate(report["status"], start=r0 + 2):
        put(f"B{i}", s["name"], 10, True).fill = fill(STATUS_TINT[s["name"]])
        put(f"C{i}", s["count"], 10, False, ink, "right")
        put(f"D{i}", s["due"], 10, False, ink, "right", "#,##0")
        put(f"F{i}", s["received"], 10, False, ink, "right", "#,##0")
        put(f"G{i}", s["missing"], 10, False, ink, "right", "#,##0")
    chart = BarChart()
    chart.type = "bar"
    chart.style = 2
    chart.title = "Giá trị đơn theo trạng thái"
    chart.legend = None
    chart.y_axis.numFmt = "#,##0"
    chart.y_axis.majorGridlines = None
    data = Reference(ws, min_col=4, min_row=r0 + 1, max_row=r0 + 5)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=2, min_row=r0 + 2, max_row=r0 + 5))
    series = chart.series[0]
    for idx, s in enumerate(report["status"]):
        point = DataPoint(idx=idx)
        point.graphicalProperties.solidFill = STATUS_HEX[s["name"]]
        point.graphicalProperties.line.noFill = True
        series.dPt.append(point)
    chart.gapWidth = 60
    chart.height, chart.width = 6.2, 12.5
    ws.add_chart(chart, f"I{r0}")

    # Tiền về theo giờ
    r1 = r0 + 8
    put(f"B{r1}", "TIỀN VỀ THEO GIỜ", 11, True)
    put(f"B{r1 + 1}", "Giờ", 9, True, muted).border = Border(bottom=Side(style="thin", color="FF" + line))
    put(f"C{r1 + 1}", "Số tiền", 9, True, muted, "right").border = Border(bottom=Side(style="thin", color="FF" + line))
    put(f"D{r1 + 1}", "Giao dịch", 9, True, muted, "right").border = Border(bottom=Side(style="thin", color="FF" + line))
    hours = report["hourly"]
    for i, h in enumerate(hours, start=r1 + 2):
        put(f"B{i}", f"{h['hour']:02d}:00")
        put(f"C{i}", h["amount"], 10, False, ink, "right", "#,##0")
        put(f"D{i}", h["count"], 10, False, ink, "right")
    if hours:
        col_chart = BarChart()
        col_chart.type = "col"
        col_chart.style = 2
        col_chart.title = "Tiền về theo giờ (theo ảnh đã ghép)"
        col_chart.legend = None
        col_chart.y_axis.numFmt = "#,##0"
        col_chart.add_data(Reference(ws, min_col=3, min_row=r1 + 1, max_row=r1 + 1 + len(hours)), titles_from_data=True)
        col_chart.set_categories(Reference(ws, min_col=2, min_row=r1 + 2, max_row=r1 + 1 + len(hours)))
        col_chart.series[0].graphicalProperties.solidFill = "2A78D6"
        col_chart.series[0].graphicalProperties.line.noFill = True
        col_chart.gapWidth = 80
        col_chart.height, col_chart.width = 6.2, 12.5
        ws.add_chart(col_chart, f"F{r1}")

    # Việc cần làm
    r2 = max(r1 + 3 + len(hours), r1 + 14)
    put(f"B{r2}", "VIỆC CẦN LÀM", 11, True)
    todo_headers = ("Trạng thái", "Mã đơn", "Khách hàng", "SĐT", "Còn thiếu", "Việc cần làm")
    for col, h in zip("BCDFGH", todo_headers):
        put(f"{col}{r2 + 1}", h, 9, True, muted, "right" if h == "Còn thiếu" else "left").border = Border(
            bottom=Side(style="thin", color="FF" + line))
    row = r2 + 2
    for status in ("Trả thiếu", "Chưa trả", "Cần kiểm tra"):
        for o in (o for o in report["orders"] if o["status"] == status):
            put(f"B{row}", status, 10, True).fill = fill(STATUS_TINT[status])
            put(f"C{row}", o["code"])
            put(f"D{row}", o["customer"])
            put(f"F{row}", o["phone"])
            put(f"G{row}", o["missing"], 10, False, ink, "right", "#,##0")
            put(f"H{row}", o["action"], 10, False, muted)
            row += 1
    if row == r2 + 2:
        put(f"B{row}", "Không có việc tồn. Mọi đơn đã đủ theo ảnh.", 10, False, muted)
        row += 1

    p = report["proofs"]
    put(f"B{row + 1}", (f"Ảnh: {p['total']} · ghép tự động {p['kinds']['Đã ghép']} ({p['auto_rate']:.0%}) · "
                        f"trùng {p['kinds']['Trùng']} · cần xem {p['total'] - p['kinds']['Đã ghép'] - p['kinds']['Trùng']}"), 9, False, muted)
    note_row = row + 2
    if report.get("invoices"):
        kinds = Counter(i["kind"] for i in report["invoices"])
        put(f"B{note_row}", (f"Hóa đơn điện tử: {len(report['invoices'])} số · khớp {kinds['Khớp']} · lệch {kinds['Lệch']} · "
                             f"tải trùng {kinds['Trùng']} · chỉ có ở một bên {kinds['Một bên']}"), 9, False, muted)
        note_row += 1
    put(f"B{note_row}", ("Số tiền lấy từ ảnh khách gửi, chưa phải xác nhận tiền đã vào tài khoản. "
                         "Xem sao kê trước khi giao đơn lớn. Agent không gửi tin nhắn."), 9, False, muted)
    ws.sheet_properties.tabColor = "2A78D6"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# ---------------- HTML ----------------

def html_dashboard(path: Path, report: dict) -> Path:
    payload = json.dumps(report, ensure_ascii=False, default=str).replace("</", "<\\/")
    html = TEMPLATE.read_text(encoding="utf-8").replace("/*__REPORT__*/null", payload)
    path.write_text(html, encoding="utf-8")
    return path
