# -*- coding: utf-8 -*-
"""Read contacts.xlsx -> produce knowledge.txt (markdown table) for the bot prompt."""
import openpyxl, io

wb = openpyxl.load_workbook("contacts.xlsx")
ws = wb.active
rows = list(ws.iter_rows(values_only=True))
header = [str(h) for h in rows[0]]
data = rows[1:]

lines = []
lines.append("| " + " | ".join(header) + " |")
lines.append("| " + " | ".join(["---"]*len(header)) + " |")
for r in data:
    lines.append("| " + " | ".join("" if c is None else str(c) for c in r) + " |")
table = "\n".join(lines)

prompt = (
    "Bạn là bot hỗ trợ IT nội bộ cho nhân viên ngân hàng. "
    "Khi nhân viên hỏi bằng tiếng Việt, hãy đối chiếu câu hỏi với BẢNG ĐẦU MỐI dưới đây "
    "(chọn dòng có từ khoá/mô tả khớp nhất về ý nghĩa, kể cả khi nhân viên dùng từ đồng nghĩa, "
    "viết tắt, hoặc gõ sai). Trả lời ngắn gọn bằng tiếng Việt theo đúng định dạng:\n"
    "Phòng ban: <tên phòng>\nNgười đầu mối: <tên> — <chức vụ>\nKênh liên hệ: <email> | <Teams> | ĐT: <số>\n"
    "Nếu không khớp dòng nào, trả: 'Chưa xác định được đầu mối, vui lòng liên hệ P.Công nghệ Thông tin — Teams: IT Support'.\n\n"
    "BẢNG ĐẦU MỐI:\n" + table
)

with open("knowledge.txt", "w", encoding="utf-8") as f:
    f.write(prompt)

print(f"knowledge.txt written: {len(prompt)} chars, {len(data)} rows.")
print("--- first 600 chars ---")
print(prompt[:600])
