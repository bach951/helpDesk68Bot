# -*- coding: utf-8 -*-
"""Build a COMPACT knowledge table (essential columns only) to fit env var size limit."""
import openpyxl

wb = openpyxl.load_workbook("contacts.xlsx")
ws = wb.active
rows = list(ws.iter_rows(values_only=True))[1:]  # skip header

# Compact: one line per row, pipe-separated, only essential fields
# Tu_khoa | Phong_ban | Nguoi_dau_moi | Chuc_vu | Email | Teams | Dien_thoai
lines = []
for r in rows:
    stt, nhom, tukhoa, mota, phong, nguoi, chucvu, email, teams, dt = r
    lines.append(f"{tukhoa} | {phong} | {nguoi} | {chucvu} | {email} | {teams} | {dt}")

table = "\n".join(lines)

prompt = (
    "Ban la bot IT helpdesk noi bo. Nhan vien hoi bang tieng Viet -> "
    "chon dong co 'tu khoa' khop nhat (tu dong nghia/viet tat/go sai cung OK). "
    "Tra loi dung dinh dang:\n"
    "Phong ban: <phong>\nNguoi dau moi: <ten> - <chuc vu>\nKenh: <email> | <Teams> | DT: <so>\n"
    "Khong khop -> 'Chua xac dinh, lien he P.CNTT - Teams: IT Support'.\n\n"
    "BANG (tu_khoa | phong | nguoi | chuc_vu | email | teams | dt):\n" + table
)

with open("knowledge_compact.txt", "w", encoding="utf-8") as f:
    f.write(prompt)
print(f"knowledge_compact.txt: {len(prompt)} chars, {len(lines)} rows")
