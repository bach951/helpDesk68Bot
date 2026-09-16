import os
import logging
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
import httpx
import openpyxl

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("it-helpdesk-bot")

TELEGRAM_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
LLM_BASE_URL = os.environ["LLM_BASE_URL"].rstrip("/")
LLM_API_KEY = os.environ["LLM_API_KEY"]
LLM_MODEL = os.environ["LLM_MODEL"]
CONTACTS_PATH = os.environ.get("CONTACTS_PATH", "contacts.xlsx")
TG_API = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}"
GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")

SYSTEM_INSTRUCTIONS = (
    "Ban la bot ho tro IT noi bo cho nhan vien ngan hang. "
    "Khi nhan vien hoi bang tieng Viet, hay doi chieu cau hoi voi BANG DAU MOI duoi day "
    "(chon dong co tu khoa/mo ta khop nhat ve y nghia, ke ca khi dung tu dong nghia, "
    "viet tat, hoac go sai).\n\n"
    "QUYET DINH 2 loai:\n"
    "1. Mac dinh (tra cuu/hoi/bao loi/yeu cau): tra loi ngan gon. "
    "LUON tra ve TAT CA phong ban co lien quan, khong chi 1. "
    "Vi du: 'SWIFT bi NAK' -> ca P.Thanh toan Quoc te VA P.He thong Core Banking. "
    "Moi phong 1 block dung dinh dang:\n"
    "   Phong ban: <ten phong>\n"
    "   Nguoi dau moi: <ten> - <chuc vu>\n"
    "   Kenh lien he: <email> | <Teams> | DT: <so>\n"
    "   Cuoi moi block them 1 dong: [CONTACT:email-phong-ban|ten-phong-ban]\n"
    "2. Neu user hoi ve ONBOARDING / NHAN VIEN MOI / SETUP / 'toi moi vao' / 'can setup gi' "
    "-> tra loi day du huong dan tung buoc, moi buoc ghi ro phong ban + email + Teams "
    "(lay tu BANG DAU MOI). Dinh dang:\n"
    "   Buoc 1: Tai khoan AD + Email Outlook -> lien he <phong> (<email> | <Teams>)\n"
    "   Buoc 2: Cai dat VPN -> lien he <phong> (<email> | <Teams>)\n"
    "   Buoc 3: Cap phat may tinh/laptop -> lien he <phong> (<email> | <Teams>)\n"
    "   Buoc 4: Tai khoan Core Banking -> lien he <phong> (<email> | <Teams>)\n"
    "   Buoc 5: Cai Microsoft Teams + join channel phong ban\n"
    "   Buoc 6: Map may in mang (lien he IT)\n"
    "   Buoc 7: Token/OTP giao dich -> lien he <phong> (<email> | <Teams>)\n"
    "   Buoc 8: He thong cham cong -> lien he <phong> (<email> | <Teams>)\n"
    "   Buoc 9: Phan mem chuyen dung (ERP/CRM) -> lien he <phong> (<email> | <Teams>)\n"
    "   Buoc 10: Tai khoan intranet/portal -> lien he <phong> (<email> | <Teams>)\n"
    "   Sau khi hoan thanh, lien he P.Nhan su de bao danh sach ho so.\n\n"
    "Neu khong khop dong nao, tra: "
    "'Chua xac dinh duoc dau moi, vui long lien he P.Ho tro Cong nghe - Teams: IT Support'.\n\n"
    "BANG DAU MOI:\n"
)

app = FastAPI(title="IT Helpdesk Bot")
_cache = {"mtime": 0.0, "table": ""}
_states = {}


def load_contacts_table() -> str:
    mtime = os.path.getmtime(CONTACTS_PATH)
    if mtime != _cache["mtime"]:
        wb = openpyxl.load_workbook(CONTACTS_PATH, read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        header = [str(h) for h in rows[0]]
        lines = ["| " + " | ".join(header) + " |"]
        lines.append("| " + " | ".join(["---"] * len(header)) + " |")
        for r in rows[1:]:
            lines.append("| " + " | ".join("" if c is None else str(c) for c in r) + " |")
        _cache["mtime"] = mtime
        _cache["table"] = "\n".join(lines)
        log.info("Reloaded contacts: %d rows", len(rows) - 1)
    return _cache["table"]


async def ask_llm(question: str) -> str:
    system_prompt = SYSTEM_INSTRUCTIONS + load_contacts_table()
    payload = {
        "model": LLM_MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": question},
        ],
        "temperature": 0,
    }
    async with httpx.AsyncClient(timeout=60) as client:
        resp = await client.post(
            f"{LLM_BASE_URL}/chat/completions",
            headers={"Authorization": f"Bearer {LLM_API_KEY}"},
            json=payload,
        )
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]


async def telegram_send(chat_id: int, text: str) -> None:
    async with httpx.AsyncClient(timeout=30) as client:
        await client.post(
            f"{TG_API}/sendMessage",
            json={"chat_id": chat_id, "text": text, "parse_mode": "HTML"},
        )


async def download_telegram_photo(file_id: str) -> bytes:
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.get(f"{TG_API}/getFile", params={"file_id": file_id})
        file_path = resp.json()["result"]["file_path"]
        resp2 = await client.get(
            f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{file_path}"
        )
        return resp2.content


def send_email(to_addr: str, from_email: str, subject: str, body: str,
               attachment: bytes | None = None) -> None:
    full_body = f"Email nay duoc gui thay mat cho {from_email} boi IT Helpdesk Bot.\n\n{body}"
    if attachment:
        msg = MIMEMultipart()
        msg["Subject"] = subject
        msg["From"] = GMAIL_USER
        msg["To"] = to_addr
        msg["Reply-To"] = from_email
        msg.attach(MIMEText(full_body, "plain", "utf-8"))
        img = MIMEImage(attachment)
        img.add_header("Content-Disposition", "attachment", filename="screenshot.png")
        msg.attach(img)
    else:
        msg = MIMEText(full_body, "plain", "utf-8")
        msg["Subject"] = subject
        msg["From"] = GMAIL_USER
        msg["To"] = to_addr
        msg["Reply-To"] = from_email
    with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as srv:
        srv.starttls()
        srv.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        srv.send_message(msg)
    log.info("Email sent: %s -> %s reply-to=%s subject=%s attach=%s",
             GMAIL_USER, to_addr, from_email, subject, bool(attachment))


@app.get("/health")
def health():
    return {"status": "healthy"}


@app.post("/telegram/webhook")
async def telegram_webhook(request: Request):
    data = await request.json()
    msg = data.get("message") or data.get("edited_message") or {}
    text = (msg.get("text") or msg.get("caption") or "").strip()
    chat_id = (msg.get("chat") or {}).get("id")
    photo_sizes = msg.get("photo")
    if chat_id is None:
        return JSONResponse({"ok": True})
    if not text and not photo_sizes:
        return JSONResponse({"ok": True})
    log.info("Incoming: chat_id=%s text=%s photo=%s", chat_id, text[:120], bool(photo_sizes))

    st = _states.get(chat_id)

    if text.lower() == "/cancel":
        _states.pop(chat_id, None)
        await telegram_send(chat_id, "Đã hủy. Gửi câu hỏi mới để bắt đầu lại.")
        return JSONResponse({"ok": True})

    if st and st["state"] == "AWAITING_EMAIL_DECISION":
        contacts = st.get("contacts", [])
        if len(contacts) == 1 and text.lower() in ("gửi email", "gui email", "email", "co", "có", "yes", "y", "1"):
            st["state"] = "AWAITING_EMAIL"
            st["email"] = contacts[0]["email"]
            st["dept"] = contacts[0]["dept"]
            await telegram_send(
                chat_id,
                f"Nập email của bạn (để {st['dept'] or 'IT'} reply lại cho bạn) hoặc /cancel:",
            )
            return JSONResponse({"ok": True})
        elif text.isdigit() and 1 <= int(text) <= len(contacts):
            idx = int(text) - 1
            st["state"] = "AWAITING_EMAIL"
            st["email"] = contacts[idx]["email"]
            st["dept"] = contacts[idx]["dept"]
            await telegram_send(
                chat_id,
                f"Đã chọn: <b>{contacts[idx]['dept']}</b> ({contacts[idx]['email']})\n"
                f"Nập email của bạn (để {st['dept']} reply lại cho bạn) hoặc /cancel:",
            )
            return JSONResponse({"ok": True})
        else:
            _states.pop(chat_id, None)
            st = None

    if st and st["state"] == "AWAITING_EMAIL":
        st["state"] = "AWAITING_TITLE"
        st["from_email"] = text
        await telegram_send(
            chat_id,
            f"Email của bạn: <b>{text}</b>\n"
            f"Nập tiêu đề email (hoặc /cancel để hủy):",
        )
        return JSONResponse({"ok": True})

    if st and st["state"] == "AWAITING_TITLE":
        st["state"] = "AWAITING_CONTENT"
        st["title"] = text
        await telegram_send(
            chat_id,
            f"Tiêu đề: <b>{text}</b>\nNhập nội dung email hoặc gửi ảnh chụp màn hình lỗi (hoặc /cancel):",
        )
        return JSONResponse({"ok": True})

    if st and st["state"] == "AWAITING_CONTENT":
        title = st["title"]
        to_addr = st["email"]
        from_email = st["from_email"]
        dept = st.get("dept", "")
        _states.pop(chat_id, None)
        attachment = None
        if photo_sizes:
            file_id = photo_sizes[-1]["file_id"]
            try:
                attachment = await download_telegram_photo(file_id)
            except Exception as exc:
                log.exception("Photo download error")
                await telegram_send(chat_id, f"Tải ảnh thất bại: {type(exc).__name__}")
                return JSONResponse({"ok": True})
        if not text:
            text = "(Xem ảnh đính kèm)"
        try:
            send_email(to_addr, from_email, title, text, attachment)
            attach_note = " + ảnh đính kèm 📸" if attachment else ""
            await telegram_send(
                chat_id,
                f"Đã gửi email ✅{attach_note}\n<b>Tới:</b> {dept} ({to_addr})\n"
                f"<b>Reply-to:</b> {from_email}\n<b>Tiêu đề:</b> {title}",
            )
        except Exception as exc:
            log.exception("Email send error")
            await telegram_send(chat_id, f"Gửi email thất bại: {type(exc).__name__}: {exc}")
        return JSONResponse({"ok": True})

    if photo_sizes and not text:
        await telegram_send(
            chat_id,
            "📸 Bạn gửi ảnh nhưng chưa mô tả vấn đề. Hãy gõ mô tả vấn đề trước, "
            "sau đó gửi ảnh khi bot hỏi nội dung email.",
        )
        return JSONResponse({"ok": True})

    try:
        answer = await ask_llm(text)
    except Exception as exc:
        log.exception("LLM error")
        answer = f"Xử lý lỗi, vui lòng thử lại sau. ({type(exc).__name__})"

    import re as _re
    contacts = []
    for m in _re.finditer(r"\[CONTACT:([^\]]+)\]", answer):
        payload = m.group(1).strip()
        if "|" in payload:
            email, dept = payload.split("|", 1)
            contacts.append({"email": email.strip(), "dept": dept.strip()})
        else:
            contacts.append({"email": payload, "dept": ""})

    if contacts:
        display = _re.sub(r"\[CONTACT:[^\]]+\]", "", answer).strip()
        if len(contacts) == 1:
            _states[chat_id] = {"state": "AWAITING_EMAIL_DECISION", "contacts": contacts}
            await telegram_send(
                chat_id,
                f"{display}\n\nBạn có muốn gửi email tới <b>{contacts[0]['dept'] or contacts[0]['email']}</b>? "
                f"Gõ \"gửi email\" hoặc bỏ qua.",
            )
        else:
            _states[chat_id] = {"state": "AWAITING_EMAIL_DECISION", "contacts": contacts}
            options = "\n".join(
                f"  {i+1}. {c['dept']} ({c['email']})" for i, c in enumerate(contacts)
            )
            await telegram_send(
                chat_id,
                f"{display}\n\nGửi email tới phòng nào? Gõ số (1-{len(contacts)}) hoặc bỏ qua:\n{options}",
            )
    else:
        await telegram_send(chat_id, answer)

    log.info("Replied: chat_id=%s len=%d", chat_id, len(answer))
    return JSONResponse({"ok": True})


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
