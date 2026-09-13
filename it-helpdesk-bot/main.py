import os
import logging
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

SYSTEM_INSTRUCTIONS = (
    "Ban la bot ho tro IT noi bo cho nhan vien ngan hang. "
    "Khi nhan vien hoi bang tieng Viet, hay doi chieu cau hoi voi BANG DAU MOI duoi day "
    "(chon dong co tu khoa/mo ta khop nhat ve y nghia, ke ca khi dung tu dong nghia, "
    "viet tat, hoac go sai). Tra loi ngan gon bang tieng Viet dung dinh dang:\n"
    "Phong ban: <ten phong>\n"
    "Nguoi dau moi: <ten> - <chuc vu>\n"
    "Kenh lien he: <email> | <Teams> | DT: <so>\n"
    "Neu khong khop dong nao, tra: "
    "'Chua xac dinh duoc dau moi, vui long lien he P.Cong nghe Thong tin - Teams: IT Support'.\n\n"
    "BANG DAU MOI:\n"
)

app = FastAPI(title="IT Helpdesk Bot")
_cache = {"mtime": 0.0, "table": ""}


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


@app.get("/health")
def health():
    return {"status": "healthy"}


@app.post("/telegram/webhook")
async def telegram_webhook(request: Request):
    data = await request.json()
    msg = data.get("message") or data.get("edited_message") or {}
    text = (msg.get("text") or "").strip()
    chat_id = (msg.get("chat") or {}).get("id")
    if not text or chat_id is None:
        return JSONResponse({"ok": True})
    log.info("Incoming: chat_id=%s text=%s", chat_id, text[:120])
    try:
        answer = await ask_llm(text)
    except Exception as exc:
        log.exception("LLM error")
        answer = f"Xu ly loi, vui long thu lai sau. ({type(exc).__name__})"
    await telegram_send(chat_id, answer)
    log.info("Replied: chat_id=%s len=%d", chat_id, len(answer))
    return JSONResponse({"ok": True})


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
