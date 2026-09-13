# IT Helpdesk Bot — Telegram bot trên GreenNode AgentBase

Bot Telegram trả lời câu hỏi IT nội bộ bằng tiếng Việt. Nhân viên hỏi
"máy tính lỗi liên hệ ai", "SWIFT message bị NAK thì hỏi ai"… → bot trả
**Phòng ban / Người đầu mối / Kênh liên hệ (email – Teams – ĐT)**.

## Kiến trúc

```
Telegram user ──webhook──> Custom Agent (FastAPI, port 8080)
                                │
                                ├─ đọc contacts.xlsx (100 dòng) ← file Excel đầu mối
                                ├─ gọi LLM GLM-5.3 Flash (GreenNode AIP, OpenAI-compatible)
                                └─ trả lời qua Telegram Bot API
```

Code đọc `contacts.xlsx` **trực tiếp** lúc runtime — không bơm knowledge vào env var.
Cập nhật Excel = sửa file + rebuild + redeploy.

## Repo layout

```
it-helpdesk-bot/          # code bot (FastAPI)
  main.py                 # server: /health + /telegram/webhook
  requirements.txt
  Dockerfile
  contacts.xlsx           # 100 dòng đầu mối (sample)
  .env.example            # template env vars (copy → .env, điền secret)
contacts.xlsx             # file Excel nguồn (copy vào it-helpdesk-bot/ khi build)
generate_contacts.py      # script sinh contacts.xlsx (sửa data → rerun)
agent/                    # greennode-agentbase-skills (deploy/monitor scripts)
.greennode.json.example   # template IAM credentials
```

## Prerequisites

1. **Docker** (Docker Desktop) — để build image.
2. **GreenNode account** + IAM service account (`AgentBaseFullAccess`, `AiPlatformFullAccess`).
3. **Telegram bot token** — tạo qua @BotFather (`/newbot`).
4. **GreenNode AIP API key** — tạo tại https://aiplatform.console.vngcloud.vn (OpenAI-compatible).

## Setup (sau khi clone)

```bash
git clone https://github.com/bach951/helpDesk68Bot.git
cd helpDesk68Bot

# 1. GreenNode IAM credentials (cho skill scripts: deploy, monitor)
cp .greennode.json.example .greennode.json
#   → sửa .greennode.json: điền client_id + client_secret

# 2. Bot env vars
cd it-helpdesk-bot
cp .env.example .env
#   → sửa .env: điền TELEGRAM_BOT_TOKEN, LLM_API_KEY
#     (LLM_BASE_URL + LLM_MODEL đã có sẵn giá trị default)
```

### `.env` cần điền

| Variable | Giá trị | Lấy từ đâu |
|----------|---------|------------|
| `TELEGRAM_BOT_TOKEN` | `123456:ABC...` | @BotFather |
| `LLM_BASE_URL` | `https://maas-llm-aiplatform-hcm.api.vngcloud.vn/v1` | GreenNode AIP |
| `LLM_API_KEY` | (96 ký tự) | GreenNode AIP console → API keys |
| `LLM_MODEL` | `z-ai/glm-5.3-flash-thirdparty` | GreenNode AIP → models |

## Chạy local (test)

```bash
cd it-helpdesk-bot
pip install -r requirements.txt
python main.py                    # http://localhost:8080
# test health:
curl http://localhost:8080/health # → {"status":"healthy"}
```

## Deploy lên GreenNode AgentBase

Dùng skill `/agentbase-deploy` (scripts trong `agent/`):

```bash
# 1. Docker login AgentBase Container Registry
bash agent/agentbase/scripts/cr.sh credentials docker-login

# 2. Build + push image
docker build --platform linux/amd64 -t vcr.vngcloud.vn/<repo>/it-helpdesk-bot:v1 .
docker push vcr.vngcloud.vn/<repo>/it-helpdesk-bot:v1

# 3. Create runtime
bash agent/agentbase/scripts/runtime.sh create \
  --name it-helpdesk-bot \
  --image "vcr.vngcloud.vn/<repo>/it-helpdesk-bot:v1" \
  --flavor runtime-s2-general-2x4 \
  --env-file .env --from-cr --poc false

# 4. Set Telegram webhook → <endpoint-url>/telegram/webhook
curl -X POST "https://api.telegram.org/bot<TOKEN>/setWebhook" \
  -H "Content-Type: application/json" \
  -d '{"url":"<ENDPOINT_URL>/telegram/webhook"}'
```

## Cập nhật Excel đầu mối

```bash
# Sửa data trong generate_contacts.py → rerun:
python generate_contacts.py              # sinh lại contacts.xlsx
cp contacts.xlsx it-helpdesk-bot/contacts.xlsx
# Rebuild + redeploy (xem Deploy)
```

## Lưu ý bảo mật

**Không bao giờ commit** các file secret (đã có trong `.gitignore`):
- `.greennode.json` — IAM credentials
- `.env` — Telegram token + LLM API key
- `telegram-channel.json` — bot token

Chỉ commit `.env.example` + `.greennode.json.example` (template với placeholder).
