# -*- coding: utf-8 -*-
"""Build .env with SYSTEM_PROMPT=<flattened knowledge> on a single line."""
with open("knowledge.txt", "r", encoding="utf-8") as f:
    knowledge = f.read()
flat = knowledge.replace("\r", "").replace("\n", "\\n")
with open(".env", "w", encoding="utf-8") as f:
    f.write("SYSTEM_PROMPT=" + flat + "\n")
print(f".env written: SYSTEM_PROMPT length={len(flat)} chars (single line)")
