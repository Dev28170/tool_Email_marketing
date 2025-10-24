# office_mailer.py
"""
Async Office365 mailer (MVP)
- Uses Azure AD (client credentials/app-only OR delegated token you supply)
- Sends emails via Microsoft Graph
- Supports attachments (small direct, >4MB upload session stub)
- Placeholder replacement hook that can call an LLM (user-provided OpenAI key)
- Concurrency with per-tenant semaphore and retry/backoff for 429
"""

import os
import asyncio
import base64
import json
import time
from typing import List, Optional, Dict, Any
import aiohttp
import requests
from msal import ConfidentialClientApplication
from dataclasses import dataclass

# --- Configuration from env ---
TENANT_ID = os.getenv("AZURE_TENANT_ID")
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
SCOPE = os.getenv("SCOPE", "https://graph.microsoft.com/.default")
OPENAI_KEY = os.getenv("OPENAI_API_KEY", None)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# --- Simple token manager (client credentials flow) ---
class TokenManager:
    def __init__(self, tenant_id: str, client_id: str, client_secret: str, scope: str = "https://graph.microsoft.com/.default"):
        self.app = ConfidentialClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret
        )
        self.scope = [scope]
        self._cache = {}

    def get_app_token(self) -> str:
        """Get app-only token (client credentials)."""
        token_entry = self.app.acquire_token_silent(self.scope, account=None)
        if not token_entry:
            token_entry = self.app.acquire_token_for_client(scopes=self.scope)
        if "access_token" not in token_entry:
            raise RuntimeError(f"Failed to acquire token: {token_entry}")
        return token_entry["access_token"]

# --- Placeholder replacement (AI optional) ---
async def replace_placeholders(body: str, placeholders: Dict[str, str]) -> str:
    """
    placeholders: mapping from placeholder label without brackets -> prompt or static text.
    Example: {"PROJECT_DETAILS": "Please write a short status update about project X..."}
    If OPENAI_KEY is set, call OpenAI completion to generate text; otherwise use fallback static text.
    """
    async def gen_from_openai(prompt: str) -> str:
        # Minimal OpenAI call using aiohttp (text-davinci style is illustrative; adapt to whichever model & API version you use)
        if not OPENAI_KEY:
            return "(AI_KEY_NOT_PROVIDED) " + prompt[:200]
        url = "https://api.openai.com/v1/chat/completions"
        headers = {"Authorization": f"Bearer {OPENAI_KEY}", "Content-Type": "application/json"}
        payload = {
            "model": "gpt-4o-mini",  # change if unavailable
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": 200,
        }
        async with aiohttp.ClientSession() as session:
            async with session.post(url, headers=headers, json=payload, timeout=30) as resp:
                if resp.status != 200:
                    text = await resp.text()
                    return f"(OPENAI_ERROR {resp.status}) {text[:200]}"
                data = await resp.json()
                # conservative path: extract text safely
                try:
                    return data["choices"][0]["message"]["content"]
                except Exception:
                    return "(OPENAI_PARSE_ERROR)"
    # Replace all placeholders like [PROJECT_DETAILS]
    result = body
    for key, prompt_or_text in placeholders.items():
        token = f"[{key}]"
        if token in result:
            # decide whether prompt_or_text is prompt or static text
            if prompt_or_text.strip().endswith(".") or len(prompt_or_text.split()) > 5:
                # treat as prompt -> attempt AI
                generated = await gen_from_openai(prompt_or_text)
                result = result.replace(token, generated)
            else:
                result = result.replace(token, prompt_or_text)
    return result

# --- Mail sending primitives ---
@dataclass
class EmailMessage:
    sender_upn: str
    to: List[str]
    subject: str
    body_html: str
    bcc: Optional[List[str]] = None
    attachments: Optional[List[Dict[str, Any]]] = None  # each: {"name":..., "bytes": b'...'}

class GraphClient:
    def __init__(self, token_mgr: TokenManager):
        self.token_mgr = token_mgr

    def _headers(self):
        token = self.token_mgr.get_app_token()
        return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    async def send_mail(self, msg: EmailMessage, save_to_sent: bool = True):
        url = f"{GRAPH_BASE}/users/{msg.sender_upn}/sendMail"
        payload = {
            "message": {
                "subject": msg.subject,
                "body": {"contentType": "HTML", "content": msg.body_html},
                "toRecipients": [{"emailAddress": {"address": r}} for r in msg.to],
            },
            "saveToSentItems": save_to_sent
        }
        if msg.bcc:
            payload["message"]["bccRecipients"] = [{"emailAddress": {"address": r}} for r in msg.bcc]
        if msg.attachments:
            # For small attachments (<4MB) include inline as fileAttachment
            attachments_payload = []
            for att in msg.attachments:
                b64 = base64.b64encode(att["bytes"]).decode("utf-8")
                attachments_payload.append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": att["name"],
                    "contentBytes": b64
                })
            payload["message"]["attachments"] = attachments_payload

        # use aiohttp for non-blocking request
        async with aiohttp.ClientSession() as session:
            headers = self._headers()
            async with session.post(url, headers=headers, json=payload) as resp:
                text = await resp.text()
                if resp.status in (200, 202):
                    return {"ok": True, "status": resp.status}
                # handle throttling
                if resp.status == 429:
                    retry_after = int(resp.headers.get("Retry-After", "5"))
                    raise ThrottleException(retry_after, text)
                raise RuntimeError(f"Send failed {resp.status}: {text}")

    async def upload_large_attachment(self, sender_upn: str, filename: str, file_bytes: bytes) -> str:
        """
        Create upload session and upload file in chunks; return attachment id to include in message.
        (This is illustrative — full chunk logic should resume on error and handle ranges.)
        """
        # Create upload session
        url = f"{GRAPH_BASE}/users/{sender_upn}/messages/createUploadSession"
        payload = {"AttachmentItem": {"attachmentType":"file","name": filename, "size": len(file_bytes)}}
        headers = self._headers()
        # create session (POST)
        r = requests.post(url, headers=headers, json=payload)
        if r.status_code not in (200,201):
            raise RuntimeError(f"Failed create upload session: {r.status_code} {r.text}")
        session_obj = r.json()
        upload_url = session_obj.get("uploadUrl")
        if not upload_url:
            raise RuntimeError("No uploadUrl in session response")
        # simple single-chunk upload if small enough
        # For brevity, we upload in one PUT (works if server accepts)
        put_headers = {"Content-Length": str(len(file_bytes)), "Content-Range": f"bytes 0-{len(file_bytes)-1}/{len(file_bytes)}"}
        put_resp = requests.put(upload_url, headers=put_headers, data=file_bytes)
        if put_resp.status_code in (200,201):
            # upload finished, Graph returns the attachment metadata — user must attach to message before send or send draft
            return put_resp.json().get("id", "")
        else:
            raise RuntimeError(f"Upload failed {put_resp.status_code}: {put_resp.text}")

class ThrottleException(Exception):
    def __init__(self, retry_after: int, message="throttled"):
        super().__init__(message)
        self.retry_after = retry_after

async def worker_send(graph: GraphClient, email: EmailMessage, semaphore: asyncio.Semaphore, max_retries: int = 5):
    backoff = 1.0
    attempt = 0
    async with semaphore:
        while attempt < max_retries:
            try:
                res = await graph.send_mail(email)
                return {"ok": True, "status": res}
            except ThrottleException as e:
                await asyncio.sleep(e.retry_after + 0.5)
            except Exception as e:
                attempt += 1
                await asyncio.sleep(backoff)
                backoff *= 2
        return {"ok": False, "error": f"Failed after {max_retries} retries"}

async def main_test():
    if not (TENANT_ID and CLIENT_ID and CLIENT_SECRET):
        raise RuntimeError("Set AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET in env")

    token_mgr = TokenManager(TENANT_ID, CLIENT_ID, CLIENT_SECRET, scope=SCOPE)
    graph = GraphClient(token_mgr)

    sender = os.getenv("TEST_SENDER_UPN")  # e.g. "alice@yourdomain.com"
    receiver = os.getenv("TEST_RECIPIENT", "recipient@example.com")
    if not sender:
        raise RuntimeError("Set TEST_SENDER_UPN env var")

    body_template = "<p>Hello, here is the update: [PROJECT_DETAILS]</p>"
    placeholders = {"PROJECT_DETAILS": "Write a 2-sentence friendly status update about project Athena: current phase, next milestone."}
    body_filled = await replace_placeholders(body_template, placeholders)

    email = EmailMessage(
        sender_upn=sender,
        to=[receiver],
        subject="Test send from Office365 module",
        body_html=body_filled,
        bcc=None,
        attachments=None
    )

    # concurrency: up to 10 concurrent sends for this tenant
    semaphore = asyncio.Semaphore(10)

    # spawn several sends (simulate multiple accounts by varying sender_upn)
    tasks = [worker_send(graph, email, semaphore) for _ in range(3)]
    results = await asyncio.gather(*tasks)
    print("Results:", results)

if __name__ == "__main__":
    asyncio.run(main_test())
