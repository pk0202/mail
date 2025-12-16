import os
import json
from typing import List, Optional

import msal
import requests
from fastapi import FastAPI, HTTPException, Header
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field
from dotenv import load_dotenv

load_dotenv(override=True)

# ---------- Config ----------

CLIENT_ID = os.getenv("CLIENT_ID1")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID", "common")

# For app-permissions, you generally want .default here
#SCOPES = [os.getenv("SCOPES") or "https://graph.microsoft.com/.default"]

SCOPES = ["https://graph.microsoft.com/.default"]

print("DEBUG CLIENT_ID:", CLIENT_ID)
print("DEBUG TENANT_ID:", TENANT_ID)
print("DEBUG SCOPES:", SCOPES)

TOKEN_CACHE_FILE = os.getenv("TOKEN_CACHE_FILE", "token_cache.bin")
EMAIL_API_KEY = os.getenv("EMAIL_API_KEY")
DEFAULT_SENDER = os.getenv("DEFAULT_SENDER")  # e.g. E_DWS_P_CQD@OPTUM.COM

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH_SENDMAIL_URL_TEMPLATE = "https://graph.microsoft.com/v1.0/users/{user_id_or_upn}/sendMail"

app = FastAPI(title="Graph Email API (App-only / Client Credentials)")


# ---------- Models ----------

class AttachmentIn(BaseModel):
    name: str
    content_base64: str = Field(..., description="Base64-encoded file bytes")
    content_type: Optional[str] = Field(
        default="application/octet-stream",
        description="MIME type, e.g. text/plain, application/pdf"
    )


class SendEmailRequest(BaseModel):
    from_email: Optional[str] = Field(
        default=None,
        description="Sender email address (if omitted, DEFAULT_SENDER env var is used)"
    )
    to: List[str] = Field(..., description="List of TO email addresses")
    cc: Optional[List[str]] = Field(default=None, description="List of CC email addresses")
    subject: str = Field(..., description="Subject (plain text)")
    body_html: str = Field(..., description="Body as HTML")
    importance: Optional[str] = Field(
        default="normal",
        description="Importance: high | normal | low"
    )
    attachments: Optional[List[AttachmentIn]] = Field(
        default=None,
        description="Optional list of attachments"
    )


# ---------- Token cache helpers ----------

def load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    return cache


def save_cache(cache: msal.SerializableTokenCache):
    if cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())


def build_confidential_client_app(cache=None):
    """
    MSAL Confidential Client for client-credentials (app-only) flow.
    """
    if not CLIENT_ID or not CLIENT_SECRET:
        raise Exception("CLIENT_ID or CLIENT_SECRET is not configured.")
    return msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=cache
    )


def get_app_access_token() -> str:
    """
    App-only auth using client credentials flow.
    No user interaction, suitable for backend services.
    """
    cache = load_cache()
    app_msal = build_confidential_client_app(cache)

    # Try cache first
    result = app_msal.acquire_token_silent(SCOPES, account=None)

    if not result:
        # Fall back to client credentials
        result = app_msal.acquire_token_for_client(scopes=SCOPES)

    if "access_token" in result:
        save_cache(cache)
        return result["access_token"]
    else:
        raise Exception("Could not obtain access token: %s" % json.dumps(result, indent=2))


# ---------- Graph sendMail helper ----------

def send_graph_mail(
    access_token: str,
    sender: str,
    to_emails: List[str],
    cc_emails: Optional[List[str]],
    subject: str,
    body_html: str,
    importance: str = "normal",
    attachments: Optional[List[AttachmentIn]] = None,
):
    # Normalize importance
    importance = (importance or "normal").lower()
    if importance not in ["high", "normal", "low"]:
        importance = "normal"

    to_recipients = [
        {"emailAddress": {"address": addr.strip()}}
        for addr in to_emails if addr.strip()
    ]

    cc_recipients = []
    if cc_emails:
        cc_recipients = [
            {"emailAddress": {"address": addr.strip()}}
            for addr in cc_emails if addr.strip()
        ]

    message = {
        "message": {
            "subject": subject,
            "importance": importance,
            "body": {
                "contentType": "HTML",
                "content": body_html
            },
            "toRecipients": to_recipients,
        },
        "saveToSentItems": True,
    }

    if cc_recipients:
        message["message"]["ccRecipients"] = cc_recipients

    if attachments:
        graph_attachments = []
        for att in attachments:
            graph_attachments.append({
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": att.name,
                "contentType": att.content_type or "application/octet-stream",
                "contentBytes": att.content_base64
            })
        message["message"]["attachments"] = graph_attachments

    graph_url = GRAPH_SENDMAIL_URL_TEMPLATE.format(user_id_or_upn=sender)

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    resp = requests.post(graph_url, headers=headers, json=message)
    return resp


# ---------- FastAPI endpoints ----------

@app.get("/")
def root():
    return {
        "status": "ok",
        "message": "Graph Email API (app-only client credentials)",
        "endpoints": ["/send-email"]
    }


@app.post("/send-email")
def send_email_api(
    payload: SendEmailRequest,
    x_api_key: Optional[str] = Header(None)
):
    # Simple shared-secret API key for callers
    if EMAIL_API_KEY and x_api_key != EMAIL_API_KEY:
        raise HTTPException(status_code=401, detail="Unauthorized")

    if not payload.to:
        raise HTTPException(status_code=400, detail="At least one TO recipient is required.")

    sender = payload.from_email or DEFAULT_SENDER
    if not sender:
        raise HTTPException(
            status_code=400,
            detail="Sender email is required. Provide from_email in payload or set DEFAULT_SENDER env var."
        )

    try:
        access_token = get_app_access_token()
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Failed to get Graph access token (client credentials). Error: {e}"
        )

    resp = send_graph_mail(
        access_token=access_token,
        sender=sender,
        to_emails=payload.to,
        cc_emails=payload.cc,
        subject=payload.subject,
        body_html=payload.body_html,
        importance=payload.importance or "normal",
        attachments=payload.attachments,
    )

    if resp.status_code != 202:
        raise HTTPException(
            status_code=resp.status_code,
            detail={"graph_error": resp.text}
        )

    return JSONResponse(
        content={
            "status": "OK",
            "message": "Email sent",
        },
        status_code=200
    )
