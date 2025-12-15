import os
import json
import base64
from typing import List, Optional

import msal
import requests
from fastapi import FastAPI, HTTPException, Header
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field
from dotenv import load_dotenv

load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID", "common")
SCOPES = (os.getenv("SCOPES") or "Mail.Send").split(",")
TOKEN_CACHE_FILE = os.getenv("TOKEN_CACHE_FILE", "token_cache.bin")
EMAIL_API_KEY = os.getenv("EMAIL_API_KEY")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH_SENDMAIL_URL = "https://graph.microsoft.com/v1.0/me/sendMail"

app = FastAPI(title="Graph Email API (Personal Delegated)")


# ---------- Models ----------

class AttachmentIn(BaseModel):
    name: str
    content_base64: str = Field(..., description="Base64-encoded file bytes")
    content_type: Optional[str] = Field(
        default="application/octet-stream",
        description="MIME type, e.g. text/plain, application/pdf"
    )


class SendEmailRequest(BaseModel):
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


def build_public_client_app(cache=None):
    return msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache
    )


def get_access_token() -> str:
    """
    Delegated auth using public client + device code flow.
    For local testing only (console interaction).
    """
    if not CLIENT_ID:
        raise Exception("CLIENT_ID is not configured.")

    cache = load_cache()
    app_msal = build_public_client_app(cache)

    accounts = app_msal.get_accounts()
    result = None

    if accounts:
        # Try silent token
        result = app_msal.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        # Start device code flow (will print URL+code in console)
        flow = app_msal.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise Exception("Failed to create device flow: %s" % json.dumps(flow, indent=2))

        print("==============================================")
        print("To sign in, open this URL in a browser:")
        print(flow["verification_uri"])
        print("Then enter this code:")
        print(flow["user_code"])
        print("==============================================")

        result = app_msal.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        save_cache(cache)
        return result["access_token"]
    else:
        raise Exception("Could not obtain access token: %s" % json.dumps(result, indent=2))


# ---------- Graph sendMail helper ----------

def send_graph_mail(
    access_token: str,
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
        "saveToSentItems": "true"
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

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    resp = requests.post(GRAPH_SENDMAIL_URL, headers=headers, json=message)
    return resp


# ---------- FastAPI endpoints ----------

@app.get("/")
def root():
    return {
        "status": "ok",
        "message": "Graph Email API (personal delegated)",
        "endpoints": ["/send-email"]
    }


@app.post("/send-email")
def send_email_api(
    payload: SendEmailRequest,
    x_api_key: Optional[str] = Header(None)
):
    # Simple shared-secret API key for local tests
    if EMAIL_API_KEY and x_api_key != EMAIL_API_KEY:
        raise HTTPException(status_code=401, detail="Unauthorized")

    if not payload.to:
        raise HTTPException(status_code=400, detail="At least one TO recipient is required.")

    try:
        access_token = get_access_token()
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Failed to get Graph access token (check console for device login URL/code). Error: {e}"
        )

    resp = send_graph_mail(
        access_token=access_token,
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
