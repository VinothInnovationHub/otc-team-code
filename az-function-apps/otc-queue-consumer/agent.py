"""otc-queue-consumer — the PROCESSING half of this project's logic: reads a
queued request, calls the existing generate_otc backend (via APIM), finalizes
the generated document into whichever store is active (Blob or SharePoint),
updates the tracking table + SharePoint list, and sends the proactive Teams
notification. The INTAKE half (form, uploads, submit, search, the bot itself)
lives in the separate App Service app (app_services/msbotservice/src/agent.py)
— this Function App has no bot of its own and never talks to a Teams
conversation except via the raw Bot Connector REST call in
send_proactive_card() below.

Deliberately a standalone deployable unit with its own copy of the shared
storage/tracking/SharePoint helpers (not an import across deployable units) —
this mirrors this team's own convention of duplicating CommonHandler.py /
Constants.py / CosmosClient.py across their separate az-function-apps/
folders, rather than introducing a shared private package to maintain.

Hosted on Azure Functions (Consumption), triggered by the Storage Queue that
the App Service app's agent.py::enqueue_request() writes to. Azure Functions
is used here (not App Service) because there is no first-party App-Service-
native queue-trigger binding — a Function App's queue trigger is the standard,
low-maintenance way to run background work alongside an App-Service-hosted
web app, and this team's own codebase already follows exactly this split
(mi-extrct is a separate, small, single-purpose Function App too).
"""

import json
import logging
import sys
import uuid
from datetime import datetime, timedelta, timezone
from os import environ
from typing import Any

import aiohttp
import requests
from azure.core.exceptions import ResourceNotFoundError
from azure.data.tables import TableClient
from azure.identity import ClientSecretCredential, DefaultAzureCredential, ManagedIdentityCredential
from azure.storage.blob import BlobSasPermissions, BlobServiceClient, ContentSettings, generate_blob_sas

from Constants import ConstMessages, Defaults, EnvVariables

logger = logging.getLogger("otc_queue_consumer.agent")

# ---------------------------------------------------------------------------
# CONFIG — every external dependency, one place, env vars only. Kept complete
# here (including a few settings only the App Service app actually uses, such
# as STORAGE_QUEUE_NAME) so this file stays a self-documenting reference of
# every setting this project needs, matching .env.example.
# ---------------------------------------------------------------------------
APIM_BASE_URL = environ.get(EnvVariables.APIM_BASE_URL, "").rstrip("/")
APIM_SUBSCRIPTION_KEY = environ.get(EnvVariables.APIM_SUBSCRIPTION_KEY, "")

STORAGE_ACCOUNT_NAME = environ.get(EnvVariables.STORAGE_ACCOUNT_NAME, "")
STORAGE_UPLOADS_CONTAINER = environ.get(EnvVariables.STORAGE_UPLOADS_CONTAINER, Defaults.STORAGE_UPLOADS_CONTAINER)
STORAGE_GENERATED_CONTAINER = environ.get(EnvVariables.STORAGE_GENERATED_CONTAINER, Defaults.STORAGE_GENERATED_CONTAINER)
AZURE_STORAGE_CONNECTION_STRING = environ.get(EnvVariables.AZURE_STORAGE_CONNECTION_STRING, "")
CONVERSATION_STORE_TABLE_NAME = environ.get(EnvVariables.CONVERSATION_STORE_TABLE_NAME, Defaults.CONVERSATION_STORE_TABLE_NAME)
TRACKING_TABLE_NAME = environ.get(EnvVariables.TRACKING_TABLE_NAME, Defaults.TRACKING_TABLE_NAME)

USE_SHAREPOINT = environ.get(EnvVariables.USE_SHAREPOINT, "false").strip().lower() == "true"
SHAREPOINT_SITE_URL = environ.get(EnvVariables.SHAREPOINT_SITE_URL, "")
SHAREPOINT_LIST_NAME = environ.get(EnvVariables.SHAREPOINT_LIST_NAME, Defaults.SHAREPOINT_LIST_NAME)
SHAREPOINT_DRIVE_ID = environ.get(EnvVariables.SHAREPOINT_DRIVE_ID, "")
SHAREPOINT_SITE_ID = environ.get(EnvVariables.SHAREPOINT_SITE_ID, "")
SHAREPOINT_TENANT_ID = environ.get(EnvVariables.SHAREPOINT_TENANT_ID, "")
SHAREPOINT_CLIENT_ID = environ.get(EnvVariables.SHAREPOINT_CLIENT_ID, "")
SHAREPOINT_CLIENT_SECRET = environ.get(EnvVariables.SHAREPOINT_CLIENT_SECRET, "")


def _sharepoint_configured() -> bool:
    return bool(SHAREPOINT_SITE_ID and SHAREPOINT_TENANT_ID and SHAREPOINT_CLIENT_ID and SHAREPOINT_CLIENT_SECRET)


# Bot identity, reused for the proactive-message Bot Connector call only -
# this Function App has no AgentApplication/bot conversation handling of its
# own, so it doesn't need the full microsoft-agents SDK, just these two ids.
MICROSOFT_APP_ID = environ.get("MicrosoftAppId", "") or environ.get("MICROSOFT_APP_ID", "")
AZURE_CLIENT_ID = environ.get("AZURE_CLIENT_ID", "")

BOT_CONNECTOR_SCOPE = "https://api.botframework.com/.default"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"

if not APIM_BASE_URL:
    print("Warning: APIM_BASE_URL not set. generate_otc calls will fail.", file=sys.stderr)
if not STORAGE_ACCOUNT_NAME and not AZURE_STORAGE_CONNECTION_STRING:
    print("Warning: STORAGE_ACCOUNT_NAME not set. Storage calls will fail.", file=sys.stderr)


# ---------------------------------------------------------------------------
# Storage clients — managed identity by default.
# ---------------------------------------------------------------------------
def _blob_service_client() -> BlobServiceClient:
    if AZURE_STORAGE_CONNECTION_STRING:
        return BlobServiceClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING)
    return BlobServiceClient(f"https://{STORAGE_ACCOUNT_NAME}.blob.core.windows.net", credential=DefaultAzureCredential())


def _table_client(table_name: str) -> TableClient:
    if AZURE_STORAGE_CONNECTION_STRING:
        return TableClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING, table_name)
    return TableClient(f"https://{STORAGE_ACCOUNT_NAME}.table.core.windows.net", table_name, credential=DefaultAzureCredential())


def get_request_conversation_reference(request_id: str) -> dict[str, Any] | None:
    client = _table_client(CONVERSATION_STORE_TABLE_NAME)
    try:
        entity = client.get_entity(request_id, request_id)
    except ResourceNotFoundError:  # no row yet, or table not yet created - both 404
        return None
    return json.loads(entity["conversationReferenceJson"])


def get_tracking_record(request_id: str) -> dict[str, Any] | None:
    client = _table_client(TRACKING_TABLE_NAME)
    try:
        entity = client.get_entity("request", request_id)
    except ResourceNotFoundError:
        return None
    return dict(entity)


def update_tracking_status(request_id: str, status: str, **extra: Any) -> None:
    client = _table_client(TRACKING_TABLE_NAME)
    entity = {"PartitionKey": "request", "RowKey": request_id, "status": status, "updatedAt": datetime.now(timezone.utc).isoformat()}
    entity.update(extra)
    client.update_entity(entity, mode="merge")

    if _sharepoint_configured():
        try:
            record = get_tracking_record(request_id)
            item_id = record.get("sharepointListItemId", "") if record else ""
            if item_id:
                sharepoint_update_list_item(item_id, status=status, download_url=extra.get("downloadUrl"))
        except Exception as exc:  # noqa: BLE001 - mirror only, must not block the core status update
            logger.error(json.dumps({"event": "sharepoint_list_update_error", "requestId": request_id, "error": str(exc)}))


# ---------------------------------------------------------------------------
# SharePoint — Graph calls for document download/upload and the tracking
# list's status updates. (List item CREATION happens on the App Service side,
# at submit time — this file only ever updates an existing item.)
# ---------------------------------------------------------------------------
def _graph_token() -> str:
    if SHAREPOINT_TENANT_ID and SHAREPOINT_CLIENT_ID and SHAREPOINT_CLIENT_SECRET:
        credential = ClientSecretCredential(
            tenant_id=SHAREPOINT_TENANT_ID, client_id=SHAREPOINT_CLIENT_ID, client_secret=SHAREPOINT_CLIENT_SECRET
        )
    else:
        credential = DefaultAzureCredential()
    return credential.get_token(GRAPH_SCOPE).token


_sharepoint_list_id_cache: str | None = None


def _sharepoint_ensure_list() -> str:
    global _sharepoint_list_id_cache
    if _sharepoint_list_id_cache:
        return _sharepoint_list_id_cache
    if not SHAREPOINT_SITE_ID:
        raise RuntimeError("SHAREPOINT_SITE_ID is not configured.")
    token = _graph_token()
    lookup = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/lists",
        headers={"Authorization": f"Bearer {token}"},
        params={"$filter": f"displayName eq '{SHAREPOINT_LIST_NAME}'"},
        timeout=30,
    )
    lookup.raise_for_status()
    existing = lookup.json().get("value", [])
    if not existing:
        raise RuntimeError(f"SharePoint list '{SHAREPOINT_LIST_NAME}' not found - it should have been created at submit time.")
    _sharepoint_list_id_cache = existing[0]["id"]
    return _sharepoint_list_id_cache


def sharepoint_update_list_item(item_id: str, status: str | None = None, download_url: str | None = None) -> None:
    """Patches an existing SharePoint list item's fields as a request's status
    changes - mirrors update_tracking_status()'s Table Storage write."""
    body: dict[str, str] = {}
    if status is not None:
        body["Status"] = status
    if download_url is not None:
        body["DownloadUrl"] = download_url
    if not body:
        return
    list_id = _sharepoint_ensure_list()
    token = _graph_token()
    response = requests.patch(
        f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/lists/{list_id}/items/{item_id}/fields",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=body,
        timeout=30,
    )
    response.raise_for_status()


def _sharepoint_download(path: str) -> bytes:
    """Reads a file's bytes straight out of the SharePoint drive via Graph."""
    token = _graph_token()
    response = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{SHAREPOINT_DRIVE_ID}/root:/{path}:/content",
        headers={"Authorization": f"Bearer {token}"},
        timeout=60,
    )
    response.raise_for_status()
    return response.content


def _stage_sharepoint_upload_for_processing(blob_path: str) -> None:
    """generate_otc (the existing, unmodified backend) can only read source
    documents from Blob Storage - so when USE_SHAREPOINT is true, this stages
    a transient copy of each uploaded document into Blob Storage at the same
    relative path, purely so generate_otc's existing contract keeps working
    unmodified. SharePoint remains the system of record the user/tab sees;
    Blob Storage here is just the AI pipeline's working cache."""
    content = _sharepoint_download(blob_path)
    client = _blob_service_client()
    client.get_blob_client(container=STORAGE_UPLOADS_CONTAINER, blob=blob_path).upload_blob(content, overwrite=True)


def upload_generated_document(content: bytes, file_name: str) -> str:
    """Uploads the generated .docx to whichever store the toggle selects,
    returns a URL frontend.py can show as the Download link."""
    if USE_SHAREPOINT:
        if not SHAREPOINT_SITE_URL or not SHAREPOINT_DRIVE_ID:
            raise RuntimeError("USE_SHAREPOINT is true but SharePoint is not fully configured.")
        token = _graph_token()
        path = f"generated/{file_name}"
        response = requests.put(
            f"https://graph.microsoft.com/v1.0/drives/{SHAREPOINT_DRIVE_ID}/root:/{path}:/content",
            headers={"Authorization": f"Bearer {token}"},
            data=content,
            timeout=60,
        )
        response.raise_for_status()
        return response.json().get("webUrl", "")

    client = _blob_service_client()
    blob_client = client.get_blob_client(container=STORAGE_GENERATED_CONTAINER, blob=file_name)
    content_type = Defaults.CONTENT_TYPES.get(
        file_name.rsplit(".", 1)[-1].lower() if "." in file_name else "", "application/octet-stream"
    )
    blob_client.upload_blob(content, overwrite=True, content_settings=ContentSettings(content_type=content_type))
    sas_token = generate_blob_sas(
        account_name=client.account_name,
        container_name=STORAGE_GENERATED_CONTAINER,
        blob_name=file_name,
        account_key=None,
        user_delegation_key=client.get_user_delegation_key(
            datetime.now(timezone.utc) - timedelta(minutes=5), datetime.now(timezone.utc) + timedelta(days=7)
        ),
        permission=BlobSasPermissions(read=True),
        expiry=datetime.now(timezone.utc) + timedelta(days=7),
    )
    return f"{blob_client.url}?{sas_token}"


# ---------------------------------------------------------------------------
# generate_otc — called via APIM. Payload/response field names below were
# CONFIRMED (not assumed) against the real backend in the proving project
# (teams-bot-final-product) via a live Application Insights traceback; copy
# this file's call shape as-is unless this team's own generate_otc-equivalent
# backend uses different field names, in which case adjust
# _build_generate_otc_payload() and process_request()'s response parsing below.
# ---------------------------------------------------------------------------
async def call_generate_otc(payload: dict[str, Any]) -> dict[str, Any]:
    if not APIM_BASE_URL:
        raise RuntimeError("APIM_BASE_URL environment variable not set.")

    headers = {"Content-Type": "application/json"}
    if APIM_SUBSCRIPTION_KEY:
        headers["Ocp-Apim-Subscription-Key"] = APIM_SUBSCRIPTION_KEY

    async with aiohttp.ClientSession() as session:
        async with session.post(
            f"{APIM_BASE_URL}/otc/otc/generate",
            headers=headers,
            json=payload,
            timeout=aiohttp.ClientTimeout(total=Defaults.GENERATE_OTC_TIMEOUT_SECONDS),
        ) as response:
            body_text = await response.text()
            if response.status not in (200, 202):
                logger.error(json.dumps({"event": "generate_otc_error", "status": response.status, "body": body_text[:500]}))
                raise RuntimeError(f"generate_otc failed: HTTP {response.status}")
            return json.loads(body_text)


def _build_generate_otc_payload(message: dict[str, Any]) -> dict[str, Any]:
    """Translate this app's own submit payload shape (documents.investment_memo.path /
    documents.financial_model.path) into generate_otc's flat expected shape
    (investmentMemoBlobPath / financialModelBlobPath)."""
    documents = message.get("documents") or {}
    payload = dict(message)
    payload["investmentMemoBlobPath"] = (documents.get("investment_memo") or {}).get("path", "")
    payload["financialModelBlobPath"] = (documents.get("financial_model") or {}).get("path", "")
    return payload


# ---------------------------------------------------------------------------
# Proactive messaging — direct Bot Connector REST call with an MSI-acquired
# token (there is no Python binding for the SDK's server-to-server Connector
# client). This Function App has no bot conversation handling of its own -
# this is the one place it talks to Teams at all.
# ---------------------------------------------------------------------------
def _bot_connector_token() -> str:
    credential = ManagedIdentityCredential(client_id=MICROSOFT_APP_ID) if MICROSOFT_APP_ID else DefaultAzureCredential()
    return credential.get_token(BOT_CONNECTOR_SCOPE).token


def send_proactive_card(conversation_reference: dict[str, Any], card_content: dict[str, Any]) -> None:
    service_url = conversation_reference["serviceUrl"]
    conversation_id = conversation_reference["conversation"]["id"]
    activity = {
        "type": "message",
        "id": str(uuid.uuid4()),
        "serviceUrl": service_url,
        "channelId": conversation_reference["channelId"],
        "conversation": conversation_reference["conversation"],
        "from": conversation_reference.get("agent") or conversation_reference.get("bot") or {"id": MICROSOFT_APP_ID},
        "recipient": conversation_reference.get("user"),
        "locale": conversation_reference.get("locale"),
        "attachments": [{"contentType": "application/vnd.microsoft.card.adaptive", "content": card_content}],
    }
    token = _bot_connector_token()
    response = requests.post(
        f"{service_url.rstrip('/')}/v3/conversations/{conversation_id}/activities",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        data=json.dumps(activity, default=str),
        timeout=30,
    )
    response.raise_for_status()


def build_result_card(tracking: dict[str, Any]) -> dict[str, Any]:
    facts = [
        {"title": "Customer Name", "value": tracking.get("customerName", "")},
        {"title": "Request Type", "value": tracking.get("requestType", "")},
        {"title": "Facility Type", "value": tracking.get("facilityType", "")},
        {"title": "Request ID", "value": tracking.get("requestId", "")},
        {"title": "Completed", "value": tracking.get("updatedAt", "")},
    ]
    return {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.5",
        "body": [
            {"type": "TextBlock", "text": "✔ Draft OTC Generated Successfully", "weight": "Bolder", "size": "Large", "color": "Good", "wrap": True},
            {"type": "FactSet", "facts": facts},
            {"type": "TextBlock", "text": f"{tracking.get('fileName', '')} ({tracking.get('fileSizeKb', 0)} KB)", "wrap": True, "isSubtle": True},
        ],
        "actions": [{"type": "Action.OpenUrl", "title": "Download Draft OTC", "url": tracking.get("downloadUrl", "")}],
    }


# ---------------------------------------------------------------------------
# The full pipeline a queue message triggers — see function_app.py's queue
# trigger. Centralized here so function_app.py stays pure hosting plumbing.
# ---------------------------------------------------------------------------
async def process_request(message: dict[str, Any]) -> None:
    request_id = message["requestId"]
    try:
        update_tracking_status(request_id, ConstMessages.STATUS_IN_REVIEW)
        update_tracking_status(request_id, ConstMessages.STATUS_PROCESSING)

        if USE_SHAREPOINT:
            # generate_otc can only read from Blob Storage - stage transient
            # copies before calling it.
            documents = message.get("documents") or {}
            for doc in (documents.get("investment_memo"), documents.get("financial_model")):
                if doc and doc.get("path"):
                    _stage_sharepoint_upload_for_processing(doc["path"])

        result = await call_generate_otc(_build_generate_otc_payload(message))

        # generate_otc's exact response shape may vary by backend - try every
        # plausible key for each field (snake_case, camelCase, and a couple of
        # likely alternates) and log clearly if none match, so a contract
        # mismatch is immediately diagnosable instead of silently producing an
        # empty download link.
        inner = result.get("result") if isinstance(result.get("result"), dict) else result

        def _first(*keys: str) -> Any:
            for key in keys:
                if inner.get(key) not in (None, ""):
                    return inner[key]
            return None

        file_name = _first("file_name", "fileName", "filename", "name") or f"{request_id}.docx"
        content_b64 = _first("content_base64", "contentBase64", "content", "fileContent", "base64")
        download_url_field = _first("download_url", "downloadUrl", "documentUrl", "document_url", "url", "fileUrl")
        file_size_kb_field = _first("file_size_kb", "fileSizeKb", "size_kb", "sizeKb")

        if content_b64:
            import base64

            content = base64.b64decode(content_b64)
            download_url = upload_generated_document(content, file_name)
            file_size_kb = round(len(content) / 1024)
        elif download_url_field:
            if USE_SHAREPOINT:
                # If generate_otc places its own output directly in Blob
                # Storage regardless of this app's toggle, mirror that output
                # into SharePoint here to keep SharePoint the system of
                # record the user/tab sees when the toggle is on.
                content = requests.get(download_url_field, timeout=60).content
                download_url = upload_generated_document(content, file_name)
                file_size_kb = round(len(content) / 1024)
            else:
                download_url = download_url_field
                file_size_kb = file_size_kb_field or 0
        else:
            logger.error(
                json.dumps(
                    {
                        "event": "generate_otc_unrecognized_response_shape",
                        "requestId": request_id,
                        "responseKeys": list(inner.keys()) if isinstance(inner, dict) else type(inner).__name__,
                    }
                )
            )
            raise RuntimeError(
                "generate_otc returned a response with no recognized content/download-url field "
                f"(keys seen: {list(inner.keys()) if isinstance(inner, dict) else inner!r})"
            )

        update_tracking_status(
            request_id,
            ConstMessages.STATUS_GENERATED,
            downloadUrl=download_url,
            fileName=file_name,
            fileSizeKb=file_size_kb,
        )
        update_tracking_status(request_id, ConstMessages.STATUS_COMPLETED)

        conversation_reference = get_request_conversation_reference(request_id)
        if conversation_reference:
            tracking = get_tracking_record(request_id) or {}
            send_proactive_card(conversation_reference, build_result_card(tracking))

    except Exception as exc:  # noqa: BLE001 - must not crash the queue trigger
        logger.error(json.dumps({"event": "process_request_error", "requestId": request_id, "error": str(exc)}))
        update_tracking_status(request_id, ConstMessages.STATUS_FAILED, errorMessage=str(exc))
