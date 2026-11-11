"""otc-team-code (App Service side) — intake logic + the Teams bot: storage
(Blob/SharePoint toggle), the request-tracking table, document-upload session
issuance, the queue hand-off, and bot wiring. frontend.py is presentation-only
and calls into this module's HTTP routes (wired in routes.py / start_server.py).

Ported from teams-bot-final-product/app_services/msbotservice/src/agent.py,
which was built and proven working on Azure Functions. The ONLY change for
App Service hosting is the adapter: CloudAdapter (from
microsoft_agents.hosting.aiohttp) in place of a Functions-specific adapter —
everything else below is identical, unmodified business logic, since none of
it ever depended on Azure Functions bindings in the first place.

This file intentionally holds only the INTAKE half of the original agent.py:
config, storage clients, the tracking table, upload-session issuance (both
Blob and SharePoint), the SharePoint list's create/search paths, the queue
hand-off, and the bot handlers. The PROCESSING half (the actual generate_otc
call, document finalization, SharePoint list status updates, and proactive
messaging) lives in az-function-apps/otc-queue-consumer/agent.py instead — a
separate, independently-deployable Function App, following the same pattern
already used in this codebase (mi-extrct / sql-connect-func each carry their
own copy of shared logic rather than importing across deployable units).

EVERYTHING CONFIGURABLE lives in the CONFIG section below — one place, env vars
only, no hardcoded URLs/names/keys. See Constants.py for the env var names and
defaults.

SHAREPOINT TOGGLE: USE_SHAREPOINT=true switches document storage to a SharePoint
document library via Microsoft Graph; USE_SHAREPOINT=false (default) uses Blob
Storage. Independently of this toggle, the SharePoint tracking LIST is written
to (via sharepoint_create_list_item() below) whenever the four SHAREPOINT_*
auth settings are configured — see _sharepoint_configured().
"""

import json
import logging
import sys
from datetime import datetime, timedelta, timezone
from os import environ
from random import randint
from typing import Any

import requests
from azure.core.exceptions import ResourceExistsError, ResourceNotFoundError
from azure.data.tables import TableClient
from azure.identity import ClientSecretCredential, DefaultAzureCredential
from azure.storage.blob import BlobSasPermissions, BlobServiceClient, generate_blob_sas
from azure.storage.queue import QueueClient

from microsoft_agents.hosting.aiohttp import CloudAdapter
from microsoft_agents.hosting.core import (
    Authorization,
    AgentApplication,
    TurnState,
    TurnContext,
    MemoryStorage,
)
from microsoft_agents.authentication.msal import MsalConnectionManager
from microsoft_agents.activity import load_configuration_from_env

from Constants import ConstMessages, Defaults, EnvVariables

logger = logging.getLogger("otc_team_code.agent")

agents_sdk_config = load_configuration_from_env(environ)

# ---------------------------------------------------------------------------
# CONFIG — every external dependency, one place, env vars only. Kept complete
# here (including a few settings only the queue-consumer Function App actually
# uses, such as APIM_*) so this file stays a self-documenting reference of
# every setting this project needs, matching .env.example.
# ---------------------------------------------------------------------------
STORAGE_ACCOUNT_NAME = environ.get(EnvVariables.STORAGE_ACCOUNT_NAME, "")
STORAGE_UPLOADS_CONTAINER = environ.get(EnvVariables.STORAGE_UPLOADS_CONTAINER, Defaults.STORAGE_UPLOADS_CONTAINER)
STORAGE_GENERATED_CONTAINER = environ.get(EnvVariables.STORAGE_GENERATED_CONTAINER, Defaults.STORAGE_GENERATED_CONTAINER)
AZURE_STORAGE_CONNECTION_STRING = environ.get(EnvVariables.AZURE_STORAGE_CONNECTION_STRING, "")
CONVERSATION_STORE_TABLE_NAME = environ.get(EnvVariables.CONVERSATION_STORE_TABLE_NAME, Defaults.CONVERSATION_STORE_TABLE_NAME)
TRACKING_TABLE_NAME = environ.get(EnvVariables.TRACKING_TABLE_NAME, Defaults.TRACKING_TABLE_NAME)
STORAGE_QUEUE_NAME = environ.get(EnvVariables.STORAGE_QUEUE_NAME, Defaults.STORAGE_QUEUE_NAME)
APP_BASE_URL = environ.get(EnvVariables.APP_BASE_URL, "").rstrip("/")
UPLOAD_SAS_TTL_MINUTES = int(environ.get(EnvVariables.UPLOAD_SAS_TTL_MINUTES, str(Defaults.UPLOAD_SAS_TTL_MINUTES)))

# The boolean toggle. Read once at import time; "true" (any case) -> SharePoint,
# anything else -> Blob Storage.
USE_SHAREPOINT = environ.get(EnvVariables.USE_SHAREPOINT, "false").strip().lower() == "true"
SHAREPOINT_SITE_URL = environ.get(EnvVariables.SHAREPOINT_SITE_URL, "")
SHAREPOINT_LIST_NAME = environ.get(EnvVariables.SHAREPOINT_LIST_NAME, Defaults.SHAREPOINT_LIST_NAME)
SHAREPOINT_DRIVE_ID = environ.get(EnvVariables.SHAREPOINT_DRIVE_ID, "")
# Raw Graph site ID ("host,siteGuid,webGuid"), needed for the Lists API -
# distinct from SHAREPOINT_SITE_URL, which is just the human-readable URL.
SHAREPOINT_SITE_ID = environ.get(EnvVariables.SHAREPOINT_SITE_ID, "")

# If SharePoint lives in a different Entra tenant than this app (common when
# this bot's own tenant has no SharePoint/SPO license), a managed identity
# cannot cross tenant boundaries - set these three to use a tenant-scoped app
# registration + client secret instead (see _graph_token()). Leave all three
# blank to use this app's own identity for Graph calls (same-tenant case).
SHAREPOINT_TENANT_ID = environ.get(EnvVariables.SHAREPOINT_TENANT_ID, "")
SHAREPOINT_CLIENT_ID = environ.get(EnvVariables.SHAREPOINT_CLIENT_ID, "")
SHAREPOINT_CLIENT_SECRET = environ.get(EnvVariables.SHAREPOINT_CLIENT_SECRET, "")


def _sharepoint_configured() -> bool:
    """Whether the SharePoint *list* (tracking/visibility mirror) can be
    written to. Deliberately independent of USE_SHAREPOINT, which only
    selects where DOCUMENTS are stored (Blob vs SharePoint) - the tracking
    list is written in both storage modes whenever these four settings are
    present, so SharePoint stays a live record of every request regardless of
    which backend holds the actual files."""
    return bool(SHAREPOINT_SITE_ID and SHAREPOINT_TENANT_ID and SHAREPOINT_CLIENT_ID and SHAREPOINT_CLIENT_SECRET)


GRAPH_SCOPE = "https://graph.microsoft.com/.default"

if not STORAGE_ACCOUNT_NAME and not AZURE_STORAGE_CONNECTION_STRING:
    print("Warning: STORAGE_ACCOUNT_NAME not set. Storage calls will fail.", file=sys.stderr)

STORAGE = MemoryStorage()
CONNECTION_MANAGER = MsalConnectionManager(**agents_sdk_config)
ADAPTER = CloudAdapter(connection_manager=CONNECTION_MANAGER)
AUTHORIZATION = Authorization(STORAGE, CONNECTION_MANAGER, **agents_sdk_config)


# ---------------------------------------------------------------------------
# Storage clients — managed-identity (DefaultAzureCredential) by default; set
# AZURE_STORAGE_CONNECTION_STRING only as an escape hatch.
# ---------------------------------------------------------------------------
def _blob_service_client() -> BlobServiceClient:
    if AZURE_STORAGE_CONNECTION_STRING:
        return BlobServiceClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING)
    return BlobServiceClient(f"https://{STORAGE_ACCOUNT_NAME}.blob.core.windows.net", credential=DefaultAzureCredential())


def _table_client(table_name: str) -> TableClient:
    if AZURE_STORAGE_CONNECTION_STRING:
        return TableClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING, table_name)
    return TableClient(f"https://{STORAGE_ACCOUNT_NAME}.table.core.windows.net", table_name, credential=DefaultAzureCredential())


def _queue_client() -> QueueClient:
    if AZURE_STORAGE_CONNECTION_STRING:
        return QueueClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING, STORAGE_QUEUE_NAME)
    return QueueClient(f"https://{STORAGE_ACCOUNT_NAME}.queue.core.windows.net", STORAGE_QUEUE_NAME, credential=DefaultAzureCredential())


# ---------------------------------------------------------------------------
# Request ID — "OTC-2026-004521": year + 6-digit reference number, assigned
# once at creation and FIXED for the request's lifetime (status is tracked
# separately via STATUS_* and shown via the badge/tracker, never by changing
# this ID) - it is also the storage path prefix for uploaded documents (Blob
# and SharePoint), the Table Storage RowKey, the search key, and the generated
# document's download link.
# ---------------------------------------------------------------------------
def build_request_id() -> str:
    year = datetime.now(timezone.utc).year
    suffix = randint(0, 999999)
    return f"OTC-{year}-{suffix:06d}"


# ---------------------------------------------------------------------------
# Conversation references (for proactive messaging) — saved per-user on every
# bot turn, re-saved under the request_id at submit time so the queue consumer
# (a separate Function App, no bot of its own) can look it up by request_id.
# ---------------------------------------------------------------------------
_USER_CONV_PARTITION = "user-conversation"


def save_user_conversation_reference(teams_user_id: str, conversation_reference: dict[str, Any]) -> None:
    client = _table_client(CONVERSATION_STORE_TABLE_NAME)
    try:
        client.create_table()
    except ResourceExistsError:
        pass
    client.upsert_entity(
        {
            "PartitionKey": _USER_CONV_PARTITION,
            "RowKey": teams_user_id,
            "conversationReferenceJson": json.dumps(conversation_reference, default=str),
            "updatedAt": datetime.now(timezone.utc).isoformat(),
        }
    )


def get_user_conversation_reference(teams_user_id: str) -> dict[str, Any] | None:
    client = _table_client(CONVERSATION_STORE_TABLE_NAME)
    try:
        entity = client.get_entity(_USER_CONV_PARTITION, teams_user_id)
    except ResourceNotFoundError:  # no row yet, or table not yet created - both 404
        return None
    return json.loads(entity["conversationReferenceJson"])


def save_request_conversation_reference(request_id: str, conversation_reference: dict[str, Any]) -> None:
    client = _table_client(CONVERSATION_STORE_TABLE_NAME)
    try:
        client.create_table()
    except ResourceExistsError:
        pass
    client.upsert_entity(
        {
            "PartitionKey": request_id,
            "RowKey": request_id,
            "conversationReferenceJson": json.dumps(conversation_reference, default=str),
            "createdAt": datetime.now(timezone.utc).isoformat(),
        }
    )


# ---------------------------------------------------------------------------
# Tracking table — one row per request, updated as it moves through the
# pipeline. Backs both the live side-panel progress and the search-by-ID box
# in frontend.py. Status transitions happen in the queue consumer, not here;
# this module only creates the row and reads it back.
# ---------------------------------------------------------------------------
def create_tracking_record(request_id: str, fields: dict[str, str], requested_by: str, teams_user_id: str) -> None:
    client = _table_client(TRACKING_TABLE_NAME)
    try:
        client.create_table()
    except ResourceExistsError:
        pass
    now = datetime.now(timezone.utc).isoformat()
    entity: dict[str, Any] = {
        "PartitionKey": "request",
        "RowKey": request_id,
        "requestId": request_id,
        "status": ConstMessages.STATUS_CREATED,
        "customerName": fields.get("customerName", ""),
        "requestType": fields.get("requestType", ""),
        "facilityType": fields.get("facilityType", ""),
        "region": fields.get("region", ""),
        "country": fields.get("country", ""),
        "requestedBy": requested_by,
        "teamsUserId": teams_user_id,
        "storageMode": "sharepoint" if USE_SHAREPOINT else "blob",
        "createdAt": now,
        "updatedAt": now,
        "downloadUrl": "",
        "fileName": "",
        "fileSizeKb": 0,
        "errorMessage": "",
        "sharepointListItemId": "",
    }
    if _sharepoint_configured():
        try:
            # The SharePoint list tracks every request regardless of storage
            # mode (see _sharepoint_configured()'s docstring) - a visibility
            # mirror for business users, not the backend's source of truth
            # (Table Storage is), so a failure here must not block submission.
            entity["sharepointListItemId"] = sharepoint_create_list_item(request_id, fields)
        except Exception as exc:  # noqa: BLE001
            logger.error(json.dumps({"event": "sharepoint_list_create_error", "requestId": request_id, "error": str(exc)}))
    client.upsert_entity(entity)


def get_tracking_record(request_id: str) -> dict[str, Any] | None:
    client = _table_client(TRACKING_TABLE_NAME)
    try:
        entity = client.get_entity("request", request_id)
    except ResourceNotFoundError:  # no row yet, or table not yet created - both 404
        return None
    return dict(entity)


def search_tracking(query: str) -> list[dict[str, Any]]:
    """Search by request ID (exact/prefix) or customer name (substring,
    case-insensitive). Searches both Table Storage (always populated for every
    request) and the SharePoint list (also always written to, independent of
    storage mode), merging by request ID. Table Storage has no native
    substring filter, so customer-name matching is done client-side after a
    partition scan — fine at demo volume, not meant to scale."""
    client = _table_client(TRACKING_TABLE_NAME)
    query_upper = query.strip().upper()
    query_lower = query.strip().lower()
    results = []
    try:
        # query_entities() is lazy (returns an ItemPaged) - the service call, and
        # any ResourceNotFoundError for a not-yet-created table, happens on the
        # first iteration below, not on this call - so the loop must be inside
        # the try, not just the call.
        for entity in client.query_entities("PartitionKey eq 'request'"):
            row = dict(entity)
            if query_upper in row.get("RowKey", "").upper() or query_lower in row.get("customerName", "").lower():
                results.append(row)
    except ResourceNotFoundError:  # table not created yet - no requests submitted at all
        pass

    if _sharepoint_configured():
        seen_ids = {r.get("RowKey", "") for r in results}
        for sp_row in sharepoint_search_list(query):
            if sp_row["requestId"] not in seen_ids:
                results.append(sp_row)

    results.sort(key=lambda r: r.get("createdAt", ""), reverse=True)
    return results[:25]


# ---------------------------------------------------------------------------
# Document storage — the SharePoint/Blob boolean toggle (upload-session
# issuance only; the generated document is uploaded by the queue consumer).
# ---------------------------------------------------------------------------
def issue_upload_sas(session_id: str, file_name: str) -> dict[str, str]:
    """Short-lived, create+write-only user-delegation SAS so frontend.py can PUT a
    file straight to Blob Storage — file bytes never pass through this app."""
    now = datetime.now(timezone.utc)
    starts_on = now - timedelta(minutes=5)
    expires_on = now + timedelta(minutes=UPLOAD_SAS_TTL_MINUTES)

    blob_path = f"pending/{session_id}/{file_name}"
    client = _blob_service_client()
    blob_client = client.get_blob_client(container=STORAGE_UPLOADS_CONTAINER, blob=blob_path)

    user_delegation_key = client.get_user_delegation_key(starts_on, expires_on)
    sas_token = generate_blob_sas(
        account_name=client.account_name,
        container_name=STORAGE_UPLOADS_CONTAINER,
        blob_name=blob_path,
        user_delegation_key=user_delegation_key,
        permission=BlobSasPermissions(create=True, write=True),
        start=starts_on,
        expiry=expires_on,
    )
    return {
        "sasUrl": f"{blob_client.url}?{sas_token}",
        "blobPath": blob_path,
        "expiresOn": expires_on.isoformat(),
        "storageMode": "blob",
    }


def _graph_token() -> str:
    """Uses a tenant-scoped app registration (client credentials) when
    SharePoint lives in a different Entra tenant than this app (see
    SHAREPOINT_TENANT_ID above); otherwise falls back to this app's own
    managed identity for the same-tenant case."""
    if SHAREPOINT_TENANT_ID and SHAREPOINT_CLIENT_ID and SHAREPOINT_CLIENT_SECRET:
        credential = ClientSecretCredential(
            tenant_id=SHAREPOINT_TENANT_ID, client_id=SHAREPOINT_CLIENT_ID, client_secret=SHAREPOINT_CLIENT_SECRET
        )
    else:
        credential = DefaultAzureCredential()
    return credential.get_token(GRAPH_SCOPE).token


def issue_sharepoint_upload_session(session_id: str, file_name: str) -> dict[str, str]:
    """SharePoint equivalent of issue_upload_sas() — a Graph upload session the
    browser can PUT bytes to directly (same 'bytes never touch this app' shape).
    Unlike a Blob SAS PUT, Graph's upload session requires a Content-Range
    header on the PUT even for a single-shot upload - frontend.py reads the
    storageMode field below to decide whether to send it (confirmed via a
    live HTTP 400 when it was missing in the original project - fixed, not
    assumed)."""
    if not SHAREPOINT_SITE_URL or not SHAREPOINT_DRIVE_ID:
        raise RuntimeError(
            "USE_SHAREPOINT is true but SHAREPOINT_SITE_URL/SHAREPOINT_DRIVE_ID are not configured."
        )
    token = _graph_token()
    path = f"pending/{session_id}/{file_name}"
    response = requests.post(
        f"https://graph.microsoft.com/v1.0/drives/{SHAREPOINT_DRIVE_ID}/root:/{path}:/createUploadSession",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
        timeout=30,
    )
    response.raise_for_status()
    body = response.json()
    return {
        "sasUrl": body["uploadUrl"],
        "blobPath": path,
        "expiresOn": body.get("expirationDateTime", ""),
        "storageMode": "sharepoint",
    }


_sharepoint_list_id_cache: str | None = None


def _sharepoint_ensure_list() -> str:
    """Finds (or creates) the SharePoint *list* that tracks requests - distinct
    from the document library this project writes files to (a List in the
    Lists app, not a folder of files). Looked up by display name and
    auto-created with a matching column set on first use so no manual
    SharePoint setup is needed beyond the site itself. Cached in-process;
    re-discovered on cold start."""
    global _sharepoint_list_id_cache
    if _sharepoint_list_id_cache:
        return _sharepoint_list_id_cache
    if not SHAREPOINT_SITE_ID:
        raise RuntimeError(
            "SHAREPOINT_SITE_ID is not configured (needed for the Lists API, distinct from SHAREPOINT_DRIVE_ID)."
        )
    token = _graph_token()
    headers = {"Authorization": f"Bearer {token}"}

    lookup = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/lists",
        headers=headers,
        params={"$filter": f"displayName eq '{SHAREPOINT_LIST_NAME}'"},
        timeout=30,
    )
    lookup.raise_for_status()
    existing = lookup.json().get("value", [])
    if existing:
        _sharepoint_list_id_cache = existing[0]["id"]
        return _sharepoint_list_id_cache

    create = requests.post(
        f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/lists",
        headers={**headers, "Content-Type": "application/json"},
        json={
            "displayName": SHAREPOINT_LIST_NAME,
            "list": {"template": "genericList"},
            "columns": [
                {"name": "RequestId", "text": {}},
                {"name": "CustomerName", "text": {}},
                {"name": "RequestType", "text": {}},
                {"name": "FacilityType", "text": {}},
                {"name": "Status", "text": {}},
                {"name": "DownloadUrl", "text": {}},
            ],
        },
        timeout=30,
    )
    create.raise_for_status()
    _sharepoint_list_id_cache = create.json()["id"]
    return _sharepoint_list_id_cache


def sharepoint_create_list_item(request_id: str, fields: dict[str, str]) -> str:
    """Creates one row in the SharePoint tracking list for a new request.
    Returns the new item's id, which callers store (tracking table's
    sharepointListItemId field) so the queue consumer can PATCH it directly
    later instead of relying on Graph's unreliable custom-column filtering."""
    list_id = _sharepoint_ensure_list()
    token = _graph_token()
    response = requests.post(
        f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/lists/{list_id}/items",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={
            "fields": {
                "Title": request_id,
                "RequestId": request_id,
                "CustomerName": fields.get("customerName", ""),
                "RequestType": fields.get("requestType", ""),
                "FacilityType": fields.get("facilityType", ""),
                "Status": ConstMessages.STATUS_CREATED,
            }
        },
        timeout=30,
    )
    response.raise_for_status()
    return response.json()["id"]


def sharepoint_search_list(query: str) -> list[dict[str, Any]]:
    """Searches the SharePoint tracking list directly via Graph, independent of
    Table Storage - proves the list is a genuinely queryable tracking surface,
    not just a write-only mirror. Client-side substring filter after a full
    list scan: Graph's list-item $filter on custom (non-indexed) columns is
    unreliable. Degrades to an empty list on any failure rather than breaking
    the whole /api/search response."""
    try:
        list_id = _sharepoint_ensure_list()
        token = _graph_token()
        response = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/lists/{list_id}/items",
            headers={"Authorization": f"Bearer {token}"},
            params={"$expand": "fields", "$top": "200"},
            timeout=30,
        )
        response.raise_for_status()
    except Exception as exc:  # noqa: BLE001 - search must degrade gracefully, not 500 the endpoint
        logger.error(json.dumps({"event": "sharepoint_list_search_error", "error": str(exc)}))
        return []

    query_upper = query.strip().upper()
    query_lower = query.strip().lower()
    results = []
    for item in response.json().get("value", []):
        item_fields = item.get("fields", {})
        request_id = item_fields.get("RequestId", "")
        customer_name = item_fields.get("CustomerName", "")
        if query_upper in request_id.upper() or query_lower in customer_name.lower():
            results.append(
                {
                    "requestId": request_id,
                    "RowKey": request_id,
                    "customerName": customer_name,
                    "requestType": item_fields.get("RequestType", ""),
                    "facilityType": item_fields.get("FacilityType", ""),
                    "status": item_fields.get("Status", ""),
                    "downloadUrl": item_fields.get("DownloadUrl", ""),
                    "createdAt": "",
                    "source": "sharepoint",
                }
            )
    return results


def issue_upload_session(session_id: str, file_name: str) -> dict[str, str]:
    """The one entry point routes.py's /api/sas route calls — branches on the
    boolean toggle so nothing above this line needs to know which backend is active."""
    return issue_sharepoint_upload_session(session_id, file_name) if USE_SHAREPOINT else issue_upload_sas(session_id, file_name)


# ---------------------------------------------------------------------------
# Queue hand-off — Storage Queue (not Service Bus): avoid Service Bus unless
# there's a concrete reason it's required. A Storage Queue on the existing
# storage account needs no new resource type.
# ---------------------------------------------------------------------------
def enqueue_request(request_id: str, payload: dict[str, Any]) -> None:
    """The queue-consumer Function App's host.json declares
    QueuesOptions.MessageEncoding "Base64" (the WebJobs Queue trigger
    extension's own default) — the message body MUST be base64-encoded or the
    trigger fails to decode it at the binding level, before the queue
    consumer's code ever runs (no exception, no "Executing" trace — just
    silent, instant poisoning). azure-storage-queue's send_message() does NOT
    base64-encode by default, unlike the .NET SDK conventions WebJobs
    assumes, so this has to be done explicitly here."""
    import base64

    client = _queue_client()
    try:
        client.create_queue()
    except ResourceExistsError:
        pass
    body = json.dumps({"requestId": request_id, **payload}, default=str)
    client.send_message(base64.b64encode(body.encode("utf-8")).decode("ascii"))


# ---------------------------------------------------------------------------
# Bot handlers — minimal. The real UI is frontend.py's tab; the bot's only job
# is the welcome message and remembering each user's conversationReference for
# proactive messaging later (sent by the queue consumer, not this app).
# ---------------------------------------------------------------------------
AGENT_APP = AgentApplication[TurnState](storage=STORAGE, adapter=ADAPTER, authorization=AUTHORIZATION, **agents_sdk_config)


def _remember_user(context: TurnContext) -> None:
    user_id = context.activity.from_property.id if context.activity.from_property else None
    if not user_id:
        return
    conversation_reference = context.activity.get_conversation_reference()
    save_user_conversation_reference(user_id, conversation_reference.model_dump(mode="json", by_alias=True))


@AGENT_APP.conversation_update("membersAdded")
async def on_members_added(context: TurnContext, _state: TurnState):
    _remember_user(context)
    await context.send_activity(
        "Welcome to the OTC Assistant. Open the **New OTC Request** tab above to submit a request and track it live."
    )
    return True


@AGENT_APP.activity("message")
async def on_message(context: TurnContext, _state: TurnState):
    _remember_user(context)
    await context.send_activity("Open the **New OTC Request** tab above to submit or track a request.")


@AGENT_APP.error
async def on_error(context: TurnContext, error: Exception):
    print(f"\n [on_turn_error] unhandled error: {error}", file=sys.stderr)
    await context.send_activity("The bot encountered an error.")
