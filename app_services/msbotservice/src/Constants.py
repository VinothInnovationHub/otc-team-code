class ConstMessages:
    FILENAME_AGENT = "agent"
    FILENAME_ROUTES = "routes"

    STATUS_CREATED = "Request Created"
    STATUS_IN_REVIEW = "In Review"
    STATUS_PROCESSING = "AI Processing"
    STATUS_GENERATED = "Document Generated"
    STATUS_COMPLETED = "Completed"
    STATUS_FAILED = "Failed"


class EnvVariables:
    # Matches rg-otc-agent-dev's existing naming exactly (Container App + Function
    # App env vars already use these names) so this app's settings can reuse the
    # same values rather than inventing a second naming convention.
    APIM_BASE_URL = "APIM_BASE_URL"
    APIM_SUBSCRIPTION_KEY = "APIM_SUBSCRIPTION_KEY"
    STORAGE_ACCOUNT_NAME = "STORAGE_ACCOUNT_NAME"
    STORAGE_UPLOADS_CONTAINER = "STORAGE_UPLOADS_CONTAINER"
    STORAGE_GENERATED_CONTAINER = "STORAGE_GENERATED_CONTAINER"
    AZURE_STORAGE_CONNECTION_STRING = "AZURE_STORAGE_CONNECTION_STRING"
    CONVERSATION_STORE_TABLE_NAME = "CONVERSATION_STORE_TABLE_NAME"
    TRACKING_TABLE_NAME = "TRACKING_TABLE_NAME"
    STORAGE_QUEUE_NAME = "STORAGE_QUEUE_NAME"
    APP_BASE_URL = "APP_BASE_URL"
    UPLOAD_SAS_TTL_MINUTES = "UPLOAD_SAS_TTL_MINUTES"

    # The boolean toggle: "true" -> SharePoint document library, "false" -> Blob
    # Storage. Read as a string env var ("true"/"false") since that's how App
    # Settings are stored; parsed once at import time (see agent.py USE_SHAREPOINT).
    USE_SHAREPOINT = "USE_SHAREPOINT"
    SHAREPOINT_SITE_URL = "SHAREPOINT_SITE_URL"
    # Display name of the SharePoint *list* (not the document library) that
    # tracks requests - auto-created with this name on first use if missing.
    SHAREPOINT_LIST_NAME = "SHAREPOINT_LIST_NAME"
    SHAREPOINT_DRIVE_ID = "SHAREPOINT_DRIVE_ID"
    # The raw Graph site ID (e.g. "host,siteGuid,webGuid") - distinct from
    # SHAREPOINT_SITE_URL, needed for the Lists API (/sites/{siteId}/lists/...).
    SHAREPOINT_SITE_ID = "SHAREPOINT_SITE_ID"

    # The SharePoint site may live in a DIFFERENT Microsoft Entra tenant than the
    # one hosting this app - a managed identity cannot cross tenant boundaries, so
    # the SharePoint branch needs its own tenant-scoped app registration + client
    # secret in that case, separate from every other Azure/Graph call in this
    # project. If SharePoint lives in the SAME tenant as this app, these three can
    # be left unset and the app's own identity is used instead (see agent.py's
    # _graph_token()).
    SHAREPOINT_TENANT_ID = "SHAREPOINT_TENANT_ID"
    SHAREPOINT_CLIENT_ID = "SHAREPOINT_CLIENT_ID"
    SHAREPOINT_CLIENT_SECRET = "SHAREPOINT_CLIENT_SECRET"


class Defaults:
    STORAGE_UPLOADS_CONTAINER = "otc-uploads"
    STORAGE_GENERATED_CONTAINER = "otc-generated"
    CONVERSATION_STORE_TABLE_NAME = "ConversationReferences"
    TRACKING_TABLE_NAME = "OtcRequestTracking"
    STORAGE_QUEUE_NAME = "otc-request-queue"
    SHAREPOINT_LIST_NAME = "OTC Requests"
    UPLOAD_SAS_TTL_MINUTES = 30
    GENERATE_OTC_TIMEOUT_SECONDS = 90

    # File extension -> Content-Type for the generated document upload, so a
    # browser download/open shows the right app icon instead of a generic
    # application/octet-stream default.
    CONTENT_TYPES = {
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "doc": "application/msword",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "xls": "application/vnd.ms-excel",
        "pdf": "application/pdf",
    }


# Required fields the tab (frontend.py) submits alongside the two document
# references. Must match frontend.py exactly: only these three carry a required
# asterisk and client-side check there — region/country are optional metadata
# (plain dropdowns, no asterisk, no required check). Keep both in sync by hand.
REQUIRED_FIELDS = ("customerName", "requestType", "facilityType")
