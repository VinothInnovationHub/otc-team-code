"""aiohttp route handlers — pure hosting plumbing, the App-Service equivalent of
teams-bot-final-product's function_app.py. Every real decision (storage toggle,
tracking, the queue hand-off) lives in agent.py; this file only parses the
HTTP request, validates input, and formats the JSON response.
"""

import json
import logging

from aiohttp.web import Request, Response, json_response

import agent
import frontend
from Constants import REQUIRED_FIELDS

logger = logging.getLogger("otc_team_code.routes")

ALLOWED_UPLOAD_EXTENSIONS = {".pdf", ".doc", ".docx", ".xls", ".xlsx"}


def _json(payload: dict, status: int = 200) -> Response:
    # default=str mirrors the Functions version's json.dumps(..., default=str) -
    # Table Storage entity values (datetimes etc.) aren't natively JSON-serializable.
    return Response(text=json.dumps(payload, default=str), status=status, content_type="application/json")


async def form(req: Request) -> Response:
    return Response(text=frontend.render_page(), content_type="text/html")


async def sas(req: Request) -> Response:
    session_id = req.query.get("sessionId", "")
    file_name = req.query.get("fileName", "")
    if not session_id or not file_name:
        return _json({"error": "sessionId and fileName are required"}, status=400)
    extension = "." + file_name.rsplit(".", 1)[-1].lower() if "." in file_name else ""
    if extension not in ALLOWED_UPLOAD_EXTENSIONS:
        # Enforced server-side, not just by the file picker's `accept` attribute in
        # frontend.py — that's trivially bypassed (drag-and-drop, devtools).
        return _json(
            {"error": f"Unsupported file type '{extension or file_name}'. Allowed: PDF, DOC/DOCX, XLS/XLSX."},
            status=400,
        )
    try:
        result = agent.issue_upload_session(session_id, file_name)
    except Exception as exc:  # noqa: BLE001
        logger.error(json.dumps({"event": "sas_error", "error": str(exc)}))
        return _json({"error": str(exc)}, status=502)
    return _json(result)


async def submit(req: Request) -> Response:
    try:
        payload = await req.json()
    except Exception:
        return _json({"error": "Invalid JSON"}, status=400)

    missing = [f for f in REQUIRED_FIELDS if not payload.get(f)]
    if missing:
        return _json({"error": f"Missing required fields: {', '.join(missing)}"}, status=400)

    request_id = agent.build_request_id()
    requested_by = payload.get("requestedBy", "")
    teams_user_id = payload.get("teamsUserId", "")

    agent.create_tracking_record(request_id, payload, requested_by, teams_user_id)

    conversation_reference = agent.get_user_conversation_reference(teams_user_id) if teams_user_id else None
    if conversation_reference:
        agent.save_request_conversation_reference(request_id, conversation_reference)

    agent.enqueue_request(request_id, payload)

    return _json({"requestId": request_id, "status": "submitted"})


async def status(req: Request) -> Response:
    request_id = req.query.get("id", "")
    if not request_id:
        return _json({"error": "id is required"}, status=400)
    record = agent.get_tracking_record(request_id)
    if not record:
        return _json({"error": "not found"}, status=404)
    # frontend.py's pollStatus() reads data.record, not the bare body - keep this
    # wrapped (same convention as /api/search's {"results": [...]}), not a bare object.
    return _json({"record": record})


async def search(req: Request) -> Response:
    query = req.query.get("q", "")
    results = agent.search_tracking(query) if query else []
    return _json({"results": results})
