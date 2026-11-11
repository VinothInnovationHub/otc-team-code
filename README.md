# otc-team-code

OTC Assistant, re-packaged to run on **App Service + Azure Functions + APIM** —
this team's own hosting pattern — instead of the Azure Functions-only hosting
used in the original `teams-bot-final-product` project (that project was
**not touched** to produce this; everything here is a separate, independent
copy).

## Why this folder exists

The original project was built on Azure Functions end-to-end because the
subscription it was built in has zero App Service quota. This team's own
codebase (reviewed directly to produce this port) already hosts its bot on
App Service via `microsoft_agents.hosting.aiohttp` (`CloudAdapter` +
`start_agent_process`, a plain aiohttp web server) and runs background work
in small, separate Azure Functions (e.g. their `mi-extrct` function). This
folder follows exactly that same split:

- **`app_services/msbotservice/`** — the bot + the Teams tab's web UI, hosted
  on App Service. Mirrors their own `app_services/msbotservice/src/`
  structure (`main.py`, `start_server.py`, `agent.py`) file-for-file.
- **`az-function-apps/otc-queue-consumer/`** — a small, single-purpose
  Function App that does the actual AI processing (calls the existing
  `generate_otc`-equivalent backend via APIM, finalizes the generated
  document, updates tracking/SharePoint, sends the proactive Teams
  notification). This is the one piece that stays on Azure Functions:
  App Service has no first-party native queue-trigger binding, so
  background/queue-triggered work belongs in a Function App — exactly the
  pattern this team already uses.

## What's proven working vs. what's new here

Every piece of **business logic** below was built and verified working
end-to-end (real uploads, real AI processing, real generated documents, real
Teams notifications, both Blob and SharePoint storage modes) in the original
Azure-Functions-hosted project. **Only the hosting/routing layer changed**
for this port — `agent.py`'s actual logic is copied with no behavioral
changes, aside from the two files (`agent.py` in each deployable) being split
into an "intake" half and a "processing" half instead of one combined file,
matching this team's own convention of giving each separate Function App its
own focused copy of shared logic (their `mi-extrct` and `sql-connect-func`
each carry their own `CommonHandler.py`/`Constants.py`/`CosmosClient.py`
rather than importing across deployable units — this port follows the same
approach for the same reason: each deployable stays self-contained).

**Not yet proven in App Service specifically** (no App Service quota was
available to deploy-test this port against): the hosting/routing swap itself.
The business logic underneath it was proven on a different host (Azure
Functions); the aiohttp wiring in `start_server.py`/`routes.py` mirrors this
team's own already-working `main.py`/`start_server.py` pattern closely enough
that it should work as-is, but this team should be the first to actually run
it end-to-end.

## Folder structure

```
otc-team-code/
├── app_services/msbotservice/
│   ├── requirements.txt
│   └── src/
│       ├── main.py          # entrypoint - mirrors this team's own main.py exactly
│       ├── start_server.py  # aiohttp server bootstrap - mirrors this team's own,
│       │                    #   extended with this project's HTTP routes
│       ├── routes.py        # NEW: aiohttp handlers for /api/form, /api/sas,
│       │                    #   /api/submit, /api/status, /api/search
│       ├── agent.py         # intake logic: storage, tracking, upload-session
│       │                    #   issuance, SharePoint list create/search, bot wiring
│       ├── frontend.py      # the entire Teams tab UI - byte-identical to the
│       │                    #   proven original, since it has zero hosting
│       │                    #   dependency (pure HTML/CSS/JS returned as a string)
│       ├── Constants.py      # every config name/default in one place
│       └── .env.example      # every environment variable this app needs
└── az-function-apps/otc-queue-consumer/
    ├── function_app.py       # the queue trigger - pure routing, calls agent.py
    ├── agent.py              # processing logic: generate_otc call, document
    │                         #   finalization, status updates, proactive message
    ├── Constants.py
    ├── host.json
    ├── requirements.txt
    └── .env.example
```

## Setup

1. **Fill in both `.env.example` files** with this team's real values (Azure
   subscription resources, APIM endpoint + key, storage account, bot identity,
   SharePoint details if using it) and set them as Application Settings on the
   respective App Service / Function App. The two `.env.example` files call
   out exactly which settings must match each other (same storage account,
   same `USE_SHAREPOINT`, same bot identity).
2. **Confirm the App Service startup command.** This team's own
   `app_services/msbotservice/src/main.py` uses relative imports
   (`from .agent import ...`), implying it's launched as a Python module
   (e.g. `python -m msbotservice.src.main` or similar) rather than a bare
   script. Use whatever startup command this team's own existing App Service
   bot is already configured with — it isn't visible in the reviewed source
   files, so it isn't guessed here.
3. **Deploy the Function App** (`az-function-apps/otc-queue-consumer/`)
   exactly as this team already deploys their own Function Apps (e.g.
   `mi-extrct`, `sql-connect-func`) — same tooling, same conventions.
4. **Grant the two apps' identities** Storage Blob/Table Data Contributor on
   the business-data storage account, and (if using SharePoint) Graph
   `Sites.Selected` scoped to the one SharePoint site, exactly as described
   in each `.env.example`.
5. **Test both storage modes** by toggling `USE_SHAREPOINT` on both apps
   together (they must always agree) and submitting a real request through
   the tab.

## What generate_otc contract this expects

`agent.py` (queue consumer) calls the backend via
`POST {APIM_BASE_URL}/otc/otc/generate` with a flat payload including
`investmentMemoBlobPath` / `financialModelBlobPath`, and tolerantly parses a
response containing either base64 file content or a direct download URL
(several naming conventions accepted — see `process_request()`'s field
lookups). If this team's own backend uses different field names or a
different path, adjust `_build_generate_otc_payload()` and the response
parsing in `process_request()` accordingly — everything else is unaffected.
