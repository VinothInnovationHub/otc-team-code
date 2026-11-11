# OTC Assistant — Tech & Flow Explainer

Nov 11, 2026 · @Vinothkumar

## At a Glance

Read the whole document for the full picture, or use just this section if you need the headline answers right now.

- **Two codebases, same logic.** `teams-bot-final-product` (Azure Functions) is proven and live. `otc-team-code` (App Service + a small Function App) is the exact same business logic, re-hosted to match how another team runs their own bots — built and verified to compile, but not yet deployed or run anywhere.
- **Files always land in company-owned storage** — a Blob container or a SharePoint site the business controls — never in any individual employee's personal OneDrive. See "Document Upload Approach" below for why that distinction matters.
- **Sharing this app across companies** (a merger, or handing the install package to a separate organization): Blob storage mode works everywhere with zero extra setup; SharePoint mode needs one extra permission step; the Adaptive Card + paperclip pattern breaks down once a separate tenant is involved. See "Cross-Organization Scenarios" below.
- **Debugging:** work out which of the 15 numbered steps failed, then go straight to the one function responsible for it. See "Debugging Method" at the end.

The sections below follow one reading order: a plain-English glossary of the Azure pieces first, then each codebase's full step-by-step flow, then the storage-approach decision and what it means in practice, and finally a recommendation and a debugging guide.

## Live Walkthrough (Video)

A real end-to-end run against the live deployment: opening the Teams UI, uploading the two documents, watching the live progress tracker through processing, getting the download link once the draft is ready, and using the search-by-ID feature.

&#91;image: Live walkthrough — Teams UI, document upload, processing, download link, search\]

## Complete End-to-End Architecture

The full picture in one diagram — every component actually running in the live deployment, exactly as the demo above shows it.

&#91;embedded content: Complete end-to-end architecture — every component in the live, proven deployment\]

## Overview

Both codebases do the exact same job: a user fills a form in Microsoft Teams, uploads two documents (Investment Memo, Financial Model), and an AI drafts an Offer-to-Close document for them — with live progress tracking and a Teams notification when it's done. The business logic is identical between them; only *where the code physically runs* differs.

- **`teams-bot-final-product`** — everything hosted on Azure Functions. This is the proven, live, working system.
- **`otc-team-code`** — the same logic re-packaged to run on App Service + a small separate Function App, matching how another team already hosts their own bots. Built to hand over to them; not yet deployed anywhere.

This document also covers a design decision that deserves a deliberate choice, not a default one: **how the two uploaded documents reach storage** — Microsoft Teams' built-in Adaptive Card + paperclip attachment flow, versus the direct-to-storage web tab approach this project uses. The two have materially different consequences for where sensitive business documents actually end up living.

## The Building Blocks

Before the flow makes sense, it helps to know what each Azure piece actually is — think of these as physical "departments" in a company, each doing one job.

| Azure piece | What it actually is | Plain-English analogy |
| --- | --- | --- |
| Blob Storage | A place to store files | A filing cabinet for documents |
| Table Storage | A place to store structured records (rows with columns) | A spreadsheet — one row per request, with a Status column |
| Queue Storage | A place to drop a note saying "please do this job" | A ticket dispenser — you take a number, someone processes it later |
| Function App / App Service | A place where your Python code actually runs | The office building where your staff (code) sit |
| Bot Service | Teams' own registry that lets your code send/receive chat messages | The reception desk that connects outside visitors (Teams users) to your office |
| APIM (API Management) | A gatekeeper in front of another service | A receptionist who checks your ID (subscription key) before letting you through to a specific department |
| generate\_otc | The actual AI worker (Document Intelligence reads your files, OpenAI writes the draft) | The specialist who does the real work, sitting behind APIM's front desk |
| SharePoint / Graph API | Microsoft's document/list system, reached via its own "Graph" API | An alternative filing cabinet + a shared spreadsheet, in a different Microsoft building entirely |

## teams-bot-final-product

**Tech stack**: Python, hosted entirely on **Azure Functions** (serverless — Azure runs your code only when something triggers it: an HTTP request, or a queue message). Uses Microsoft's own `microsoft-agents` SDK for the Teams bot plumbing, plain HTML/CSS/JavaScript for the tab (no frontend framework, no build step), and the Azure Python SDKs for talking to storage.

&#91;embedded content: teams-bot-final-product — one Function App runs the whole flow\]

**The request flow, step by step** (the one thing worth memorizing — every bug lives somewhere in this chain):

1. User opens the tab → sees the form
2. User fills the form, picks 2 files, clicks Submit
3. Browser asks the server for a safe upload link (a short-lived SAS token for Blob, or a Graph upload session for SharePoint)
4. Browser uploads the 2 files **directly** to storage — the server's own code never touches the file bytes
5. Browser posts the form data to `/api/submit`
6. Server creates a row in Table Storage (status = "Request Created") and, if SharePoint is configured, a matching row in a SharePoint list
7. Server drops a message on the Storage Queue: "process request OTC-2026-xxxxxx"
8. Server replies immediately with the request ID — the browser starts polling `/api/status` every 4 seconds

*(the rest happens in the background, independent of the browser)*

9. The Queue trigger wakes the Function App up when the message arrives
10. It calls `generate_otc` **through APIM** (APIM checks the subscription key first, then forwards the call)
11. `generate_otc` reads the 2 files from Blob, runs Document Intelligence + Azure OpenAI, returns a result
12. The app uploads the generated `.docx` to whichever store is active (Blob or SharePoint)
13. The tracking row is updated to "Completed" with a real download link
14. A proactive Adaptive Card is sent into the Teams chat: "your draft is ready"
15. The browser's next poll sees "Completed" → shows the checkmark animation and download button

**Status: proven, live, working** — every step above has been tested end-to-end with real files, both storage modes.

## otc-team-code

**Tech stack**: the same Python business logic, but hosted differently. The bot and the tab run on **App Service** (an always-on web server, not serverless — your code runs continuously via `aiohttp`, Python's own async web framework) using Microsoft's first-party `microsoft-agents-hosting-aiohttp` adapter. A separate, small **Function App** handles the background processing, because App Service has no built-in "wake up when a queue message arrives" feature — only Azure Functions has that.

&#91;embedded content: otc-team-code — App Service (always-on) handles intake; a Function App (event-triggered) handles processing\]

**Why split into two pieces?** Steps 1–8 (form, upload link, submit) need something that's always listening for web requests — that's what App Service is for. Steps 9–15 (process the queue message, call the AI, update status, notify Teams) only need to run occasionally, triggered by a queue message — exactly what Azure Functions is built for. Running background work like this *inside* an App Service app is possible but awkward (App Service wants to go idle/scale down when nothing is actively serving web traffic, which fights a long-running background worker). A small dedicated Function App avoids that fight entirely.

**The flow itself is identical** to teams-bot-final-product's 15 steps above — only *which box* runs steps 1–8 versus 9–15 has changed. One extra wrinkle: here the SharePoint tracking-list *row is created* by App Service (at submit time) but *updated* by the Queue Consumer (as the AI processing progresses) — both sides talk to SharePoint, just at different moments.

**Status: built and verified to compile; not yet deployed or run anywhere** — there was no App Service quota available to actually test this hosting swap against. The business logic underneath it was already proven on the other host.

## Document Upload Approach: Adaptive Card vs Direct-to-Storage

This is a real design fork worth deciding deliberately. Microsoft Teams bots can collect file uploads two very different ways.

**Adaptive Card + paperclip**: the bot sends a card with an "Upload" button; the user clicks it, then attaches the file using Teams' native paperclip icon — the same gesture used for sharing a file in any chat. Teams silently uploads that file to **the user's own personal OneDrive** and hands the bot a reference to it.

**Direct-to-storage (this project's approach)**: a web tab shows a drag-and-drop dropzone; the browser asks the server for a short-lived upload link, then uploads the file **directly to the organization's own storage** (a Blob container or a SharePoint site the business controls) — never to any individual's personal account.

&#91;embedded content: paperclip uploads land in a personal OneDrive; this project's tab lands in organization-owned storage\]

### The core problem with paperclip uploads

It isn't the UX gesture that's the issue — it's **where the file physically ends up**. A paperclip upload lands in that one employee's personal OneDrive, not in any storage the business owns or controls. Concretely:

- If that employee leaves the company, their OneDrive can be disabled or deleted — and the source document (an Investment Memo, a Financial Model) can become permanently unreachable, even though the AI-generated draft and the tracking record still exist.
- The document is now subject to **that individual's own OneDrive settings** — retention policy, sharing settings, DLP scanning — not the organization's central storage policy. For financial/compliance documents, this is a real governance gap, not a theoretical one.
- There's no predictable, centrally-indexed location (`pending/{request-id}/{filename}`, the convention this project uses) — files are scattered across however many employees have ever submitted a request, each in their own OneDrive, under whatever name Teams happened to give it.
- Re-processing, auditing, or bulk-exporting source documents later means going back through each individual's sharing permissions rather than one storage account with one access policy.

### Merits and demerits, end to end

|  | Adaptive Card + Paperclip | Direct-to-Storage (this project) |
| --- | --- | --- |
| **Where files land** | Each user's personal OneDrive | Organization-owned Blob container or SharePoint site |
| **Engineering effort** | Low — no SAS tokens, no CORS, no web tab to build/host | Higher — needs upload-link issuance, CORS configuration, a hosted tab |
| **Governance / compliance** | Weak — subject to each individual's personal storage settings; breaks if that person leaves | Strong — one centrally-owned, access-controlled location for every request |
| **Long-term retrieval / audit** | Hard — scattered across many personal accounts, no predictable path | Easy — one account, one predictable path per request ID |
| **UX richness** | Chat-only: plain text prompts, no live visual feedback, a 2-step dance per file (click button, then attach) | Rich: drag-and-drop, instant file-type/size validation, thumbnails, a live progress tracker, animations |
| **Native fit for a pure chat bot** | Very natural — no tab, everything stays in the conversation | Needs a tab registered in the Teams manifest — more moving parts if the business wants chat-only |
| **Scalability of the AI pipeline** | Harder — the backend needs Graph access tied to each individual sharing event | Simple — the backend reads from one known storage location with its own service identity |
| **File bytes through your own server?** | No (goes straight to OneDrive via Teams) | No (goes straight to Blob/SharePoint via a direct upload link) — same efficiency, different destination |

The one thing both approaches share: in neither case does the file pass through the bot's own server code. The difference that actually matters is **whose storage it lands in** — an individual's, or the organization's.

## Cross-Organization Scenarios: Mergers, Guests & Separate-Tenant Installs

A common real-world situation: **Company A** acquires **Company B**. Company B's employees need to use this Teams app, but they sit in a different identity/tenant context — either as guests inside Company A's tenant, or still fully inside Company B's own, separate tenant. This section works through both cases, for all three upload approaches, because the right answer genuinely differs by scenario.

### Scenario 1: Company B users are Guests inside Company A's tenant

**The one real blocker here isn't our code — it's Teams admin policy.** Microsoft Teams applies separate app permission policies to Guest-type users than to full members, and many tenants restrict custom/line-of-business apps for guests by default. This needs Company A's Teams admin to explicitly allow the app for guests (Teams Admin Center → Teams apps → Permission policies). Nothing in our code needs to change for this.

&#91;embedded content: Scenario 1 — guest access runs on our own service credentials, gated once by Teams admin policy\]

**Why our approach handles guests cleanly once that's allowed**: uploads (Blob or SharePoint), processing, tracking, search, and proactive notifications all authenticate using **our own service's credentials** (a managed identity or app registration) — never the end user's own permissions or license. A guest's Teams user ID is just as valid a key for tracking/notification as a full member's. So once a guest can open the app at all, everything downstream works identically to a full member.

**One nuance, SharePoint mode only**: if the final download link is a direct SharePoint `webUrl`, the guest clicking it is now using *their own* browser session, not our service's credentials — they need ordinary read access granted on that one SharePoint site. Blob mode has no equivalent step at all (a SAS link is self-authorizing).

### Scenario 2: Company B installs the Teams app package into their own, separate tenant

&#91;embedded content: Scenario 2 — Blob mode crosses the tenant boundary cleanly; Graph-based paperclip retrieval hits a hard consent wall\]

This is a bigger architectural question than Scenario 1, because here there is **no existing trust relationship at all** between the two tenants.

**The bot and the tab still work with no code changes needed.** Teams bot messaging routes through Microsoft's own Bot Framework Connector — a global, Microsoft-hosted layer that doesn't care which tenant installed the app, only that the `botId` resolves to a real, reachable messaging endpoint. The tab is just a public HTTPS page loaded in an iframe — any tenant's Teams client can load it. Because our manifest declares no Teams SSO (`webApplicationInfo`), there's no per-tenant token-exchange dependency either. The only requirement is that **Company B's own Teams admin** allows custom apps for their users — a policy decision on their side, not ours.

**Blob storage mode: fully tenant-agnostic, no extra steps.** Upload, processing, and the download link are all handled by our own server-issued credentials — completely unaffected by which organization installed the app.

**SharePoint storage mode: processing still works, but the final link needs real setup.** The upload/processing path is still server-issued and unaffected. But the final SharePoint `webUrl` now points a Company B user at a site in a tenant they have **zero prior relationship with** — SharePoint's **external sharing** must be enabled on that site/tenant, and that specific user explicitly granted access. This is a bigger lift than Scenario 1's nuance, because there's no existing guest relationship to build on.

**Practical guidance for this scenario**: default to Blob mode for any cross-tenant handover like this. It sidesteps the external-sharing question entirely. (The dedicated "Recommendation" section near the end of this document covers the broader, project-wide choice between upload approaches — this note is specific to the cross-tenant case.)

### What happens if we'd built this on Adaptive Card + paperclip instead — Scenario 2 specifically

This is where the paperclip approach stops being merely less convenient and starts being **architecturally broken**, not just worse UX.

**The upload step still works** — Teams uploads the attached file to the user's own OneDrive, in their own tenant, regardless of which tenant our bot is registered in.

**The retrieval step is where it breaks**, and there are only two ways to attempt it:

- **Fetch via Microsoft Graph, using our own app's permissions — a hard, architectural dead end.** Graph application permissions are consented **per tenant**. Our app has admin consent only inside Company A's tenant; it has no registration or consent relationship with Company B's tenant at all. There is no configuration fix for this — an app consented in one tenant cannot call a different tenant's Graph resources without that second tenant's own separate admin consent. This isn't the network-level "firewall" issue from before; it's a step earlier — the request is never valid in the first place.
- **Fetch via the pre-signed `downloadUrl`** Teams includes in the attachment payload, skipping Graph entirely. This might work, since the link doesn't depend on our app's own tenant consent — but it's fragile in two ways: it can still be blocked by the same kind of SharePoint/OneDrive network-location restriction described earlier, and these links are typically **short-lived** (commonly expiring within hours). Any delay before processing — a queue backlog, a slow AI run — risks the link simply expiring before we ever fetch it, for reasons that have nothing to do with the AI pipeline itself.

**Our direct-to-storage approach has no equivalent failure mode at all**: the file is already sitting safely in our own storage account the moment it's uploaded, with zero dependency on the uploading user's tenant and no link-expiry race to lose.

### Summary across both scenarios

|  | Scenario 1: Guest in Company A's tenant | Scenario 2: Installed in Company B's own tenant |
| --- | --- | --- |
| Bot + Tab | Works — needs Company A's Teams admin to allow guests to use custom apps | Works — needs Company B's own Teams admin to allow custom apps |
| Blob storage mode | Fully works, no extra steps | Fully works, no extra steps |
| SharePoint storage mode | Works — guest needs read access on the one site for the final link | Works for processing — final link needs external sharing enabled + an explicit grant |
| Adaptive Card + Paperclip | Upload works; Graph-based retrieval is blocked by SharePoint network restrictions even with a correct app registration | Upload works; Graph-based retrieval is a hard architectural dead end (cross-tenant consent); `downloadUrl` retrieval is possible but fragile (expiry + the same network-restriction risk) |

## Recommendation

**Keep the direct-to-storage approach** this project already uses, rather than switching to Adaptive Card + paperclip, for any request that involves real business documents (financial models, customer data, anything with a retention/compliance requirement). The extra engineering cost — SAS tokens, CORS, a hosted tab — is a one-time build cost. The governance gap in the paperclip approach (documents scattered across individual employees' personal OneDrive accounts, with no central ownership or retention control) is an ongoing, compounding risk that gets worse with every additional user who submits a request.

The Adaptive Card + paperclip pattern is a reasonable choice for genuinely low-stakes, throwaway attachments (a screenshot for a support ticket, say) where nobody will ever need to centrally audit or reliably re-fetch the file later. It is not a good fit for the documents this system handles.

## The Complete Picture: One Pipeline, No Timeouts, Gated End to End

Everything above comes together into one simple shape. This is the diagram to put in front of anyone — technical or not — who asks "so how does this actually work, and is it safe?"

&#91;embedded content: The complete pipeline, instant for the user and gated end to end\]

**Why there's no timeout risk, ever.** The user's request is acknowledged in milliseconds — the green path above. The slow part (reading two documents, running Document Intelligence, generating a draft with Azure OpenAI) happens entirely in the background, kicked off by a queue message rather than by the user's own request waiting on hold. A slow AI run, a busy day, a transient retry — none of it can ever time out a user's click, because nothing the user does is ever connected, end to end, to the slow part. The live tracker and the Teams notification are how the user finds out it's done, not how they wait for it.

**Why it's secure by design, not as an afterthought.** Every hop in this diagram is gated or scoped, never open: APIM checks a subscription key before anything reaches the backend; the backend never holds a human's password or a long-lived secret — it authenticates to Azure itself using a managed identity; and the one thing end users ever touch directly — the upload link — is a short-lived, purpose-built token that expires and can't be reused, never a standing door into the storage account. The same shape holds whether the final document lands in Blob storage or in SharePoint.

**A real fix already applied, as an example of this in practice.** A review of the tab's own HTML/JS (`frontend.py`, shared byte-for-byte by both projects) found one genuine gap: the Customer Name field — free text the user types in the form — was being written back onto the page through `innerHTML` without escaping, in the live-tracking panel and in search results. A crafted customer name could have run as script in the browser of anyone who later viewed that request or found it in a search. Both repositories have been patched with an `esc()` helper applied at every point user-supplied text is written into the page, and the fix is live on GitHub in both `teams-bot-final-product` and `otc-team-code`.

## Debugging Method: "Which Step Broke?"

When something's wrong, don't look at the whole codebase — work out which of the 15 steps (in the teams-bot-final-product section above) failed, then go straight to the one function responsible for it.

| Symptom | Which step | Where to look |
| --- | --- | --- |
| Tab shows a blank page or error | Step 1 | `frontend.py` / the route serving `/api/form` |
| "Could not prepare upload" error | Step 3 | `issue_upload_sas` / `issue_sharepoint_upload_session` in `agent.py` |
| "Upload failed (HTTP ...)" | Step 4 | Check CORS (Blob) or the `Content-Range` header (SharePoint) |
| Submit button spins forever / errors | Step 5–8 | `/api/submit` route + `create_tracking_record` + `enqueue_request` |
| Status stuck on "Request Created" forever | Step 9 | The queue trigger never fired — check queue binding settings, or just wait (cold start can take \~30–60s) |
| Status flips to "Failed" | Step 10–12 | Read the `errorMessage` field in the tracking row — it's the real exception, not a guess |
| Status says "Completed" but no Teams message | Step 14 | `send_proactive_card` — usually a missing/stale conversation reference |
| Everything works but the SharePoint list is empty | The independent list mirror | `_sharepoint_configured()` — check all 4 SharePoint settings are set |

**The habit to build**: always ask "which numbered step is this?" first. Once you know the step, you know the exact function — and from there it's just reading that one function's code and the real error message it logs, not guessing across the whole project.
