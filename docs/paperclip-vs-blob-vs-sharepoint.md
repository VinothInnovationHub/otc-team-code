# Paperclip vs Blob vs SharePoint — Checkpoint Comparison

Nov 11, 2026 · @Vinothkumar

## Executive Summary

Three ways exist to get the two source documents (Investment Memo, Financial Model) into storage: **Adaptive Card + paperclip** (lands in the uploading employee's own personal OneDrive), **direct-to-Blob** (a short-lived upload token, lands in an organization-owned storage account), and **direct-to-SharePoint** (a Graph upload session, lands in an organization-owned site). Checkpoint by checkpoint — Graph API permission scope, what happens if the uploading employee leaves, link/token expiry, network exposure, and cross-tenant sharing — the paperclip approach has one root-cause weakness the other two don't: it ties a business document's custody to one individual's personal account instead of the organization's own infrastructure. Every problem in Table 1 traces back to that single fact.

**Recommendation**: default to **Blob storage** — zero Graph API dependency, smallest permission footprint. Use **SharePoint** specifically when the business wants native Microsoft 365 governance tooling (retention labels, DLP, version history) on top. Rule out **paperclip + personal OneDrive** for any document with a retention, audit, or compliance requirement — which an Investment Memo and a Financial Model clearly are.

## Side-by-Side Comparison

| Checkpoint | Paperclip + Personal OneDrive | Blob Storage (SAS) | SharePoint (Sites.Selected) |
| --- | --- | --- | --- |
| Where the file lands | The uploading employee's own personal OneDrive | An organization-owned Blob container | An organization-owned SharePoint site/library |
| Graph API permission scope | `Files.Read.All` — tenant-wide read over **every** user's OneDrive and SharePoint. No per-user scoping exists for a personal drive | None — Blob isn't part of Graph at all; access is Azure RBAC (managed identity), scoped to one storage account | `Sites.Selected` — explicitly granted to **exactly one site**, nothing else in the tenant |
| Uploading user leaves the company | Source document can become permanently unreachable once their OneDrive is suspended/deleted under offboarding policy | No impact — never tied to any individual's identity | No impact — the site belongs to the organization, not the contributor |
| Guest user sharing | Guests often have no provisioned OneDrive in the host tenant at all, so the upload itself can fail or behave inconsistently; even when it works, retrieval still needs `Files.Read.All` against an account that may not persist | Non-issue — upload/download run on the bot's own service identity, completely independent of the requester's guest/member status | Processing is unaffected (service identity, `Sites.Selected`); the final `webUrl` link needs the guest explicitly granted read access on that one site — a one-time, scoped grant |
| Attachment / file-size handling | Governed by Teams' and OneDrive's own client-side attachment behavior — outside this app's control, and can behave inconsistently across desktop/web/mobile Teams clients | Governed entirely by this app's own upload code against Blob directly — consistent across every platform, no meaningful practical limit for this document type | Handled via Microsoft Graph's resumable upload session, explicitly chunked by the app — consistent, but needs deliberate handling (e.g. a `Content-Range` header) that Blob doesn't require |
| Download link / token lifetime | Pre-signed `downloadUrl` is short-lived (commonly hours) — can expire before a delayed background job ever fetches it | Only the SAS token for one operation is short-lived (\~30 min); the file itself persists indefinitely and a fresh SAS can be minted anytime | Re-authenticated Graph call each time (no fixed expiry), or a `webUrl` gated by the viewer's own site permission |
| Network / conditional-access exposure | Subject to tenant-wide SharePoint/OneDrive network-restriction policy — can block the backend even with a correct app registration | Storage account's own firewall/private-endpoint rules, set directly by IT on that one resource | Same tenant-wide network-restriction exposure as paperclip, but with a far smaller permission footprint |
| Auditability / governance | Scattered across however many employees have ever submitted a request, under each person's own settings | One storage account, one access policy, standard diagnostic logging | Native SharePoint versioning, retention labels, DLP, eDiscovery |
| Cross-tenant sharing (separate org) | Architecturally blocked — Graph app-only consent is per-tenant, with no standing in another tenant | Fully tenant-agnostic — nothing in the path is tied to any tenant but the app's own | Processing works cross-tenant; the final link needs External Sharing enabled plus an explicit per-user grant |
| Security point of view | Broad, tenant-wide permission just to read 2 files; the document sits entirely outside IT's security perimeter, under personal retention/DLP settings | Smallest footprint of the three — no Graph dependency, standard Azure RBAC, file never leaves org-controlled infrastructure | Scoped to one site (meaningfully smaller blast radius than `Files.Read.All`), plus native DLP/retention — larger footprint than pure Blob, much smaller than paperclip |
| Setup cost | Low — no SAS tokens, no CORS, no hosted tab | Moderate — SAS issuance endpoint, CORS config, a hosted tab (one-time cost, already paid here) | Moderate-to-higher — site/list provisioning, `Sites.Selected` grant, Graph upload-session handling |

## Recommendation — Best-Practice Projection

Checkpoint by checkpoint, **paperclip + personal OneDrive is the only one of the three with a structural flaw**: it ties a business document's custody, reachability, and security boundary to one employee's personal account rather than the organization's own infrastructure. Every problem in Table 1 — the `Files.Read.All` permission scope, the employee-leaves failure mode, the `downloadUrl` expiry race, the cross-tenant dead end — traces back to that one root cause, and none of it is fixable by configuration. It is a property of the approach, not a bug to patch.

Tables 2 and 3 share the same organizational-custody foundation; they differ only in surface area and tooling:

1. **Default to Blob storage** for the general case. It has zero Graph API dependency, the smallest permission footprint, and the least operational surface area of the three.
2. **Choose SharePoint** specifically when the business wants native SharePoint/Microsoft 365 governance tooling (retention labels, DLP, co-authoring) on top of the document — and is willing to accept the modest extra setup (`Sites.Selected` provisioning, Graph upload-session handling) and the one network-policy dependency it still shares with the paperclip approach.
3. **Rule out paperclip + personal OneDrive** for any document type carrying a retention, compliance, or audit requirement — which, for an Investment Memo or a Financial Model, this clearly is.

This project already implements option 1 (Blob) and option 2 (SharePoint) as an interchangeable toggle, with identical business logic either way — so adopting this recommendation costs nothing further to build.
