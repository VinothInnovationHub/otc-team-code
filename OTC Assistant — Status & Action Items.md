Status Tracker · OTC Assistant

# Status & action items — quick update

| Owner | Area | Status | Outcome | Next action |
| --- | --- | --- | --- | --- |
| Subham | Network scenarios — all 3 architectures tried | Testing complete | All 3 blocked at the network layer. Card+paperclip: outbound to SharePoint blocked. WebView/Tab: inbound blocked, DNS doesn't resolve off GVPN. Hosted agent: upload works, download blocked. Same root cause across all three. | Raise the 3 recommended fixes (allow SharePoint :443 outbound, direct Blob SAS upload/download, DNS resolve for Tab path) with Network + Security for sign-off. |
| Shruti | Code / PR | Pushed, awaiting test | PR approved and merged. | Run functional testing against the pushed build; report pass/fail. |
| Ashish / Mieshen (TAV) | Approval | Pending | Checked with Ryan — approval not yet received. | Follow up with Ryan directly on sign-off status; escalate if no response by next check-in. |