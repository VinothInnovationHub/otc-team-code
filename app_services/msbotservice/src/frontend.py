"""The "New OTC Request" full-page Teams tab — presentation only.

Progress lives in a side panel (not a top stepper), there's a search-by-ID box,
and status is polled from /api/status after submit — because real processing
goes through a queue -> generate_otc (Document Intelligence + Azure OpenAI),
which takes real time, unlike a stub that replies in milliseconds. Nothing
here talks to Document Intelligence/OpenAI/SharePoint/Blob directly; it only
calls this app's own /api/sas, /api/submit, /api/status routes (see
routes.py), all of which delegate to agent.py.

Served as-is from GET /api/form (wired in routes.py / start_server.py); there
is no separate .html file on disk. This file is hosting-agnostic — it is pure
HTML/CSS/JS returned as a string, unchanged whether served by Azure Functions
or, as here, an aiohttp route on App Service.
"""

PAGE_HTML = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>New OTC Request</title>
<script src="https://res.cdn.office.net/teams-js/2.19.0/js/MicrosoftTeams.min.js"></script>
<style>
  :root {
    --brand: #5b5fc7; --brand-dark: #464775; --brand-light: #eef0fc;
    --ink: #242424; --ink-soft: #616161; --line: #e1dfdd;
    --good: #0f7b0f; --good-bg: #eaf7ea;
    --bad: #a4262c; --bad-bg: #fdf2f3;
    --warn: #986f0b; --warn-bg: #fff8e5;
    --bg: #f7f7fb;
  }
  * { box-sizing: border-box; }
  body { font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, sans-serif; margin: 0; color: var(--ink); background: var(--bg); }

  .topbar { display: flex; align-items: center; gap: 12px; padding: 16px 24px; background: #fff; border-bottom: 1px solid var(--line); flex-wrap: wrap; }
  .topbar h1 { font-size: 17px; font-weight: 700; margin: 0; flex: none; }
  .search-box { flex: 1; min-width: 220px; max-width: 360px; display: flex; gap: 6px; margin-left: auto; }
  .search-box input { flex: 1; padding: 8px 12px; border: 1px solid var(--line); border-radius: 8px; font-size: 13px; }
  .search-box button { padding: 8px 14px; border: none; border-radius: 8px; background: var(--brand); color: #fff; font-weight: 600; font-size: 13px; cursor: pointer; }
  .search-box button:hover { background: var(--brand-dark); }
  .new-btn { padding: 8px 14px; border: 1px solid var(--line); border-radius: 8px; background: #fff; color: var(--ink); font-weight: 600; font-size: 13px; cursor: pointer; }
  .new-btn:hover { border-color: var(--brand); color: var(--brand-dark); }

  .layout { display: grid; grid-template-columns: 240px 1fr; gap: 0; min-height: calc(100vh - 57px); }
  .side-panel { background: #fff; border-right: 1px solid var(--line); padding: 24px 18px; }
  .side-panel .sp-title { font-size: 11px; font-weight: 700; letter-spacing: .06em; text-transform: uppercase; color: var(--ink-soft); margin-bottom: 10px; }
  .sp-meta { display: flex; justify-content: space-between; font-size: 11.5px; color: var(--ink-soft); font-weight: 600; margin-bottom: 8px; min-height: 15px; }
  .sp-meta .meta-percent { color: var(--brand-dark); }
  .progress-bar-track { height: 4px; background: var(--line); border-radius: 999px; margin-bottom: 22px; overflow: hidden; }
  .progress-bar-fill { height: 100%; width: 0%; border-radius: 999px; background: linear-gradient(90deg, var(--brand), var(--brand-dark)); transition: width .5s ease; }
  .progress-bar-fill.failed { background: var(--bad); }
  .track-step { display: flex; gap: 10px; padding-bottom: 22px; position: relative; }
  .track-step:not(:last-child)::before { content: ""; position: absolute; left: 11px; top: 24px; bottom: -2px; width: 2px; background: var(--line); }
  .track-step.done:not(:last-child)::before { background: var(--good); }
  .track-dot { width: 24px; height: 24px; border-radius: 50%; background: #fff; border: 2px solid var(--line); flex: none; display: flex; align-items: center; justify-content: center; font-size: 11px; font-weight: 700; color: var(--ink-soft); z-index: 1; position: relative; }
  .track-step.done .track-dot { background: var(--good); border-color: var(--good); color: #fff; }
  .track-step.active .track-dot { border-color: var(--brand); color: var(--brand); box-shadow: 0 0 0 4px var(--brand-light); }
  .track-step.active .track-dot::after { content: ""; position: absolute; top: -3px; left: -3px; right: -3px; bottom: -3px; border-radius: 50%; border: 2px solid transparent; border-top-color: var(--brand); animation: spin 0.9s linear infinite; }
  .track-step.failed .track-dot { background: var(--bad); border-color: var(--bad); color: #fff; }
  .track-label { font-size: 12.5px; font-weight: 600; color: var(--ink-soft); padding-top: 3px; }
  .track-step.done .track-label, .track-step.active .track-label { color: var(--ink); }
  .track-step.failed .track-label { color: var(--bad); }

  .main { padding: 28px 32px; max-width: 720px; }
  .panel { background: #fff; border: 1px solid var(--line); border-radius: 12px; padding: 24px; position: relative; overflow: hidden; transition: opacity .28s ease, transform .28s ease; }
  .panel.panel-enter { opacity: 0; transform: translateY(10px); }
  .panel-head h2 { font-size: 16px; margin: 0 0 4px; }
  .panel-head p { margin: 0; font-size: 12.5px; color: var(--ink-soft); }

  .welcome-panel { text-align: center; padding: 56px 40px; }
  .welcome-badge { width: 72px; height: 72px; border-radius: 20px; background: linear-gradient(135deg, var(--brand), var(--brand-dark)); color: #fff; font-size: 32px; display: flex; align-items: center; justify-content: center; margin: 0 auto 22px; box-shadow: 0 8px 24px -8px rgba(91,95,199,.55); }
  .welcome-title { font-size: 24px; font-weight: 800; margin: 0 0 8px; letter-spacing: -.01em; }
  .welcome-sub { font-size: 14px; color: var(--ink-soft); margin: 0 auto 28px; max-width: 420px; line-height: 1.5; }
  .welcome-features { display: flex; justify-content: center; gap: 10px; flex-wrap: wrap; margin-bottom: 32px; }
  .welcome-feature { display: flex; align-items: center; gap: 7px; font-size: 12.5px; font-weight: 600; color: var(--brand-dark); background: var(--brand-light); border-radius: 999px; padding: 8px 14px; }
  .welcome-cta { max-width: 320px; margin: 0 auto; padding: 13px 20px; font-size: 14.5px; box-shadow: 0 6px 18px -6px rgba(91,95,199,.5); }

  .section-title { font-size: 13px; font-weight: 700; margin: 20px 0 10px; }
  .section-title:first-of-type { margin-top: 0; }
  label { display: block; font-size: 12.5px; font-weight: 600; margin: 14px 0 6px; }
  input[type="text"], select { width: 100%; padding: 10px 12px; font-size: 14px; border: 1px solid var(--line); border-radius: 8px; background: #fff; color: var(--ink); font-family: inherit; }
  input[type="text"]:focus, select:focus { outline: none; border-color: var(--brand); box-shadow: 0 0 0 3px var(--brand-light); }
  .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 0 14px; }
  .req { color: var(--bad); }

  .dropzone { border: 1.5px dashed var(--line); border-radius: 10px; padding: 18px; margin-top: 6px; text-align: center; cursor: pointer; }
  .dropzone:hover, .dropzone.drag { border-color: var(--brand); background: var(--brand-light); }
  .dropzone .dz-label { font-size: 13px; font-weight: 600; color: var(--brand-dark); }
  .dropzone .dz-hint { font-size: 11.5px; color: var(--ink-soft); margin-top: 4px; }
  .dropzone input[type="file"] { display: none; }

  .file-chip { display: none; align-items: center; gap: 12px; border: 1px solid var(--good); background: var(--good-bg); border-radius: 10px; padding: 10px 12px; margin-top: 6px; }
  .file-chip.show { display: flex; }
  .file-chip .fc-thumb { width: 40px; height: 52px; border-radius: 5px; background: #fff; border: 1px solid var(--line); flex: none; position: relative; overflow: hidden; box-shadow: 0 2px 6px -2px rgba(0,0,0,.12); }
  .file-chip .fc-thumb::before { content: ""; position: absolute; top: 0; left: 0; right: 0; height: 13px; background: var(--good); }
  .file-chip .fc-thumb .fc-lines { position: absolute; top: 19px; left: 6px; right: 6px; }
  .file-chip .fc-thumb .fc-lines span { display: block; height: 2.5px; background: var(--line); border-radius: 2px; margin-bottom: 4px; }
  .file-chip .fc-thumb .fc-lines span:nth-child(2) { width: 75%; }
  .file-chip .fc-thumb .fc-lines span:nth-child(3) { width: 55%; }
  .file-chip .fc-thumb.fc-thumb-grid .fc-lines { display: grid; grid-template-columns: repeat(3, 1fr); gap: 2px; top: 18px; }
  .file-chip .fc-thumb.fc-thumb-grid .fc-lines span { height: 7px; margin-bottom: 2px; border-radius: 1px; }
  .file-chip .fc-thumb .fc-badge { position: absolute; bottom: 4px; left: 4px; right: 4px; font-size: 7px; font-weight: 800; letter-spacing: .03em; text-align: center; color: #fff; background: var(--good); border-radius: 2px; padding: 1px 0; }
  .file-chip .fc-main { flex: 1; min-width: 0; }
  .file-chip .fc-name { font-size: 12.5px; font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .file-chip .fc-size { font-size: 11.5px; color: var(--good); font-weight: 700; }
  .file-chip .fc-remove { background: none; border: none; color: var(--ink-soft); font-size: 16px; cursor: pointer; width: 22px; height: 22px; border-radius: 50%; flex: none; }
  .file-chip .fc-remove:hover { background: #fff; color: var(--bad); }
  .file-error { font-size: 11.5px; color: var(--bad); margin-top: 6px; display: none; }

  .btn { font-size: 14px; font-weight: 700; padding: 11px 20px; border-radius: 8px; cursor: pointer; border: none; width: 100%; margin-top: 20px; }
  .btn-primary { background: var(--brand); color: #fff; }
  .btn-primary:hover { background: var(--brand-dark); }
  .btn-primary:disabled { background: #c8c6c4; cursor: not-allowed; }
  .btn-good { background: var(--good); color: #fff; text-align: center; display: block; text-decoration: none; }
  .form-error { display: none; background: var(--bad-bg); color: var(--bad); border: 1px solid #f3c6c9; border-radius: 8px; padding: 10px 12px; font-size: 13px; margin-top: 14px; }

  .status-badge { display: inline-flex; align-items: center; gap: 6px; padding: 5px 12px; border-radius: 999px; font-size: 12px; font-weight: 700; }
  .badge-spinner { width: 11px; height: 11px; border-radius: 50%; border: 2px solid currentColor; border-right-color: transparent; animation: spin 0.7s linear infinite; flex: none; }
  .status-badge.Request-Created, .status-badge.In-Review { background: var(--brand-light); color: var(--brand-dark); }
  .status-badge.AI-Processing { background: var(--warn-bg); color: var(--warn); }
  .status-badge.Document-Generated, .status-badge.Completed { background: var(--good-bg); color: var(--good); }
  .status-badge.Failed { background: var(--bad-bg); color: var(--bad); }

  .success-check { display: block; margin: 4px auto 14px; }
  .success-check-circle { stroke: var(--good); stroke-width: 3; fill: none; stroke-dasharray: 145; stroke-dashoffset: 145; }
  .success-check-mark { stroke: var(--good); stroke-width: 4; fill: none; stroke-linecap: round; stroke-linejoin: round; stroke-dasharray: 30; stroke-dashoffset: 30; }
  .success-check.playing .success-check-circle { animation: draw-circle .5s ease forwards; }
  .success-check.playing .success-check-mark { animation: draw-check .35s ease forwards .45s; }
  @keyframes draw-circle { to { stroke-dashoffset: 0; } }
  @keyframes draw-check { to { stroke-dashoffset: 0; } }

  .confetti-burst { position: absolute; inset: 0; pointer-events: none; overflow: visible; z-index: 5; }
  .confetti-piece { position: absolute; top: 30px; width: 7px; height: 7px; border-radius: 2px; opacity: 0; animation: confetti-fall 1.1s ease-out forwards; }
  @keyframes confetti-fall { 0% { opacity: 1; transform: translate(0,0) rotate(0deg); } 100% { opacity: 0; transform: translate(var(--dx), 170px) rotate(var(--dr)); } }

  .review-row { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid var(--line); font-size: 13.5px; }
  .review-row:last-child { border-bottom: none; }
  .review-row .k { color: var(--ink-soft); }
  .review-row .v { font-weight: 600; text-align: right; }

  .search-results { display: flex; flex-direction: column; gap: 8px; margin-top: 16px; }
  .search-result { border: 1px solid var(--line); border-radius: 10px; padding: 12px 14px; cursor: pointer; display: flex; justify-content: space-between; align-items: center; gap: 10px; }
  .search-result:hover { border-color: var(--brand); background: var(--brand-light); }
  .search-result .sr-id { font-size: 12px; font-family: Consolas, monospace; color: var(--ink-soft); }
  .search-result .sr-name { font-weight: 600; font-size: 13.5px; }

  .hidden { display: none !important; }
  .spin { animation: spin 0.9s linear infinite; }
  @keyframes spin { to { transform: rotate(360deg); } }
</style>
</head>
<body>

<div class="topbar">
  <h1>\U0001F4CB OTC Assistant</h1>
  <div class="search-box">
    <input type="text" id="searchInput" placeholder="Search by Request ID or customer name">
    <button id="searchBtn">Search</button>
  </div>
  <button class="new-btn" id="newRequestBtn">+ New Request</button>
</div>

<div class="layout">
  <div class="side-panel" id="sidePanel">
    <div class="sp-title">Request Progress</div>
    <div class="sp-meta" id="spMeta"><span class="meta-percent" id="metaPercent"></span><span id="metaElapsed"></span></div>
    <div class="progress-bar-track"><div class="progress-bar-fill" id="progressBarFill"></div></div>
    <div id="trackSteps">
      <div class="track-step" data-stage="Request Created"><div class="track-dot">1</div><div class="track-label">Request Created</div></div>
      <div class="track-step" data-stage="In Review"><div class="track-dot">2</div><div class="track-label">In Review</div></div>
      <div class="track-step" data-stage="AI Processing"><div class="track-dot">3</div><div class="track-label">AI Processing</div></div>
      <div class="track-step" data-stage="Document Generated"><div class="track-dot">4</div><div class="track-label">Document Generated</div></div>
      <div class="track-step" data-stage="Completed"><div class="track-dot">5</div><div class="track-label">Completed</div></div>
    </div>
  </div>

  <div class="main">

    <div class="panel welcome-panel" id="welcomePanel">
      <div class="welcome-badge">📋</div>
      <h2 class="welcome-title">OTC Assistant</h2>
      <p class="welcome-sub">AI-drafted Offer-to-Close documents, tracked live from request to completion.</p>
      <div class="welcome-features">
        <div class="welcome-feature"><span>📄</span>AI document drafting</div>
        <div class="welcome-feature"><span>📊</span>Live status tracking</div>
        <div class="welcome-feature"><span>🔍</span>Instant search by ID</div>
      </div>
      <button class="btn btn-primary welcome-cta" id="welcomeStartBtn">+ Start New OTC Request</button>
    </div>

    <div class="panel hidden" id="formPanel">
      <div class="panel-head">
        <h2>Create OTC Request</h2>
        <p>Provide customer details and upload the required documents.</p>
      </div>

      <p class="section-title">\U0001F464 Customer Information</p>
      <label for="customerName">Customer Name <span class="req">*</span></label>
      <input type="text" id="customerName" placeholder="Enter customer name">

      <div class="grid-2">
        <div>
          <label for="requestType">Request Type <span class="req">*</span></label>
          <select id="requestType">
            <option value="">Select request type</option>
            <option value="newDeal">New Deal</option>
            <option value="amendment">Amendment</option>
            <option value="renewal">Renewal</option>
          </select>
        </div>
        <div>
          <label for="facilityType">Facility Type <span class="req">*</span></label>
          <select id="facilityType">
            <option value="">Select facility type</option>
            <option value="termLoan">Term Loan</option>
            <option value="revolvingCredit">Revolving Credit Facility</option>
            <option value="overdraft">Overdraft</option>
          </select>
        </div>
      </div>

      <div class="grid-2">
        <div>
          <label for="region">Region</label>
          <select id="region">
            <option value="">Select region</option>
            <option value="UK">UK</option>
            <option value="Europe">Europe</option>
            <option value="Americas">Americas</option>
            <option value="APAC">Asia Pacific</option>
          </select>
        </div>
        <div>
          <label for="country">Country</label>
          <select id="country">
            <option value="">Select country</option>
            <option value="GB">United Kingdom</option>
            <option value="US">United States</option>
            <option value="DE">Germany</option>
            <option value="FR">France</option>
          </select>
        </div>
      </div>

      <p class="section-title">\U0001F4C4 Supporting Documents</p>
      <p style="margin:-4px 0 0;font-size:12.5px;color:var(--ink-soft);">PDF, Word or Excel only.</p>

      <label>Investment Memo (IM) <span class="req">*</span></label>
      <div class="dropzone" id="imZone">
        <div class="dz-label">\U0001F4E4 Click to upload or drag and drop</div>
        <div class="dz-hint">PDF or Word — .pdf, .doc, .docx</div>
        <input type="file" id="imFile" accept=".pdf,.doc,.docx">
      </div>
      <div class="file-chip" id="imChip">
        <div class="fc-thumb">
          <div class="fc-lines"><span></span><span></span><span></span></div>
          <div class="fc-badge" id="imBadge">DOC</div>
        </div>
        <div class="fc-main"><div class="fc-name" id="imName">—</div><div class="fc-size" id="imSize">—</div></div>
        <button type="button" class="fc-remove" id="imRemove" title="Remove">×</button>
      </div>
      <div class="file-error" id="imError"></div>

      <label>Financial Model Document <span class="req">*</span></label>
      <div class="dropzone" id="fmZone">
        <div class="dz-label">\U0001F4E4 Click to upload or drag and drop</div>
        <div class="dz-hint">Excel only — .xls, .xlsx</div>
        <input type="file" id="fmFile" accept=".xls,.xlsx">
      </div>
      <div class="file-chip" id="fmChip">
        <div class="fc-thumb fc-thumb-grid">
          <div class="fc-lines"><span></span><span></span><span></span><span></span><span></span><span></span></div>
          <div class="fc-badge" id="fmBadge">XLS</div>
        </div>
        <div class="fc-main"><div class="fc-name" id="fmName">—</div><div class="fc-size" id="fmSize">—</div></div>
        <button type="button" class="fc-remove" id="fmRemove" title="Remove">×</button>
      </div>
      <div class="file-error" id="fmError"></div>

      <div class="form-error" id="formError"></div>
      <button class="btn btn-primary" id="submitBtn">Submit Request</button>
    </div>

    <div class="panel hidden" id="trackingPanel">
      <svg class="success-check hidden" id="successCheck" viewBox="0 0 52 52" width="56" height="56">
        <circle class="success-check-circle" cx="26" cy="26" r="23"/>
        <path class="success-check-mark" d="M14 27l7 7 16-16"/>
      </svg>
      <div class="panel-head">
        <h2 id="trackingTitle">Tracking Request</h2>
        <p id="trackingSub"></p>
      </div>
      <div style="margin:14px 0;"><span class="status-badge" id="statusBadge">—</span></div>
      <div id="trackingReview"></div>
      <a class="btn btn-good hidden" id="trackingDownload" href="#" target="_blank">⬇️ Download Draft OTC</a>
      <p id="trackingNote" style="font-size:12.5px;color:var(--ink-soft);margin-top:14px;"></p>
    </div>

    <div class="panel hidden" id="searchPanel">
      <div class="panel-head">
        <h2>Search Results</h2>
        <p id="searchSub"></p>
      </div>
      <div class="search-results" id="searchResults"></div>
    </div>

  </div>
</div>

<script>
(function () {
  var STAGES = ["Request Created", "In Review", "AI Processing", "Document Generated", "Completed"];
  var ALLOWED_IM = [".pdf", ".doc", ".docx"];
  var ALLOWED_FM = [".xls", ".xlsx"];
  var fileMeta = { im: null, fm: null };
  var pollTimer = null;

  function apiBase() { return window.location.origin + "/api"; }
  function show(id) { document.getElementById(id).classList.remove("hidden"); }
  function hide(id) { document.getElementById(id).classList.add("hidden"); }

  // Every value interpolated into innerHTML below passes through this first -
  // customerName etc. are free-text user input round-tripped from storage,
  // so without escaping, a crafted name (e.g. "<img src=x onerror=...>")
  // would execute as script for anyone who later views or searches for it.
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }

  function showPanel(name) {
    ["welcomePanel", "formPanel", "trackingPanel", "searchPanel"].forEach(function (p) { hide(p); });
    var el = document.getElementById(name);
    el.classList.add("panel-enter");
    show(name);
    // Double rAF: lets the browser paint the opacity:0 starting state before
    // removing the class, so the transition actually runs instead of being
    // batched into the same frame as the display:none -> block flip.
    requestAnimationFrame(function () {
      requestAnimationFrame(function () { el.classList.remove("panel-enter"); });
    });
  }

  function fmtSize(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return Math.round(bytes / 1024) + " KB";
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
  }

  function extOf(name) {
    var i = name.lastIndexOf(".");
    return i === -1 ? "" : name.slice(i).toLowerCase();
  }

  function wireDropzone(zoneId, fileInputId, key, chipPrefix, allowed) {
    var zone = document.getElementById(zoneId);
    var input = document.getElementById(fileInputId);
    var chip = document.getElementById(chipPrefix + "Chip");
    var nameEl = document.getElementById(chipPrefix + "Name");
    var sizeEl = document.getElementById(chipPrefix + "Size");
    var badgeEl = document.getElementById(chipPrefix + "Badge");
    var errorEl = document.getElementById(chipPrefix + "Error");
    var removeBtn = document.getElementById(chipPrefix + "Remove");

    function setFile(file) {
      if (!file) return;
      if (allowed.indexOf(extOf(file.name)) === -1) {
        errorEl.textContent = "Unsupported file type. Allowed: " + allowed.join(", ");
        errorEl.style.display = "block";
        return;
      }
      errorEl.style.display = "none";
      fileMeta[key] = file;
      nameEl.textContent = file.name;
      sizeEl.textContent = fmtSize(file.size);
      badgeEl.textContent = extOf(file.name).replace(".", "").toUpperCase();
      chip.classList.add("show");
      zone.classList.add("hidden");
    }

    zone.addEventListener("click", function () { input.click(); });
    input.addEventListener("change", function (e) { setFile(e.target.files[0]); });
    zone.addEventListener("dragover", function (e) { e.preventDefault(); zone.classList.add("drag"); });
    zone.addEventListener("dragleave", function () { zone.classList.remove("drag"); });
    zone.addEventListener("drop", function (e) {
      e.preventDefault();
      zone.classList.remove("drag");
      if (e.dataTransfer.files[0]) setFile(e.dataTransfer.files[0]);
    });
    removeBtn.addEventListener("click", function () {
      fileMeta[key] = null;
      input.value = "";
      chip.classList.remove("show");
      zone.classList.remove("hidden");
      errorEl.style.display = "none";
    });
  }
  wireDropzone("imZone", "imFile", "im", "im", ALLOWED_IM);
  wireDropzone("fmZone", "fmFile", "fm", "fm", ALLOWED_FM);

  function fields() {
    return {
      customerName: document.getElementById("customerName").value.trim(),
      requestType: document.getElementById("requestType").value,
      facilityType: document.getElementById("facilityType").value,
      region: document.getElementById("region").value,
      country: document.getElementById("country").value,
    };
  }

  function setTrackSteps(status, failed) {
    var idx = STAGES.indexOf(status);
    document.querySelectorAll(".track-step").forEach(function (el) {
      var s = el.getAttribute("data-stage");
      var sIdx = STAGES.indexOf(s);
      el.classList.remove("done", "active", "failed");
      if (failed && sIdx === idx) { el.classList.add("failed"); }
      else if (sIdx < idx || (sIdx === idx && status === "Completed")) { el.classList.add("done"); }
      else if (sIdx === idx) { el.classList.add("active"); }
    });

    var percent = idx === -1 ? 0 : Math.round(((idx + 1) / STAGES.length) * 100);
    var fill = document.getElementById("progressBarFill");
    fill.style.width = percent + "%";
    fill.classList.toggle("failed", !!failed);
    document.getElementById("metaPercent").textContent = idx === -1 ? "" : percent + "% complete";
  }

  var elapsedTimer = null;
  function stopElapsedTimer() {
    if (elapsedTimer) { clearInterval(elapsedTimer); elapsedTimer = null; }
  }
  function fmtElapsed(totalSecs) {
    var m = Math.floor(totalSecs / 60), s = totalSecs % 60;
    return (m > 0 ? m + "m " : "") + s + "s";
  }
  // Ticks live while a request is in flight; freezes at the real elapsed
  // total (updatedAt - createdAt) once it reaches a terminal status, rather
  // than continuing to count up after the work is actually done.
  function updateElapsed(createdAtIso, terminal, updatedAtIso) {
    stopElapsedTimer();
    var startMs = new Date(createdAtIso).getTime();
    var elapsedEl = document.getElementById("metaElapsed");
    function render() {
      var endMs = terminal ? new Date(updatedAtIso).getTime() : Date.now();
      elapsedEl.textContent = fmtElapsed(Math.max(0, Math.floor((endMs - startMs) / 1000)));
    }
    render();
    if (!terminal) { elapsedTimer = setInterval(render, 1000); }
  }

  var celebratedRequestId = null;
  function celebrateCompletion() {
    var check = document.getElementById("successCheck");
    check.classList.remove("hidden", "playing");
    void check.offsetWidth; // force reflow so the draw-in animation replays
    check.classList.add("playing");

    var panel = document.getElementById("trackingPanel");
    var burst = document.createElement("div");
    burst.className = "confetti-burst";
    var colors = ["#5b5fc7", "#0f7b0f", "#986f0b", "#464775", "#2b88d8"];
    for (var i = 0; i < 22; i++) {
      var piece = document.createElement("span");
      piece.className = "confetti-piece";
      piece.style.left = (42 + Math.random() * 16) + "%";
      piece.style.background = colors[i % colors.length];
      piece.style.setProperty("--dx", Math.round(Math.random() * 220 - 110) + "px");
      piece.style.setProperty("--dr", Math.round(Math.random() * 520 - 260) + "deg");
      piece.style.animationDelay = (Math.random() * 0.15) + "s";
      burst.appendChild(piece);
    }
    panel.appendChild(burst);
    setTimeout(function () { burst.remove(); }, 1600);
  }

  function renderTrackingReview(record) {
    var rows = [
      ["Customer Name", record.customerName || "—"],
      ["Request Type", record.requestType || "—"],
      ["Facility Type", record.facilityType || "—"],
      ["Requested By", record.requestedBy || "—"],
      ["Request ID", record.requestId || record.RowKey || "—"],
      ["Created", record.createdAt || "—"],
      ["Last Updated", record.updatedAt || "—"],
    ];
    document.getElementById("trackingReview").innerHTML = rows.map(function (r) {
      return '<div class="review-row"><span class="k">' + esc(r[0]) + '</span><span class="v">' + esc(r[1]) + "</span></div>";
    }).join("");
  }

  function applyTracking(record) {
    var status = record.status || "Request Created";
    document.getElementById("trackingTitle").textContent = record.customerName ? ("Request for " + record.customerName) : "Tracking Request";
    document.getElementById("trackingSub").textContent = (record.requestId || record.RowKey || "");
    var badge = document.getElementById("statusBadge");
    var inProgress = ["In Review", "AI Processing", "Document Generated"].indexOf(status) !== -1;
    badge.innerHTML = (inProgress ? '<span class="badge-spinner"></span>' : "") + status;
    badge.className = "status-badge " + status.replace(/ /g, "-");
    setTrackSteps(status, status === "Failed");
    renderTrackingReview(record);

    var terminal = status === "Completed" || status === "Failed";
    if (record.createdAt) { updateElapsed(record.createdAt, terminal, record.updatedAt); }

    var dl = document.getElementById("trackingDownload");
    var note = document.getElementById("trackingNote");
    if (status === "Completed" && record.downloadUrl) {
      dl.href = record.downloadUrl;
      dl.classList.remove("hidden");
      note.textContent = "";
      var thisId = record.requestId || record.RowKey;
      if (celebratedRequestId !== thisId) {
        celebratedRequestId = thisId;
        celebrateCompletion();
      }
    } else if (status === "Failed") {
      dl.classList.add("hidden");
      note.textContent = "Something went wrong: " + (record.errorMessage || "unknown error") + ". Contact support with the Request ID above.";
      document.getElementById("successCheck").classList.add("hidden");
    } else {
      dl.classList.add("hidden");
      note.textContent = "This usually takes a few minutes. You'll also get a message in this chat once it's ready.";
      document.getElementById("successCheck").classList.add("hidden");
    }
  }

  function pollStatus(requestId) {
    if (pollTimer) clearInterval(pollTimer);
    function tick() {
      fetch(apiBase() + "/status?id=" + encodeURIComponent(requestId))
        .then(function (r) { return r.json(); })
        .then(function (data) {
          if (!data.record) return;
          applyTracking(data.record);
          if (data.record.status === "Completed" || data.record.status === "Failed") {
            clearInterval(pollTimer);
          }
        })
        .catch(function () {});
    }
    tick();
    pollTimer = setInterval(tick, 4000);
  }

  function resetForm() {
    document.getElementById("customerName").value = "";
    ["requestType", "facilityType", "region", "country"].forEach(function (id) { document.getElementById(id).value = ""; });
    ["im", "fm"].forEach(function (p) {
      fileMeta[p] = null;
      document.getElementById(p + "Chip").classList.remove("show");
      document.getElementById(p + "Zone").classList.remove("hidden");
      document.getElementById(p + "Error").style.display = "none";
    });
    document.getElementById("formError").style.display = "none";
    setTrackSteps("", false);
    stopElapsedTimer();
    document.getElementById("metaElapsed").textContent = "";
    document.getElementById("successCheck").classList.add("hidden");
  }

  function startNewRequest() {
    if (pollTimer) clearInterval(pollTimer);
    resetForm();
    showPanel("formPanel");
  }
  document.getElementById("newRequestBtn").addEventListener("click", startNewRequest);
  document.getElementById("welcomeStartBtn").addEventListener("click", startNewRequest);

  document.getElementById("submitBtn").addEventListener("click", function () {
    var errorEl = document.getElementById("formError");
    var f = fields();
    var missing = [];
    if (!f.customerName) missing.push("Customer Name");
    if (!f.requestType) missing.push("Request Type");
    if (!f.facilityType) missing.push("Facility Type");
    if (!fileMeta.im) missing.push("Investment Memo");
    if (!fileMeta.fm) missing.push("Financial Model Document");
    if (missing.length) {
      errorEl.textContent = "Please provide: " + missing.join(", ") + ".";
      errorEl.style.display = "block";
      return;
    }
    errorEl.style.display = "none";
    document.getElementById("submitBtn").disabled = true;
    document.getElementById("submitBtn").textContent = "Submitting…";

    var sessionId = (window.crypto && window.crypto.randomUUID) ? window.crypto.randomUUID() : String(Date.now());

    function uploadOne(file, label) {
      return fetch(apiBase() + "/sas?sessionId=" + encodeURIComponent(sessionId) + "&fileName=" + encodeURIComponent(file.name))
        .then(function (r) { return r.json(); })
        .then(function (sas) {
          if (!sas.sasUrl) { throw new Error(sas.error || "Could not prepare upload for " + file.name); }
          var headers = { "Content-Type": file.type || "application/octet-stream" };
          if (sas.storageMode === "sharepoint") {
            // Graph's upload session rejects a single-shot PUT with HTTP 400
            // unless Content-Range is set, even when the whole file fits in
            // one request (confirmed live, not assumed) - a plain Blob SAS
            // PUT doesn't use this header at all, so it's sent only here.
            headers["Content-Range"] = "bytes 0-" + (file.size - 1) + "/" + file.size;
          } else {
            headers["x-ms-blob-type"] = "BlockBlob";
          }
          return fetch(sas.sasUrl, { method: "PUT", headers: headers, body: file })
            .then(function (uploadRes) {
              if (!uploadRes.ok) { throw new Error("Upload failed for " + file.name + " (HTTP " + uploadRes.status + ")"); }
              return { name: file.name, path: sas.blobPath, label: label };
            });
        });
    }

    Promise.all([uploadOne(fileMeta.im, "investment_memo"), uploadOne(fileMeta.fm, "financial_model")])
      .then(function (uploaded) {
        var payload = Object.assign({}, f, {
          requestedBy: (window.__requesterName || ""),
          teamsUserId: (window.__teamsUserId || ""),
          documents: {
            investment_memo: { document_type: "investment_memo", name: uploaded[0].name, path: uploaded[0].path },
            financial_model: { document_type: "financial_model", name: uploaded[1].name, path: uploaded[1].path },
          },
        });
        return fetch(apiBase() + "/submit", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
      })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        document.getElementById("submitBtn").disabled = false;
        document.getElementById("submitBtn").textContent = "Submit Request";
        if (!data.requestId) { throw new Error(data.error || "No request id returned"); }
        showPanel("trackingPanel");
        applyTracking({ requestId: data.requestId, status: "Request Created", customerName: f.customerName, requestType: f.requestType, facilityType: f.facilityType });
        pollStatus(data.requestId);
      })
      .catch(function (err) {
        document.getElementById("submitBtn").disabled = false;
        document.getElementById("submitBtn").textContent = "Submit Request";
        errorEl.textContent = "Could not submit: " + err.message + ". Please try again.";
        errorEl.style.display = "block";
      });
  });

  function runSearch() {
    var q = document.getElementById("searchInput").value.trim();
    if (!q) return;
    fetch(apiBase() + "/search?q=" + encodeURIComponent(q))
      .then(function (r) { return r.json(); })
      .then(function (data) {
        var results = data.results || [];
        document.getElementById("searchSub").textContent = results.length + " result(s) for \\"" + q + "\\"";
        document.getElementById("searchResults").innerHTML = results.map(function (r) {
          var id = r.requestId || r.RowKey;
          return '<div class="search-result" data-id="' + esc(id) + '">' +
            '<div><div class="sr-name">' + esc(r.customerName || "—") + '</div><div class="sr-id">' + esc(id) + '</div></div>' +
            '<span class="status-badge ' + esc((r.status || "").replace(/ /g, "-")) + '">' + esc(r.status || "—") + '</span></div>';
        }).join("") || '<p style="font-size:13px;color:var(--ink-soft);">No matching requests.</p>';
        document.querySelectorAll(".search-result").forEach(function (el) {
          el.addEventListener("click", function () {
            var id = el.getAttribute("data-id");
            showPanel("trackingPanel");
            pollStatus(id);
          });
        });
        showPanel("searchPanel");
      })
      .catch(function () {});
  }
  document.getElementById("searchBtn").addEventListener("click", runSearch);
  document.getElementById("searchInput").addEventListener("keydown", function (e) { if (e.key === "Enter") runSearch(); });

  window.__requesterName = "";
  window.__teamsUserId = "";
  if (window.microsoftTeams) {
    microsoftTeams.app.initialize()
      .then(function () { return microsoftTeams.app.getContext(); })
      .then(function (context) {
        var name = (context && context.user && context.user.displayName) || "";
        var uid = (context && context.user && context.user.id) || "";
        window.__teamsUserId = uid;
        if (name) window.__requesterName = name;
      })
      .catch(function () {});
  }
})();
</script>
</body>
</html>
"""


def render_page() -> str:
    return PAGE_HTML
