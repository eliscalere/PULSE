/**
 * PULSE Tickets — v1.0.0
 * Standalone issue-tracker companion to PULSE.
 * Reads from the "PULSE Issues" SharePoint list.
 */
(function () {
  "use strict";

  const LIST_TITLE  = "PULSE Issues";
  const STATUSES    = ["New", "Triaged", "In Progress", "Resolved", "Closed"];
  const STATUS_TABS = ["All", ...STATUSES];
  const ISSUE_TYPES = [
    "Bug / error", "Data / saving", "Access / permissions",
    "Display / usability", "Feature Request", "Other"
  ];

  let tickets = [];
  const state = { status: "All", type: "All", search: "" };
  const bootStartedAt = performance.now();

  /* ── Helpers ── */
  function escHtml(v) {
    const n = document.createElement("div");
    n.textContent = v == null ? "" : String(v);
    return n.innerHTML;
  }
  function fmtDate(v) {
    if (!v) return "—";
    try { return new Date(v).toLocaleDateString([], { year: "numeric", month: "short", day: "numeric" }); }
    catch (_) { return String(v).slice(0, 10); }
  }
  function fmtDateTime(v) {
    if (!v) return "—";
    try { return new Date(v).toLocaleString([], { year: "numeric", month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }); }
    catch (_) { return String(v); }
  }
  function $(sel, root) { return (root || document).querySelector(sel); }
  function $all(sel, root) { return Array.from((root || document).querySelectorAll(sel)); }

  function safeJson(v) {
    if (!v) return null;
    if (typeof v === "object") return v;
    try { return JSON.parse(v); } catch (_) { return null; }
  }

  /* ── Normalise SP item ── */
  function normalise(item) {
    const logs      = safeJson(item.LogsJson)        || [];
    const diag      = safeJson(item.DiagnosticsJson) || {};
    const errCodes  = safeJson(item.ErrorCodesJson)  || [];
    const ssUrl     = item.ScreenshotFileUrl || item.ScreenshotServerRelativeUrl || item.ScreenshotDataUrl || "";
    return Object.assign({}, item, {
      Status:        item.IssueStatus    || item.Status      || "New",
      ReportedBy:    item.ReportedByName || item.ReportedBy  || "Unknown",
      Type:          item.IssueType      || item.Type        || "Other",
      Description:   item.Description   || item.Summary      || "",
      logs,
      diag,
      errCodes,
      screenshotSrc: ssUrl
    });
  }

  /* ── Status / type display ── */
  function statusPill(s) {
    const cls = "sp--" + String(s || "New").toLowerCase().replace(/[\s/]+/g, "-");
    return `<span class="status-pill ${cls}">${escHtml(s || "New")}</span>`;
  }
  function typeBadge(t) {
    return `<span class="type-badge">${escHtml(t || "Other")}</span>`;
  }

  /* ── Filtering ── */
  function filtered() {
    const q = state.search.toLowerCase();
    return tickets.filter(t => {
      if (state.status !== "All" && t.Status !== state.status) return false;
      if (state.type   !== "All" && t.Type   !== state.type)   return false;
      if (q && !String(t.Title      || "").toLowerCase().includes(q) &&
               !String(t.IssueCode  || "").toLowerCase().includes(q) &&
               !String(t.ReportedBy || "").toLowerCase().includes(q)) return false;
      return true;
    });
  }

  /* ── SP helpers ── */
  function getSiteUrl() {
    if (window._spPageContextInfo && window._spPageContextInfo.webAbsoluteUrl)
      return window._spPageContextInfo.webAbsoluteUrl.replace(/\/$/, "");
    // APP_CONFIG is a const in app-config.js — access it without window.
    try {
      if (typeof APP_CONFIG !== "undefined" && APP_CONFIG.manualSharePointSiteUrl)
        return APP_CONFIG.manualSharePointSiteUrl.replace(/\/$/, "");
    } catch (_) {}
    // Fallback: reconstruct from _spPageContextInfo.webServerRelativeUrl (same
    // chain the main PULSE app uses when webAbsoluteUrl is missing).
    if (window._spPageContextInfo && window._spPageContextInfo.webServerRelativeUrl)
      return (window.location.origin + window._spPageContextInfo.webServerRelativeUrl).replace(/\/$/, "");
    return window.location.origin;
  }
  function odata(v) { return String(v).replace(/'/g, "''"); }
  function spFetch(url, opts) {
    return fetch(url, Object.assign({ credentials: "same-origin" }, opts, {
      headers: Object.assign({ Accept: "application/json;odata=nometadata", "Cache-Control": "no-cache" }, opts && opts.headers)
    }));
  }
  async function getDigest() {
    const r = await spFetch(`${getSiteUrl()}/_api/contextinfo`, { method: "POST" });
    if (!r.ok) throw new Error(`Digest error (${r.status})`);
    return (await r.json()).FormDigestValue;
  }
  async function spPatch(id, body) {
    const digest = await getDigest();
    const r = await fetch(`${getSiteUrl()}/_api/web/lists/getbytitle('${odata(LIST_TITLE)}')/items(${id})`, {
      method: "POST", credentials: "same-origin",
      headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json;odata=nometadata", "X-RequestDigest": digest, "IF-MATCH": "*", "X-HTTP-Method": "MERGE" },
      body: JSON.stringify(body)
    });
    if (!r.ok) throw new Error(`Update failed (${r.status})`);
  }
  async function spDel(id) {
    const digest = await getDigest();
    const r = await fetch(`${getSiteUrl()}/_api/web/lists/getbytitle('${odata(LIST_TITLE)}')/items(${id})`, {
      method: "POST", credentials: "same-origin",
      headers: { "X-RequestDigest": digest, "IF-MATCH": "*", "X-HTTP-Method": "DELETE" }
    });
    if (!r.ok && r.status !== 404) throw new Error(`Delete failed (${r.status})`);
  }
  async function spCreate(body) {
    const digest = await getDigest();
    const r = await fetch(`${getSiteUrl()}/_api/web/lists/getbytitle('${odata(LIST_TITLE)}')/items`, {
      method: "POST", credentials: "same-origin",
      headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json;odata=nometadata", "X-RequestDigest": digest },
      body: JSON.stringify(body)
    });
    if (!r.ok) { const txt = await r.text().catch(() => ""); throw new Error(`Create failed (${r.status}): ${txt.slice(0, 200)}`); }
    return r.json();
  }

  /* ── Toast ── */
  function toast(msg, type) {
    const stack = $("#tkt-toast-stack") || document.body;
    const el = document.createElement("div");
    el.className = `tkt-toast${type ? " tkt-toast--" + type : ""}`;
    el.textContent = msg;
    stack.appendChild(el);
    setTimeout(() => { el.style.opacity = "0"; setTimeout(() => el.remove(), 300); }, 3000);
  }

  /* ── Loader ── */
  function dismissLoader() {
    const el = document.getElementById("pulse-boot-loader");
    if (!el) return;
    const wait = Math.max(0, 1000 - (performance.now() - bootStartedAt));
    setTimeout(() => {
      el.animate([{ opacity: 1, transform: "scale(1)" }, { opacity: 0, transform: "scale(.985)" }],
        { duration: 360, easing: "cubic-bezier(.4,0,.2,1)", fill: "forwards" })
        .onfinish = () => el.remove();
    }, wait);
  }

  /* ── Hash routing ── */
  function currentIssueId() {
    const hash = location.hash.replace(/^#\/?/, "");
    return hash && /^ISS-/i.test(hash) ? hash.toUpperCase() : null;
  }
  function navigateTo(id) {
    location.hash = id ? "#/" + id : "#/";
  }

  /* ── Render dispatcher ── */
  function render() {
    const id = currentIssueId();
    if (id) {
      const t = tickets.find(x => String(x.IssueCode || "").toUpperCase() === id);
      if (t) { renderDetail(t); return; }
    }
    renderList();
  }

  /* ══════════════════════════════════════════════
     LIST VIEW — matches PULSE app tickets page
     ══════════════════════════════════════════════ */
  function renderList() {
    const rows = filtered();
    const counts = STATUS_TABS.reduce((acc, s) => {
      acc[s] = s === "All" ? tickets.length : tickets.filter(t => t.Status === s).length;
      return acc;
    }, {});

    const view = $("#tkt-view");
    view.innerHTML = `
      <div class="tkt-list-page">
        <section class="tkt-controls">
          <div class="tkt-search-wrap">
            <svg class="tkt-search-icon" viewBox="0 0 16 16" fill="none" aria-hidden="true"><circle cx="6.5" cy="6.5" r="4.75" stroke="currentColor" stroke-width="1.5"/><path d="M10.5 10.5L14 14" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>
            <input type="search" class="tkt-search" id="tkt-search" placeholder="Search by ID or title" value="${escHtml(state.search)}" autocomplete="off">
          </div>
          <div class="filter-pills" id="tkt-status-tabs">
            ${STATUS_TABS.map(s => `<button class="filter-pill${state.status === s ? " active" : ""}" data-status="${escHtml(s)}">${escHtml(s)}<span>${counts[s]}</span></button>`).join("")}
          </div>
          <select class="tkt-type-filter select-aewttr" id="tkt-type-filter">
            <option value="All">All Types</option>
            ${ISSUE_TYPES.map(t => `<option value="${escHtml(t)}"${state.type === t ? " selected" : ""}>${escHtml(t)}</option>`).join("")}
          </select>
        </section>

        <section class="tkt-table-shell">
          <div class="tkt-table-meta"><strong>${rows.length}</strong> of ${tickets.length} issue${tickets.length === 1 ? "" : "s"}</div>
          <div class="tkt-table-scroll">
            <table class="tkt-table">
              <thead><tr><th>Issue</th><th>Title</th><th>Type</th><th>Status</th><th>Page</th><th>Reported By</th><th>Date</th></tr></thead>
              <tbody>
                ${rows.length ? rows.map(t => `
                  <tr data-code="${escHtml(t.IssueCode || "")}" style="cursor:pointer;">
                    <td><span class="tkt-id">${escHtml(t.IssueCode || "ISSUE-" + t.Id)}</span></td>
                    <td><strong>${escHtml(t.Title || "Untitled")}</strong></td>
                    <td>${typeBadge(t.Type)}</td>
                    <td>${statusPill(t.Status)}</td>
                    <td class="tkt-page-cell">${escHtml(t.PageTitle || "—")}</td>
                    <td>${escHtml(t.ReportedBy || "—")}</td>
                    <td class="tkt-date-cell">${fmtDate(t.OccurredAt || t.Created)}</td>
                  </tr>`).join("") : `
                  <tr><td colspan="7"><div class="tkt-empty"><strong>No issues found</strong><span>Adjust filters or create a new issue.</span></div></td></tr>`}
              </tbody>
            </table>
          </div>
        </section>
      </div>`;

    $all("tr[data-code]", view).forEach(r => r.addEventListener("click", () => navigateTo(r.dataset.code)));
    $all(".filter-pill", view).forEach(b => b.addEventListener("click", () => { state.status = b.dataset.status; render(); }));
    $("#tkt-type-filter", view).addEventListener("change", e => { state.type = e.target.value; render(); });
    $("#tkt-search", view).addEventListener("input", e => { state.search = e.target.value; render(); });
  }

  /* ══════════════════════════════════════════════
     DETAIL VIEW — tabbed full-page view
     ══════════════════════════════════════════════ */
  function renderDetail(t) {
    const view = $("#tkt-view");
    const ssrc = t.screenshotSrc || "";
    const logs  = Array.isArray(t.logs) ? t.logs : [];
    const diag  = t.diag || {};
    const errs  = Array.isArray(t.errCodes) ? t.errCodes : [];
    const appLogs = Array.isArray(diag.sharePointDebug) ? diag.sharePointDebug : [];

    const errorLogs = logs.filter(l => l.level === "error" || l.level === "warn");
    const infoLogs  = logs.filter(l => l.level !== "error" && l.level !== "warn");

    function logRow(l) {
      return `<div class="tkt-log-item tkt-log--${escHtml(l.level || "info")}">
        <span class="tkt-log-level">${escHtml(l.level || "log")}</span>
        <span class="tkt-log-msg">${escHtml(l.message || "")}</span>
        <span class="tkt-log-ts">${l.ts ? fmtDateTime(l.ts) : ""}</span>
      </div>`;
    }
    function appLogRow(l) {
      return `<div class="tkt-log-item tkt-log--${escHtml(l.level || "info")}">
        <span class="tkt-log-level">${escHtml(l.level || "log")}</span>
        <span class="tkt-log-msg">${escHtml((l.category ? "[" + l.category + "] " : "") + (l.message || ""))}</span>
        <span class="tkt-log-ts">${l.ts ? fmtDateTime(l.ts) : ""}</span>
      </div>`;
    }

    // Tab definitions — only show tabs that have content
    const TABS = [
      { id: "overview",   label: "Overview" },
      { id: "screenshot", label: "Screenshot", hide: !ssrc },
      { id: "browser-logs", label: `Browser Logs${logs.length ? " (" + logs.length + ")" : ""}`, hide: !logs.length },
      { id: "app-logs",   label: `App Logs${appLogs.length ? " (" + appLogs.length + ")" : ""}`, hide: !appLogs.length },
      { id: "error-codes", label: `Error Codes${errs.length ? " (" + errs.length + ")" : ""}`, hide: !errs.length && !diag.browser },
    ].filter(t => !t.hide);

    const activeTab = TABS[0].id;

    view.innerHTML = `
      <div class="tkt-detail-page">
        <div class="tkt-detail-topbar">
          <button class="tkt-back-btn" id="tktd-back">&#8592; All issues</button>
          <div class="tkt-detail-topbar-actions">
            ${t.Status !== "Resolved" && t.Status !== "Closed" ? `<button class="tkt-btn-outline" id="tktd-resolve">Resolve</button>` : ""}
            <button class="tkt-btn-danger" id="tktd-delete">Delete</button>
          </div>
        </div>

        <div class="tkt-detail-layout">
          <!-- Main: tabs + panels -->
          <div class="tkt-detail-main">
            <div class="tkt-detail-header">
              <span class="tkt-id">${escHtml(t.IssueCode || "ISSUE-" + t.Id)}</span>
              <h1 class="tkt-detail-title">${escHtml(t.Title || "Untitled")}</h1>
            </div>

            <div class="tkt-tabs" id="tktd-tabs">
              ${TABS.map(tab => `<button class="tkt-tab${tab.id === activeTab ? " active" : ""}" data-tab="${tab.id}">${escHtml(tab.label)}</button>`).join("")}
            </div>

            <!-- Overview -->
            <div class="tkt-tab-panel${activeTab === "overview" ? " active" : ""}" data-panel="overview">
              <div class="tkt-detail-section">
                <div class="tkt-section-label">Description</div>
                <p class="tkt-detail-text">${escHtml(t.Description || "No description provided.")}</p>
              </div>
              ${t.ExpectedBehavior ? `
              <div class="tkt-detail-section">
                <div class="tkt-section-label">Expected result</div>
                <p class="tkt-detail-text">${escHtml(t.ExpectedBehavior)}</p>
              </div>` : ""}
              ${t.AdditionalContext ? `
              <div class="tkt-detail-section">
                <div class="tkt-section-label">Additional context / updates</div>
                <p class="tkt-detail-text">${escHtml(t.AdditionalContext)}</p>
              </div>` : ""}
              ${t.ResolutionNote ? `
              <div class="tkt-detail-section">
                <div class="tkt-section-label">Resolution note</div>
                <p class="tkt-detail-text" style="color:var(--aewttr-green)">${escHtml(t.ResolutionNote)}</p>
              </div>` : ""}
              <div class="tkt-detail-section tkt-update-section">
                <div class="tkt-section-label">Add update</div>
                <textarea class="tkt-textarea" id="tktd-update-text" rows="3" placeholder="Add a note, resolution step, or follow-up…"></textarea>
                <div class="tkt-update-actions">
                  <button class="tkt-btn-primary" id="tktd-post-btn">Post update</button>
                </div>
                <div class="tkt-form-msg" id="tktd-msg"></div>
              </div>
            </div>

            <!-- Screenshot -->
            ${ssrc ? `
            <div class="tkt-tab-panel" data-panel="screenshot">
              <div class="tkt-detail-section">
                <div class="tkt-section-label">Captured screenshot</div>
                <div class="tkt-screenshot-wrap">
                  <img src="${escHtml(ssrc)}" alt="Captured screenshot" class="tkt-screenshot-img" id="tktd-screenshot">
                  <span class="tkt-screenshot-hint">Click to enlarge</span>
                </div>
                ${t.ScreenshotFileName ? `<p class="tkt-screenshot-filename">${escHtml(t.ScreenshotFileName)}</p>` : ""}
              </div>
            </div>` : ""}

            <!-- Browser Logs -->
            ${logs.length ? `
            <div class="tkt-tab-panel" data-panel="browser-logs">
              ${errorLogs.length ? `
              <div class="tkt-detail-section">
                <div class="tkt-section-label">Errors &amp; warnings <span class="tkt-badge tkt-badge--red">${errorLogs.length}</span></div>
                <div class="tkt-log-list">${errorLogs.map(logRow).join("")}</div>
              </div>` : ""}
              ${infoLogs.length ? `
              <div class="tkt-detail-section">
                <div class="tkt-section-label">Info / debug <span class="tkt-badge">${infoLogs.length}</span></div>
                <div class="tkt-log-list">${infoLogs.map(logRow).join("")}</div>
              </div>` : ""}
            </div>` : ""}

            <!-- App Logs -->
            ${appLogs.length ? `
            <div class="tkt-tab-panel" data-panel="app-logs">
              <div class="tkt-detail-section">
                <div class="tkt-section-label">PULSE app debug log <span class="tkt-badge">${appLogs.length}</span></div>
                <div class="tkt-log-list">${appLogs.map(appLogRow).join("")}</div>
              </div>
            </div>` : ""}

            <!-- Error Codes + Diagnostics -->
            <div class="tkt-tab-panel" data-panel="error-codes">
              ${errs.length ? `
              <div class="tkt-detail-section">
                <div class="tkt-section-label">Error codes <span class="tkt-badge tkt-badge--red">${errs.length}</span></div>
                <div class="tkt-error-codes">${errs.map(e => `<span class="tkt-error-chip">${escHtml(e)}</span>`).join("")}</div>
              </div>` : `<div class="tkt-detail-section"><p class="tkt-detail-text" style="color:var(--aewttr-muted)">No error codes captured.</p></div>`}
              ${diag.browser || diag.viewport || diag.capturedAt ? `
              <div class="tkt-detail-section">
                <div class="tkt-section-label">Client diagnostics</div>
                <div class="tkt-diag-grid">
                  ${diag.browser       ? `<div class="tkt-diag-kv"><span>Browser</span><span>${escHtml(String(diag.browser).slice(0, 120))}</span></div>` : ""}
                  ${diag.viewport      ? `<div class="tkt-diag-kv"><span>Viewport</span><span>${escHtml(diag.viewport.width + "×" + diag.viewport.height + " @" + (diag.viewport.devicePixelRatio || 1) + "x")}</span></div>` : ""}
                  ${diag.mode         ? `<div class="tkt-diag-kv"><span>App mode</span><span>${escHtml(diag.mode)}</span></div>` : ""}
                  ${diag.capturedAt   ? `<div class="tkt-diag-kv"><span>Captured</span><span>${escHtml(fmtDateTime(diag.capturedAt))}</span></div>` : ""}
                  ${diag.detectedSiteUrl ? `<div class="tkt-diag-kv"><span>Site URL</span><span>${escHtml(diag.detectedSiteUrl)}</span></div>` : ""}
                  ${diag.hasPageContext != null ? `<div class="tkt-diag-kv"><span>SP context</span><span>${diag.hasPageContext ? "Present" : "Missing"}</span></div>` : ""}
                </div>
              </div>` : ""}
              ${diag.page ? `
              <div class="tkt-detail-section">
                <div class="tkt-section-label">Page context</div>
                <div class="tkt-diag-grid">
                  ${diag.page.title  ? `<div class="tkt-diag-kv"><span>Title</span><span>${escHtml(diag.page.title)}</span></div>` : ""}
                  ${diag.page.route  ? `<div class="tkt-diag-kv"><span>Route</span><code>${escHtml(diag.page.route)}</code></div>` : ""}
                  ${diag.page.url    ? `<div class="tkt-diag-kv"><span>URL</span><span style="word-break:break-all">${escHtml(String(diag.page.url).slice(0, 200))}</span></div>` : ""}
                </div>
              </div>` : ""}
            </div>
          </div>

          <!-- Sidebar -->
          <aside class="tkt-detail-aside">
            <div class="tkt-meta-card">
              <div class="tkt-meta-row">
                <span class="tkt-meta-label">Status</span>
                <select class="tkt-meta-select select-aewttr" id="tktd-status">
                  ${STATUSES.map(s => `<option${t.Status === s ? " selected" : ""}>${escHtml(s)}</option>`).join("")}
                </select>
              </div>
              <div class="tkt-meta-row">
                <span class="tkt-meta-label">Type</span>
                <span>${typeBadge(t.Type)}</span>
              </div>
              <div class="tkt-meta-row">
                <span class="tkt-meta-label">Reporter</span>
                <span class="tkt-meta-value">${escHtml(t.ReportedBy || "—")}</span>
              </div>
              ${t.ReporterEmail ? `
              <div class="tkt-meta-row">
                <span class="tkt-meta-label">Email</span>
                <a href="mailto:${escHtml(t.ReporterEmail)}" class="tkt-meta-link">${escHtml(t.ReporterEmail)}</a>
              </div>` : ""}
              <div class="tkt-meta-row">
                <span class="tkt-meta-label">Reported</span>
                <span class="tkt-meta-value">${fmtDateTime(t.OccurredAt || t.Created)}</span>
              </div>
              ${t.PageTitle ? `
              <div class="tkt-meta-row">
                <span class="tkt-meta-label">Page</span>
                <span class="tkt-meta-value">${escHtml(t.PageTitle)}</span>
              </div>` : ""}
              ${t.Route ? `
              <div class="tkt-meta-row">
                <span class="tkt-meta-label">Route</span>
                <code class="tkt-meta-code">${escHtml(t.Route)}</code>
              </div>` : ""}
              ${t.PageUrl ? `
              <div class="tkt-meta-row">
                <span class="tkt-meta-label">URL</span>
                <a href="${escHtml(t.PageUrl)}" target="_blank" rel="noopener" class="tkt-meta-link" style="font-size:11px;word-break:break-all;">${escHtml(String(t.PageUrl).replace(/^https?:\/\/[^/]+/, "").slice(0, 60))}…</a>
              </div>` : ""}
            </div>

            ${ssrc ? `
            <div class="tkt-meta-thumb-card" id="tktd-thumb-card">
              <div class="tkt-section-label" style="padding:12px 14px 8px;margin:0;">Screenshot</div>
              <img src="${escHtml(ssrc)}" alt="Screenshot thumbnail" class="tkt-meta-thumb" id="tktd-thumb">
            </div>` : ""}

            ${errs.length || logs.filter(l => l.level === "error").length ? `
            <div class="tkt-meta-card tkt-meta-card--warn">
              <div class="tkt-meta-row" style="border:0;padding:10px 14px;">
                <span class="tkt-meta-label">Errors captured</span>
                <span class="tkt-meta-value" style="color:var(--aewttr-red)">${errs.length} code${errs.length !== 1 ? "s" : ""}, ${errorLogs.length} log${errorLogs.length !== 1 ? "s" : ""}</span>
              </div>
            </div>` : ""}
          </aside>
        </div>
      </div>`;

    /* ── Tab switching ── */
    $all(".tkt-tab", view).forEach(btn => btn.addEventListener("click", () => {
      $all(".tkt-tab", view).forEach(b => b.classList.remove("active"));
      $all(".tkt-tab-panel", view).forEach(p => p.classList.remove("active"));
      btn.classList.add("active");
      const panel = $(`[data-panel="${btn.dataset.tab}"]`, view);
      if (panel) panel.classList.add("active");
    }));

    /* ── Back ── */
    $("#tktd-back", view).addEventListener("click", () => navigateTo(null));

    /* ── Status change ── */
    $("#tktd-status", view).addEventListener("change", async (e) => {
      const newStatus = e.target.value;
      if (newStatus === t.Status) return;
      try {
        await spPatch(t.Id, { IssueStatus: newStatus });
        t.IssueStatus = newStatus; t.Status = newStatus;
        toast(`Status updated to ${newStatus}`, "success");
        renderDetail(t);
      } catch (err) { toast(err.message, "error"); e.target.value = t.Status; }
    });

    /* ── Resolve ── */
    const resolveBtn = $("#tktd-resolve", view);
    if (resolveBtn) resolveBtn.addEventListener("click", async () => {
      resolveBtn.disabled = true; resolveBtn.textContent = "Resolving…";
      try {
        await spPatch(t.Id, { IssueStatus: "Resolved" });
        t.IssueStatus = "Resolved"; t.Status = "Resolved";
        toast(`${t.IssueCode || "Issue"} resolved`, "success");
        renderDetail(t);
      } catch (err) { resolveBtn.disabled = false; resolveBtn.textContent = "Resolve"; toast(err.message, "error"); }
    });

    /* ── Post update ── */
    const postBtn = $("#tktd-post-btn", view);
    if (postBtn) postBtn.addEventListener("click", async () => {
      const text = ($("#tktd-update-text", view) || {}).value && $("#tktd-update-text", view).value.trim();
      if (!text) return;
      const btn = postBtn;
      const msg = $("#tktd-msg", view);
      btn.disabled = true; btn.textContent = "Posting…"; if (msg) msg.textContent = "";
      const stamp = new Date().toLocaleDateString([], { month: "short", day: "numeric", year: "numeric" });
      const who = (window._spPageContextInfo && window._spPageContextInfo.userDisplayName) || "Admin";
      const appended = (t.AdditionalContext ? t.AdditionalContext + "\n\n" : "") + `[${stamp} — ${who}] ${text}`;
      try {
        await spPatch(t.Id, { AdditionalContext: appended });
        t.AdditionalContext = appended;
        toast("Update posted", "success");
        renderDetail(t);
      } catch (err) {
        btn.disabled = false; btn.textContent = "Post update";
        if (msg) { msg.className = "tkt-form-msg tkt-form-msg--error"; msg.textContent = err.message; }
      }
    });

    /* ── Delete ── */
    const deleteBtn = $("#tktd-delete", view);
    deleteBtn.addEventListener("click", async () => {
      if (deleteBtn.dataset.confirmed !== "1") {
        deleteBtn.dataset.confirmed = "1"; deleteBtn.textContent = "Confirm delete";
        setTimeout(() => { if (deleteBtn.dataset.confirmed === "1") { deleteBtn.dataset.confirmed = ""; deleteBtn.textContent = "Delete"; } }, 3000);
        return;
      }
      deleteBtn.disabled = true; deleteBtn.textContent = "Deleting…";
      try {
        await spDel(t.Id);
        tickets = tickets.filter(x => x.Id !== t.Id);
        toast(`${t.IssueCode || "Issue"} deleted`, "success");
        navigateTo(null);
      } catch (err) { deleteBtn.disabled = false; deleteBtn.textContent = "Delete"; deleteBtn.dataset.confirmed = ""; toast(err.message, "error"); }
    });

    /* ── Screenshot lightbox (main panel + sidebar thumb) ── */
    function openLightbox() {
      const lb = document.createElement("div");
      lb.className = "tkt-lightbox";
      lb.innerHTML = `<div class="tkt-lightbox-backdrop"></div><div class="tkt-lightbox-content"><img src="${escHtml(ssrc)}" alt="Screenshot"><button class="tkt-lightbox-close" aria-label="Close">&times;</button></div>`;
      document.body.appendChild(lb);
      const close = () => lb.remove();
      lb.querySelector(".tkt-lightbox-backdrop").addEventListener("click", close);
      lb.querySelector(".tkt-lightbox-close").addEventListener("click", close);
      const esc = e => { if (e.key === "Escape") { close(); document.removeEventListener("keydown", esc); } };
      document.addEventListener("keydown", esc);
    }
    const ssImg = $("#tktd-screenshot", view);
    if (ssImg) ssImg.addEventListener("click", openLightbox);
    const thumb = $("#tktd-thumb", view);
    if (thumb) thumb.addEventListener("click", openLightbox);
  }

  /* ── New issue modal ── */
  function openNewIssueModal() {
    const root = $("#tkt-modal-root");
    root.innerHTML = `
      <div class="tkt-modal-backdrop">
        <div class="tkt-modal">
          <div class="tkt-modal-head">
            <div><span class="tkt-workspace-label">Support request</span><h3>New Issue</h3></div>
            <button class="tkt-modal-close" id="nim-close">&times;</button>
          </div>
          <div class="tkt-modal-body">
            <div class="tkt-form-row"><label>Title <span class="tkt-req">*</span></label><input class="tkt-input" id="nim-title" placeholder="Brief description of the issue"></div>
            <div class="tkt-form-grid">
              <div class="tkt-form-row"><label>Type</label><select class="select-aewttr" id="nim-type">${ISSUE_TYPES.map(t => `<option>${escHtml(t)}</option>`).join("")}</select></div>
              <div class="tkt-form-row"><label>Status</label><select class="select-aewttr" id="nim-status">${STATUSES.map(s => `<option${s === "New" ? " selected" : ""}>${escHtml(s)}</option>`).join("")}</select></div>
            </div>
            <div class="tkt-form-row"><label>Description <span class="tkt-req">*</span></label><textarea class="tkt-textarea" id="nim-desc" rows="4" placeholder="What happened? What were you trying to do?"></textarea></div>
            <div class="tkt-form-row"><label>Expected result <span class="tkt-optional">optional</span></label><textarea class="tkt-textarea" id="nim-expected" rows="2"></textarea></div>
            <div class="tkt-form-row"><label>Additional context <span class="tkt-optional">optional</span></label><textarea class="tkt-textarea" id="nim-context" rows="2"></textarea></div>
            <div class="tkt-form-msg" id="nim-msg"></div>
          </div>
          <div class="tkt-modal-foot">
            <button class="tkt-btn-ghost" id="nim-cancel">Cancel</button>
            <button class="tkt-btn-primary" id="nim-submit">Submit Issue</button>
          </div>
        </div>
      </div>`;

    const close = () => { root.innerHTML = ""; };
    $("#nim-close",  root).addEventListener("click", close);
    $("#nim-cancel", root).addEventListener("click", close);
    root.querySelector(".tkt-modal-backdrop").addEventListener("pointerdown", e => { if (e.target.classList.contains("tkt-modal-backdrop")) close(); });

    $("#nim-submit", root).addEventListener("click", async () => {
      const title = ($("#nim-title", root).value || "").trim();
      const desc  = ($("#nim-desc",  root).value || "").trim();
      const msg   = $("#nim-msg", root);
      if (!title) { msg.className = "tkt-form-msg tkt-form-msg--error"; msg.textContent = "Title is required."; return; }
      if (!desc)  { msg.className = "tkt-form-msg tkt-form-msg--error"; msg.textContent = "Description is required."; return; }
      const btn = $("#nim-submit", root);
      btn.disabled = true; btn.textContent = "Submitting…"; msg.textContent = "";
      const now = new Date().toISOString();
      const issueCode = "ISS-" + now.slice(0, 10).replace(/-/g, "") + "-" + Math.random().toString(36).slice(2, 7).toUpperCase();
      try {
        const ctx = window._spPageContextInfo || {};
        await spCreate({
          Title:            title.slice(0, 255),
          IssueCode:        issueCode,
          IssueType:        $("#nim-type",     root).value,
          IssueStatus:      $("#nim-status",   root).value,
          Description:      desc.slice(0, 18000),
          ExpectedBehavior: ($("#nim-expected", root).value || "").trim().slice(0, 12000),
          AdditionalContext:($("#nim-context",  root).value || "").trim().slice(0, 12000),
          OccurredAt:       now,
          ReportedByName:   (ctx.userDisplayName || "").slice(0, 255) || "Reporter",
          ReporterEmail:    (ctx.userEmail || "").slice(0, 255),
          LogsJson:         "[]",
          ErrorCodesJson:   "[]",
          DiagnosticsJson:  "{}"
        });
        close();
        toast(`${issueCode} submitted`, "success");
        await load();
      } catch (err) {
        btn.disabled = false; btn.textContent = "Submit Issue";
        msg.className = "tkt-form-msg tkt-form-msg--error"; msg.textContent = err.message;
      }
    });
  }

  /* ── Load from SharePoint ── */
  async function load() {
    const select = [
      "Id", "Title", "Created",
      "IssueCode", "IssueStatus", "IssueType",
      "ReportedByName", "ReporterEmail",
      "OccurredAt", "Description", "ExpectedBehavior", "AdditionalContext",
      "PageTitle", "Route", "PageUrl",
      "ErrorCodesJson", "LogsJson", "DiagnosticsJson",
      "ScreenshotDataUrl", "ScreenshotFileUrl", "ScreenshotServerRelativeUrl", "ScreenshotFileName",
      "ResolutionNote"
    ].join(",");

    try {
      const url = `${getSiteUrl()}/_api/web/lists/getbytitle('${odata(LIST_TITLE)}')/items?$top=500&$select=${select}&$orderby=Created desc`;
      const r = await spFetch(url);
      if (!r.ok) {
        const txt = await r.text().catch(() => "");
        let detail = txt.slice(0, 300);
        try { const p = JSON.parse(txt); detail = (p.error && p.error.message && p.error.message.value) || detail; } catch (_) {}
        throw new Error(`SharePoint ${r.status}${detail ? ": " + detail : ""}`);
      }
      const payload = await r.json();
      tickets = (payload.value || []).map(normalise);
      render();
    } catch (err) {
      const view = $("#tkt-view");
      view.innerHTML = `<div class="tkt-load-error"><strong>Could not load issues</strong><p>${escHtml(err.message)}</p></div>`;
      console.error("[PULSE Tickets]", err);
    } finally {
      dismissLoader();
    }
  }

  /* ── Boot ── */
  function boot() {
    const newBtn = $("#btn-new-issue");
    if (newBtn) newBtn.addEventListener("click", openNewIssueModal);
    window.addEventListener("hashchange", render);
    load();
    setInterval(load, 60000);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
}());
