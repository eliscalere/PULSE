PAGE_RENDERERS.admin = function (parts) {
  const sub = parts[0];
  if (!sub) return renderAdminLauncher();
  renderAdminPage(sub, parts[1]);
};

function canManageSharePointSetup() {
  const setupStatus = window.AEWTTR.setupStatus;

  if (canCurrentUserAccessAdmin()) {
    return true;
  }

  if (window.AEWTTR.mode !== "sharepoint") {
    return false;
  }

  return !!(setupStatus && !setupStatus.ready);
}

const ADMIN_TOOLS = [
  { key: "lists", title: "Lists", sub: "Manage selectable Portfolios, End Item Configs, Locations, and ATO values.", icon: "bx-list-ul" },
  { key: "groups", title: "Groups", sub: "Create people groups used in pickers app-wide. Project members sync into a group named after the project.", icon: "bx-group" },
  { key: "spsetup", title: "SharePoint Setup", sub: "Check list/column status, pull the SharePoint people directory, and run setup.", icon: "bx-data" },
  { key: "notifications", title: "Notification Settings", sub: "Choose your own delivery channels, areas, and message tone.", icon: "bx-bell", externalRoute: "notification-settings" },
  { key: "users", title: "Users", sub: "All users directory, assign people to PULSE, manage AEWTTR roles.", icon: "bx-id-card", externalRoute: "users" },
  { key: "meetingadmins", title: "Meeting Admins", sub: "Designate who can start and facilitate weekly and project meetings.", icon: "bx-user-check" },
  { key: "issues", title: "Issues", sub: "Review user-submitted issue reports, diagnostics, and captured page snapshots.", icon: "bx-message-error" },
  { key: "settings", title: "Settings", sub: "Manage application-wide controls, including the CUI marking bar.", icon: "bx-slider-alt" },
  { key: "activitylog", title: "Activity Log", sub: "Review every user action recorded across the app.", icon: "bx-history", externalRoute: "logs" }
];

function renderAdminLauncher() {
  setTopbar("Admin", "SharePoint setup, users, and activity log.", "");
  $("#page-content").innerHTML = `
    <div class="tile-grid">
      ${ADMIN_TOOLS.map(t => `
        <div class="app-tile" data-go="${t.key}"${t.externalRoute ? ` data-route="${t.externalRoute}"` : ""}>
          <div class="tile-icon"><i class="bx ${t.icon}"></i></div>
          <div class="tile-title">${t.title}</div>
          <div class="tile-sub" data-help>${t.sub}</div>
          <button class="btn-aewttr btn-aewttr-sm tile-open" data-go="${t.key}"${t.externalRoute ? ` data-route="${t.externalRoute}"` : ""}${tip(`Open ${t.title}`)}>Open →</button>
        </div>`).join("")}
    </div>
  `;
  $all("[data-go]", $("#page-content")).forEach((button) => button.addEventListener("click", () => {
    if (button.dataset.route) {
      navigate(button.dataset.route);
      return;
    }
    navigate("admin/" + button.dataset.go);
  }));
}

function renderAdminPage(sub, recordId) {
  const tool = ADMIN_TOOLS.find(t => t.key === sub);
  if (!tool) { navigate("admin"); return; }
  if (tool.externalRoute) { navigate(tool.externalRoute); return; }
  const reviewingIssue = sub === "issues" && recordId;
  setTopbar(reviewingIssue ? "Issue review" : tool.title, reviewingIssue ? "Admin → Issues" : "Admin", `<button class="btn-aewttr-outline" id="btn-back-admin"${tip(reviewingIssue ? "Return to issue reports" : "Return to Admin home")}><i class="bx bx-arrow-back"></i> ${reviewingIssue ? "Issues" : "Admin"}</button>`);
  $("#page-content").innerHTML = `<div id="admin-body"></div>`;
  $("#btn-back-admin").addEventListener("click", () => navigate(reviewingIssue ? "admin/issues" : "admin"));
  const body = $("#admin-body");

  if (sub === "lists") return drawListsConfig(body);
  if (sub === "groups") return drawGroupsConfig(body);
  if (sub === "meetingadmins") return drawMeetingAdminsConfig(body);
  if (sub === "issues") return reviewingIssue ? drawIssueReviewPage(body, decodeURIComponent(recordId)) : drawIssueReports(body);
  if (sub === "settings") return drawAppSettings(body);
  if (sub === "spsetup") return drawSpSetup(body);
  if (sub === "releases") return drawReleases(body);
}

function drawIssueReports(body) {
  if (!canCurrentUserAccessAdmin()) {
    body.innerHTML = requireAdmin(() => "");
    return;
  }
  const db = window.AEWTTR.db;
  const render = () => {
    const issues = (db.issues || []).slice().sort((a, b) => String(b.createdAt || "").localeCompare(String(a.createdAt || "")));
    const counts = ["New", "Triaged", "In Progress", "Resolved"].map((status) => ({
      status,
      count: issues.filter((issue) => (issue.status || "New") === status).length
    }));
    body.innerHTML = `
      <div class="issues-admin-wrap">
        <div class="issues-admin-summary">
          <div><h3>Issue reports</h3><p>${issues.length ? "Open a report for a focused review with separate tabs for the capture, diagnostics, logs, and error codes." : "Reports submitted with the top-bar Report issue button will appear here."}</p></div>
          <div class="issues-admin-counts">
            ${counts.map(({ status, count }) => `<div class="issues-admin-count">${count}<span>${escapeHtml(status)}</span></div>`).join("")}
          </div>
        </div>
        <div class="issues-admin-list">
          ${issues.length ? issues.map((issue) => `
            <button type="button" class="issue-admin-card" data-issue-id="${escapeHtml(issue.id)}">
              <div class="issue-admin-card-head"><span class="issue-admin-card-code">${escapeHtml(issue.id || "ISSUE")}</span><span class="issue-status ${typeof issueStatusClass === "function" ? issueStatusClass(issue.status) : ""}">${escapeHtml(issue.status || "New")}</span></div>
              <h4>${escapeHtml(issue.title || issue.type || "Issue report")}</h4>
              <p>${escapeHtml(issue.description || "No description provided.")}</p>
              <div class="issue-admin-card-foot"><span><i class="bx bx-user"></i>${escapeHtml(issue.reportedBy || "Unknown")}</span><span>${escapeHtml(typeof formatIssueReportDate === "function" ? formatIssueReportDate(issue.createdAt) : issue.createdAt || "—")}</span></div>
            </button>`).join("") : `<div class="issue-admin-empty"><i class="bx bx-message-error"></i><strong>No issues reported yet</strong><div style="margin-top:4px;">Each submitted report includes page context, diagnostics, and any tab photo the reporter attached.</div></div>`}
        </div>
      </div>
    `;
    $all("[data-issue-id]", body).forEach((button) => button.addEventListener("click", () => {
      navigate(`admin/issues/${button.dataset.issueId}`);
    }));
  };
  render();
}

function issueReviewJson(value) {
  try { return JSON.stringify(value == null ? {} : value, null, 2); } catch (e) { return String(value || ""); }
}

function issueReviewValue(value, fallback) {
  const text = String(value == null ? "" : value).trim();
  return text || fallback || "—";
}

function openIssueScreenshotViewer(issue) {
  const screenshot = issue && typeof safeIssueScreenshotSource === "function" ? safeIssueScreenshotSource(issue) : "";
  if (!screenshot) return;
  const originalFileUrl = typeof safeIssueScreenshotFileUrl === "function" ? safeIssueScreenshotFileUrl(issue.screenshotFileUrl) : "";
  const modal = openModal(`
    <div class="issue-photo-viewer">
      <div class="aewttr-modal-head"><div><h3>Screenshot</h3><p>${escapeHtml(issue.id || "Issue report")} · ${escapeHtml(issue.pageTitle || "Captured page")}</p></div><button class="aewttr-modal-close" type="button" id="issue-photo-close" aria-label="Close screenshot viewer">&times;</button></div>
      <div class="issue-photo-toolbar" role="toolbar" aria-label="Screenshot zoom controls">
        <button class="btn-aewttr-outline btn-aewttr-sm" type="button" id="issue-photo-zoom-out" aria-label="Zoom out"><i class="bx bx-zoom-out"></i></button>
        <button class="btn-aewttr-outline btn-aewttr-sm" type="button" id="issue-photo-reset">Fit</button>
        <button class="btn-aewttr-outline btn-aewttr-sm" type="button" id="issue-photo-zoom-in" aria-label="Zoom in"><i class="bx bx-zoom-in"></i></button>
        <button class="btn-aewttr-outline btn-aewttr-sm" type="button" id="issue-photo-original"><i class="bx bx-image"></i> Original pixels</button>
        <span id="issue-photo-zoom-label" aria-live="polite">100%</span>
        ${originalFileUrl ? `<a class="btn-aewttr-outline btn-aewttr-sm issue-photo-original-link" href="${escapeHtml(originalFileUrl)}" target="_blank" rel="noopener"><i class="bx bx-link-external"></i> Open original</a>` : ""}
        <span class="issue-photo-toolbar-hint">Use the controls or double-click the image to zoom.</span>
      </div>
      <div class="issue-photo-viewport" id="issue-photo-viewport" tabindex="0"><img id="issue-photo-image" src="${escapeHtml(screenshot)}" alt="Full screenshot attached to ${escapeHtml(issue.id || "this issue")}"></div>
    </div>
  `, { wide: true, className: "issue-photo-viewer-modal" });
  const image = $("#issue-photo-image", modal);
  const viewport = $("#issue-photo-viewport", modal);
  const label = $("#issue-photo-zoom-label", modal);
  let zoom = 1;
  const applyZoom = (next, resetScroll) => {
    zoom = Math.max(0.25, Math.min(8, Math.round(next * 4) / 4));
    image.style.width = `${zoom * 100}%`;
    label.textContent = `${Math.round(zoom * 100)}%`;
    $("#issue-photo-zoom-out", modal).disabled = zoom <= 0.25;
    $("#issue-photo-zoom-in", modal).disabled = zoom >= 8;
    if (resetScroll) { viewport.scrollTop = 0; viewport.scrollLeft = 0; }
  };
  const showOriginalPixels = () => {
    const availableWidth = Math.max(1, viewport.clientWidth - 32);
    applyZoom(Math.max(1, image.naturalWidth / availableWidth), true);
  };
  $("#issue-photo-close", modal).addEventListener("click", closeModal);
  $("#issue-photo-zoom-out", modal).addEventListener("click", () => applyZoom(zoom - 0.25));
  $("#issue-photo-zoom-in", modal).addEventListener("click", () => applyZoom(zoom + 0.25));
  $("#issue-photo-reset", modal).addEventListener("click", () => applyZoom(1, true));
  $("#issue-photo-original", modal).addEventListener("click", showOriginalPixels);
  image.addEventListener("dblclick", () => applyZoom(zoom === 1 ? 2 : 1, zoom !== 1));
  modal.addEventListener("keydown", (event) => {
    if (event.key === "+" || event.key === "=") applyZoom(zoom + 0.25);
    if (event.key === "-") applyZoom(zoom - 0.25);
    if (event.key === "0") applyZoom(1, true);
  });
  image.addEventListener("load", () => applyZoom(1, true), { once: true });
  applyZoom(1, true);
}

function drawIssueReviewPage(body, issueId) {
  if (!canCurrentUserAccessAdmin()) {
    body.innerHTML = requireAdmin(() => "");
    return;
  }
  const db = window.AEWTTR.db;
  const issue = (db.issues || []).find((item) => item.id === issueId);
  if (!issue) {
    body.innerHTML = `<div class="issue-admin-empty"><i class="bx bx-search-alt"></i><strong>Issue not found</strong><div style="margin-top:4px;">It may have been deleted or is no longer available in this site.</div><button class="btn-aewttr-outline btn-aewttr-sm" type="button" id="issue-review-return" style="margin-top:14px;">Back to Issues</button></div>`;
    $("#issue-review-return", body).addEventListener("click", () => navigate("admin/issues"));
    return;
  }

  const tabs = [
    { id: "overview", label: "Overview", icon: "bx-file" },
    { id: "capture", label: "Tab photo", icon: "bx-camera" },
    { id: "diagnostics", label: "Diagnostics", icon: "bx-pulse" },
    { id: "logs", label: "Browser logs", icon: "bx-terminal" },
    { id: "codes", label: "Error codes", icon: "bx-error-circle" }
  ];
  let activeTab = "overview";
  let autoSaveTimer = null;
  let changeVersion = 0;
  let savedVersion = 0;
  let saveInFlight = false;

  const setAutoSaveState = (state, message) => {
    const indicator = $("#issue-review-autosave", body);
    const detail = $("#issue-review-save-status", body);
    if (indicator) {
      indicator.className = `issue-review-autosave issue-review-autosave--${state}`;
      indicator.innerHTML = `<i class="bx ${state === "saving" ? "bx-loader-alt bx-spin" : state === "error" ? "bx-error-circle" : "bx-cloud"}"></i> ${escapeHtml(message)}`;
    }
    if (detail) detail.textContent = state === "error" ? message : "";
  };

  const flushAutoSave = async () => {
    if (saveInFlight || savedVersion >= changeVersion) return;
    clearTimeout(autoSaveTimer);
    autoSaveTimer = null;
    const savingVersion = changeVersion;
    saveInFlight = true;
    setAutoSaveState("saving", "Saving changes…");
    try {
      await Repo.save("issue", issue, { immediate: true });
      savedVersion = Math.max(savedVersion, savingVersion);
      if (typeof logUserAction === "function") logUserAction({ action: "Update", area: "Admin", summary: `Auto-saved issue ${issue.id} as ${issue.status}`, recordId: issue.id });
      setAutoSaveState("saved", "Saved automatically");
    } catch (error) {
      savedVersion = changeVersion;
      setAutoSaveState("error", typeof issueSaveMessage === "function" ? issueSaveMessage(error) : "Could not save changes.");
    } finally {
      saveInFlight = false;
      if (savedVersion < changeVersion) autoSaveTimer = setTimeout(flushAutoSave, 350);
      else unregisterAutosaveFlusher(flushAutoSave);
    }
  };

  const queueAutoSave = () => {
    changeVersion += 1;
    clearTimeout(autoSaveTimer);
    setAutoSaveState("saving", "Saving changes…");
    registerAutosaveFlusher(flushAutoSave);
    autoSaveTimer = setTimeout(flushAutoSave, 650);
  };

  const render = () => {
    const screenshot = typeof safeIssueScreenshotSource === "function" ? safeIssueScreenshotSource(issue) : "";
    const diagnostics = issue.diagnostics || {};
    const logs = Array.isArray(issue.logs) ? issue.logs : [];
    const codes = Array.isArray(issue.errorCodes) ? issue.errorCodes : [];
    const metadata = issue.screenshotMeta || {};
    const tabContent = {
      overview: `
        <div class="issue-review-grid">
          <section class="issue-review-section"><h3>Reporter notes</h3><p>${escapeHtml(issueReviewValue(issue.description, "No description was provided."))}</p>${issue.expectedBehavior ? `<div class="issue-review-note"><strong>Expected instead</strong><p>${escapeHtml(issue.expectedBehavior)}</p></div>` : ""}${issue.additionalContext ? `<div class="issue-review-note"><strong>Additional context</strong><p>${escapeHtml(issue.additionalContext)}</p></div>` : ""}</section>
          <aside class="issue-review-section issue-review-facts"><h3>Report details</h3><dl><div><dt>Reporter</dt><dd>${escapeHtml(issueReviewValue(issue.reportedBy, "Unknown"))}</dd></div><div><dt>Submitted</dt><dd>${escapeHtml(typeof formatIssueReportDate === "function" ? formatIssueReportDate(issue.createdAt) : issue.createdAt)}</dd></div><div><dt>Page</dt><dd>${escapeHtml(issueReviewValue(issue.pageTitle))}</dd></div><div><dt>Route</dt><dd><code>${escapeHtml(issueReviewValue(issue.route))}</code></dd></div></dl></aside>
        </div>
        <section class="issue-review-section issue-review-resolution"><h3>Resolution</h3><div class="issue-review-resolution-grid"><div class="form-row"><label for="issue-review-status">Status</label><select class="select-aewttr" id="issue-review-status">${["New", "Triaged", "In Progress", "Resolved", "Closed"].map((status) => `<option${status === (issue.status || "New") ? " selected" : ""}>${status}</option>`).join("")}</select></div><div class="form-row"><label for="issue-review-note">Admin notes <span class="field-optional">Optional</span></label><textarea class="input-aewttr" id="issue-review-note" rows="3" placeholder="Resolution, workaround, or next step…">${escapeHtml(issue.resolutionNote || "")}</textarea></div></div><div class="issue-review-save-status" id="issue-review-save-status" aria-live="polite"></div></section>`,
      capture: `<section class="issue-review-section"><h3>Captured tab photo</h3>${screenshot ? `<button class="issue-review-photo-button" type="button" id="issue-review-open-photo" aria-label="Open full-resolution screenshot viewer"><img class="issue-review-photo" src="${escapeHtml(screenshot)}" alt="Native browser capture attached to this issue"><span><i class="bx bx-zoom-in"></i> Open full resolution</span></button><dl class="issue-review-capture-meta"><div><dt>Source</dt><dd>${escapeHtml(issueReviewValue(metadata.source, "Native browser tab capture"))}</dd></div><div><dt>Captured</dt><dd>${escapeHtml(typeof formatIssueReportDate === "function" ? formatIssueReportDate(metadata.capturedAt) : metadata.capturedAt)}</dd></div><div><dt>Resolution</dt><dd>${escapeHtml(metadata.width && metadata.height ? `${metadata.width} × ${metadata.height}` : "—")}</dd></div><div><dt>File</dt><dd>${escapeHtml(issueReviewValue(issue.screenshotFileName, issue.screenshotFileUrl ? "PULSE App Data image" : "Legacy inline image"))}</dd></div><div><dt>Storage</dt><dd>${escapeHtml(issueReviewValue(metadata.storage, issue.screenshotFileUrl ? "PULSE App Data / AEWTTR-PULSE / Issue Screenshots" : "Issue record"))}</dd></div></dl>` : `<div class="issue-screenshot-empty"><i class="bx bx-camera-off"></i><span>No native tab photo was attached to this report.</span></div>`}</section>`,
      diagnostics: `<section class="issue-review-section"><h3>Environment diagnostics</h3><p class="issue-review-intro">Captured at report time, including the page, browser, viewport, SharePoint errors, and PULSE debug entries.</p><pre class="issue-review-code">${escapeHtml(issueReviewJson(diagnostics))}</pre></section>`,
      logs: `<section class="issue-review-section"><h3>Browser logs</h3><p class="issue-review-intro">${logs.length ? `${logs.length} recent entry${logs.length === 1 ? "" : "ies"} captured with this report.` : "No browser log entries were captured."}</p>${logs.length ? `<ol class="issue-review-log-list">${logs.map((entry) => `<li><span class="issue-review-log-level issue-review-log-level--${escapeHtml(String(entry.level || "info").toLowerCase())}">${escapeHtml(entry.level || "info")}</span><div><time>${escapeHtml(typeof formatIssueReportDate === "function" ? formatIssueReportDate(entry.ts) : entry.ts || "")}</time><code>${escapeHtml(issueReviewValue(entry.message))}</code></div></li>`).join("")}</ol>` : ""}</section>`,
      codes: `<section class="issue-review-section"><h3>Error codes</h3><p class="issue-review-intro">HTTP and application error codes identified from the captured browser and SharePoint logs.</p>${codes.length ? `<div class="issue-review-code-list">${codes.map((code) => `<span>${escapeHtml(code)}</span>`).join("")}</div>` : `<div class="issue-review-empty-state"><i class="bx bx-check-circle"></i><span>No error codes were detected in the captured logs.</span></div>`}</section>`
    };
    body.innerHTML = `
      <div class="issue-review-page">
        <header class="issue-review-header"><div><p class="issue-report-kicker">${escapeHtml(issue.id || "Issue report")}</p><h2>${escapeHtml(issue.title || issue.type || "Issue report")}</h2><div class="issue-review-header-meta"><span class="issue-type-pill">${escapeHtml(issue.type || "Other")}</span><span class="issue-status ${typeof issueStatusClass === "function" ? issueStatusClass(issue.status) : ""}">${escapeHtml(issue.status || "New")}</span><span><i class="bx bx-user"></i> ${escapeHtml(issueReviewValue(issue.reportedBy, "Unknown reporter"))}</span></div></div><div class="issue-review-actions"><span class="issue-review-autosave issue-review-autosave--saved" id="issue-review-autosave"><i class="bx bx-cloud"></i> Saved automatically</span>${["Resolved", "Closed"].includes(issue.status) ? "" : `<button class="btn-aewttr" type="button" id="issue-review-resolve"><i class="bx bx-check"></i> Resolve</button>`}<button class="btn-danger-outline" type="button" id="issue-review-delete"><i class="bx bx-trash"></i> Delete</button></div></header>
        <nav class="issue-review-tabs" aria-label="Issue review sections">${tabs.map((tab) => `<button type="button" class="issue-review-tab${activeTab === tab.id ? " active" : ""}" data-issue-tab="${tab.id}" aria-selected="${activeTab === tab.id}"><i class="bx ${tab.icon}"></i>${tab.label}</button>`).join("")}</nav>
        <div class="issue-review-panel">${tabContent[activeTab]}</div>
      </div>`;

    $all("[data-issue-tab]", body).forEach((button) => button.addEventListener("click", () => { activeTab = button.dataset.issueTab; render(); }));
    const openPhoto = $("#issue-review-open-photo", body);
    if (openPhoto) openPhoto.addEventListener("click", () => openIssueScreenshotViewer(issue));
    const status = $("#issue-review-status", body);
    const note = $("#issue-review-note", body);
    if (status) status.addEventListener("change", () => {
      issue.status = status.value;
      const chip = body.querySelector(".issue-review-header .issue-status");
      if (chip) {
        chip.className = `issue-status ${typeof issueStatusClass === "function" ? issueStatusClass(issue.status) : ""}`;
        chip.textContent = issue.status;
      }
      queueAutoSave();
    });
    if (note) note.addEventListener("input", () => { issue.resolutionNote = note.value; queueAutoSave(); });
    const resolve = $("#issue-review-resolve", body);
    if (resolve) resolve.addEventListener("click", async () => {
      clearTimeout(autoSaveTimer);
      if (note) issue.resolutionNote = note.value.trim();
      issue.status = "Resolved";
      resolve.disabled = true;
      try {
        await Repo.save("issue", issue, { immediate: true });
        savedVersion = changeVersion;
        if (typeof logUserAction === "function") logUserAction({ action: "Resolve", area: "Admin", summary: `Resolved issue ${issue.id}`, recordId: issue.id });
        toast("Issue marked resolved.", "success");
        render();
      } catch (error) {
        toast(typeof issueSaveMessage === "function" ? issueSaveMessage(error) : "Could not resolve this issue.", "error");
        resolve.disabled = false;
      }
    });
    $("#issue-review-delete", body).addEventListener("click", async () => {
      const confirmed = await confirmDialog({ title: "Delete issue", message: `Delete ${issue.id || "this issue"}? This cannot be undone.`, confirmLabel: "Delete", danger: true });
      if (!confirmed) return;
      clearTimeout(autoSaveTimer);
      const index = db.issues.indexOf(issue);
      if (index < 0) return navigate("admin/issues");
      db.issues.splice(index, 1);
      try {
        await Repo.remove("issue", issue);
        let screenshotCleanupFailed = false;
        if (window.AEWTTR.mode === "sharepoint" && issue.screenshotServerRelativeUrl && typeof sharePointAdapter !== "undefined" && typeof sharePointAdapter.deleteIssueScreenshot === "function") {
          try {
            await sharePointAdapter.deleteIssueScreenshot(window.AEWTTR.siteUrl, issue.screenshotServerRelativeUrl);
          } catch (cleanupError) {
            screenshotCleanupFailed = true;
            console.warn("PULSE could not remove the issue screenshot file.", cleanupError);
          }
        }
        if (window.AEWTTR.mode !== "sharepoint" && typeof aewttrSaveStore === "function") aewttrSaveStore();
        if (typeof logUserAction === "function") logUserAction({ action: "Delete", area: "Admin", summary: `Deleted issue ${issue.id}`, recordId: issue.id });
        toast(screenshotCleanupFailed ? "Issue deleted. Its screenshot file could not be removed from PULSE App Data." : "Issue and screenshot deleted.", screenshotCleanupFailed ? "error" : "success");
        navigate("admin/issues");
      } catch (error) {
        delete issue._deleted;
        db.issues.splice(index, 0, issue);
        toast(typeof issueSaveMessage === "function" ? issueSaveMessage(error) : "Could not delete this issue.", "error");
        render();
      }
    });
  };
  render();
}

function drawAppSettings(body) {
  body.innerHTML = requireAdmin(() => {
    const settings = typeof getAppSettings === "function" ? getAppSettings() : { cuiMarkingEnabled: false };
    return `
      <div class="form-page-grid">
        <div class="aewttr-card aewttr-card-pad">
          <div class="side-panel-title">Display markings</div>
          <p style="margin-top:0;color:var(--aewttr-muted);font-size:13px;line-height:1.5;">Turn on a persistent green CUI bar at the top and bottom of every PULSE page. This setting applies to everyone who opens the app.</p>
          <label class="weekly-roster-option" style="margin:16px 0 14px;">
            <input type="checkbox" id="app-settings-cui" ${settings.cuiMarkingEnabled ? "checked" : ""}>
            <div><strong>Show CUI marking bar</strong><span>Displays a green “CUI” bar above and below the workspace.</span></div>
          </label>
          <div style="display:flex;align-items:center;gap:10px;">
            <button class="btn-aewttr" id="app-settings-save"${tip("Save app settings")}><i class="bx bx-save"></i> Save settings</button>
            <span id="app-settings-status" style="font-size:12px;color:var(--aewttr-muted);" aria-live="polite"></span>
          </div>
        </div>
        <div class="aewttr-card aewttr-card-pad form-page-side">
          <div class="side-panel-title">SharePoint setup</div>
          <p style="margin:0;color:var(--aewttr-muted);font-size:12.5px;line-height:1.55;">The setting is stored in the single-row <code>PULSE App Settings</code> list. Run <strong>Admin → SharePoint Setup</strong> after publishing this package to create the list and its column on the site.</p>
        </div>
      </div>
    `;
  });
  if (!canCurrentUserAccessAdmin()) return;
  const save = $("#app-settings-save", body);
  const status = $("#app-settings-status", body);
  save.addEventListener("click", async () => {
    const settings = typeof getAppSettings === "function" ? getAppSettings() : null;
    if (!settings) return;
    const previous = !!settings.cuiMarkingEnabled;
    settings.cuiMarkingEnabled = $("#app-settings-cui", body).checked;
    save.disabled = true;
    status.textContent = "Saving…";
    try {
      await Repo.save("appSettings", settings, { immediate: true });
      if (typeof applyCuiMarking === "function") applyCuiMarking();
      if (typeof logUserAction === "function") logUserAction({ action: "Update", area: "Admin", summary: `CUI marking bar turned ${settings.cuiMarkingEnabled ? "on" : "off"}` });
      status.textContent = "Saved";
      toast("Application settings saved.", "success");
    } catch (e) {
      settings.cuiMarkingEnabled = previous;
      status.textContent = (e && (e.friendly || e.message)) || "Could not save settings.";
      toast("Could not save application settings.", "error");
    } finally {
      save.disabled = false;
    }
  });
}

function drawMeetingAdminsConfig(body) {
  if (!canCurrentUserAccessAdmin()) {
    body.innerHTML = `<div class="lock-block"><i class="bx bx-lock-alt"></i><h3>Admin access required</h3><p>Only admins can manage meeting facilitators.</p></div>`;
    return;
  }
  const cfg = getLocationConfig();
  if (!Array.isArray(cfg.meetingAdminEmails)) cfg.meetingAdminEmails = [];

  function render() {
    const list = cfg.meetingAdminEmails;
    body.innerHTML = `
      <div class="admin-list-section">
        <p class="admin-section-desc">People listed here can start and facilitate weekly and project meetings, in addition to system admins. Enter the exact name or email address as it appears in PULSE.</p>
        <div class="admin-list-items" id="meeting-admin-list">
          ${list.length === 0
            ? `<div class="admin-list-empty">No meeting facilitators configured. Admins can always facilitate.</div>`
            : list.map((entry, i) => `
              <div class="admin-list-row">
                <span class="admin-list-label"><i class="bx bx-user" style="margin-right:6px;opacity:.5;"></i>${escapeHtml(entry)}</span>
                <button class="btn-aewttr-ghost btn-aewttr-sm admin-list-remove" data-idx="${i}"${tip("Remove")}><i class="bx bx-x"></i></button>
              </div>`).join("")}
        </div>
        <div class="admin-list-add-row" style="display:flex;gap:8px;margin-top:12px;">
          <input class="input-aewttr" id="meeting-admin-input" placeholder="Name or email address…" style="flex:1;">
          <button class="btn-aewttr btn-aewttr-sm" id="meeting-admin-add"><i class="bx bx-plus"></i> Add</button>
        </div>
        <div style="margin-top:12px;display:flex;align-items:center;gap:12px;">
          <button class="btn-aewttr" id="meeting-admin-save">Save changes</button>
          <span id="meeting-admin-status" style="font-size:13px;color:var(--aewttr-muted);"></span>
        </div>
      </div>
    `;

    $all(".admin-list-remove", body).forEach((btn) => {
      btn.addEventListener("click", () => {
        const idx = parseInt(btn.dataset.idx, 10);
        cfg.meetingAdminEmails.splice(idx, 1);
        render();
      });
    });

    const input = $("#meeting-admin-input", body);
    const addBtn = $("#meeting-admin-add", body);
    const saveBtn = $("#meeting-admin-save", body);
    const statusEl = $("#meeting-admin-status", body);

    function addEntry() {
      const val = (input && input.value || "").trim();
      if (!val) return;
      if (cfg.meetingAdminEmails.includes(val)) { if (statusEl) statusEl.textContent = "Already in the list."; return; }
      cfg.meetingAdminEmails.push(val);
      if (input) input.value = "";
      if (statusEl) statusEl.textContent = "";
      render();
    }

    if (addBtn) addBtn.addEventListener("click", addEntry);
    if (input) input.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); addEntry(); } });

    if (saveBtn) saveBtn.addEventListener("click", async () => {
      saveBtn.disabled = true;
      if (statusEl) statusEl.textContent = "Saving…";
      try {
        await Repo.save("locationConfig", cfg);
        if (statusEl) { statusEl.textContent = "Saved."; setTimeout(() => { if (statusEl) statusEl.textContent = ""; }, 2500); }
      } catch (e) {
        if (statusEl) statusEl.textContent = "Save failed — try again.";
      } finally {
        saveBtn.disabled = false;
      }
    });
  }

  render();
}

function drawSpSetup(body) {
  if (!canManageSharePointSetup()) {
    body.innerHTML = `
      <div class="lock-block">
        <i class="bx bx-lock-alt"></i>
        <h3>SharePoint setup access required</h3>
        <p>This setup can be run by an app admin, or while the SharePoint lists are not yet provisioned.</p>
      </div>`;
    return;
  }

  const mode = window.AEWTTR.mode;
  const siteUrl = window.AEWTTR.siteUrl || "(not detected)";
  const currentUser = window.AEWTTR.currentSpUser;
  const setupStatus = window.AEWTTR.setupStatus;
  const lastError = (typeof getLastSpError === "function" && getLastSpError()) || window.AEWTTR.spBootError;

  body.innerHTML = `
    <div class="aewttr-card aewttr-card-pad spsetup-card">
      <div class="side-panel-title">SharePoint Setup</div>
      <div class="spsetup-status-strip">
        <div class="spsetup-status-item"><span class="k">Mode</span><span class="v">${mode === "sharepoint" ? `<span class="kc-badge">SharePoint</span>` : `<span class="kc-badge">Local fallback (dev)</span>`}</span></div>
        <div class="spsetup-status-item"><span class="k">Site</span><span class="v" title="${escapeHtml(siteUrl)}">${escapeHtml(siteUrl)}</span></div>
        <div class="spsetup-status-item"><span class="k">Current user</span><span class="v">${currentUser ? escapeHtml(currentUser.displayName + " (" + currentUser.email + ")") : "n/a — local fallback mode"}</span></div>
        <div class="spsetup-status-item"><span class="k">Connection</span><span class="v">${mode === "sharepoint" ? (lastError ? `<span style="color:var(--aewttr-danger,#c0392b);">Error</span>` : "Connected") : "Not in SharePoint"}</span></div>
        <div class="spsetup-status-item" id="spsetup-role-summary"><span class="k">Your role</span><span class="v">${mode === "sharepoint" && currentUser ? "Checking…" : "n/a"}</span></div>
      </div>

      <div class="spsetup-actions">
        <button class="btn-aewttr-outline btn-aewttr-sm" id="spsetup-check"${tip("Check whether SharePoint lists and columns exist")}>Check Setup</button>
        <button class="btn-aewttr btn-aewttr-sm" id="spsetup-run"${tip("Create/update lists and columns, seed config rows, sync roles, and repair legacy required columns")}>Run SharePoint Setup</button>
        <button class="btn-aewttr-outline btn-aewttr-sm" id="spsetup-pull-people"${tip("Import site users, Members/Owners, and permission grants into the PULSE people directory (does not assign anyone to PULSE)")}><i class="bx bx-user-plus"></i> Pull SharePoint people</button>
        <button class="btn-aewttr-outline btn-aewttr-sm" id="spsetup-sample-project"${tip("Create a new SAMPLE project with full metadata, people, tracker divider/tasks, risks, notes, and a board — safe to re-run (unique name each click)")}><i class="bx bx-flask"></i> Generate sample project</button>
      </div>
      <p id="spsetup-pull-status" style="font-size:12px;color:var(--aewttr-muted);margin:8px 0 0;"></p>
      <p id="spsetup-sample-status" style="font-size:12px;color:var(--aewttr-muted);margin:4px 0 0;"></p>

      <div id="spsetup-results"></div>

      <div class="side-panel-title" style="margin-top:16px;">List Status</div>
      <div id="spsetup-list-status" class="spsetup-list-grid">${setupStatus ? renderListStatusGrid(setupStatus) : `<p style="font-size:12.5px;color:var(--aewttr-muted);grid-column:1/-1;">Click "Check Setup" to check.</p>`}</div>

      ${lastError ? `
        <div class="spsetup-error-block">
          <strong>Last SharePoint error</strong>
          <div style="font-size:12px;margin-top:4px;">${escapeHtml(lastError.friendly + (lastError.status ? " (HTTP " + lastError.status + ")" : ""))}</div>
          ${lastError.recordedAt ? `<div style="font-size:11px;color:var(--aewttr-muted);margin-top:2px;">Recorded ${escapeHtml(fmtRelativeTime(lastError.recordedAt))}</div>` : ""}
          <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="spsetup-clear-error" style="margin-top:6px;padding:2px 8px;"${tip("Clear this recorded error")}>Clear</button>
        </div>
      ` : ""}
      <p style="font-size:11.5px;color:var(--aewttr-muted);margin-top:14px;">This app uses the currently authenticated SharePoint user — no separate login. SharePoint permissions control site/list access; app roles control which in-app tools are visible.</p>
    </div>
  `;
  wireSpSetupButtons(body);
  if (mode === "sharepoint" && siteUrl && currentUser) refreshRoleSummary(body);

  // A background refresh that succeeds clears the stale error at the source
  // (sharepoint-adapter.js's sendRequest), but if you're already sitting on
  // this page when that happens, the "Connection: Error" badge would stay
  // stuck-looking until you navigated away and back. Re-draw in place so it
  // reflects reality without disturbing whatever you're doing.
  if (window.AEWTTR._spSetupLiveRefreshHandler) {
    window.removeEventListener("pulse:data-refreshed", window.AEWTTR._spSetupLiveRefreshHandler);
  }
  window.AEWTTR._spSetupLiveRefreshHandler = () => {
    if (!document.querySelector(".spsetup-card")) return;
    if (document.querySelector(".aewttr-modal-backdrop")) return;
    const active = document.activeElement;
    if (active && (active.tagName === "INPUT" || active.tagName === "TEXTAREA" || active.tagName === "SELECT" || active.isContentEditable)) return;
    drawSpSetup(body);
  };
  window.addEventListener("pulse:data-refreshed", window.AEWTTR._spSetupLiveRefreshHandler);
}

async function refreshRoleSummary(body) {
  const el = $("#spsetup-role-summary .v", body);
  if (!el) return;
  try {
    const roleInfo = await sharePointAdapter.getCurrentUserRole(window.AEWTTR.siteUrl, window.AEWTTR.currentSpUser);
    const matched = roleInfo.matchedBy || roleInfo.source;
    el.innerHTML = matched && matched !== "Default"
      ? `${escapeHtml(roleInfo.role)} <span style="color:var(--aewttr-muted);font-size:11px;">(matched by ${escapeHtml(matched)})</span>`
      : `<span style="color:var(--aewttr-warning,#9a6700);">${escapeHtml(roleInfo.role)} — not matched in PULSE App Roles, using default</span>`;
  } catch (e) {
    el.innerHTML = `<span style="color:var(--aewttr-danger,#c0392b);">Could not check</span>`;
  }
}

/* Compact chip grid — was a tall kv-list (one multi-line row per list, up
   to 15 lists) that pushed everything else off screen. Each chip shows
   just name + status; the specific missing/mismatched fields live in a
   tooltip instead of always-rendered text, so a fully-provisioned site
   reads as a quick, glanceable grid rather than a wall of green text. */
function renderListStatusGrid(status) {
  return status.lists.map((l) => {
    const missing = l.missingFields || [];
    const mismatches = l.typeMismatchFields || [];
    const suspicious = l.suspiciousFields || [];
    const state = !l.exists || missing.length || mismatches.length ? "missing" : suspicious.length ? "warn" : "ok";
    const badge = !l.exists ? "Missing" : (missing.length || mismatches.length) ? "Needs repair" : suspicious.length ? "Review" : "Ready";
    const issues = [];
    if (missing.length) issues.push(`Missing: ${missing.join(", ")}`);
    if (mismatches.length) issues.push(`Type warnings: ${mismatches.join("; ")}`);
    if (suspicious.length) issues.push(`Store warnings: ${suspicious.join("; ")}`);
    const shortName = l.listTitle.replace(/^PULSE /, "");
    return `
      <div class="spsetup-list-chip ${state}"${issues.length ? tip(issues.join(" · ")) : ""}>
        <span class="name">${escapeHtml(shortName)}</span>
        <span class="badge">${escapeHtml(badge)}</span>
      </div>
    `;
  }).join("");
}

function wireSampleProjectButton(body) {
  const sampleBtn = $("#spsetup-sample-project", body);
  const sampleStatus = $("#spsetup-sample-status", body);
  if (!sampleBtn) return;

  if (!canManageSharePointSetup() && !canCurrentUserAccessAdmin()) {
    sampleBtn.disabled = true;
    return;
  }

  sampleBtn.addEventListener("click", async () => {
    if (typeof generateSampleProject !== "function") {
      toast("Sample project generator is not loaded.", "error");
      return;
    }
    sampleBtn.disabled = true;
    const original = sampleBtn.innerHTML;
    sampleBtn.innerHTML = `<i class="bx bx-loader-alt bx-spin"></i> Generating…`;
    if (sampleStatus) sampleStatus.textContent = "Starting sample project…";
    try {
      const result = await generateSampleProject({
        navigateAfter: true,
        onProgress: (p) => {
          if (sampleStatus && p && p.message) sampleStatus.textContent = p.message;
        }
      });
      const msg = `Created ${result.id} — ${result.name}`;
      if (sampleStatus) sampleStatus.textContent = msg;
      toast(msg, "success");
    } catch (e) {
      const friendly = (e && e.friendly) || e.message || "Could not generate sample project";
      if (sampleStatus) sampleStatus.textContent = friendly;
      toast(friendly, "error");
      sampleBtn.disabled = false;
      sampleBtn.innerHTML = original;
    }
  });
}

function wireSpSetupButtons(body) {
  const checkBtn = $("#spsetup-check", body);
  const runBtn = $("#spsetup-run", body);
  const pullBtn = $("#spsetup-pull-people", body);
  const pullStatus = $("#spsetup-pull-status", body);
  const sampleBtn = $("#spsetup-sample-project", body);
  const clearErrorBtn = $("#spsetup-clear-error", body);
  if (clearErrorBtn) clearErrorBtn.addEventListener("click", () => { clearLastSpError(); drawSpSetup(body); });

  wireSampleProjectButton(body);

  if (!canManageSharePointSetup()) {
    checkBtn.disabled = true;
    runBtn.disabled = true;
    if (pullBtn) pullBtn.disabled = true;
    if (sampleBtn) sampleBtn.disabled = true;
    $("#spsetup-results", body).innerHTML = `<p style="font-size:12.5px;color:var(--aewttr-muted);margin-top:12px;">You need app admin access to run setup after the app is provisioned.</p>`;
    return;
  }
  if (window.AEWTTR.mode !== "sharepoint") {
    checkBtn.disabled = true;
    runBtn.disabled = true;
    if (pullBtn) pullBtn.disabled = true;
    // Sample project still works in local fallback — only SP list setup is disabled.
    $("#spsetup-results", body).innerHTML = `<p style="font-size:12.5px;color:var(--aewttr-muted);margin-top:12px;">Setup check/run only work when hosted inside SharePoint. You can still Generate sample project in local mode.</p>`;
    return;
  }
  checkBtn.addEventListener("click", async () => {
    checkBtn.disabled = true;
    checkBtn.textContent = "Checking…";
    try {
      const status = await getSetupStatus(window.AEWTTR.siteUrl);
      window.AEWTTR.setupStatus = status;
      $("#spsetup-list-status", body).innerHTML = renderListStatusGrid(status);
      toast(status.ready ? "All lists ready" : "Some lists need setup — see below", status.ready ? "success" : "error");
    } catch (e) {
      toast((e && e.friendly) || "Could not check setup status", "error");
    }
    checkBtn.disabled = false;
    checkBtn.textContent = "Check Setup";
  });
  runBtn.addEventListener("click", async () => {
    runBtn.disabled = true;
    runBtn.textContent = "Running…";
    $("#spsetup-results", body).innerHTML = "";
    try {
      const result = await runSharePointSetup(window.AEWTTR.siteUrl);
      window.AEWTTR.setupStatus = result.status;
      $("#spsetup-list-status", body).innerHTML = renderListStatusGrid(result.status);
      $("#spsetup-results", body).innerHTML = `<div class="aewttr-card aewttr-card-pad" style="font-size:12px;margin-top:12px;"><strong>Setup results</strong><ul style="margin:8px 0 0;padding-left:18px;">${result.log.map(l => `<li>${escapeHtml(l)}</li>`).join("")}</ul></div>`;
      toast(result.ok ? "SharePoint setup complete" : "Setup finished with errors — see results", result.ok ? "success" : "error");
      if (typeof sharePointAdapter.bootstrapAdminIfNeeded === "function" && window.AEWTTR.currentSpUser) {
        try {
          const boot = await sharePointAdapter.bootstrapAdminIfNeeded(window.AEWTTR.siteUrl, window.AEWTTR.currentSpUser);
          if (boot && boot.bootstrapped) toast("Bootstrap Admin created for you so PULSE stays reachable.", "success");
        } catch (e) { /* ignore */ }
      }
    } catch (e) {
      const details = formatSetupError(e);
      $("#spsetup-results", body).innerHTML = `<div class="aewttr-card aewttr-card-pad" style="font-size:12px;color:var(--aewttr-danger,#c0392b);white-space:pre-wrap;margin-top:12px;">${escapeHtml(details)}</div>`;
      toast((e && e.friendly) || "SharePoint setup failed", "error");
    }
    runBtn.disabled = false;
    runBtn.textContent = "Run SharePoint Setup";
  });

  if (pullBtn) {
    pullBtn.addEventListener("click", async () => {
      if (!canCurrentUserAccessAdmin() && !canManageSharePointSetup()) {
        toast("You do not have permission to pull the SharePoint directory.", "error");
        return;
      }
      pullBtn.disabled = true;
      const original = pullBtn.innerHTML;
      pullBtn.innerHTML = `<i class="bx bx-loader-alt bx-spin"></i> Pulling…`;
      if (pullStatus) pullStatus.textContent = "Starting SharePoint people pull…";
      try {
        const result = await sharePointAdapter.pullSharePointPeopleDirectory(window.AEWTTR.siteUrl, {
          onProgress: (p) => {
            if (pullStatus && p && p.message) pullStatus.textContent = p.message;
          }
        });
        const msg = `Imported ${result.total} people (${result.added} new). Assign them to PULSE from Users → All users.`;
        if (pullStatus) pullStatus.textContent = msg;
        toast(msg, "success");
      } catch (e) {
        const friendly = (e && e.friendly) || e.message || "People pull failed";
        if (pullStatus) pullStatus.textContent = friendly;
        toast(friendly, "error");
      }
      pullBtn.disabled = false;
      pullBtn.innerHTML = original;
    });
  }
}

function formatSetupError(error) {
  if (!error) return "Unknown error.";
  if (typeof error === "string") return error;

  const lines = [];
  lines.push(error.friendly || error.message || "Unknown error.");
  if (error.status) lines.push("HTTP status: " + error.status);
  if (error.cause) lines.push("Cause: " + error.cause);
  if (error.raw) lines.push("Raw response: " + error.raw);
  return lines.join("\n\n");
}

function drawNotificationConfig(body) {
  body.innerHTML = requireAdmin(() => {
    const cfg = typeof notificationConfig === "function" ? notificationConfig() : { teamsEnabled: true, alsoSendEmail: true };
    return `
      <div class="form-page-grid">
        <div class="aewttr-card aewttr-card-pad">
          <div class="side-panel-title">Outbound Notifications</div>
          <p style="font-size:13px;color:var(--aewttr-muted);margin-top:0;">
            PULSE can email people today using SharePoint's built-in mail endpoint. Teams chat messages need a small Power Automate flow that watches the <strong>PULSE Notifications</strong> list.
          </p>
          <div class="form-row" style="margin-top:14px;">
            <label style="display:flex;align-items:center;gap:8px;font-weight:600;">
              <input type="checkbox" id="nc-teams-enabled" ${cfg.teamsEnabled ? "checked" : ""}>
              Send Teams chat messages
            </label>
          </div>
          <div class="form-row">
            <label style="display:flex;align-items:center;gap:8px;font-weight:600;">
              <input type="checkbox" id="nc-also-email" ${cfg.alsoSendEmail !== false ? "checked" : ""}>
              Also send email notifications
            </label>
          </div>
          <button class="btn-aewttr" id="nc-save"${tip("Save notification settings")}><i class="bx bx-check"></i> Save Settings</button>
        </div>
        <div class="aewttr-card aewttr-card-pad form-page-side">
          <div class="side-panel-title">Power Automate setup</div>
          <ol style="font-size:12.5px;color:var(--aewttr-text);line-height:1.55;padding-left:18px;margin:0;">
            <li>Run <strong>Admin → SharePoint Setup → Run SharePoint Setup</strong> so <code>PULSE Notifications</code> exists (includes <code>EmailBodyHtml</code>).</li>
            <li>Create a cloud flow triggered by <strong>When an item is created</strong> in <code>PULSE Notifications</code>.</li>
            <li>Split <code>ToEmails</code> on <code>;</code> and loop each recipient email.</li>
            <li>Use <strong>Post adaptive card in a chat or channel</strong> (Teams).</li>
            <li>Set <strong>Adaptive Card</strong> to: <code>json(triggerOutputs()?['body/BodyHtml'])</code></li>
            <li>For email (optional): <strong>Send an email (V2)</strong> with <code>EmailBodyHtml</code> as HTML body.</li>
            <li>Fallback plain text is still in <code>TeamsText</code> if needed.</li>
            <li>Update <code>DeliveryStatus</code> to <code>Sent</code> or <code>Failed</code>.</li>
          </ol>
          <div class="side-panel-title" style="margin-top:16px;">What sends notifications</div>
          <ul style="font-size:12.5px;color:var(--aewttr-text);line-height:1.55;padding-left:18px;margin:8px 0 0;">
            <li><strong>Travel</strong> — admins on submit; finance on admin approval; requester on status changes (approved, denied, cancelled, finalized).</li>
            <li><strong>Document review</strong> — pending reviewers when a doc is submitted, a revision is uploaded, or they are added; daily reminders until they approve or request changes (only reviewers still marked Pending).</li>
          </ul>
          <div class="side-panel-title" style="margin-top:16px;">Daily doc-review reminders (optional)</div>
          <p style="font-size:12px;color:var(--aewttr-muted);margin:8px 0 0;">
            While PULSE is open, reminders run automatically every 24 hours per pending reviewer. For reminders when nobody has the app open, add a second scheduled Power Automate flow that reads <code>PULSE Document Review</code>, parses <code>ReviewersJson</code>, and queues rows in <code>PULSE Notifications</code> for reviewers with <code>decision: Pending</code> and <code>lastNotifiedAt</code> older than 24 hours.
          </p>
          <p style="font-size:12px;color:var(--aewttr-muted);margin:12px 0 0;">
            <code>BodyHtml</code> is Adaptive Card JSON (v1.4) — area label, subject, preview text, fact rows, and an Open in PULSE button.
          </p>
        </div>
      </div>
    `;
  });

  $("#nc-save", body).addEventListener("click", async () => {
    const db = window.AEWTTR.db;
    if (!db.notificationConfig) db.notificationConfig = { teamsEnabled: true, alsoSendEmail: true };
    db.notificationConfig.teamsEnabled = $("#nc-teams-enabled", body).checked;
    db.notificationConfig.alsoSendEmail = $("#nc-also-email", body).checked;
    try {
      if (typeof reanchorConfig === "function") reanchorConfig("notificationConfig", db.notificationConfig);
      await Repo.save("notificationConfig", db.notificationConfig);
      toast("Notification settings saved.", "success");
      if (typeof notifyLocalDataChanged === "function") notifyLocalDataChanged("notification-config");
      if (typeof logUserAction === "function") {
        logUserAction({
          action: "Update",
          area: "Admin",
          summary: `Updated notification settings — Teams ${db.notificationConfig.teamsEnabled ? "on" : "off"}, email ${db.notificationConfig.alsoSendEmail ? "on" : "off"}`
        });
      }
    } catch (e) {
      toast((e && e.message) || "Could not save notification settings.", "error");
    }
  });
}

function drawGroupsConfig(body) {
  body.innerHTML = requireAdmin(() => {
    const groups = typeof loadPulseGroups === "function" ? loadPulseGroups() : [];
    const sorted = groups.slice().sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")));
    return `
      <div class="aewttr-card aewttr-card-pad" style="margin-bottom:16px;">
        <div class="toolbar-row" style="margin-bottom:0;">
          <div>
            <div class="side-panel-title" style="margin:0 0 4px;">People groups</div>
            <div style="font-size:12px;color:var(--aewttr-muted);" data-help>Groups appear in every people picker as a single unit. Document Review expands them into individuals. Project teams sync automatically into a group named after the project.</div>
          </div>
          <button class="btn-aewttr" id="btn-add-group"${tip("Create a new group")}><i class="bx bx-plus"></i> New group</button>
        </div>
      </div>
      <div class="aewttr-card">
        <table class="aewttr-table">
          <thead><tr><th>Name</th><th>Source</th><th>Members</th><th></th></tr></thead>
          <tbody>
            ${sorted.length ? sorted.map((g) => `
              <tr style="cursor:default;" data-group-id="${escapeHtml(g.id)}">
                <td><strong>${escapeHtml(g.name)}</strong></td>
                <td>${g.source === "project" ? `<span class="kc-badge">Project${g.projectId ? ` · ${escapeHtml(g.projectId)}` : ""}</span>` : `<span class="kc-badge">Manual</span>`}</td>
                <td style="font-size:12.5px;color:var(--aewttr-muted);">${(g.members || []).length
                  ? escapeHtml((g.members || []).map((m) => m.name).join(", "))
                  : "No members"}</td>
                <td style="text-align:right;white-space:nowrap;">
                  <button class="btn-aewttr-outline btn-aewttr-sm" data-edit-group="${escapeHtml(g.id)}"${tip(`Edit ${escapeHtml(g.name)}`)}>Edit</button>
                  ${g.source === "project" ? "" : `<button class="btn-danger-outline btn-aewttr-sm" data-delete-group="${escapeHtml(g.id)}"${tip(`Delete ${escapeHtml(g.name)}`)}>Delete</button>`}
                </td>
              </tr>`).join("") : `<tr style="cursor:default;"><td colspan="4"><div class="empty-state">No groups yet — create one, or add people to a project to auto-create its group.</div></td></tr>`}
          </tbody>
        </table>
      </div>`;
  });
  if (!window.AEWTTR.db.user.isAdmin) return;
  $("#btn-add-group", body).addEventListener("click", () => openAdminGroupEditor(null, () => drawGroupsConfig(body)));
  $all("[data-edit-group]", body).forEach((btn) => btn.addEventListener("click", () => {
    const group = loadPulseGroups().find((g) => g.id === btn.dataset.editGroup);
    openAdminGroupEditor(group, () => drawGroupsConfig(body));
  }));
  $all("[data-delete-group]", body).forEach((btn) => btn.addEventListener("click", async () => {
    const group = loadPulseGroups().find((g) => g.id === btn.dataset.deleteGroup);
    if (!group || group.source === "project") return;
    const ok = await confirmDialog({
      title: "Delete group",
      message: `Delete "${group.name}"? This cannot be undone.`,
      confirmLabel: "Delete",
      danger: true
    });
    if (!ok) return;
    try {
      await deletePulseGroup(group.id);
      toast("Group deleted.", "success");
      drawGroupsConfig(body);
    } catch (e) {
      toast((e && e.friendly) || (e && e.message) || "Could not delete group.", "error");
    }
  }));
}

function openAdminGroupEditor(group, onDone) {
  const isProject = !!(group && group.source === "project");
  const members = ((group && group.members) || []).map((m) => ({
    id: m.id || "",
    name: m.name,
    email: m.email || ""
  }));
  const modal = openModal(`
    <div class="aewttr-modal-head">
      <h3>${group ? (isProject ? "Edit project group" : "Edit group") : "New group"}</h3>
      <button class="aewttr-modal-close">&times;</button>
    </div>
    <div class="aewttr-modal-body">
      ${isProject ? `<p class="docreview-hint" style="margin-top:0;">This group syncs from project <strong>${escapeHtml(group.projectId || "")}</strong>. Renaming keeps the link; membership still updates when project people change.</p>` : ""}
      <div class="form-row"><label>Group name</label>
        <input class="input-aewttr" id="ag-name" value="${escapeHtml(group ? group.name : "")}" placeholder="e.g. Range Ops reviewers" ${isProject ? "" : ""}>
      </div>
      <div class="form-row" style="margin-bottom:0;">
        <label>Members</label>
        <div class="traveler-picker traveler-picker--inline">
          <div id="ag-selected" class="traveler-chip-list"></div>
          <input class="input-aewttr" id="ag-input" placeholder="Search people…">
          <div id="ag-suggestions" class="traveler-suggestions"></div>
        </div>
      </div>
    </div>
    <div class="aewttr-modal-foot">
      <button class="btn-aewttr-ghost" id="ag-cancel">Cancel</button>
      <button class="btn-aewttr" id="ag-save">${group ? "Save" : "Create"}</button>
    </div>
  `);
  wirePeoplePicker(modal, members, { mount: "ag-selected", input: "ag-input", suggestions: "ag-suggestions" }, {
    includeGroups: false,
    allowManualEmail: true
  });
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#ag-cancel", modal).addEventListener("click", closeModal);
  $("#ag-save", modal).addEventListener("click", async () => {
    const name = $("#ag-name", modal).value.trim();
    if (!name) { toast("Group name is required.", "error"); return; }
    if (!members.length) { toast("Add at least one member.", "error"); return; }
    const me = (typeof currentUserIdentity === "function" ? currentUserIdentity() : (window.AEWTTR.db.user || {}));
    const saveBtn = $("#ag-save", modal);
    saveBtn.disabled = true;
    try {
      await upsertPulseGroup({
        id: (group && group.id) || uid("GRP"),
        name,
        members: members.map((m) => ({ id: m.id || "", name: m.name, email: m.email || "" })),
        source: (group && group.source) || "manual",
        projectId: (group && group.projectId) || "",
        createdBy: (group && group.createdBy) || me.name || "",
        createdByEmail: (group && group.createdByEmail) || me.email || "",
        createdDate: (group && group.createdDate) || (typeof todayIsoDate === "function" ? todayIsoDate() : new Date().toISOString().slice(0, 10)),
        _spId: group && group._spId
      });
      closeModal();
      toast(group ? "Group updated." : "Group created.", "success");
      if (onDone) onDone();
    } catch (e) {
      toast((e && e.friendly) || (e && e.message) || "Could not save group.", "error");
      saveBtn.disabled = false;
    }
  });
}

function drawListsConfig(body) {
  const cfg = getLocationConfig();

  const SECTIONS = [
    { id: "portfolios",      title: "Portfolios",         hint: "Shown in the Portfolio picker on projects.",              icon: "bx-briefcase",      list: () => cfg.portfolios      || (cfg.portfolios      = []) },
    { id: "contractors",     title: "Contractors",        hint: "Contractor company picker on projects and contractor travel.", icon: "bx-buildings",      list: () => cfg.contractors     || (cfg.contractors     = []), useKnown: typeof getKnownContractorNames === "function" ? getKnownContractorNames : null },
    { id: "contracts",       title: "Contracts",          hint: "Contract number/name picker on the funding tab.",         icon: "bx-file-blank",     list: () => cfg.contracts       || (cfg.contracts       = []) },
    { id: "endItems",        title: "End Item Configs",   hint: "Config End Item picker on projects.",                     icon: "bx-chip",           list: () => cfg.configEndItems  || (cfg.configEndItems  = []) },
    { id: "programs",        title: "Program Offices",    hint: "Program Office picker on the funding tab.",               icon: "bx-buildings",      list: () => cfg.programs        || (cfg.programs        = []), useKnown: typeof getKnownProgramNames === "function" ? getKnownProgramNames : null },
    { id: "atos",            title: "ATO Values",         hint: "ATO picker on the classification tab.",                   icon: "bx-shield-alt-2",   list: () => cfg.atos            || (cfg.atos            = []), useKnown: typeof getKnownAtoNames === "function" ? getKnownAtoNames : null },
    { id: "locations",       title: "Ranges / Locations", hint: "Location picker on projects.",                            icon: "bx-map-pin",        list: () => cfg.locations       || (cfg.locations       = []) },
    { id: "taskOrders",      title: "Task Orders",        hint: "Task Order picker on the funding tab.",                   icon: "bx-list-check",     list: () => cfg.taskOrders      || (cfg.taskOrders      = []), useKnown: typeof getKnownTaskOrderNames === "function" ? getKnownTaskOrderNames : null },
    { id: "fundingTypes",    title: "Funding Types",      hint: "Funding Type picker on the funding tab.",                 icon: "bx-dollar-circle",  list: () => cfg.fundingTypes    || (cfg.fundingTypes    = []), useKnown: typeof getKnownFundingTypeNames === "function" ? getKnownFundingTypeNames : null },
    { id: "fiscalYears",     title: "Fiscal Years",       hint: "Fiscal Year picker on the funding tab.",                  icon: "bx-calendar",       list: () => cfg.fiscalYears     || (cfg.fiscalYears     = []), useKnown: typeof getKnownFiscalYearNames === "function" ? getKnownFiscalYearNames : null },
    { id: "fundingStatuses", title: "Funding Statuses",   hint: "Funding Status picker on the funding tab.",               icon: "bx-trending-up",    list: () => cfg.fundingStatuses || (cfg.fundingStatuses = []), useKnown: typeof getKnownFundingStatusNames === "function" ? getKnownFundingStatusNames : null },
  ];

  const SECTION_META = {};
  const SINGULAR = {
    portfolios: "Portfolio", contractors: "Contractor Company", contracts: "Contract",
    endItems: "End Item Config", programs: "Program Office", atos: "ATO", locations: "Location",
    taskOrders: "Task Order", fundingTypes: "Funding Type", fiscalYears: "Fiscal Year", fundingStatuses: "Funding Status",
  };
  const PLACEHOLDER = {
    portfolios: "e.g. AEWTTR", contractors: "e.g. Harbor Systems LLC", contracts: "e.g. N00024-26-C-0440",
    endItems: "e.g. AN/ALQ-217", programs: "e.g. PEO IEWS", atos: "e.g. ATO-2024-001",
    locations: "e.g. White Sands Missile Range", taskOrders: "e.g. TO-2024-001",
    fundingTypes: "e.g. RDT&E", fiscalYears: "e.g. FY26", fundingStatuses: "e.g. Fully Funded",
  };
  SECTIONS.forEach(s => {
    SECTION_META[s.id] = { singular: SINGULAR[s.id], label: `${SINGULAR[s.id]} name`, placeholder: PLACEHOLDER[s.id], list: s.list };
  });

  let activeSection = window.AEWTTR.state && window.AEWTTR.state.adminListsSection || SECTIONS[0].id;
  /* Which contractor rows currently show their employee sub-list. Kept on
     window.AEWTTR.state (like adminListsSection above) rather than a plain
     local Set, so adding or renaming an employee — which redraws the whole
     page through saveAndRefresh — doesn't close the row the admin is looking
     at. */
  if (!window.AEWTTR.state.adminExpandedContractors) window.AEWTTR.state.adminExpandedContractors = [];
  const expandedContractors = new Set(window.AEWTTR.state.adminExpandedContractors);

  function getItems(s) {
    return s.useKnown ? s.useKnown() : (s.list() || []);
  }

  /* useKnown sections display the union of the stored catalog and whatever is
     actually in use on live records, so an item can appear here without being
     in the raw stored array — e.g. a contractor typed directly onto a project
     that nobody has ever added from this page. Rename/Remove only make sense
     against something this page actually stores, so this decides which a row
     gets: the stored index if there is one to act on, else a note that it
     comes from live data and reappears on its own while still referenced. */
  function storedIndex(s, name) {
    const key = normalizeContractorName ? normalizeContractorName(name).toLowerCase() : String(name).trim().toLowerCase();
    return s.list().findIndex((n) => String(n).trim().toLowerCase() === key);
  }

  function contractorRowExtra(company) {
    if (!getKnownContractorPeople) return "";
    const isOpen = expandedContractors.has(company);
    const key = typeof contractorPeopleKey === "function" ? contractorPeopleKey(company) : company.toLowerCase();
    const people = getKnownContractorPeople(company);
    const storedPeople = (cfg.contractorPeople && cfg.contractorPeople[key]) || [];
    const findStored = (name) => storedPeople.findIndex((p) => normalizeContractorName(p.name).toLowerCase() === name.toLowerCase());
    return `
      <button type="button" class="al-row-expand${isOpen ? " is-open" : ""}" data-people-toggle="${escapeHtml(company)}" aria-label="${isOpen ? "Hide" : "Show"} employees">
        <i class="bx bx-chevron-${isOpen ? "down" : "right"}"></i> ${people.length} ${people.length === 1 ? "employee" : "employees"}
      </button>
      ${isOpen ? `
      <div class="al-subpanel">
        ${people.length ? people.map((p) => {
          const idx = findStored(p.name);
          return `<div class="al-subrow">
            <span class="al-subrow-name">${escapeHtml(p.name)}${p.email ? ` <small>${escapeHtml(p.email)}</small>` : ""}</span>
            ${idx >= 0 ? `<div class="al-row-actions">
              <button class="btn-aewttr-ghost btn-aewttr-sm" data-people-rename="${escapeHtml(company)}" data-people-idx="${idx}"><i class="bx bx-pencil"></i></button>
              <button class="btn-danger-outline btn-aewttr-sm" data-people-remove="${escapeHtml(company)}" data-people-idx="${idx}"><i class="bx bx-trash"></i></button>
            </div>` : `<span class="al-subrow-derived">from project or travel data</span>`}
          </div>`;
        }).join("") : `<div class="al-subrow al-subrow-empty">No employees on file yet.</div>`}
        <button type="button" class="btn-aewttr-outline btn-aewttr-sm" data-people-add="${escapeHtml(company)}" style="margin-top:8px;"><i class="bx bx-plus"></i> Add Employee</button>
      </div>` : ""}`;
  }

  function renderPanel(sectionId) {
    const s = SECTIONS.find(x => x.id === sectionId) || SECTIONS[0];
    const items = getItems(s);
    const isContractors = s.id === "contractors";
    return `
      <div class="al-panel-header">
        <div>
          <div class="al-panel-title"><i class="bx ${escapeHtml(s.icon)}"></i> ${escapeHtml(s.title)}</div>
          <div class="al-panel-hint">${escapeHtml(s.hint)}</div>
        </div>
        <button class="btn-aewttr btn-aewttr-sm" data-lists-add="${escapeHtml(s.id)}"><i class="bx bx-plus"></i> Add ${escapeHtml(SINGULAR[s.id] || s.title)}</button>
      </div>
      <div class="al-panel-list">
        ${items.length ? items.map((item) => {
          const idx = storedIndex(s, item);
          /* Each contractor's row and its (possibly open) employee sub-panel
             share one wrapper, so they stay stacked together in source order
             — sibling <div>s here would all lay out at the .al-panel-list
             flex level and the sub-panels would end up grouped at the bottom
             instead of under their own row. */
          return `
          <div class="al-item">
            <div class="al-row${isContractors ? " al-row--expandable" : ""}">
              <div class="al-row-icon"><i class="bx bx-tag-alt"></i></div>
              <span class="al-row-name">${escapeHtml(item)}</span>
              ${idx >= 0 ? `<div class="al-row-actions">
                <button class="btn-aewttr-ghost btn-aewttr-sm" data-lists-rename="${escapeHtml(s.id)}" data-lists-idx="${idx}"><i class="bx bx-pencil"></i> Rename</button>
                <button class="btn-danger-outline btn-aewttr-sm" data-lists-remove="${escapeHtml(s.id)}" data-lists-idx="${idx}"><i class="bx bx-trash"></i></button>
              </div>` : `<span class="al-subrow-derived">from project${isContractors ? " or travel" : ""} data</span>`}
            </div>
            ${isContractors ? contractorRowExtra(item) : ""}
          </div>`;
        }).join("") : `
          <div class="al-empty-state">
            <i class="bx bx-inbox" style="font-size:28px;color:var(--aewttr-muted);"></i>
            <p>No ${escapeHtml(s.title.toLowerCase())} yet.</p>
            <button class="btn-aewttr-outline btn-aewttr-sm" data-lists-add="${escapeHtml(s.id)}"><i class="bx bx-plus"></i> Add first entry</button>
          </div>`}
      </div>`;
  }

  body.innerHTML = requireAdmin(() => `
    <div class="al-shell">
      <nav class="al-nav">
        ${SECTIONS.map(s => {
          const count = getItems(s).length;
          return `<button type="button" class="al-nav-item${s.id === activeSection ? " is-active" : ""}" data-al-section="${escapeHtml(s.id)}">
            <i class="bx ${escapeHtml(s.icon)} al-nav-icon"></i>
            <span class="al-nav-label">${escapeHtml(s.title)}</span>
            <span class="al-nav-count">${count}</span>
          </button>`;
        }).join("")}
      </nav>
      <div class="al-panel" id="al-panel">
        ${renderPanel(activeSection)}
      </div>
    </div>
  `);

  if (!window.AEWTTR.db.user.isAdmin) return;

  const saveAndRefresh = async () => {
    if (typeof reanchorConfig === "function") reanchorConfig("locationConfig", cfg);
    await Repo.save("locationConfig", cfg);
    if (typeof notifyLocalDataChanged === "function") notifyLocalDataChanged("location-config");
    drawListsConfig(body);
  };

  // nav switching — re-render panel in-place without full redraw
  $all("[data-al-section]", body).forEach((btn) => {
    btn.addEventListener("click", () => {
      activeSection = btn.dataset.alSection;
      if (window.AEWTTR.state) window.AEWTTR.state.adminListsSection = activeSection;
      $all("[data-al-section]", body).forEach(b => b.classList.toggle("is-active", b.dataset.alSection === activeSection));
      const panel = $("#al-panel", body);
      if (panel) panel.innerHTML = renderPanel(activeSection);
      wireListEvents();
    });
  });

  function wireListEvents() {
    $all("[data-lists-add]", body).forEach((btn) => {
      btn.addEventListener("click", async () => {
        const meta = SECTION_META[btn.dataset.listsAdd];
        if (!meta) return;
        const name = (await promptDialog({ title: `Add ${meta.singular}`, label: meta.label, placeholder: meta.placeholder }) || "").trim();
        if (!name) return;
        const list = meta.list();
        if (list.some((n) => n.toLowerCase() === name.toLowerCase())) { toast(`That ${meta.singular.toLowerCase()} already exists.`, "error"); return; }
        list.push(name);
        list.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
        await saveAndRefresh();
        toast(`${meta.singular} added`, "success");
      });
    });

    $all("[data-lists-rename]", body).forEach((btn) => {
      btn.addEventListener("click", async () => {
        const meta = SECTION_META[btn.dataset.listsRename];
        if (!meta) return;
        const idx = Number(btn.dataset.listsIdx);
        const list = meta.list();
        const current = list[idx];
        const name = (await promptDialog({ title: `Rename ${meta.singular}`, label: "Name", value: current }) || "").trim();
        if (!name || name === current) return;
        if (list.some((n, i) => i !== idx && n.toLowerCase() === name.toLowerCase())) { toast("That name already exists.", "error"); return; }
        list[idx] = name;
        list.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
        await saveAndRefresh();
        toast("Renamed", "success");
      });
    });

    $all("[data-lists-remove]", body).forEach((btn) => {
      btn.addEventListener("click", async () => {
        const meta = SECTION_META[btn.dataset.listsRemove];
        if (!meta) return;
        const idx = Number(btn.dataset.listsIdx);
        const list = meta.list();
        const item = list[idx];
        const ok = await confirmDialog({ title: `Remove ${meta.singular}`, message: `Remove "${item}" from the selectable list? Projects that already use it keep it until edited.`, confirmLabel: "Remove", danger: true });
        if (!ok) return;
        list.splice(idx, 1);
        await saveAndRefresh();
        toast("Removed", "success");
      });
    });

    $all("[data-people-toggle]", body).forEach((btn) => {
      btn.addEventListener("click", () => {
        const company = btn.dataset.peopleToggle;
        if (expandedContractors.has(company)) expandedContractors.delete(company);
        else expandedContractors.add(company);
        window.AEWTTR.state.adminExpandedContractors = [...expandedContractors];
        const panel = $("#al-panel", body);
        if (panel) panel.innerHTML = renderPanel(activeSection);
        wireListEvents();
      });
    });

    function contractorPeopleStore(company) {
      const key = typeof contractorPeopleKey === "function" ? contractorPeopleKey(company) : company.toLowerCase();
      if (!cfg.contractorPeople || typeof cfg.contractorPeople !== "object") cfg.contractorPeople = {};
      if (!Array.isArray(cfg.contractorPeople[key])) cfg.contractorPeople[key] = [];
      return cfg.contractorPeople[key];
    }

    $all("[data-people-add]", body).forEach((btn) => {
      btn.addEventListener("click", async () => {
        const company = btn.dataset.peopleAdd;
        const name = (await promptDialog({ title: "Add Employee", label: "Full name", placeholder: "e.g. Dana Reyes" }) || "").trim();
        if (!name) return;
        const store = contractorPeopleStore(company);
        if (store.some((p) => p.name.toLowerCase() === name.toLowerCase())) { toast("That employee already exists.", "error"); return; }
        const email = (await promptDialog({ title: "Email", label: "Email (optional)", placeholder: "name@company.com" }) || "").trim();
        store.push({ name, email });
        store.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));
        expandedContractors.add(company);
        window.AEWTTR.state.adminExpandedContractors = [...expandedContractors];
        await saveAndRefresh();
        toast("Employee added", "success");
      });
    });

    $all("[data-people-rename]", body).forEach((btn) => {
      btn.addEventListener("click", async () => {
        const company = btn.dataset.peopleRename;
        const idx = Number(btn.dataset.peopleIdx);
        const store = contractorPeopleStore(company);
        const current = store[idx];
        if (!current) return;
        const name = (await promptDialog({ title: "Rename Employee", label: "Full name", value: current.name }) || "").trim();
        if (!name) return;
        if (store.some((p, i) => i !== idx && p.name.toLowerCase() === name.toLowerCase())) { toast("That employee already exists.", "error"); return; }
        current.name = name;
        store.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));
        expandedContractors.add(company);
        window.AEWTTR.state.adminExpandedContractors = [...expandedContractors];
        await saveAndRefresh();
        toast("Renamed", "success");
      });
    });

    $all("[data-people-remove]", body).forEach((btn) => {
      btn.addEventListener("click", async () => {
        const company = btn.dataset.peopleRemove;
        const idx = Number(btn.dataset.peopleIdx);
        const store = contractorPeopleStore(company);
        const person = store[idx];
        if (!person) return;
        const ok = await confirmDialog({ title: "Remove Employee", message: `Remove "${person.name}" from ${company}'s employee list? Travel requests that already used them keep their name.`, confirmLabel: "Remove", danger: true });
        if (!ok) return;
        store.splice(idx, 1);
        expandedContractors.add(company);
        window.AEWTTR.state.adminExpandedContractors = [...expandedContractors];
        await saveAndRefresh();
        toast("Removed", "success");
      });
    });
  }

  wireListEvents();
}

function drawLocationConfig(body) {
  body.innerHTML = requireAdmin(() => {
    const cfg = getLocationConfig();
    return `
      <div class="aewttr-card aewttr-card-pad" style="margin-bottom:16px;">
        <div class="toolbar-row" style="margin-bottom:0;">
          <div>
            <div class="side-panel-title" style="margin:0 0 4px;">Ranges / Locations</div>
            <div style="font-size:12px;color:var(--aewttr-muted);" data-help>These show up as the selectable Project Location(s) list when editing a project.</div>
          </div>
          <button class="btn-aewttr" id="btn-add-location"${tip("Add a new selectable location")}><i class="bx bx-plus"></i> Add Location</button>
        </div>
      </div>
      <div class="aewttr-card">
        <table class="aewttr-table">
          <thead><tr><th>Name</th><th></th></tr></thead>
          <tbody>
            ${cfg.locations.length ? cfg.locations.map((loc, index) => `
              <tr style="cursor:default;" data-loc-index="${index}">
                <td><strong>${escapeHtml(loc)}</strong></td>
                <td style="text-align:right;">
                  <button class="btn-aewttr-outline btn-aewttr-sm" data-rename-loc="${index}"${tip(`Rename ${escapeHtml(loc)}`)}>Rename</button>
                  <button class="btn-danger-outline btn-aewttr-sm" data-remove-loc="${index}"${tip(`Remove ${escapeHtml(loc)}`)}>Remove</button>
                </td>
              </tr>`).join("") : `<tr style="cursor:default;"><td colspan="2"><div class="empty-state">No locations yet — add one to make it selectable on projects.</div></td></tr>`}
          </tbody>
        </table>
      </div>`;
  });
  if (!window.AEWTTR.db.user.isAdmin) return;
  const cfg = getLocationConfig();
  const saveLocations = async () => {
    if (typeof reanchorConfig === "function") reanchorConfig("locationConfig", cfg);
    await Repo.save("locationConfig", cfg);
    drawLocationConfig(body);
    if (typeof notifyLocalDataChanged === "function") notifyLocalDataChanged("location-config");
  };
  $("#btn-add-location", body).addEventListener("click", async () => {
    const name = (await promptDialog({ title: "Add Location", label: "Location name", placeholder: "e.g. White Sands Missile Range" }) || "").trim();
    if (!name) return;
    if (cfg.locations.some((loc) => loc.toLowerCase() === name.toLowerCase())) {
      toast("That location already exists.", "error");
      return;
    }
    cfg.locations.push(name);
    await saveLocations();
    toast("Location added", "success");
  });
  $all("[data-rename-loc]", body).forEach((btn) => btn.addEventListener("click", async () => {
    const index = Number(btn.dataset.renameLoc);
    const current = cfg.locations[index];
    const name = (await promptDialog({ title: "Rename Location", label: "Location name", value: current }) || "").trim();
    if (!name || name === current) return;
    if (cfg.locations.some((loc, i) => i !== index && loc.toLowerCase() === name.toLowerCase())) {
      toast("That location already exists.", "error");
      return;
    }
    cfg.locations[index] = name;
    await saveLocations();
    toast("Location renamed", "success");
  }));
  $all("[data-remove-loc]", body).forEach((btn) => btn.addEventListener("click", async () => {
    const index = Number(btn.dataset.removeLoc);
    const ok = await confirmDialog({
      title: "Remove location",
      message: `Remove "${cfg.locations[index]}" from the selectable list? Projects that already have it selected keep it until edited.`,
      confirmLabel: "Remove",
      danger: true
    });
    if (!ok) return;
    cfg.locations.splice(index, 1);
    await saveLocations();
    toast("Location removed", "success");
  }));
}

function drawUserAccess(body) {
  if (window.AEWTTR.mode === "sharepoint") {
    body.innerHTML = requireAdmin(() => `
      <div class="lock-block">
        <i class="bx bx-id-card"></i>
        <h3>Manage roles from the Users page</h3>
        <p>This site's roster now comes from SharePoint site users. Add people, sync roles, and change access from <strong>Users</strong> in the main nav — this keeps one source of truth instead of two.</p>
        <button class="btn-aewttr" id="btn-goto-users"${tip("Open user access management")}>Open Users</button>
      </div>`);
    const btn = $("#btn-goto-users", body);
    if (btn) btn.addEventListener("click", () => navigate("users"));
    return;
  }
  body.innerHTML = requireAdmin(() => {
    const db = window.AEWTTR.db;
    return `
      <div class="aewttr-card aewttr-card-pad" style="margin-bottom:16px;">
        <div class="toolbar-row" style="margin-bottom:0;">
          <div>
            <div class="side-panel-title" style="margin:0 0 4px;">Team Members</div>
            <div style="font-size:12px;color:var(--aewttr-muted);" data-help>Admins can add people here, assign access, and then include them in the weekly meeting roster.</div>
          </div>
          <button class="btn-aewttr" id="btn-add-member"${tip("Add a new team member to the roster")}><i class="bx bx-user-plus"></i> Add Team Member</button>
        </div>
      </div>
      <div class="aewttr-card">
        <table class="aewttr-table">
          <thead><tr><th>Name</th><th>Role</th><th>SPO Access</th><th>Power BI Access</th><th>Weekly</th><th>Last Login</th><th></th></tr></thead>
          <tbody>
            ${db.members.map(m => `
              <tr style="cursor:default;" data-mid="${m.id}">
                <td><strong>${escapeHtml(m.name)}</strong></td>
                <td>${escapeHtml(m.role)}</td>
                <td>${escapeHtml(m.spo)}</td>
                <td>${escapeHtml(m.powerbi)}</td>
                <td>${weeklyMeetingDb().roster.includes(m.id) ? `<span class="kc-badge">In Weekly</span>` : `—`}</td>
                <td>${fmtDate(m.lastLogin)}</td>
                <td><button class="btn-aewttr-outline btn-aewttr-sm" data-edit="${m.id}"${tip(`Edit access for ${m.name}`)}>Edit</button></td>
              </tr>`).join("")}
          </tbody>
        </table>
      </div>`;
  });
  if (!window.AEWTTR.db.user.isAdmin) return;
  $("#btn-add-member", body).addEventListener("click", () => openAddUserModal(() => drawUserAccess(body)));
  $all("[data-edit]", body).forEach(b => b.addEventListener("click", () => openEditUserModal(b.dataset.edit, () => drawUserAccess(body))));
}

function openAddUserModal(onDone) {
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>Add Team Member</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body">
      <div class="form-row"><label>Name</label><input class="input-aewttr" id="au-name" placeholder="First Last"></div>
      <div class="form-row"><label>Role</label>
        <select class="select-aewttr" id="au-role">
          <option value="Employee" selected>Employee</option>
          <option value="Contractor">Contractor</option>
          <option value="Guest">Guest</option>
          <option value="Product Manager">Product Manager</option>
          <option value="Engineer">Engineer</option>
          <option value="Developer">Developer</option>
          <option value="ISSO">ISSO</option>
          <option value="Range POC">Range POC</option>
        </select>
      </div>
      <div class="form-grid-2">
        <div class="form-row"><label>SPO Access</label>
          <select class="select-aewttr" id="au-spo">${["None", "View", "Full"].map(c => `<option ${c === "View" ? "selected" : ""}>${c}</option>`).join("")}</select>
        </div>
        <div class="form-row"><label>Power BI Access</label>
          <select class="select-aewttr" id="au-pbi">${["None", "View", "Full"].map(c => `<option ${c === "View" ? "selected" : ""}>${c}</option>`).join("")}</select>
        </div>
      </div>
      <div class="form-row"><label>Admin Role</label>
        <select class="select-aewttr" id="au-admin"><option value="false" selected>Member</option><option value="true">Admin</option></select>
      </div>
      <label class="weekly-roster-option" style="margin-top:6px;">
        <input type="checkbox" id="au-weekly" checked>
        <div>
          <strong>Include in Weekly Meeting</strong>
          <span>Add this person to the EOS-style room rotation immediately.</span>
        </div>
      </label>
    </div>
    <div class="aewttr-modal-foot">
      <button class="btn-aewttr-ghost" id="au-cancel">Cancel</button>
      <button class="btn-aewttr" id="au-save">Add Member</button>
    </div>
  `, { wide: true });
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#au-cancel", modal).addEventListener("click", closeModal);
  $("#au-save", modal).addEventListener("click", () => {
    const db = window.AEWTTR.db;
    const name = $("#au-name", modal).value.trim();
    const role = $("#au-role", modal).value.trim();
    if (!name) { toast("Name is required", "error"); return; }
    const newId = "u" + (db.members.reduce((max, member) => Math.max(max, +(member.id || "").replace(/^u/, "") || 0), 0) + 1);
    db.members.push({
      id: newId,
      name,
      role: role || "Team Member",
      spo: $("#au-spo", modal).value,
      powerbi: $("#au-pbi", modal).value,
      isAdmin: $("#au-admin", modal).value === "true",
      lastLogin: new Date().toISOString().slice(0, 10)
    });
    const weekly = weeklyMeetingDb();
    if ($("#au-weekly", modal).checked && !weekly.roster.includes(newId)) weekly.roster.push(newId);
    ensureWeeklyCurrentSession();
    aewttrSaveStore();
    closeModal();
    toast("Team member added", "success");
    if (onDone) onDone();
  });
}

function openEditUserModal(id, onDone) {
  const m = window.AEWTTR.db.members.find(x => x.id === id);
  const weekly = weeklyMeetingDb();
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>Edit Access — ${escapeHtml(m.name)}</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body">
      <div class="form-row"><label>Role</label>
        <select class="select-aewttr" id="eu-role">
          ${["Employee", "Contractor", "Guest", "Product Manager", "Engineer", "Developer", "ISSO", "Range POC"].includes(m.role) ? "" : `<option value="${escapeHtml(m.role)}">${escapeHtml(m.role)}</option>`}
          ${["Employee", "Contractor", "Guest", "Product Manager", "Engineer", "Developer", "ISSO", "Range POC"].map(r => `<option ${m.role === r ? "selected" : ""}>${r}</option>`).join("")}
        </select>
      </div>
      <div class="form-grid-2">
        <div class="form-row"><label>SPO Access</label>
          <select class="select-aewttr" id="eu-spo">${["None", "View", "Full"].map(c => `<option ${m.spo === c ? "selected" : ""}>${c}</option>`).join("")}</select>
        </div>
        <div class="form-row"><label>Power BI Access</label>
          <select class="select-aewttr" id="eu-pbi">${["None", "View", "Full"].map(c => `<option ${m.powerbi === c ? "selected" : ""}>${c}</option>`).join("")}</select>
        </div>
      </div>
      <div class="form-row"><label>Admin Role</label>
        <select class="select-aewttr" id="eu-admin"><option value="false" ${m.isAdmin ? "" : "selected"}>Member</option><option value="true" ${m.isAdmin ? "selected" : ""}>Admin</option></select>
      </div>
      <label class="weekly-roster-option" style="margin-top:6px;">
        <input type="checkbox" id="eu-weekly" ${weekly.roster.includes(m.id) ? "checked" : ""}>
        <div>
          <strong>Include in Weekly Meeting</strong>
          <span>Admins can add or remove people from the room rotation here.</span>
        </div>
      </label>
    </div>
    <div class="aewttr-modal-foot">
      <button class="btn-aewttr-ghost" id="eu-cancel">Cancel</button>
      <button class="btn-aewttr" id="eu-save">Save Changes</button>
    </div>
  `);
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#eu-cancel", modal).addEventListener("click", closeModal);
  $("#eu-save", modal).addEventListener("click", () => {
    m.role = $("#eu-role", modal).value;
    m.spo = $("#eu-spo", modal).value;
    m.powerbi = $("#eu-pbi", modal).value;
    m.isAdmin = $("#eu-admin", modal).value === "true";
    const onWeekly = $("#eu-weekly", modal).checked;
    if (onWeekly && !weekly.roster.includes(m.id)) weekly.roster.push(m.id);
    if (!onWeekly) weekly.roster = weekly.roster.filter(id => id !== m.id);
    ensureWeeklyCurrentSession();
    aewttrSaveStore();
    closeModal();
    toast("Access updated", "success");
    if (onDone) onDone();
  });
}

function drawDemoData(body) {
  body.innerHTML = requireAdmin(() => `
    <div class="aewttr-card aewttr-card-pad">
      <div class="side-panel-title">Demo Portfolio Seed</div>
      <p style="margin-top:0;color:var(--aewttr-muted);">Creates sample projects with tracker tasks, status notes, and tasks assigned to <strong>${escapeHtml(window.AEWTTR.db.user.name)}</strong> for Monday demos.</p>
      <button class="btn-aewttr" id="demo-seed-btn"><i class="bx bx-data"></i> Load Demo Projects</button>
      <p style="font-size:12px;color:var(--aewttr-muted);margin-top:14px;">Safe to run multiple times — existing projects are skipped.</p>
    </div>
  `);
  if (!canCurrentUserAccessAdmin()) return;
  $("#demo-seed-btn", body).addEventListener("click", () => seedDemoPortfolio(body));
}

async function seedDemoPortfolio(body) {
  const db = window.AEWTTR.db;
  const userName = db.user.name || "Demo User";
  const today = new Date().toISOString().slice(0, 10);
  const samples = [
    { id: "PULSE-01", name: "Radar Integration", team: "Systems", priority: "High", status: "In Progress", tasks: ["Finalize interface spec", "Complete bench test", "Prepare stakeholder brief"] },
    { id: "PULSE-02", name: "Ship Check Support", team: "Operations", priority: "Medium", status: "Active", tasks: ["Coordinate travel window", "Update tracker milestones", "Review contractor deliverables"] },
    { id: "PULSE-03", name: "Training Pipeline", team: "Training", priority: "Medium", status: "Planning", tasks: ["Draft curriculum outline", "Schedule instructor sessions", "Publish readiness checklist"] },
    { id: "PULSE-04", name: "Document Control", team: "Quality", priority: "High", status: "In Review", tasks: ["Route SOP for concurrence", "Close open review comments", "Archive signed revision"] },
    { id: "PULSE-05", name: "Executive Dashboard", team: "Leadership", priority: "High", status: "On Track", tasks: ["Define executive KPIs", "Validate operations metrics", "Export portfolio snapshot"] }
  ];
  let created = 0;
  for (const sample of samples) {
    if (db.projects.some((p) => p.id === sample.id)) continue;
    const project = {
      id: sample.id,
      name: sample.name,
      team: sample.team,
      priority: sample.priority,
      tools: "Tracker, Meeting, Documents",
      stakeholders: "IPT Leadership",
      rag: "Green",
      status: sample.status
    };
    db.projects.push(project);
    const extra = ensureProjectExtra(sample.id);
    extra.history.unshift({ date: today, author: userName, note: `Demo project seeded for ${sample.name}.` });
    extra.risks = sample.status === "In Review" ? [{ text: "Pending reviewer concurrence", level: "Medium" }] : [];
    db.ganttTasks[sample.id] = sample.tasks.map((title, index) => ({
      id: uid("g"),
      title,
      assignee: index === 0 ? userName : ((db.members[index % Math.max(db.members.length, 1)] || {}).name || "Unassigned"),
      start: today,
      end: new Date(Date.now() + (index + 7) * 86400000).toISOString().slice(0, 10),
      status: index === 0 ? "In Progress" : "Not Started",
      health: index === 2 ? "At Risk" : "On Track",
      notes: "",
      subtasks: []
    }));
    if (typeof Repo !== "undefined" && Repo.save) {
      await Repo.save("project", project);
      for (const task of db.ganttTasks[sample.id]) {
        await Repo.save("actionItem", task, { projectCode: sample.id, source: "Tracker" });
      }
    }
    created += 1;
  }
  if (typeof Repo !== "undefined" && Repo.flush) await Repo.flush();
  toast(created ? `Added ${created} demo project${created === 1 ? "" : "s"}.` : "Demo projects already exist.", created ? "success" : "info");
  drawDemoData(body);
}

function drawApprovalQueue(body) {
  body.innerHTML = requireAdmin(() => {
    const db = window.AEWTTR.db;
    const pendingTravel = typeof sortTravelRequests === "function"
      ? sortTravelRequests(db.travelRequests.filter(r => r.status === "Pending" || r.status === "Pending Finance"))
      : db.travelRequests.filter(r => r.status === "Pending" || r.status === "Pending Finance");
    return `
      <div class="aewttr-card aewttr-card-pad">
        <div class="side-panel-title">Pending Approvals</div>
        ${pendingTravel.length ? `
          <table class="aewttr-table">
            <thead><tr><th>Type</th><th>Item</th><th>Requested by</th><th>Status</th><th></th></tr></thead>
            <tbody>
              ${pendingTravel.map(r => `
                <tr style="cursor:default;">
                  <td><span class="type-badge">${escapeHtml(r.formMode || r.requestType || "Travel")}</span></td>
                  <td>${r.id} — ${escapeHtml(r.destination || r.tripTitle || "Request")}</td>
                  <td>${escapeHtml(r.requester)}</td>
                  <td>${typeof statusPill === "function" ? statusPill(r.status) : escapeHtml(r.status)}</td>
                  <td><button class="btn-aewttr-outline btn-aewttr-sm" data-goto-travel="${escapeHtml(r.id)}"${tip("Open this request in Travel approvals")}>Review</button></td>
                </tr>`).join("")}
            </tbody>
          </table>
        ` : `<div class="empty-state">No pending approvals beyond travel right now.</div>`}
        <p style="font-size:12px;color:var(--aewttr-muted);margin-top:14px;" data-help>This queue is future-proofed for additional approval types beyond travel (document signatures, etc.).</p>
      </div>`;
  });
  if (!window.AEWTTR.db.user.isAdmin) return;
  $all("[data-goto-travel]", body).forEach(b => b.addEventListener("click", () => navigate("travel/approve", { tr: b.dataset.gotoTravel })));
}

function drawReleases(body) {
  if (!canCurrentUserAccessAdmin()) {
    body.innerHTML = requireAdmin(() => "");
    return;
  }

  const releases = [
    {
      id: "pulse-main",
      name: "Main PULSE Application",
      version: "v1.0.0",
      filename: "PULSE-v1.0.0.html",
      fsFilename: "AEWTTR-PULSE_v.20260727.html",
      badge: "Production Ready",
      badgeClass: "kc-badge",
      icon: "bx-rocket",
      description: "Full IPT operational workspace incorporating project management, travel tracking, weekly meeting notes with OneNote/DOCX import, document review workflows, and system setup.",
      command: "node apps/PULSE/scripts/build-travel-packages.js"
    },
    {
      id: "travel-mytravel",
      name: "PULSE Travel — My Travel",
      version: "v1.0.0",
      filename: "PULSE-Travel-MyTravel-v1.0.0.html",
      fsFilename: "Travel-MyTravel-v1.0.0.html",
      badge: "Portal Package",
      badgeClass: "kc-badge",
      icon: "bx-briefcase-alt-2",
      description: "Dedicated end-user travel portal for tracking individual trip requests, monitoring approval status, and submitting travel debriefs.",
      command: "node apps/PULSE/scripts/build-travel-packages.js"
    },
    {
      id: "travel-forms",
      name: "PULSE Travel — Request Forms",
      version: "v1.0.0",
      filename: "PULSE-Travel-Request-Forms-v1.0.0.html",
      fsFilename: "Travel-Request-Forms-v1.0.0.html",
      badge: "Form Package",
      badgeClass: "kc-badge",
      icon: "bx-file",
      description: "Streamlined single-purpose web part for filing new TDY, local travel, and leave requests directly within SharePoint.",
      command: "node apps/PULSE/scripts/build-travel-packages.js"
    },
    {
      id: "travel-calendar",
      name: "PULSE Travel — Calendar",
      version: "v1.0.0",
      filename: "PULSE-Travel-Calendar-v1.0.0.html",
      fsFilename: "Travel-Calendar-v1.0.0.html",
      badge: "Calendar Package",
      badgeClass: "kc-badge",
      icon: "bx-calendar",
      description: "Shared organizational travel and leave calendar workspace powered by FullCalendar and SharePoint list sync.",
      command: "node apps/PULSE/scripts/build-travel-packages.js"
    },
    {
      id: "pulse-code",
      name: "PULSE CODE",
      version: "v1.0.0",
      filename: "PULSE-CODE-v1.0.0.html",
      fsFilename: "PULSE-CODE-v1.0.0.html",
      badge: "Utility Package",
      badgeClass: "kc-badge",
      icon: "bx-code-alt",
      description: "In-browser code editor and file manager for SharePoint site assets and embedded web parts.",
      command: "node apps/PULSE-CODE/scripts/build-sharepoint-package.js"
    },
    {
      id: "pulse-tickets",
      name: "PULSE Tickets",
      version: "v1.0.0",
      filename: "PULSE-Tickets-v1.0.0.html",
      fsFilename: "PULSE-Tickets-v1.0.0.html",
      badge: "Support Package",
      badgeClass: "kc-badge",
      icon: "bx-support",
      description: "Standalone issue tracking and user support ticket management web part.",
      command: "node apps/PULSE/scripts/build-travel-packages.js"
    },
    {
      id: "pulse-docs",
      name: "PULSE Documentation",
      version: "v1.0.0",
      filename: "PULSE-Documentation-v1.0.0.html",
      fsFilename: "PULSE-Documentation-v1.0.0.html",
      badge: "Docs Package",
      badgeClass: "kc-badge",
      icon: "bx-book-open",
      description: "Interactive technical handoff, SOPs, architectural reference, and user guide application.",
      command: "node apps/pulse-documentation/scripts/build-sharepoint-package.mjs"
    }
  ];

  body.innerHTML = `
    <div class="aewttr-card aewttr-card-pad">
      <div class="side-panel-title">App Releases &amp; SharePoint Firepit Packages</div>
      <p style="font-size:13px;color:var(--aewttr-muted);margin:4px 0 16px;">
        PULSE packages are compiled as standalone, single-file HTML web parts with zero external network dependencies. All scripts, stylesheets, fonts, and images are bundled directly for SharePoint deployment.
      </p>

      <div class="spsetup-status-strip" style="margin-bottom:20px;">
        <div class="spsetup-status-item"><span class="k">Total Releases</span><span class="v">7 Packages</span></div>
        <div class="spsetup-status-item"><span class="k">Target Host</span><span class="v"><span class="kc-badge">SharePoint Firepit</span></span></div>
        <div class="spsetup-status-item"><span class="k">Packaging Format</span><span class="v">Single-File HTML Bundle</span></div>
        <div class="spsetup-status-item"><span class="k">Build Status</span><span class="v" style="color:var(--aewttr-success,#27ae60);"><i class="bx bx-check-circle"></i> Verified Clean</span></div>
      </div>

      <div class="release-cards-grid" style="display:grid;grid-template-columns:repeat(auto-fill, minmax(340px, 1fr));gap:16px;">
        ${releases.map((rel) => `
          <div class="aewttr-card aewttr-card-pad" style="background:var(--aewttr-bg-card,#fff);border:1px solid var(--aewttr-border,#e2e8f0);border-radius:8px;display:flex;flex-direction:column;justify-content:space-between;padding:16px;">
            <div>
              <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">
                <div style="display:flex;align-items:center;gap:8px;">
                  <i class="bx ${rel.icon}" style="font-size:20px;color:var(--aewttr-primary,#1d4ed8);"></i>
                  <strong style="font-size:14.5px;color:var(--aewttr-fg,#0f172a);">${escapeHtml(rel.name)}</strong>
                </div>
                <span class="${rel.badgeClass}" style="font-size:11px;">${escapeHtml(rel.version)}</span>
              </div>
              <p style="font-size:12.5px;color:var(--aewttr-muted,#64748b);line-height:1.45;margin-bottom:12px;">${escapeHtml(rel.description)}</p>
            </div>
            <div>
              <div style="font-size:11.5px;color:var(--aewttr-muted,#64748b);background:var(--aewttr-bg-subtle,#f8fafc);padding:6px 10px;border-radius:4px;font-family:monospace;margin-bottom:10px;word-break:break-all;">
                <div><i class="bx bx-file"></i> releases/${escapeHtml(rel.filename)}</div>
                ${rel.fsFilename !== rel.filename ? `<div style="margin-top:2px;"><i class="bx bx-folder"></i> fs-packages/${escapeHtml(rel.fsFilename)}</div>` : ""}
              </div>
              <div style="display:flex;align-items:center;gap:8px;">
                <button type="button" class="btn-aewttr-outline btn-aewttr-sm btn-copy-cmd" data-cmd="${escapeHtml(rel.command)}" style="width:100%;justify-content:center;"><i class="bx bx-copy"></i> Copy build command</button>
              </div>
            </div>
          </div>
        `).join("")}
      </div>
    </div>
  `;

  $all(".btn-copy-cmd", body).forEach((btn) => {
    btn.addEventListener("click", () => {
      const cmd = btn.dataset.cmd;
      if (cmd && navigator.clipboard) {
        navigator.clipboard.writeText(cmd).then(() => {
          if (typeof toast === "function") toast("Build command copied to clipboard!", "success");
        }).catch(() => {
          if (typeof toast === "function") toast(`Build command: ${cmd}`, "info");
        });
      }
    });
  });
}

