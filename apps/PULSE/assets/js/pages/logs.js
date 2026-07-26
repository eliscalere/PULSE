PAGE_RENDERERS.logs = function () {
  if (!canCurrentUserAccessAdmin()) {
    navigate("dashboard");
    return;
  }
  if (!window.AEWTTR.state.auditLogFilter) {
    window.AEWTTR.state.auditLogFilter = { area: "All", action: "All", search: "" };
  }
  if (typeof initAuditLogStore === "function") initAuditLogStore();

  setTopbar("Activity Log", "Append-only record of user actions across PULSE.", `
    <button class="btn-aewttr-outline btn-aewttr-sm" id="audit-refresh"><i class="bx bx-refresh"></i> Refresh</button>
    <button class="btn-aewttr-outline btn-aewttr-sm" id="audit-export"><i class="bx bx-download"></i> Export CSV</button>
  `);

  const areas = ["All", "Travel", "Documents", "Projects", "Weekly", "Tickets", "Users", "Admin", "Navigation", "System"];
  const actions = ["All", "Create", "Update", "Delete", "Navigate", "Approve", "Deny", "Submit", "Export", "Login", "Other"];
  const st = window.AEWTTR.state.auditLogFilter;

  function getLiveEntries() {
    return (window.AEWTTR.db && window.AEWTTR.db.auditLog) || [];
  }

  function getFilteredRows() {
    const search = String(st.search || "").trim().toLowerCase();
    return getLiveEntries().filter((entry) => {
      if (st.area !== "All" && entry.area !== st.area) return false;
      if (st.action !== "All" && entry.action !== st.action) return false;
      if (!search) return true;
      const hay = [
        entry.summary,
        entry.actorName,
        entry.actorEmail,
        entry.recordId,
        entry.route,
        entry.area,
        entry.action
      ].join(" ").toLowerCase();
      return hay.includes(search);
    });
  }

  function rowHtml(rows) {
    const all = getLiveEntries();
    return rows.length ? rows.map((entry, index) => `
      <tr class="audit-log-row" data-audit-index="${index}">
        <td class="audit-log-when">${escapeHtml(fmtDateTime(entry.ts))}</td>
        <td>
          <strong>${escapeHtml(entry.actorName || "Unknown")}</strong>
          <div class="audit-log-sub">${escapeHtml(entry.actorEmail || entry.actorRole || "")}</div>
        </td>
        <td><span class="type-badge">${escapeHtml(entry.action || "Other")}</span></td>
        <td>${escapeHtml(entry.area || "System")}</td>
        <td>${escapeHtml(entry.summary || "")}</td>
        <td>${entry.detail ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-sm audit-log-detail-btn" data-audit-detail="${index}">Details</button>` : ""}</td>
      </tr>
      ${entry.detail ? `<tr class="audit-log-detail-row" id="audit-detail-${index}" hidden><td colspan="6"><pre class="audit-log-detail-pre">${escapeHtml(JSON.stringify(entry.detail, null, 2))}</pre></td></tr>` : ""}
    `).join("") : `<tr><td colspan="6"><div class="empty-state">No activity logged in the last 30 days.</div></td></tr>`;
  }

  function wireDetailButtons() {
    $all("[data-audit-detail]", $("#page-content")).forEach((button) => {
      button.addEventListener("click", () => {
        const row = $(`#audit-detail-${button.dataset.auditDetail}`, $("#page-content"));
        if (row) row.hidden = !row.hidden;
      });
    });
  }

  function renderTable() {
    const rows = getFilteredRows();
    const all = getLiveEntries();
    const meta = $("#audit-log-meta", $("#page-content"));
    const tbody = $("#audit-log-tbody", $("#page-content"));
    if (meta) meta.textContent = `${rows.length} of ${all.length} entries shown (last 30 days, all users)`;
    if (tbody) {
      tbody.innerHTML = rowHtml(rows);
      wireDetailButtons();
    }
    return rows;
  }

  function setProgress(msg) {
    const el = $("#audit-progress", $("#page-content"));
    if (el) { el.textContent = msg; el.hidden = !msg; }
  }

  /* Load all numbered audit lists from SP — last 30 days, all users. */
  async function refreshFromSharePoint() {
    if (typeof isSharePointMode !== "function" || !isSharePointMode()) return;
    const siteUrl = typeof currentSiteUrl === "function" ? currentSiteUrl() : "";
    if (!siteUrl || typeof sharePointAdapter === "undefined" || typeof spItemToAuditLog !== "function") return;

    const refreshBtn = $("#audit-refresh");
    if (refreshBtn) { refreshBtn.disabled = true; refreshBtn.innerHTML = `<i class="bx bx-loader-alt bx-spin"></i> Loading…`; }

    const since = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
    const filter = `ActionTime ge datetime'${since}'`;
    const base = "PULSE Audit Log";
    const merged = [];
    let listsChecked = 0;

    for (let n = 1; n <= 20; n++) {
      const listName = n === 1 ? base : `${base} ${n}`;
      setProgress(`Scanning ${listName}… (${merged.length} entries loaded so far)`);
      try {
        const items = await sharePointAdapter.getItems(siteUrl, listName, {
          top: 5000,
          orderby: "ActionTime desc",
          filter
        });
        listsChecked++;
        if (!items || !items.length) break;
        merged.push(...items.map(spItemToAuditLog));
        // If the list returned fewer items than cap, it's not full — try next list
        // in case log rotation created another one.
      } catch (e) {
        break; // list does not exist — done
      }
    }

    setProgress("");

    if (merged.length || listsChecked > 0) {
      window.AEWTTR.db.auditLog = merged.sort((a, b) => String(b.ts).localeCompare(String(a.ts)));
    }

    if (refreshBtn) { refreshBtn.disabled = false; refreshBtn.innerHTML = `<i class="bx bx-refresh"></i> Refresh`; }
    renderTable();
  }

  $("#page-content").innerHTML = `
    <div class="toolbar-row" style="margin-bottom:14px;">
      <div class="search-box"><i class="bx bx-search"></i><input id="audit-search" placeholder="Search actor, summary, record…" value="${escapeHtml(st.search)}"></div>
      <select class="select-aewttr" id="audit-area-filter" style="max-width:170px;">
        ${areas.map((area) => `<option value="${area}" ${st.area === area ? "selected" : ""}>${area === "All" ? "All Areas" : area}</option>`).join("")}
      </select>
      <select class="select-aewttr" id="audit-action-filter" style="max-width:170px;">
        ${actions.map((action) => `<option value="${action}" ${st.action === action ? "selected" : ""}>${action === "All" ? "All Actions" : action}</option>`).join("")}
      </select>
    </div>
    <div class="audit-progress-bar" id="audit-progress" hidden></div>
    <div class="aewttr-card">
      <div class="audit-log-meta" id="audit-log-meta"></div>
      <table class="aewttr-table audit-log-table">
        <thead>
          <tr>
            <th>When</th>
            <th>Actor</th>
            <th>Action</th>
            <th>Area</th>
            <th>Summary</th>
            <th></th>
          </tr>
        </thead>
        <tbody id="audit-log-tbody"></tbody>
      </table>
    </div>
  `;

  renderTable();

  // Auto-load from SP on page open
  refreshFromSharePoint();

  $("#audit-search", $("#page-content")).addEventListener("input", (event) => {
    window.AEWTTR.state.auditLogFilter.search = event.target.value;
    renderTable();
  });
  $("#audit-area-filter", $("#page-content")).addEventListener("change", (event) => {
    window.AEWTTR.state.auditLogFilter.area = event.target.value;
    renderTable();
  });
  $("#audit-action-filter", $("#page-content")).addEventListener("change", (event) => {
    window.AEWTTR.state.auditLogFilter.action = event.target.value;
    renderTable();
  });
  $("#audit-refresh").addEventListener("click", () => refreshFromSharePoint());
  $("#audit-export").addEventListener("click", async () => {
    const rows = getFilteredRows();
    const columns = ["When", "Actor", "Email", "Role", "Action", "Area", "Summary", "RecordId", "Route"];
    const dataRows = rows.map((entry) => [
      entry.ts,
      entry.actorName,
      entry.actorEmail,
      entry.actorRole,
      entry.action,
      entry.area,
      entry.summary,
      entry.recordId,
      entry.route
    ]);
    const service = window.AEWTTR && window.AEWTTR.ExportService;
    if (!service || typeof service.exportCsv !== "function") {
      toast("CSV export is unavailable in this package.", "error");
      return;
    }
    const fileName = `pulse-activity-log-${new Date().toISOString().slice(0, 10)}.csv`;
    const popup = typeof service.reserveOpen === "function" ? service.reserveOpen(fileName, "text/csv") : null;
    const result = await service.exportCsv(fileName, columns, dataRows, {
      popup,
      folderName: "Activity Logs"
    });
    if (result.mode === "sharepoint") toast("Activity log saved and opened in SharePoint.", "success");
    else if (result.uploadError) toast("SharePoint upload failed — downloaded a local copy instead.", "warn");
    else toast("Activity log downloaded.", "success");
    if (typeof logUserAction === "function") {
      logUserAction({ action: "Export", area: "Admin", summary: `Exported ${rows.length} activity log entries to CSV` });
    }
  });
};

function fmtDateTime(iso) {
  if (!iso) return "—";
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return String(iso).slice(0, 19).replace("T", " ");
  return date.toLocaleString(undefined, {
    month: "short",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit"
  });
}
