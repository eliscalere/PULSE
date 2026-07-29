PAGE_RENDERERS.tickets = function (parts) {
  if (!window.AEWTTR.state.ticketFilter) window.AEWTTR.state.ticketFilter = "All";
  if (!window.AEWTTR.state.ticketType) window.AEWTTR.state.ticketType = "All";
  if (!window.AEWTTR.state.ticketSearch) window.AEWTTR.state.ticketSearch = "";
  const db = window.AEWTTR.db;

  // If a ticket ID is in the route path, render the full-page detail
  const routeId = parts && parts[0] && /^TKT-\d+$/i.test(parts[0]) ? parts[0].toUpperCase() : null;
  if (routeId) {
    const t = db.tickets.find(x => x.id === routeId);
    if (t) { drawTicketDetail(t); return; }
    // Not found — fall through to list
  }

  const routeIntent = typeof consumeRouteIntent === "function" ? consumeRouteIntent("tickets") : null;
  setTopbar("Support Tickets", "Log blockers, bugs, access requests, and questions.", "");

  const STATUS_TABS = ["All", "Open", "In Progress", "Resolved"];
  const TYPES = ["All", "Blocker", "Bug", "Access", "Platform", "Question", "Feature Request"];

  function draw() {
    const st = window.AEWTTR.state;
    const counts = STATUS_TABS.reduce((acc, status) => {
      acc[status] = status === "All" ? db.tickets.length : db.tickets.filter((ticket) => ticket.status === status).length;
      return acc;
    }, {});
    const rows = db.tickets.filter(t => {
      if (st.ticketFilter !== "All" && t.status !== st.ticketFilter) return false;
      if (st.ticketType !== "All" && t.type !== st.ticketType) return false;
      if (st.ticketSearch && !(t.title.toLowerCase().includes(st.ticketSearch.toLowerCase()) || t.id.toLowerCase().includes(st.ticketSearch.toLowerCase()))) return false;
      return true;
    });
    const page = $("#page-content");
    page.innerHTML = `
      <div class="tickets-page">
        <header class="tickets-page-head">
          <div class="tickets-page-copy">
            <div class="tickets-page-label">Support workspace</div>
            <h2>Keep work moving</h2>
            <p>Track blockers, bugs, access needs, platform issues, and questions from intake through resolution.</p>
          </div>
          <button class="btn-aewttr" id="btn-new-ticket"${tip("Create a new support ticket")}><i class="bx bx-plus"></i> New Ticket</button>
        </header>

        <section class="tickets-controls" aria-label="Ticket filters">
          <div class="search-box tickets-search"><i class="bx bx-search"></i><input id="tkt-search" placeholder="Search by ID or title" value="${escapeHtml(st.ticketSearch)}"></div>
          <div class="filter-pills tickets-status-filters" id="tkt-status-pills">
            ${STATUS_TABS.map(f => `<button class="filter-pill ${st.ticketFilter === f ? "active" : ""}" data-f="${f}"${tip(f === "All" ? "Show all tickets" : `Show ${f.toLowerCase()} tickets`)}>${f}<span>${counts[f]}</span></button>`).join("")}
          </div>
          <select class="select-aewttr tickets-type-filter" id="tkt-type-filter" aria-label="Filter tickets by type">
            ${TYPES.map(t => `<option value="${t}" ${st.ticketType === t ? "selected" : ""}>${t === "All" ? "All Types" : t}</option>`).join("")}
          </select>
        </section>

        <section class="tickets-table-shell">
          <div class="tickets-table-meta"><strong>${rows.length}</strong> of ${db.tickets.length} ticket${db.tickets.length === 1 ? "" : "s"}</div>
          <div class="tickets-table-scroll">
            <table class="aewttr-table tickets-table">
              <thead><tr><th>Ticket</th><th>Title</th><th>Project</th><th>Type</th><th>Status</th><th>Opened</th><th>Reporter</th></tr></thead>
              <tbody>
                ${rows.length ? rows.map(t => `
                  <tr data-id="${t.id}" style="cursor:pointer;">
                    <td><span class="tickets-id">${escapeHtml(t.id)}</span></td><td><strong>${escapeHtml(t.title)}</strong></td>
                    <td>${escapeHtml(t.project || "—")}</td>
                    <td><span class="type-badge">${escapeHtml(t.type)}</span></td>
                    <td>${statusPill(t.status)}</td>
                    <td>${fmtDate(t.opened)}</td><td>${escapeHtml(t.reporter)}</td>
                  </tr>`).join("") : `<tr><td colspan="7"><div class="tickets-empty"><strong>No tickets found</strong><span>Adjust the filters or create a new ticket.</span></div></td></tr>`}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    `;
    $all("tr[data-id]", page).forEach(r => r.addEventListener("click", () => {
      navigate("tickets/" + r.dataset.id);
    }));
    $all("[data-f]", $("#tkt-status-pills")).forEach(b => b.addEventListener("click", () => { window.AEWTTR.state.ticketFilter = b.dataset.f; draw(); }));
    $("#tkt-type-filter").addEventListener("change", (e) => { window.AEWTTR.state.ticketType = e.target.value; draw(); });
    $("#tkt-search").addEventListener("input", (e) => { window.AEWTTR.state.ticketSearch = e.target.value; draw(); });
    $("#btn-new-ticket").addEventListener("click", () => openNewTicketModal(draw));
  }
  draw();
  if (routeIntent && routeIntent.ticket) {
    const ticket = db.tickets.find((item) => item.id === routeIntent.ticket);
    if (ticket) navigate("tickets/" + ticket.id);
  }
};

function drawTicketDetail(t) {
  setTopbar("Support Tickets", "Log blockers, bugs, access requests, and questions.", "");
  const db = window.AEWTTR.db;

  function render() {
    const page = $("#page-content");
    const updates = (t.updates || []);
    page.innerHTML = `
      <div class="ticket-detail-page">
        <div class="ticket-detail-topbar">
          <button class="btn-aewttr-ghost ticket-detail-back" id="tktd-back"><i class="bx bx-arrow-back"></i> All tickets</button>
          <div class="ticket-detail-topbar-actions">
            ${t.status !== "Resolved" ? `<button class="btn-aewttr-outline" id="tktd-resolve"${tip("Mark this ticket resolved")}>Resolve</button>` : ""}
            <button class="btn-aewttr-ghost tktd-delete-btn" id="tktd-delete" style="color:var(--aewttr-red,#DC2626);"${tip("Permanently delete this ticket")}>Delete</button>
          </div>
        </div>

        <div class="ticket-detail-layout">
          <!-- Left: main content -->
          <div class="ticket-detail-main">
            <div class="ticket-detail-header">
              <span class="tickets-id">${escapeHtml(t.id)}</span>
              <h1 class="ticket-detail-title">${escapeHtml(t.title)}</h1>
            </div>

            <div class="ticket-detail-section">
              <div class="ticket-detail-section-label">Description</div>
              <p class="ticket-detail-text">${escapeHtml(t.detail || "No description provided.")}</p>
            </div>

            <div class="ticket-detail-section">
              <div class="ticket-detail-section-label">Workaround</div>
              <p class="ticket-detail-text">${escapeHtml(t.workaround || "None documented.")}</p>
            </div>

            <div class="ticket-detail-section">
              <div class="ticket-detail-section-label">ESDP / Escalation Path</div>
              <p class="ticket-detail-text">${escapeHtml(t.esdp || "None documented.")}</p>
            </div>

            <div class="ticket-detail-section ticket-timeline-section">
              <div class="ticket-detail-section-label">Activity</div>
              <div class="ticket-timeline" id="tktd-timeline">
                ${updates.length ? updates.map(u => `
                  <div class="ticket-timeline-item">
                    <div class="ticket-timeline-meta">
                      <strong>${escapeHtml(u.author)}</strong>
                      <span class="ticket-timeline-date">${fmtDate(u.date)}</span>
                    </div>
                    <div class="ticket-timeline-body">${escapeHtml(u.text)}</div>
                    ${u.screenshot ? `<div class="ticket-timeline-screenshot"><img src="${u.screenshot}" alt="Attached screenshot" class="ticket-screenshot-img" data-screenshot></div>` : ""}
                  </div>
                `).join("") : `<div class="ticket-timeline-empty">No updates yet.</div>`}
              </div>

              <div class="ticket-update-composer" id="tktd-composer">
                <textarea class="textarea-aewttr ticket-update-textarea" id="tktd-update-text" placeholder="Add an update, note, or resolution step…" rows="3"></textarea>
                <div class="ticket-update-composer-actions">
                  <div class="ticket-screenshot-attach" id="tktd-screenshot-attach">
                    <button class="btn-aewttr-ghost btn-aewttr-sm ticket-attach-btn" id="tktd-attach-btn" type="button"${tip("Attach a screenshot to this update")}><i class="bx bx-image"></i> Attach screenshot</button>
                    <div class="ticket-screenshot-preview" id="tktd-screenshot-preview" hidden></div>
                    <input type="file" id="tktd-file-input" accept="image/*" style="display:none">
                  </div>
                  <button class="btn-aewttr btn-aewttr-sm" id="tktd-post-btn"${tip("Post this update")}>Post update</button>
                </div>
              </div>
            </div>
          </div>

          <!-- Right: metadata sidebar -->
          <aside class="ticket-detail-aside">
            <div class="ticket-detail-meta-card">
              <div class="ticket-meta-row">
                <span class="ticket-meta-label">Status</span>
                <select class="select-aewttr ticket-meta-select" id="tktd-status">
                  ${["Open", "In Progress", "Resolved"].map(s => `<option ${t.status === s ? "selected" : ""}>${s}</option>`).join("")}
                </select>
              </div>
              <div class="ticket-meta-row">
                <span class="ticket-meta-label">Type</span>
                <span class="v"><span class="type-badge">${escapeHtml(t.type)}</span></span>
              </div>
              <div class="ticket-meta-row">
                <span class="ticket-meta-label">Reporter</span>
                <span class="ticket-meta-value">${escapeHtml(t.reporter)}</span>
              </div>
              <div class="ticket-meta-row">
                <span class="ticket-meta-label">Opened</span>
                <span class="ticket-meta-value">${fmtDate(t.opened)}</span>
              </div>
              ${(t.affected || []).filter(Boolean).length ? `
              <div class="ticket-meta-row">
                <span class="ticket-meta-label">Projects</span>
                <span class="ticket-meta-value">${(t.affected || []).map(x => escapeHtml(x)).join(", ")}</span>
              </div>` : ""}
            </div>

            ${t.type === "Blocker" ? `
            <button class="btn-aewttr-outline ticket-meta-action" id="tktd-nmci"${tip("Generate an NMCI access request draft")}><i class="bx bx-file"></i> Draft NMCI</button>` : ""}
          </aside>
        </div>
      </div>
    `;

    // Back
    $("#tktd-back", page).addEventListener("click", () => navigate("tickets"));

    // Status change
    $("#tktd-status", page).addEventListener("change", (e) => {
      t.status = e.target.value;
      if (typeof Repo !== "undefined") Repo.save("ticket", t); else aewttrSaveStore();
      notifyReporterOfTicketUpdate(t, `Status changed to ${t.status}`);
      toast(`${t.id} status set to ${t.status}`, "success");
      render(); // re-render to update resolve button visibility
    });

    // Resolve
    const resolveBtn = $("#tktd-resolve", page);
    if (resolveBtn) resolveBtn.addEventListener("click", () => {
      t.status = "Resolved";
      if (typeof Repo !== "undefined") Repo.save("ticket", t); else aewttrSaveStore();
      notifyReporterOfTicketUpdate(t, "Status changed to Resolved");
      toast(`${t.id} resolved`, "success");
      render();
    });

    // Delete
    const deleteBtn = $("#tktd-delete", page);
    deleteBtn.addEventListener("click", () => {
      if (deleteBtn.dataset.confirmed !== "1") {
        deleteBtn.dataset.confirmed = "1";
        deleteBtn.textContent = "Confirm Delete";
        setTimeout(() => { if (deleteBtn.dataset.confirmed === "1") { deleteBtn.dataset.confirmed = ""; deleteBtn.textContent = "Delete"; } }, 3000);
        return;
      }
      const idx = db.tickets.findIndex(x => x.id === t.id);
      if (idx !== -1) db.tickets.splice(idx, 1);
      if (typeof Repo !== "undefined") Repo.remove("ticket", t); else aewttrSaveStore();
      navigate("tickets");
      toast(`${t.id} deleted`, "success");
    });

    // NMCI
    const nmciBtn = $("#tktd-nmci", page);
    if (nmciBtn) nmciBtn.addEventListener("click", () => toast("NMCI request drafted (prototype only)", "success"));

    // Screenshot attachment
    let pendingScreenshot = null;
    const attachBtn = $("#tktd-attach-btn", page);
    const fileInput = $("#tktd-file-input", page);
    const screenshotPreview = $("#tktd-screenshot-preview", page);

    attachBtn.addEventListener("click", () => fileInput.click());

    fileInput.addEventListener("change", () => {
      const file = fileInput.files && fileInput.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (e) => {
        pendingScreenshot = e.target.result;
        screenshotPreview.hidden = false;
        screenshotPreview.innerHTML = `
          <div class="ticket-screenshot-thumb-wrap">
            <img src="${pendingScreenshot}" alt="Pending screenshot" class="ticket-screenshot-thumb">
            <button class="ticket-screenshot-remove" id="tktd-remove-shot" type="button" aria-label="Remove screenshot">&times;</button>
          </div>`;
        $("#tktd-remove-shot", page).addEventListener("click", () => {
          pendingScreenshot = null;
          screenshotPreview.hidden = true;
          screenshotPreview.innerHTML = "";
          fileInput.value = "";
          attachBtn.innerHTML = '<i class="bx bx-image"></i> Attach screenshot';
        });
        attachBtn.innerHTML = '<i class="bx bx-check"></i> Screenshot attached';
      };
      reader.readAsDataURL(file);
    });

    // Paste screenshot anywhere on the composer
    const composer = $("#tktd-composer", page);
    composer.addEventListener("paste", (e) => {
      const items = e.clipboardData && e.clipboardData.items;
      if (!items) return;
      for (const item of items) {
        if (item.type.startsWith("image/")) {
          const file = item.getAsFile();
          if (!file) continue;
          const reader = new FileReader();
          reader.onload = (ev) => {
            pendingScreenshot = ev.target.result;
            screenshotPreview.hidden = false;
            screenshotPreview.innerHTML = `
              <div class="ticket-screenshot-thumb-wrap">
                <img src="${pendingScreenshot}" alt="Pasted screenshot" class="ticket-screenshot-thumb">
                <button class="ticket-screenshot-remove" id="tktd-remove-shot" type="button" aria-label="Remove screenshot">&times;</button>
              </div>`;
            $("#tktd-remove-shot", page).addEventListener("click", () => {
              pendingScreenshot = null;
              screenshotPreview.hidden = true;
              screenshotPreview.innerHTML = "";
              attachBtn.innerHTML = '<i class="bx bx-image"></i> Attach screenshot';
            });
            attachBtn.innerHTML = '<i class="bx bx-check"></i> Screenshot attached';
          };
          reader.readAsDataURL(file);
          e.preventDefault();
          break;
        }
      }
    });

    // Lightbox for screenshots
    $all("[data-screenshot]", page).forEach(img => {
      img.style.cursor = "zoom-in";
      img.addEventListener("click", () => {
        const lb = document.createElement("div");
        lb.className = "ticket-screenshot-lightbox";
        lb.innerHTML = `<div class="ticket-lightbox-backdrop"></div><div class="ticket-lightbox-content"><img src="${img.src}" alt="Screenshot"><button class="ticket-lightbox-close" aria-label="Close">&times;</button></div>`;
        document.body.appendChild(lb);
        const close = () => lb.remove();
        lb.querySelector(".ticket-lightbox-backdrop").addEventListener("click", close);
        lb.querySelector(".ticket-lightbox-close").addEventListener("click", close);
        document.addEventListener("keydown", function esc(e) { if (e.key === "Escape") { close(); document.removeEventListener("keydown", esc); } });
      });
    });

    // Post update
    $("#tktd-post-btn", page).addEventListener("click", () => {
      const text = $("#tktd-update-text", page).value.trim();
      if (!text) { toast("Enter an update before posting.", "error"); return; }
      t.updates = t.updates || [];
      const update = { date: new Date().toISOString().slice(0, 10), author: db.user.name, text };
      if (pendingScreenshot) update.screenshot = pendingScreenshot;
      t.updates.unshift(update);
      if (typeof Repo !== "undefined") Repo.save("ticket", t); else aewttrSaveStore();
      notifyReporterOfTicketUpdate(t, "New update posted");
      render();
    });
  }

  render();
}

function openNewTicketModal(onDone) {
  return openNewTicketModalWithOptions({}, onDone);
}

function openNewTicketModalWithOptions(opts, onDone) {
  opts = opts || {};
  const db = window.AEWTTR.db;
  const modal = openModal(`
    <div class="aewttr-modal-head ticket-modal-head"><div><span class="ticket-modal-label">Support request</span><h3>New Ticket</h3></div><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body ticket-modal-body">
      <div class="form-row"><label>Title</label><input class="input-aewttr" id="nt-title"></div>
      <div class="form-grid-2">
        <div class="form-row"><label>Project</label>
          <select class="select-aewttr" id="nt-project" ${opts.lockProject ? "disabled" : ""}><option value="">None</option>${db.projects.map(p => `<option value="${p.id}" ${opts.projectId === p.id ? "selected" : ""}>${escapeHtml(p.name || "Untitled project")}</option>`).join("")}</select>
        </div>
        <div class="form-row"><label>Type</label>
          <select class="select-aewttr" id="nt-type">${["Blocker", "Bug", "Access", "Platform", "Question", "Feature Request"].map(t => `<option>${t}</option>`).join("")}</select>
        </div>
      </div>
      ${opts.lockProject && opts.projectId ? `<input type="hidden" id="nt-project-hidden" value="${escapeHtml(opts.projectId)}">` : ""}
      <div class="form-row"><label>Detail</label><textarea class="textarea-aewttr" id="nt-detail"></textarea></div>
    </div>
    <div class="aewttr-modal-foot ticket-modal-foot">
      <button class="btn-aewttr-ghost" id="nt-cancel"${tip("Close without creating a ticket")}>Cancel</button>
      <button class="btn-aewttr" id="nt-save"${tip("Create this support ticket")}>Create Ticket</button>
    </div>
  `);
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#nt-cancel", modal).addEventListener("click", closeModal);
  $("#nt-save", modal).addEventListener("click", () => {
    const title = $("#nt-title", modal).value.trim();
    if (!title) { toast("Title is required", "error"); return; }
    const projectId = ($("#nt-project-hidden", modal) ? $("#nt-project-hidden", modal).value : $("#nt-project", modal).value) || "";
    const max = db.tickets.reduce((m, t) => Math.max(m, parseInt(t.id.split("-")[1], 10) || 0), 0);
    const id = "TKT-" + String(max + 1).padStart(4, "0");
    const ticket = {
      id, title, project: projectId, type: $("#nt-type", modal).value,
      status: "Open", opened: new Date().toISOString().slice(0, 10), reporter: db.user.name,
      reporterEmail: db.user.email || "",
      detail: $("#nt-detail", modal).value, workaround: "", esdp: "", affected: [projectId].filter(Boolean),
      updates: []
    };
    db.tickets.unshift(ticket);
    if (typeof Repo !== "undefined") Repo.save("ticket", ticket); else aewttrSaveStore();
    notifyAdminsOfNewTicket(ticket);
    closeModal();
    toast(`${id} created`, "success");
    if (onDone) onDone();
  });
}

function ticketDeepLinkActionUrl(ticketId) {
  return typeof pulseAppRouteUrl === "function"
    ? pulseAppRouteUrl("tickets", ticketId ? { ticket: ticketId } : undefined)
    : pulseAppUrl();
}

function ticketNotificationFacts(t) {
  return [
    { title: "Ticket", value: `${t.id} — ${t.title}` },
    { title: "Type", value: t.type || "—" },
    { title: "Project", value: t.project || "—" },
    { title: "Status", value: t.status || "—" },
    { title: "Reporter", value: t.reporter || "—" }
  ];
}

async function notifyAdminsOfNewTicket(t) {
  if (typeof isSharePointMode !== "function" || !isSharePointMode()) return;
  const db = window.AEWTTR.db;
  const adminEmails = (db.members || []).filter((m) => m.isAdmin && m.email).map((m) => m.email);
  if (!adminEmails.length) return;
  try {
    await notifyUsers({
      to: adminEmails,
      subject: `PULSE Ticket: ${t.id} opened — ${t.title}`,
      area: "Tickets",
      kind: "action",
      preview: `${t.reporter} opened a new ${t.type} ticket${t.project ? ` on ${t.project}` : ""}.`,
      facts: ticketNotificationFacts(t),
      actionUrl: ticketDeepLinkActionUrl(t.id),
      actionTitle: "Open PULSE"
    });
  } catch (e) {
    console.warn("PULSE: new ticket admin notification failed.", e);
  }
}

async function notifyReporterOfTicketUpdate(t, changeSummary) {
  if (typeof isSharePointMode !== "function" || !isSharePointMode() || !t.reporterEmail) return;
  const db = window.AEWTTR.db;
  if (t.reporterEmail.toLowerCase() === String(db.user.email || "").toLowerCase()) return;
  try {
    await notifyUsers({
      to: [t.reporterEmail],
      subject: `PULSE Ticket: ${t.id} — ${changeSummary}`,
      area: "Tickets",
      kind: "comment",
      preview: `${db.user.name} ${changeSummary.toLowerCase()} on your ticket.`,
      facts: ticketNotificationFacts(t),
      actionUrl: ticketDeepLinkActionUrl(t.id),
      actionTitle: "Open PULSE"
    });
  } catch (e) {
    console.warn("PULSE: ticket reporter notification failed.", e);
  }
}

function openTicketModal(id, onDone) {
  // Legacy entry point used from project pages etc. — navigate to full-page detail.
  navigate("tickets/" + id);
  if (onDone) onDone();
}
