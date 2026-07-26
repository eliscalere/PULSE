PAGE_RENDERERS.people = function (parts) {
  const personId = parts && parts[0];
  if (personId) return renderPersonDetail(personId);
  // Redirect back to overview if they just hit #people
  navigate("overview");
};

function buildPersonTaskStats(member) {
  const db = window.AEWTTR.db;
  let mainTasks = 0;
  let subTasks = 0;
  let onTrack = 0;
  let atRisk = 0;
  let offTrack = 0;
  let done = 0;
  const openTasks = [];

  Object.keys(db.ganttTasks || {}).forEach((pid) => {
    const proj = (db.projects || []).find((p) => p.id === pid);
    (db.ganttTasks[pid] || []).forEach((task) => {
      const ownsTask = memberMatchesAssignee(member, task.assignee);
      if (ownsTask) {
        mainTasks += 1;
        if (task.status === "Done" || task.health === "Done") done += 1;
        else {
          openTasks.push({ ...task, projectId: pid, projectName: (proj && proj.name) || pid });
          if (task.health === "Off Track") offTrack += 1;
          else if (task.health === "At Risk") atRisk += 1;
          else onTrack += 1;
        }
      }
      (task.subtasks || []).forEach((sub) => {
        if (memberMatchesAssignee(member, sub.assignee || "")) {
          subTasks += 1;
          if (!ownsTask && !sub.done) {
            openTasks.push({
              id: `${task.id}-sub`,
              title: sub.text || "Subtask",
              health: sub.health || task.health || "On Track",
              status: sub.done ? "Done" : "Open",
              start: sub.start || task.start,
              end: sub.end || task.end,
              projectId: pid,
              projectName: (proj && proj.name) || pid,
              isSubtask: true,
              parentTitle: task.title
            });
          }
        }
      });
    });
  });

  return { mainTasks, subTasks, onTrack, atRisk, offTrack, done, openTasks };
}

function buildPersonTravel(member) {
  const email = normalizeTravelerKey(member.email || "");
  const name = normalizeTravelerKey(member.name || "");
  return sortTravelRequests(window.AEWTTR.db.travelRequests || []).filter((r) => {
    if (email && normalizeTravelerKey(r.requesterEmail || "") === email) return true;
    if (name && normalizeTravelerKey(r.requester || "") === name) return true;
    return (r.travelers || []).some((t) => {
      if (email && normalizeTravelerKey(t.email || "") === email) return true;
      return name && normalizeTravelerKey(t.name || "") === name;
    });
  });
}

function renderPeopleListHtml() {
  const db = window.AEWTTR.db;
  if (!window.AEWTTR.state.peopleView) window.AEWTTR.state.peopleView = "tiles";
  if (!window.AEWTTR.state.peopleSearch) window.AEWTTR.state.peopleSearch = "";

  const view = window.AEWTTR.state.peopleView;
  const q = String(window.AEWTTR.state.peopleSearch || "").trim().toLowerCase();
  // db.members is already the AEWTTR-associated roster (active App Roles).
  const members = (db.members || [])
    .filter((m) => !m.hiddenFromMeetings)
    .filter((m) => {
      if (!q) return true;
      return `${m.name} ${m.email || ""} ${m.role || ""}`.toLowerCase().includes(q);
    })
    .slice()
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")));

  const cards = members.map((member) => {
    const stats = buildPersonTaskStats(member);
    return { member, stats };
  });

  return `
    <div class="toolbar-row people-toolbar" style="margin-bottom:14px;flex-wrap:wrap;gap:10px;">
      <div class="filter-pills" id="people-view-toggle">
        <button type="button" class="filter-pill ${view === "tiles" ? "active" : ""}" data-people-view="tiles"${tip("Show people as tiles")}>Tiles</button>
        <button type="button" class="filter-pill ${view === "rows" ? "active" : ""}" data-people-view="rows"${tip("Show people as a list")}>Rows</button>
      </div>
      <div class="search-box" style="flex:1;max-width:320px;">
        <i class="bx bx-search"></i>
        <input type="search" id="people-search" placeholder="Search people…" value="${escapeHtml(window.AEWTTR.state.peopleSearch)}">
      </div>
    </div>
    ${cards.length ? (view === "tiles" ? `
      <div class="people-tile-grid">
        ${cards.map(({ member, stats }) => `
          <button type="button" class="people-tile" data-person-id="${escapeHtml(member.id)}">
            <div class="people-tile-avatar">${userAvatarHtml({ name: member.name, email: member.email, size: 48 })}</div>
            <div class="people-tile-body">
              <div class="people-tile-name">${escapeHtml(member.name)}</div>
              <div class="people-tile-role">${escapeHtml(member.role || "Member")}</div>
              <div class="people-tile-stats">
                <span><strong>${stats.mainTasks}</strong> tasks</span>
                <span><strong>${stats.subTasks}</strong> subtasks</span>
              </div>
            </div>
          </button>
        `).join("")}
      </div>
    ` : `
      <div class="aewttr-card">
        <table class="aewttr-table">
          <thead><tr><th>Person</th><th>Role</th><th>Main tasks</th><th>Subtasks</th><th>Health</th><th></th></tr></thead>
          <tbody>
            ${cards.map(({ member, stats }) => `
              <tr class="people-row" data-person-id="${escapeHtml(member.id)}" style="cursor:pointer;">
                <td>
                  <div class="people-row-person">
                    ${userAvatarHtml({ name: member.name, email: member.email, size: 32 })}
                    <div>
                      <strong>${escapeHtml(member.name)}</strong>
                      <div style="font-size:11.5px;color:var(--aewttr-muted);">${escapeHtml(member.email || "")}</div>
                    </div>
                  </div>
                </td>
                <td>${escapeHtml(member.role || "—")}</td>
                <td>${stats.mainTasks}</td>
                <td>${stats.subTasks}</td>
                <td>
                  <span class="people-health-chip on">${stats.onTrack} on</span>
                  <span class="people-health-chip risk">${stats.atRisk} risk</span>
                  <span class="people-health-chip off">${stats.offTrack} off</span>
                </td>
                <td><button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-person-id="${escapeHtml(member.id)}">Open</button></td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `) : `<div class="empty-state" style="padding:40px;">No people match your search.</div>`}
  `;
}

function bindPeopleListEvents(root, rerenderCallback) {
  $all("[data-people-view]", root).forEach((btn) => btn.addEventListener("click", () => {
    window.AEWTTR.state.peopleView = btn.dataset.peopleView;
    if (rerenderCallback) rerenderCallback();
  }));
  const search = $("#people-search", root);
  if (search) search.addEventListener("input", (e) => {
    window.AEWTTR.state.peopleSearch = e.target.value;
    if (rerenderCallback) rerenderCallback();
  });
  $all("[data-person-id]", root).forEach((node) => node.addEventListener("click", (e) => {
    e.preventDefault();
    navigate(`people/${node.dataset.personId}`);
  }));
}

function renderPersonDetail(personId) {
  const db = window.AEWTTR.db;
  const member = (db.members || []).find((m) => m.id === personId);
  if (!member) {
    setTopbar("People", "Person not found.");
    $("#page-content").innerHTML = `<div class="empty-state" style="padding:40px;">That person could not be found. <button class="btn-aewttr-outline btn-aewttr-sm" id="people-back">Back to People</button></div>`;
    $("#people-back").addEventListener("click", () => navigate("people"));
    return;
  }

  const stats = buildPersonTaskStats(member);
  const travel = buildPersonTravel(member);
  const upcomingTravel = travel.filter((r) => typeof isUpcomingOrCurrentTravel === "function" ? isUpcomingOrCurrentTravel(r) : r.status === "Approved" || r.status === "Pending");

  setTopbar(member.name, member.role || "Team member", `
    <button class="btn-aewttr-outline" id="people-back"${tip("Back to People")}><i class="bx bx-arrow-back"></i> People</button>
  `);
  $("#people-back").addEventListener("click", () => navigate("people"));

  $("#page-content").innerHTML = `
    <div class="person-detail-hero">
      ${userAvatarHtml({ name: member.name, email: member.email, size: 64 })}
      <div class="person-detail-hero-copy">
        <h2>${escapeHtml(member.name)}</h2>
        <div class="person-detail-meta">${escapeHtml(member.role || "Member")}${member.email ? ` · ${escapeHtml(member.email)}` : ""}</div>
      </div>
    </div>

    <div class="person-stat-grid">
      <div class="person-stat-card"><div class="k">Main tasks</div><div class="v">${stats.mainTasks}</div></div>
      <div class="person-stat-card"><div class="k">Subtasks</div><div class="v">${stats.subTasks}</div></div>
      <div class="person-stat-card person-stat-card--on"><div class="k">On track</div><div class="v">${stats.onTrack}</div></div>
      <div class="person-stat-card person-stat-card--risk"><div class="k">At risk</div><div class="v">${stats.atRisk}</div></div>
      <div class="person-stat-card person-stat-card--off"><div class="k">Off track</div><div class="v">${stats.offTrack}</div></div>
      <div class="person-stat-card"><div class="k">Travel</div><div class="v">${upcomingTravel.length}<span class="person-stat-sub"> upcoming</span></div></div>
    </div>

    <div class="person-detail-grid">
      <div class="aewttr-card aewttr-card-pad">
        <div class="side-panel-title">Open tasks</div>
        ${stats.openTasks.length ? `
          <table class="aewttr-table">
            <thead><tr><th>Task</th><th>Project</th><th>Health</th><th>Dates</th></tr></thead>
            <tbody>
              ${stats.openTasks.slice(0, 40).map((t) => `
                <tr data-open-project="${escapeHtml(t.projectId)}" style="cursor:pointer;">
                  <td>
                    <strong>${escapeHtml(t.title)}</strong>
                    ${t.isSubtask ? `<div style="font-size:11px;color:var(--aewttr-muted);">Subtask of ${escapeHtml(t.parentTitle || "")}</div>` : ""}
                  </td>
                  <td>${escapeHtml(t.projectId)}</td>
                  <td><span class="status-pill status-${String(t.health || "On Track").replace(/\s+/g, "-")}">${escapeHtml(t.health || "On Track")}</span></td>
                  <td>${fmtDate(t.start)} – ${fmtDate(t.end)}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        ` : `<div class="empty-state">No open tasks assigned.</div>`}
      </div>

      <div class="aewttr-card aewttr-card-pad">
        <div class="side-panel-title">Travel history</div>
        ${travel.length ? `
          <table class="aewttr-table">
            <thead><tr><th>TR</th><th>Destination</th><th>Dates</th><th>Status</th></tr></thead>
            <tbody>
              ${travel.slice(0, 30).map((r) => `
                <tr>
                  <td>${escapeHtml(r.id)}</td>
                  <td>${escapeHtml(r.destination || "—")}</td>
                  <td>${formatTravelDateRange(r)}</td>
                  <td>${statusPill(r.status)}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        ` : `<div class="empty-state">No travel on record.</div>`}
      </div>
    </div>
  `;

  $all("[data-open-project]", $("#page-content")).forEach((row) => row.addEventListener("click", () => {
    navigate(`projects/${row.dataset.openProject}/workspace`);
  }));
}
