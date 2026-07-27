function ganttParseDate(d) { return new Date(d + "T00:00:00"); }
function ganttAddDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function ganttIsoDate(d) { return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }
function ganttDaysBetween(a, b) { return Math.round((ganttParseDate(b) - ganttParseDate(a)) / 86400000); }

async function persistProjectCoverImage(proj, coverPreview, coverFile) {
  if (!coverFile && !coverPreview) return "";
  if (coverFile && window.AEWTTR.mode === "sharepoint" && window.AEWTTR.siteUrl && typeof sharePointAdapter.uploadProjectCoverImage === "function") {
    const uploaded = await sharePointAdapter.uploadProjectCoverImage(window.AEWTTR.siteUrl, proj.id, coverFile);
    return uploaded.fileUrl || "";
  }
  if (coverPreview && String(coverPreview).startsWith("http")) return coverPreview;
  return coverPreview || proj.coverImage || "";
}

PAGE_RENDERERS.projects = function (parts) {
  if (!parts || !parts.length) return renderProjectGallery();
  if (parts[0] === "~portfolios") return renderPortfoliosView(parts[1] ? decodeURIComponent(parts[1]) : null);
  if (parts[0] === "~eic") return renderEicView(parts[1] ? decodeURIComponent(parts[1]) : null);
  if (parts[0] === "~program") return renderProgramView(parts[1] ? decodeURIComponent(parts[1]) : null);
  const id = parts[0];
  const tab = parts[1] || "workspace";
  renderProjectDetail(id, tab, null);
};

if (!window.AEWTTR.state.projects) window.AEWTTR.state.projects = { search: "", priority: "All", scope: "mine" };

function isCurrentUserOnProject(proj) {
  const db = window.AEWTTR.db;
  const me = currentUserIdentity();
  if (!me.name && !me.email) return false;
  const people = (db.projectPeople && db.projectPeople[proj.id]) || [];
  return people.some((person) => {
    // projectPeople entries store {label, memberId}, not {name, email} —
    // resolve through db.members when there's a memberId, otherwise treat
    // the freeform label as the name.
    const member = person.memberId ? (db.members || []).find((m) => m.id === person.memberId) : null;
    const candidate = member ? { name: member.name, email: member.email } : { name: person.label, email: "" };
    return samePersonByNameOrEmail(candidate, me);
  });
}

function currentUserAsProjectPersonEntry() {
  const db = window.AEWTTR.db;
  const me = currentUserIdentity();
  if (!me.name && !me.email) return null;
  const known = (db.members || []).find((m) => samePersonByNameOrEmail({ name: m.name, email: m.email }, me));
  return known
    ? { id: uid("ppj"), type: "member", memberId: known.id, label: known.name, role: known.role || "", company: "", email: known.email || "" }
    : { id: uid("ppj"), type: "person", memberId: "", label: me.name || me.email, role: "", company: "", email: me.email || "" };
}

function renderProjectGallery() {
  const st = window.AEWTTR.state.projects;
  setTopbar("Projects", "Manage active project workspaces across teams.", `
    <button class="btn-aewttr" id="btn-new-project"${tip("Create a new project in the portfolio")}><i class="bx bx-plus"></i> New Project</button>
  `);

  function filtered() {
    // Read live db on every filter — background refresh replaces the store
    // object, so a closed-over `db` from mount time would stay stale forever.
    return (window.AEWTTR.db.projects || []).filter(p => {
      if (st.scope === "mine" && !isCurrentUserOnProject(p)) return false;
      if (st.priority !== "All" && p.priority !== st.priority) return false;
      if (st.search && !(p.name.toLowerCase().includes(st.search.toLowerCase()) || p.id.toLowerCase().includes(st.search.toLowerCase()))) return false;
      return true;
    });
  }

  function draw() {
    const rows = filtered();
    const rollup = pGroupRichStats(rows);
    $("#page-content").innerHTML = `
      ${pGroupTopNavHtml("projects")}
      <section class="projects-catalog">
        <header class="projects-catalog-head">
          <div><h2>${st.scope === "mine" ? "My project workspaces" : "All project workspaces"}</h2><p>Open a project to manage its tracker, workstreams, documents, meetings, and supporting tools.</p></div>
          <div class="projects-catalog-summary"><span><b>${rows.length}</b> projects</span><span><b>${rollup.openTasks}</b> open tasks</span><span><b>${Math.max(0, rollup.totalTasks - rollup.openTasks)}</b> closed tasks</span></div>
        </header>
        <div class="projects-catalog-controls">
          <div class="filter-pills" id="scope-pills">
            <button class="filter-pill ${st.scope === "mine" ? "active" : ""}" data-scope="mine"${tip("Show only projects you're a member of")}>My Projects</button>
            <button class="filter-pill ${st.scope === "all" ? "active" : ""}" data-scope="all"${tip("Show every project in the portfolio")}>All Projects</button>
          </div>
          <div class="projects-catalog-filters"><div class="search-box"><i class="bx bx-search"></i><input id="proj-search" placeholder="Search projects..." value="${escapeHtml(st.search)}"></div>
          <select class="select-aewttr" id="priority-filter" style="max-width:170px;">${["All", "Immediate", "High", "Medium", "Lower", "Stretch", "Exploratory", "Ongoing", "Document"].map(p => `<option value="${p}" ${st.priority === p ? "selected" : ""}>${p === "All" ? "All Priorities" : p}</option>`).join("")}</select></div>
        </div>
        <div class="projects-catalog-grid">
          ${rows.length ? rows.map(p => {
            const stats = pGroupRichStats([p]);
            const healthLabel = stats.worstHealth === "red" ? "Needs attention" : stats.worstHealth === "amber" ? "Watch" : "On track";
            const lifecycle = p.lifecycleStatus || p.status || "Active";
            return `<button type="button" class="projects-catalog-card projects-catalog-card--${stats.worstHealth}" data-id="${escapeHtml(p.id)}">
              <div class="projects-catalog-card-head"><span class="projects-catalog-health">${pgroupHealthDotHtml(stats.worstHealth, healthLabel)}${healthLabel}</span>${priorityTag(p.priority)}</div>
              <div class="projects-catalog-name">${escapeHtml(p.name)}</div><div class="projects-catalog-meta"><span>${escapeHtml(p.id)}</span><span>${escapeHtml(lifecycle)}</span></div>
              <div class="projects-catalog-signals"><span>${stats.openTasks} open task${stats.openTasks === 1 ? "" : "s"}</span><span>${Math.max(0, stats.totalTasks - stats.openTasks)} closed task${Math.max(0, stats.totalTasks - stats.openTasks) === 1 ? "" : "s"}</span><span>${stats.openRisks} open risk${stats.openRisks === 1 ? "" : "s"}</span></div>
              <div class="projects-catalog-foot"><span>Updated ${p.updated ? fmtDate(p.updated) : "—"}</span><i class="bx bx-chevron-right"></i></div>
            </button>`;
          }).join("") : `<div class="empty-state projects-catalog-empty">${st.scope === "mine" ? "No projects you're a member of match your filters." : "No projects match your filters."}</div>`}
        </div>
      </section>
    `;
    wirePGroupTopNav();
    $all(".projects-catalog-card[data-id]", $("#page-content")).forEach(r => r.addEventListener("click", () => navigate("projects/" + r.dataset.id)));
    $("#proj-search").addEventListener("input", (e) => {
      const cursor = e.target.selectionStart;
      st.search = e.target.value;
      draw();
      const el = $("#proj-search");
      if (el) { el.focus(); try { el.setSelectionRange(cursor, cursor); } catch (_) {} }
    });
    $all("[data-scope]", $("#scope-pills")).forEach(b => b.addEventListener("click", () => { st.scope = b.dataset.scope; draw(); }));
    $("#priority-filter").addEventListener("change", (e) => { st.priority = e.target.value; draw(); });
  }
  draw();

  // Live refresh for the gallery table (no prior opt-in handler). Soft
  // renderPage from wireBackgroundRefresh also covers this; keep a local
  // draw so focused search input doesn't force a full page remount.
  if (window.AEWTTR._projectGalleryLiveRefreshHandler) {
    window.removeEventListener("pulse:data-refreshed", window.AEWTTR._projectGalleryLiveRefreshHandler);
  }
  window.AEWTTR._projectGalleryLiveRefreshHandler = () => {
    if (!$("#proj-search") && !$("#scope-pills")) return;
    if (document.querySelector(".aewttr-modal-backdrop")) return;
    const active = document.activeElement;
    if (active && (active.tagName === "INPUT" || active.tagName === "TEXTAREA" || active.tagName === "SELECT" || active.isContentEditable)) return;
    const scroller = document.querySelector(".aewttr-content");
    const scrollTop = scroller ? scroller.scrollTop : 0;
    draw();
    if (scroller) scroller.scrollTop = scrollTop;
  };
  window.addEventListener("pulse:data-refreshed", window.AEWTTR._projectGalleryLiveRefreshHandler);

  $("#btn-new-project").addEventListener("click", () => openNewProjectModal());
}

/* =====================================================================
   Portfolio & End-Item-Config aggregation views
   ===================================================================== */

if (!window.AEWTTR.state.pGroupExport) window.AEWTTR.state.pGroupExport = {};

function pGroupTopNavHtml(active) {
  return `
    <div class="pgroup-top-nav">
      <button class="pgroup-top-tab ${active === "projects" ? "active" : ""}" data-pgroup-nav="projects">Projects</button>
      <button class="pgroup-top-tab ${active === "portfolios" ? "active" : ""}" data-pgroup-nav="portfolios">Portfolios</button>
      <button class="pgroup-top-tab ${active === "program" ? "active" : ""}" data-pgroup-nav="program">Programs</button>
      <button class="pgroup-top-tab ${active === "eic" ? "active" : ""}" data-pgroup-nav="eic">End Item Configs</button>
    </div>`;
}

function wirePGroupTopNav() {
  $all("[data-pgroup-nav]", $("#page-content")).forEach(btn => {
    btn.addEventListener("click", () => {
      const v = btn.dataset.pgroupNav;
      if (v === "projects") navigate("projects");
      else if (v === "portfolios") navigate("projects/~portfolios");
      else if (v === "program") navigate("projects/~program");
      else navigate("projects/~eic");
    });
  });
}

function allKnownPortfolios() {
  const names = new Set();
  (window.AEWTTR.db.projects || []).forEach(p => {
    const list = typeof projectPortfolios === "function" ? projectPortfolios(p) : (p.portfolios || []);
    list.forEach(n => { if (n) names.add(n); });
  });
  return Array.from(names).sort((a, b) => a.localeCompare(b));
}

function allKnownEicNames() {
  const names = new Set();
  (window.AEWTTR.db.projects || []).forEach(p => { if (p.configEndItem) names.add(p.configEndItem); });
  return Array.from(names).sort((a, b) => a.localeCompare(b));
}

function projectsInPortfolioGroup(name) {
  return (window.AEWTTR.db.projects || []).filter(p => {
    const list = typeof projectPortfolios === "function" ? projectPortfolios(p) : (p.portfolios || []);
    return list.includes(name);
  });
}

function projectsInEicGroup(name) {
  return (window.AEWTTR.db.projects || []).filter(p => p.configEndItem === name);
}

function allKnownProgramNames() {
  const names = new Set();
  (window.AEWTTR.db.projects || []).forEach(p => { if (p.program) names.add(p.program); });
  return Array.from(names).sort((a, b) => a.localeCompare(b));
}

function projectsInProgramGroup(name) {
  return (window.AEWTTR.db.projects || []).filter(p => p.program === name);
}

function renderProgramView(programName) {
  setTopbar("Programs", "Projects grouped by program — tasks, risks, and exports aggregated.", `<button class="btn-aewttr" id="btn-new-project"><i class="bx bx-plus"></i> New Project</button>`);
  if (programName) {
    renderGroupDetail(programName, "program", projectsInProgramGroup(programName));
    return;
  }
  const all = allKnownProgramNames();
  $("#page-content").innerHTML = `
    ${pGroupTopNavHtml("program")}
    <section class="pgroup-catalog">
      <div class="pgroup-catalog-head">
        <div><span class="pgroup-kicker">Project system</span><h2 class="pgroup-list-title">Program workspaces</h2><p>Open a program to review its projects, workload, people, risks, and reporting tools.</p></div>
        <span class="pgroup-catalog-count">${all.length} program${all.length === 1 ? "" : "s"}</span>
      </div>
      <div class="pgroup-catalog-content">${all.length ? renderGroupCardGrid(all, "program") : '<div class="empty-state" style="padding:40px;text-align:center;color:var(--aewttr-muted);">No programs set yet — assign a Program in Project Settings to group projects here.</div>'}</div>
    </section>`;
  wirePGroupTopNav();
  $all(".pgroup-card", $("#page-content")).forEach(btn => {
    btn.addEventListener("click", () => navigate(`projects/~program/${encodeURIComponent(btn.dataset.groupName)}`));
  });
  const nb = $("#btn-new-project"); if (nb) nb.addEventListener("click", () => openNewProjectModal());
}

function pGroupStats(projects) {
  const db = window.AEWTTR.db;
  let openTasks = 0, totalTasks = 0, highRisks = 0, openRisks = 0;
  projects.forEach(p => {
    function countTasks(list) {
      (list || []).forEach(t => {
        totalTasks++;
        if (t.status !== "Done" && t.status !== "Cancelled") openTasks++;
        if (t.subtasks) countTasks(t.subtasks);
      });
    }
    countTasks((db.ganttTasks && db.ganttTasks[p.id]) || []);
    const risks = typeof projectRisks === "function" ? projectRisks(p.id) : [];
    risks.forEach(r => {
      if (r.status !== "Closed") {
        openRisks++;
        if (r.rating === "Red") highRisks++;
      }
    });
  });
  return { openTasks, totalTasks, highRisks, openRisks };
}

function pGroupAllTasks(projects) {
  const db = window.AEWTTR.db;
  const tasks = [];
  projects.forEach(p => {
    function flatten(list, depth) {
      (list || []).forEach(t => {
        tasks.push(Object.assign({}, t, { _projectId: p.id, _projectName: p.name, _depth: depth || 0 }));
        if (t.subtasks) flatten(t.subtasks, (depth || 0) + 1);
      });
    }
    flatten((db.ganttTasks && db.ganttTasks[p.id]) || [], 0);
  });
  return tasks;
}

function pGroupAllRisks(projects) {
  const risks = [];
  projects.forEach(p => {
    (typeof projectRisks === "function" ? projectRisks(p.id) : []).forEach(r => {
      risks.push(Object.assign({}, r, { _projectId: p.id, _projectName: p.name }));
    });
  });
  return risks;
}

/* ---- Rich rollup stats for PM dashboard visuals ---- */

function pGroupParseDate(v) {
  if (!v) return null;
  const s = String(v);
  const d = new Date(s.length === 10 ? s + "T00:00:00" : s);
  return isNaN(d.getTime()) ? null : d;
}

function pGroupFmtDate(d) {
  if (!d) return "—";
  const iso = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  return fmtDate(iso);
}

function pGroupRichStats(projects) {
  const db = window.AEWTTR.db;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const soon = new Date(today); soon.setDate(soon.getDate() + 7);
  let openTasks = 0, totalTasks = 0, doneTasks = 0;
  let openRisks = 0, redRisks = 0, yellowRisks = 0;
  let lastActivity = null;
  const healthCounts = { green: 0, amber: 0, red: 0 };
  const projectRollups = [];
  projects.forEach(p => {
    let pTotal = 0, pDone = 0, pOpen = 0;
    function countTasks(list) {
      (list || []).forEach(t => {
        pTotal++;
        if (t.status === "Done" || t.status === "Cancelled") {
          if (t.status === "Done") pDone++;
        } else {
          pOpen++;
        }
        if (t.subtasks) countTasks(t.subtasks);
      });
    }
    countTasks((db.ganttTasks && db.ganttTasks[p.id]) || []);
    let pRed = 0, pYellow = 0, pOpenRisks = 0;
    (typeof projectRisks === "function" ? projectRisks(p.id) : []).forEach(r => {
      if (r.status === "Closed") return;
      pOpenRisks++;
      if (r.rating === "Red") pRed++;
      else if (r.rating === "Yellow") pYellow++;
    });
    const dueDate = pGroupParseDate(p.dueDate);
    const slipped = !!(dueDate && dueDate < today && p.lifecycleStatus !== "Completed");
    let health = "green";
    if (pRed > 0 || (slipped && (p.priority === "Immediate" || p.priority === "High"))) health = "red";
    else if (pYellow > 0 || slipped) health = "amber";
    openTasks += pOpen; totalTasks += pTotal; doneTasks += pDone;
    openRisks += pOpenRisks; redRisks += pRed; yellowRisks += pYellow;
    healthCounts[health]++;
    const upd = pGroupParseDate(p.updated);
    if (upd && (!lastActivity || upd > lastActivity)) lastActivity = upd;
    projectRollups.push({
      id: p.id, name: p.name, priority: p.priority, lifecycleStatus: p.lifecycleStatus,
      startDate: p.startDate, dueDate: p.dueDate, slipped,
      totalTasks: pTotal, doneTasks: pDone, openTasks: pOpen,
      openRisks: pOpenRisks, redRisks: pRed, health
    });
  });
  const worstHealth = healthCounts.red ? "red" : (healthCounts.amber ? "amber" : "green");
  return {
    openTasks, totalTasks, doneTasks,
    donePct: totalTasks ? Math.round((doneTasks / totalTasks) * 100) : 0,
    openRisks, redRisks, yellowRisks, lastActivity, healthCounts, worstHealth, projectRollups
  };
}

function pgroupHealthDotHtml(health, label) {
  const text = label || (health === "red" ? "Needs attention" : health === "amber" ? "Watch" : "On track");
  return `<span class="pgroup-health-dot pgroup-health-dot--${health}" title="${escapeHtml(text)}"></span>`;
}

function pgroupProgressBarHtml(done, total, label) {
  if (!total) {
    return `<div class="pgroup-progress-wrap"><div class="pgroup-progress-bar pgroup-progress-bar--empty"></div><span class="pgroup-progress-label">${escapeHtml(label || "No tasks")}</span></div>`;
  }
  const pctDone = (done / total) * 100;
  return `
    <div class="pgroup-progress-wrap">
      <div class="pgroup-progress-bar">
        ${done ? `<span class="pgroup-progress-seg pgroup-progress-seg--done" style="width:${pctDone}%"></span>` : ""}
      </div>
      ${label ? `<span class="pgroup-progress-label">${escapeHtml(label)}</span>` : ""}
    </div>`;
}

function pgroupHealthStripHtml(healthCounts, total) {
  if (!total) return "";
  const seg = (n, cls) => n ? `<span class="pgroup-health-seg pgroup-health-seg--${cls}" style="width:${(n / total) * 100}%"></span>` : "";
  return `
    <div class="pgroup-health-strip">
      ${seg(healthCounts.green, "green")}${seg(healthCounts.amber, "amber")}${seg(healthCounts.red, "red")}
    </div>
    <div class="pgroup-health-legend">
      <span>${pgroupHealthDotHtml("green")} ${healthCounts.green} on track</span>
      <span>${pgroupHealthDotHtml("amber")} ${healthCounts.amber} watch</span>
      <span>${pgroupHealthDotHtml("red")} ${healthCounts.red} at risk</span>
    </div>`;
}

function defaultPGroupExportConfig() {
  return {
    exportMode: "fullDeck",
    includeTasks: false,
    includeRisks: false,
    includePhotos: true,
    taskFilter: "open",
    riskFilter: "open",
    selectedProjects: [],
    selectedTaskIds: []
  };
}

function pGroupExportConfig(type, slug) {
  const key = `${type}:${slug}`;
  if (!window.AEWTTR.state.pGroupExport[key]) window.AEWTTR.state.pGroupExport[key] = defaultPGroupExportConfig();
  const cfg = window.AEWTTR.state.pGroupExport[key];
  /* Migrate the earlier three-option exporter into the two outputs people
     actually need: a single project update, or a complete briefing deck. */
  if (!cfg.exportMode) cfg.exportMode = cfg.reportLayout === "onePage" ? "projectSlide" : "fullDeck";
  if (cfg.includeTasks === undefined) cfg.includeTasks = false;
  if (cfg.includeRisks === undefined) cfg.includeRisks = false;
  if (cfg.includePhotos === undefined) cfg.includePhotos = true;
  if (!Array.isArray(cfg.selectedProjects)) cfg.selectedProjects = [];
  if (!Array.isArray(cfg.selectedTaskIds)) cfg.selectedTaskIds = [];
  return cfg;
}

function renderGroupCardGrid(groupNames, groupType) {
  if (!groupNames.length) {
    const noun = groupType === "portfolio" ? "portfolios" : groupType === "program" ? "programs" : "end item configurations";
    return `<div class="empty-state" style="margin-top:32px;">No ${noun} found. Add them via Project Settings on individual projects.</div>`;
  }
  return `
    <div class="pgroup-card-grid">
      ${groupNames.map(name => {
        const projs = groupType === "portfolio" ? projectsInPortfolioGroup(name) : groupType === "program" ? projectsInProgramGroup(name) : projectsInEicGroup(name);
        const rs = pGroupRichStats(projs);
        const healthLabel = rs.worstHealth === "red" ? "Needs attention" : rs.worstHealth === "amber" ? "Watch" : "On track";
        return `
          <button type="button" class="pgroup-card pgroup-card--${rs.worstHealth}" data-group-name="${escapeHtml(name)}" data-group-type="${escapeHtml(groupType)}">
            <div class="pgroup-card-head">
              ${pgroupHealthDotHtml(rs.worstHealth, healthLabel)}
              <span class="pgroup-card-name">${escapeHtml(name)}</span>
              <span class="pgroup-chip">${projs.length} project${projs.length !== 1 ? "s" : ""}</span>
            </div>
            ${rs.redRisks ? `<div class="pgroup-card-alert"><i class="bx bx-error-circle"></i> ${rs.redRisks} high risk${rs.redRisks !== 1 ? "s" : ""} · ${healthLabel}</div>` : ""}
            ${pgroupProgressBarHtml(rs.doneTasks, rs.totalTasks, rs.totalTasks ? `${rs.donePct}% done` : "No tasks")}
            <div class="pgroup-card-chips">
              <span class="pgroup-chip">${rs.openTasks} open task${rs.openTasks !== 1 ? "s" : ""}</span>
              <span class="pgroup-chip">${rs.openRisks} open risk${rs.openRisks !== 1 ? "s" : ""}</span>
            </div>
            ${rs.lastActivity ? `<div class="pgroup-card-foot">Last activity ${pGroupFmtDate(rs.lastActivity)}</div>` : ""}
          </button>`;
      }).join("")}
    </div>`;
}

function renderPortfoliosView(portfolioName) {
  setTopbar("Portfolios", "Projects grouped by portfolio — tasks, risks, and exports aggregated.", `<button class="btn-aewttr" id="btn-new-project"><i class="bx bx-plus"></i> New Project</button>`);
  if (portfolioName) {
    renderGroupDetail(portfolioName, "portfolio", projectsInPortfolioGroup(portfolioName));
    return;
  }
  $("#page-content").innerHTML = `
    ${pGroupTopNavHtml("portfolios")}
    <section class="pgroup-catalog">
      <div class="pgroup-catalog-head">
        <div><span class="pgroup-kicker">Project system</span><h2 class="pgroup-list-title">Portfolio workspaces</h2><p>Open a portfolio to review its delivery health, workload, people, risks, and reporting tools.</p></div>
        <span class="pgroup-catalog-count">${allKnownPortfolios().length} portfolio${allKnownPortfolios().length === 1 ? "" : "s"}</span>
      </div>
      <div class="pgroup-catalog-content">${renderGroupCardGrid(allKnownPortfolios(), "portfolio")}</div>
    </section>`;
  wirePGroupTopNav();
  $all(".pgroup-card", $("#page-content")).forEach(btn => {
    btn.addEventListener("click", () => navigate(`projects/~portfolios/${encodeURIComponent(btn.dataset.groupName)}`));
  });
  const nb = $("#btn-new-project"); if (nb) nb.addEventListener("click", () => openNewProjectModal());
}

function renderEicView(eicName) {
  setTopbar("End Item Configs", "Projects grouped by end item configuration — tasks, risks, and exports aggregated.", `<button class="btn-aewttr" id="btn-new-project"><i class="bx bx-plus"></i> New Project</button>`);
  if (eicName) {
    renderGroupDetail(eicName, "eic", projectsInEicGroup(eicName));
    return;
  }
  $("#page-content").innerHTML = `
    ${pGroupTopNavHtml("eic")}
    <section class="pgroup-catalog">
      <div class="pgroup-catalog-head">
        <div><span class="pgroup-kicker">Project system</span><h2 class="pgroup-list-title">End item configuration workspaces</h2><p>Review the projects, workload, people, risks, and reports tied to each configuration.</p></div>
        <span class="pgroup-catalog-count">${allKnownEicNames().length} configuration${allKnownEicNames().length === 1 ? "" : "s"}</span>
      </div>
      <div class="pgroup-catalog-content">${renderGroupCardGrid(allKnownEicNames(), "eic")}</div>
    </section>`;
  wirePGroupTopNav();
  $all(".pgroup-card", $("#page-content")).forEach(btn => {
    btn.addEventListener("click", () => navigate(`projects/~eic/${encodeURIComponent(btn.dataset.groupName)}`));
  });
  const nb = $("#btn-new-project"); if (nb) nb.addEventListener("click", () => openNewProjectModal());
}

function renderGroupDetail(groupName, groupType, projects) {
  const stKey = `pgroup-tab:${groupType}:${groupName}`;
  if (!window.AEWTTR.state[stKey]) window.AEWTTR.state[stKey] = "overview";
  let activeTab = window.AEWTTR.state[stKey];
  const backRoute = groupType === "portfolio" ? "projects/~portfolios" : groupType === "program" ? "projects/~program" : "projects/~eic";
  const topNavActive = groupType === "portfolio" ? "portfolios" : groupType === "program" ? "program" : "eic";
  const exportCfg = pGroupExportConfig(groupType, encodeURIComponent(groupName));
  let peopleSnapshot = [];

  function filteredTasks() {
    const tasks = pGroupAllTasks(projects);
    if (exportCfg.taskFilter === "open") return tasks.filter(t => t.status !== "Done" && t.status !== "Cancelled");
    if (exportCfg.taskFilter === "high") return tasks.filter(t => (t.priority === "Immediate" || t.priority === "High") && t.status !== "Done" && t.status !== "Cancelled");
    return tasks;
  }

  function filteredRisks() {
    const risks = pGroupAllRisks(projects);
    if (exportCfg.riskFilter === "open") return risks.filter(r => r.status !== "Closed");
    if (exportCfg.riskFilter === "high") return risks.filter(r => r.status !== "Closed" && r.rating === "Red");
    return risks;
  }

  function overviewHtml() {
    const rs = pGroupRichStats(projects);
    const attention = rs.projectRollups.filter(r => r.health === "red");
    return `
      <div class="pgroup-overview">
        ${attention.length ? `
          <div class="pgroup-attention-callout">
            <i class="bx bx-error"></i>
            <div class="pgroup-attention-text">
              <strong>${attention.length} project${attention.length !== 1 ? "s" : ""} need${attention.length === 1 ? "s" : ""} attention:</strong>
              ${attention.map(r => `<button type="button" class="pgroup-attention-link" data-proj-id="${escapeHtml(r.id)}">${escapeHtml(r.name)}</button>`).join('<span class="pgroup-attention-sep">·</span>')}
            </div>
          </div>` : ""}
        <div class="pgroup-stat-row">
          <div class="pgroup-stat"><span class="pgroup-stat-num">${projects.length}</span><span class="pgroup-stat-label">Projects</span></div>
          <div class="pgroup-stat"><span class="pgroup-stat-num">${rs.openTasks}</span><span class="pgroup-stat-label">Open tasks</span></div>
          <div class="pgroup-stat"><span class="pgroup-stat-num">${rs.doneTasks}</span><span class="pgroup-stat-label">Done</span></div>
          <div class="pgroup-stat"><span class="pgroup-stat-num">${rs.openRisks}</span><span class="pgroup-stat-label">Open risks</span></div>
          <div class="pgroup-stat ${rs.redRisks ? "pgroup-stat--red" : ""}"><span class="pgroup-stat-num">${rs.redRisks}</span><span class="pgroup-stat-label">High risks</span></div>
        </div>
        <div class="pgroup-rollup-grid">
          <div class="aewttr-card pgroup-rollup-card">
            <h4 class="pgroup-section-head">Project health</h4>
            ${pgroupHealthStripHtml(rs.healthCounts, projects.length)}
          </div>
          <div class="aewttr-card pgroup-rollup-card">
            <h4 class="pgroup-section-head">Task progress</h4>
            ${pgroupProgressBarHtml(rs.doneTasks, rs.totalTasks, rs.totalTasks ? `${rs.donePct}% complete` : "No tasks")}
            <div class="pgroup-rollup-legend">
              <span><span class="pgroup-legend-swatch pgroup-legend-swatch--done"></span>${rs.doneTasks} done</span>
              <span><span class="pgroup-legend-swatch pgroup-legend-swatch--open"></span>${rs.openTasks} open</span>
            </div>
          </div>
        </div>
        <h4 class="pgroup-section-head">Projects in this group</h4>
        <div class="aewttr-card">
          <table class="aewttr-table">
            <thead><tr><th style="width:24px;"></th><th>Name</th><th>Lifecycle</th><th>Priority</th><th>Progress</th><th>Tasks</th><th>Risks</th><th>Schedule</th></tr></thead>
            <tbody>
              ${projects.length ? rs.projectRollups.map(r => {
                const p = projects.find(pp => pp.id === r.id) || {};
                const healthText = r.health === "red" ? "Needs attention" : r.health === "amber" ? "Watch" : "On track";
                return `<tr data-proj-id="${escapeHtml(r.id)}" class="pgroup-proj-row" style="cursor:pointer;">
                  <td>${pgroupHealthDotHtml(r.health, healthText)}</td>
                  <td><strong>${escapeHtml(r.name)}</strong></td>
                  <td>${r.lifecycleStatus ? lifecyclePill(r.lifecycleStatus) : "—"}</td>
                  <td>${priorityTag(r.priority)}</td>
                  <td style="min-width:140px;">${pgroupProgressBarHtml(r.doneTasks, r.totalTasks, r.totalTasks ? `${Math.round((r.doneTasks / r.totalTasks) * 100)}%` : "—")}</td>
                  <td>${r.openTasks} open</td>
                  <td>${r.openRisks}${r.redRisks ? ` <span class="pgroup-inline-red">${r.redRisks} high</span>` : ""}</td>
                  <td>${r.startDate ? fmtDate(r.startDate) : "—"} → ${r.dueDate ? `<span class="${r.slipped ? "pgroup-inline-red" : ""}">${fmtDate(r.dueDate)}</span>` : "—"}</td>
                </tr>`;
              }).join("") : `<tr><td colspan="8"><div class="empty-state">No projects in this group.</div></td></tr>`}
            </tbody>
          </table>
        </div>
      </div>`;
  }

  function tasksHtml() {
    const tasks = filteredTasks();
    return `
      <div class="pgroup-tab-body">
        <div class="pgroup-filter-row">
          <span class="pgroup-filter-label">Show:</span>
          <select class="select-aewttr pgroup-task-filter-sel" style="max-width:220px;">
            <option value="open" ${exportCfg.taskFilter === "open" ? "selected" : ""}>Open tasks only</option>
            <option value="all" ${exportCfg.taskFilter === "all" ? "selected" : ""}>All tasks</option>
            <option value="high" ${exportCfg.taskFilter === "high" ? "selected" : ""}>High / Immediate priority only</option>
          </select>
        </div>
        <div class="aewttr-card">
          <table class="aewttr-table">
            <thead><tr><th>Task</th><th>Project</th><th>Assignee</th><th>Priority</th><th>Status</th><th>Due</th></tr></thead>
            <tbody>
              ${tasks.length ? tasks.map(t => `
                <tr class="pgroup-task-row" data-proj-id="${escapeHtml(t._projectId)}" style="cursor:pointer;">
                  <td style="padding-left:${Math.min((t._depth || 0) * 14, 42)}px;">${escapeHtml(t.name || t.title || "—")}</td>
                  <td>${escapeHtml(t._projectName || "")}</td>
                  <td>${escapeHtml(t.assignee || t.owner || "—")}</td>
                  <td>${priorityTag(t.priority)}</td>
                  <td>${escapeHtml(t.status || "")}</td>
                  <td>${t.due ? fmtDate(t.due) : "—"}</td>
                </tr>`).join("") : `<tr><td colspan="6"><div class="empty-state">No tasks match the current filter.</div></td></tr>`}
            </tbody>
          </table>
        </div>
      </div>`;
  }

  function risksHtml() {
    const risks = filteredRisks();
    return `
      <div class="pgroup-tab-body">
        <div class="pgroup-filter-row">
          <span class="pgroup-filter-label">Show:</span>
          <select class="select-aewttr pgroup-risk-filter-sel" style="max-width:220px;">
            <option value="open" ${exportCfg.riskFilter === "open" ? "selected" : ""}>Open risks only</option>
            <option value="all" ${exportCfg.riskFilter === "all" ? "selected" : ""}>All risks</option>
            <option value="high" ${exportCfg.riskFilter === "high" ? "selected" : ""}>High (Red) only</option>
          </select>
        </div>
        <div class="aewttr-card">
          <table class="aewttr-table">
            <thead><tr><th>Risk</th><th>Project</th><th>Rating</th><th>L</th><th>I</th><th>Owner</th><th>Status</th></tr></thead>
            <tbody>
              ${risks.length ? risks.map(r => `
                <tr>
                  <td>${escapeHtml(r.name || r.title || "—")}</td>
                  <td>${escapeHtml(r._projectName || "")}</td>
                  <td>${typeof riskRatingPill === "function" ? riskRatingPill(r.rating) : escapeHtml(r.rating || "")}</td>
                  <td>${r.likelihood || 1}</td>
                  <td>${r.impact || 1}</td>
                  <td>${escapeHtml(r.owner || "—")}</td>
                  <td>${escapeHtml(r.status || "Open")}</td>
                </tr>`).join("") : `<tr><td colspan="7"><div class="empty-state">No risks match the current filter.</div></td></tr>`}
            </tbody>
          </table>
        </div>
      </div>`;
  }

  function peopleHtml() {
    const roster = new Map();
    projects.forEach((project) => {
      const add = (name, role) => {
        const label = String(name || "").trim();
        if (!label || label === "Unassigned") return;
        if (!roster.has(label)) roster.set(label, { name: label, roles: new Set(), projects: new Set(), activeTasks: 0, completedTasks: 0, totalTasks: 0, tasks: [] });
        const row = roster.get(label);
        row.projects.add(project.id);
        if (role) row.roles.add(role);
      };
      ((window.AEWTTR.db.projectPeople && window.AEWTTR.db.projectPeople[project.id]) || []).forEach((entry) => add(entry.label, entry.role || entry.type));
      pGroupAllTasks([project]).forEach((task) => {
        add(task.assignee || task.owner, "");
        const row = roster.get(String(task.assignee || task.owner || "").trim());
        if (row) {
          row.totalTasks++;
          row.tasks.push(Object.assign({}, task, { _projectName: project.name || project.id }));
          if (task.status === "Done" || task.status === "Cancelled") row.completedTasks++;
          else {
            row.activeTasks++;
          }
        }
      });
    });
    const people = [...roster.values()].sort((a, b) => b.activeTasks - a.activeTasks || a.name.localeCompare(b.name));
    peopleSnapshot = people;
    return `<div class="pgroup-tab-body"><div class="pgroup-people-list">
      ${people.length ? people.map((person) => `<button type="button" class="pgroup-person-row" data-pgroup-person="${escapeHtml(person.name)}">
        ${userAvatarHtml(person.name, "", 32)}
        <span class="pgroup-person-copy"><strong>${escapeHtml(person.name)}</strong><small>${escapeHtml([...person.roles].join(", ") || "Team member")} · ${person.projects.size} project${person.projects.size === 1 ? "" : "s"}</small></span>
        <span class="pgroup-person-work"><b>${person.activeTasks}</b><small>active tasks</small></span>
        <i class="bx bx-chevron-right pgroup-person-open"></i>
      </button>`).join("") : `<div class="empty-state">No assigned people in this group yet.</div>`}
    </div></div>`;
  }

  function openGroupPersonModal(personName) {
    const person = peopleSnapshot.find((row) => row.name === personName);
    if (!person) return;
    const projectNames = [...person.projects].map((id) => (projects.find((project) => project.id === id) || {}).name || id);
    const openTasks = person.tasks.filter((task) => task.status !== "Done" && task.status !== "Cancelled").sort((a, b) => String(a.end || "9999").localeCompare(String(b.end || "9999")));
    const modal = openModal(`<div class="aewttr-modal-head"><div><span class="pgroup-kicker">Team workload</span><h3>${escapeHtml(person.name)}</h3></div><button class="aewttr-modal-close" type="button">&times;</button></div>
      <div class="pgroup-person-profile">
        <div class="pgroup-person-profile-lead">${userAvatarHtml(person.name, "", 48)}<div><strong>${escapeHtml([...person.roles].join(", ") || "Team member")}</strong><p>${projectNames.length} project${projectNames.length === 1 ? "" : "s"} in this workspace</p></div></div>
        <div class="pgroup-person-stats"><div><b>${person.activeTasks}</b><span>Open work</span></div><div><b>${person.completedTasks}</b><span>Completed</span></div><div><b>${person.totalTasks ? Math.round((person.completedTasks / person.totalTasks) * 100) : 0}%</b><span>Completion</span></div></div>
        <div class="pgroup-person-profile-section"><h4>Projects</h4><div class="pgroup-person-projects">${projectNames.map((name) => `<span>${escapeHtml(name)}</span>`).join("") || "<span>None assigned</span>"}</div></div>
        <div class="pgroup-person-profile-section"><h4>Open work</h4><div class="pgroup-person-task-list">${openTasks.length ? openTasks.map((task) => `<button type="button" data-person-task-project="${escapeHtml(task._projectId)}"><span><strong>${escapeHtml(task.name || task.title || "Untitled task")}</strong><small>${escapeHtml(task._projectName)}</small></span><span class="pgroup-person-task-meta">${escapeHtml(task.priority || "Normal")} · ${task.end ? fmtDate(task.end) : "No due date"}</span></button>`).join("") : "<div class=\"empty-state\">No open assigned work.</div>"}</div></div>
      </div>`, { wide: true });
    $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
    $all("[data-person-task-project]", modal).forEach((button) => button.addEventListener("click", () => { closeModal(); navigate(`projects/${button.dataset.personTaskProject}/tracker`); }));
  }

  function exportPanelHtml() {
    return `<div id="pgroup-reporting-mount"></div>`;
  }

  function drawGroupReportingTab(mount) {
    const groupLabel = groupType === "portfolio" ? "Portfolio" : groupType === "program" ? "Program" : "End item configuration";

    // focusedProjId = which project is shown in the editor (tracked separately from export scope)
    if (!exportCfg.focusedProjId || !projects.some(function(p) { return p.id === exportCfg.focusedProjId; })) {
      exportCfg.focusedProjId = (projects[0] && projects[0].id) || null;
    }
    const focusedProjId = exportCfg.focusedProjId;

    function selectedProjectIds() {
      if (exportCfg.exportMode === "projectSlide") return focusedProjId ? [focusedProjId] : [];
      // Full deck: empty selectedProjects means "all"
      return exportCfg.selectedProjects.length ? exportCfg.selectedProjects : projects.map(function(p) { return p.id; });
    }

    const isProjectSlide = exportCfg.exportMode === "projectSlide";
    const richStats = pGroupRichStats(projects);

    const sidebarItemsHtml = projects.map(function(proj) {
      const extra = (window.AEWTTR.db.projectExtra && window.AEWTTR.db.projectExtra[proj.id]) || {};
      const sc = (extra.reportConfig && extra.reportConfig.slideContent) || {};
      const rag = sc.overallRag || "";
      const ragDotCls = rag === "Green" ? "grp-sb-dot--green" : rag === "Amber" ? "grp-sb-dot--amber" : rag === "Red" ? "grp-sb-dot--red" : "grp-sb-dot--none";
      const isExportSelected = selectedProjectIds().includes(proj.id);
      const isFocused = proj.id === focusedProjId;
      const rollup = richStats.projectRollups.find(function(r) { return r.id === proj.id; }) || {};
      return `<button type="button" class="grp-sb-item${isFocused ? " is-focused" : ""}${isExportSelected ? " is-export-selected" : ""}" data-sb-proj-id="${escapeHtml(proj.id)}" title="${escapeHtml(proj.name || proj.id)}">
        <span class="grp-sb-dot ${ragDotCls}">●</span>
        <span class="grp-sb-name">${escapeHtml(proj.name || proj.id)}</span>
        <span class="grp-sb-meta">${rollup.openTasks || 0} tasks · ${rollup.openRisks || 0} risks</span>
      </button>`;
    }).join("");

    const exportProjectsHtml = projects.map(function(proj) {
      const ids = selectedProjectIds();
      const checked = ids.includes(proj.id);
      const ctrl = isProjectSlide
        ? `<input type="radio" name="grp-exp-proj" class="grp-exp-proj-ctrl" data-proj-id="${escapeHtml(proj.id)}" ${checked ? "checked" : ""}>`
        : `<input type="checkbox" class="grp-exp-proj-ctrl" data-proj-id="${escapeHtml(proj.id)}" ${checked ? "checked" : ""}>`;
      return `<label class="grp-exp-proj-row${checked ? " is-checked" : ""}">${ctrl}<span>${escapeHtml(proj.name || proj.id)}</span></label>`;
    }).join("");

    mount.innerHTML = `
      <div class="grp-rep-root">
        <div class="grp-rep-topbar">
          <div class="grp-rep-topbar-modes">
            <label class="grp-rep-mode-pill${!isProjectSlide ? " is-active" : ""}">
              <input type="radio" name="grp-rep-mode" value="fullDeck" ${!isProjectSlide ? "checked" : ""}>
              <i class="bx bxs-slideshow"></i> Full deck
            </label>
            <label class="grp-rep-mode-pill${isProjectSlide ? " is-active" : ""}">
              <input type="radio" name="grp-rep-mode" value="projectSlide" ${isProjectSlide ? "checked" : ""}>
              <i class="bx bx-file"></i> Single slide
            </label>
          </div>
          <div class="grp-rep-topbar-scope">
            <span class="grp-rep-topbar-scope-label">Export scope:</span>
            <div class="grp-exp-proj-list">
              ${exportProjectsHtml || '<span class="grp-rep-topbar-scope-label">No projects</span>'}
            </div>
            ${!isProjectSlide ? `<div class="grp-exp-scope-actions"><button type="button" class="pgroup-export-link" data-select-projects="all">All</button><button type="button" class="pgroup-export-link" data-select-projects="none">Clear</button></div>` : ""}
          </div>
          <div class="grp-rep-topbar-opts">
            <label><input type="checkbox" id="exp-photos" ${exportCfg.includePhotos ? "checked" : ""}> Photos</label>
            <label class="${isProjectSlide ? "is-disabled" : ""}"><input type="checkbox" id="exp-risks" ${exportCfg.includeRisks ? "checked" : ""} ${isProjectSlide ? "disabled" : ""}> Risks</label>
            <label class="${isProjectSlide ? "is-disabled" : ""}"><input type="checkbox" id="exp-tasks" ${exportCfg.includeTasks ? "checked" : ""} ${isProjectSlide ? "disabled" : ""}> Work</label>
          </div>
          <div class="grp-rep-topbar-action">
            <span class="pgroup-export-save-status" id="pgroup-exp-status" aria-live="polite"></span>
            <button class="btn-aewttr" id="btn-gen-pptx" ${projects.length ? "" : "disabled"}>
              <i class="bx bxs-slideshow"></i> Generate PowerPoint
            </button>
          </div>
        </div>

        <div class="grp-rep-workspace">
          <aside class="grp-rep-proj-sidebar">
            <div class="grp-rep-proj-sidebar-head">
              <span>${escapeHtml(groupLabel)}</span>
              <span class="grp-sb-count">${projects.length}</span>
            </div>
            <div class="grp-rep-proj-sidebar-list">
              ${sidebarItemsHtml || '<div class="empty-state">No projects.</div>'}
            </div>
          </aside>
          <div class="grp-rep-proj-editor" id="grp-rep-proj-editor"></div>
        </div>
      </div>`;

    // Load focused project into right panel
    const editorPanel = mount.querySelector("#grp-rep-proj-editor");
    function loadEditorProject(pid) {
      const proj = projects.find(function(p) { return p.id === pid; });
      if (!proj || !editorPanel) return;
      drawProjectReporting(editorPanel, proj);
      // Scroll editor to top
      editorPanel.scrollTop = 0;
    }
    if (focusedProjId) {
      loadEditorProject(focusedProjId);
    } else if (projects.length && editorPanel) {
      editorPanel.innerHTML = '<div class="grp-rep-editor-empty"><i class="bx bx-arrow-back"></i><p>Select a project from the list to edit its reporting content.</p></div>';
    }

    // Wire sidebar clicks — only updates which project is shown in the editor
    mount.querySelectorAll("[data-sb-proj-id]").forEach(function(btn) {
      btn.addEventListener("click", function() {
        const pid = btn.dataset.sbProjId;
        exportCfg.focusedProjId = pid;
        // Update active states
        mount.querySelectorAll("[data-sb-proj-id]").forEach(function(b) {
          b.classList.toggle("is-focused", b.dataset.sbProjId === pid);
        });
        loadEditorProject(pid);
      });
    });

    // Wire mode toggle
    mount.querySelectorAll("input[name=\"grp-rep-mode\"]").forEach(function(inp) {
      inp.addEventListener("change", function() {
        if (!inp.checked) return;
        exportCfg.exportMode = inp.value;
        drawGroupReportingTab(mount);
      });
    });

    // Wire export scope selectors (only affects which projects are exported, not the editor)
    mount.querySelectorAll(".grp-exp-proj-ctrl").forEach(function(ctrl) {
      ctrl.addEventListener("change", function() {
        if (isProjectSlide) {
          // In single-slide mode: selecting a project also focuses it
          exportCfg.focusedProjId = ctrl.dataset.projId;
          drawGroupReportingTab(mount);
        } else {
          const checked = Array.from(mount.querySelectorAll(".grp-exp-proj-ctrl:checked")).map(function(c) { return c.dataset.projId; });
          exportCfg.selectedProjects = checked.length === projects.length ? [] : checked;
        }
      });
    });
    mount.querySelectorAll("[data-select-projects]").forEach(function(btn) {
      btn.addEventListener("click", function() {
        exportCfg.selectedProjects = btn.dataset.selectProjects === "all" ? [] : projects.map(function(p) { return p.id; });
        drawGroupReportingTab(mount);
      });
    });

    // Wire content option checkboxes
    [["exp-photos","includePhotos"],["exp-risks","includeRisks"],["exp-tasks","includeTasks"]].forEach(function(pair) {
      const el = mount.querySelector("#" + pair[0]);
      if (el) el.addEventListener("change", function() { exportCfg[pair[1]] = el.checked; });
    });

    // Wire Generate PowerPoint
    const genBtn = mount.querySelector("#btn-gen-pptx");
    if (genBtn) {
      genBtn.addEventListener("click", async function() {
        const statusEl = mount.querySelector("#pgroup-exp-status");
        const api = window.AEWTTR.ProjectPptxExport;
        if (!api || typeof api.exportGroupStatusPptx !== "function") { toast("PowerPoint group export is unavailable in this package.", "error"); return; }
        const ids = selectedProjectIds();
        const sel = projects.filter(function(p) { return ids.includes(p.id); });
        if (!sel.length) { toast("Select at least one project.", "warn"); return; }
        genBtn.disabled = true;
        genBtn.innerHTML = `<i class="bx bx-loader-alt bx-spin"></i> Generating…`;
        if (statusEl) { statusEl.textContent = "Building…"; statusEl.dataset.state = "saving"; }
        const openTarget = reserveStatusExportOpen("PULSE-status-briefing.pptx");
        try {
          const reportCfg = Object.assign({}, exportCfg, { groupName, groupType, exportMode: exportCfg.exportMode, popup: openTarget });
          const result = await api.exportGroupStatusPptx(sel, reportCfg);
          if (result.mode === "sharepoint" && result.fileUrl) {
            toastStatusPptxSaved(result.fileUrl, result.fileName);
            if (statusEl) { statusEl.textContent = "Saved to SharePoint"; statusEl.dataset.state = "saved"; }
          } else {
            if (statusEl) { statusEl.textContent = "Downloaded"; statusEl.dataset.state = "saved"; }
            toast("PowerPoint downloaded", "success");
          }
        } catch (e) {
          closeStatusExportOpen(openTarget);
          console.warn("PULSE: group PPTX export failed", e);
          if (statusEl) { statusEl.textContent = "Export failed"; statusEl.dataset.state = "error"; }
          toast(`Export failed: ${(e && e.message) || "unknown error"}`, "error");
        }
        genBtn.disabled = false;
        genBtn.innerHTML = `<i class="bx bxs-slideshow"></i> Generate PowerPoint`;
        setTimeout(function() { if (statusEl && statusEl.isConnected) { statusEl.textContent = ""; delete statusEl.dataset.state; } }, 5000);
      });
    }
  }

  function tabContentHtml() {
    if (activeTab === "tasks") return tasksHtml();
    if (activeTab === "risks") return risksHtml();
    if (activeTab === "people") return peopleHtml();
    if (activeTab === "export") return exportPanelHtml();
    return overviewHtml();
  }

  function draw() {
    const stats = pGroupStats(projects);
    const richStats = pGroupRichStats(projects);
    const openTaskCount = pGroupAllTasks(projects).filter(t => t.status !== "Done" && t.status !== "Cancelled").length;
    const openRiskCount = pGroupAllRisks(projects).filter(r => r.status !== "Closed").length;
    const groupLabel = groupType === "portfolio" ? "Portfolio" : groupType === "program" ? "Program" : "End item configuration";
    const healthLabel = richStats.worstHealth === "red" ? "Needs attention" : richStats.worstHealth === "amber" ? "Watch" : "On track";
    $("#page-content").innerHTML = `
      ${pGroupTopNavHtml(topNavActive)}
      <section class="pgroup-workspace">
        <aside class="pgroup-workspace-sidebar">
          <button class="pgroup-back-btn" id="pgroup-back"><i class="bx bx-chevron-left"></i> All ${groupType === "portfolio" ? "portfolios" : groupType === "program" ? "programs" : "end item configs"}</button>
          <div class="pgroup-workspace-identity">
            <span class="pgroup-kicker">${groupLabel}</span>
            <h2 class="pgroup-detail-name">${escapeHtml(groupName)}</h2>
            <span class="pgroup-workspace-health pgroup-workspace-health--${richStats.worstHealth}">${pgroupHealthDotHtml(richStats.worstHealth, healthLabel)}${healthLabel}</span>
          </div>
          <div class="pgroup-workspace-summary">
            <span><b>${projects.length}</b> projects</span><span><b>${stats.openTasks}</b> open tasks</span><span><b>${stats.openRisks}</b> open risks</span>
          </div>
          <nav class="pgroup-workspace-nav" aria-label="${groupLabel} workspace views">
            <button class="pgroup-workspace-nav-btn ${activeTab === "overview" ? "active" : ""}" data-detail-tab="overview"><i class="bx bx-grid-alt"></i> Overview</button>
            <button class="pgroup-workspace-nav-btn ${activeTab === "tasks" ? "active" : ""}" data-detail-tab="tasks"><i class="bx bx-task"></i> Workload <span>${openTaskCount}</span></button>
            <button class="pgroup-workspace-nav-btn ${activeTab === "people" ? "active" : ""}" data-detail-tab="people"><i class="bx bx-group"></i> People</button>
            <button class="pgroup-workspace-nav-btn ${activeTab === "risks" ? "active" : ""}" data-detail-tab="risks"><i class="bx bx-shield-quarter"></i> Risks <span>${openRiskCount}</span></button>
            <button class="pgroup-workspace-nav-btn ${activeTab === "export" ? "active" : ""}" data-detail-tab="export"><i class="bx bxs-slideshow"></i> Reporting</button>
          </nav>
        </aside>
        <main class="pgroup-workspace-main">
          ${activeTab !== "export" ? `<header class="pgroup-workspace-head">
            <div><span class="pgroup-kicker">${activeTab === "tasks" ? "Workload" : activeTab === "people" ? "Team capacity" : activeTab === "risks" ? "Risk register" : "Delivery overview"}</span><h3>${activeTab === "tasks" ? "Work across this " + groupLabel.toLowerCase() : activeTab === "people" ? "People and assigned work" : activeTab === "risks" ? "Risks across projects" : groupLabel + " delivery at a glance"}</h3></div>
            <div class="pgroup-workspace-actions"><button class="btn-aewttr btn-aewttr-outline" data-detail-tab="export"><i class="bx bxs-slideshow"></i> Reporting</button><button class="btn-aewttr" id="pgroup-new-project"><i class="bx bx-plus"></i> New Project</button></div>
          </header>` : ""}
          <div class="pgroup-workspace-content pgroup-detail-body${activeTab === "export" ? " pgroup-reporting-full" : ""}">${tabContentHtml()}</div>
        </main>
      </section>`;

    wirePGroupTopNav();
    $all("[data-detail-tab]", $("#page-content")).forEach(btn => {
      btn.addEventListener("click", () => {
        activeTab = btn.dataset.detailTab;
        window.AEWTTR.state[stKey] = activeTab;
        draw();
      });
    });
    const backBtn = $("#pgroup-back", $("#page-content"));
    if (backBtn) backBtn.addEventListener("click", () => navigate(backRoute));
    const newProjectBtn = $("#pgroup-new-project", $("#page-content"));
    if (newProjectBtn) newProjectBtn.addEventListener("click", () => openNewProjectModal());

    $all(".pgroup-proj-row[data-proj-id]", $("#page-content")).forEach(row => {
      row.addEventListener("click", () => navigate(`projects/${row.dataset.projId}`));
    });
    $all(".pgroup-attention-link[data-proj-id]", $("#page-content")).forEach(el => {
      el.addEventListener("click", ev => { ev.stopPropagation(); navigate(`projects/${el.dataset.projId}`); });
    });
    $all("[data-pgroup-person]", $("#page-content")).forEach((button) => {
      button.addEventListener("click", () => openGroupPersonModal(button.dataset.pgroupPerson));
    });

    const taskSel = $(".pgroup-task-filter-sel", $("#page-content"));
    if (taskSel) taskSel.addEventListener("change", () => { exportCfg.taskFilter = taskSel.value; draw(); });
    const riskSel = $(".pgroup-risk-filter-sel", $("#page-content"));
    if (riskSel) riskSel.addEventListener("change", () => { exportCfg.riskFilter = riskSel.value; draw(); });

    if (activeTab === "export") {
      const mount = $("#pgroup-reporting-mount", $("#page-content"));
      if (mount) {
        drawGroupReportingTab(mount);
        const mainEl = mount.closest(".pgroup-workspace-main");
        if (mainEl) {
          const topOffset = mainEl.getBoundingClientRect().top;
          mainEl.style.minHeight = Math.max(500, window.innerHeight - topOffset - 16) + "px";
          mainEl.style.display = "flex";
          mainEl.style.flexDirection = "column";
        }
      }
    }
  }

  function wireExportPanel() {
    const checkMap = [
      ["exp-summaries", "includeSummaries"],
      ["exp-tasks", "includeTasks"],
      ["exp-risks", "includeRisks"],
      ["exp-photos", "includePhotos"]
    ];
    checkMap.forEach(([id, key]) => {
      const el = $(`#${id}`, $("#page-content"));
      if (el) el.addEventListener("change", () => { exportCfg[key] = el.checked; });
    });
    const tf = $("#exp-task-filter", $("#page-content"));
    if (tf) tf.addEventListener("change", () => { exportCfg.taskFilter = tf.value; });
    const rf = $("#exp-risk-filter", $("#page-content"));
    if (rf) rf.addEventListener("change", () => { exportCfg.riskFilter = rf.value; });
    $all("input[name=\"exp-layout\"]", $("#page-content")).forEach((input) => input.addEventListener("change", () => { if (input.checked) exportCfg.reportLayout = input.value; }));
    $all(".exp-proj-cb", $("#page-content")).forEach(cb => {
      cb.addEventListener("change", () => {
        const checked = $all(".exp-proj-cb:checked", $("#page-content")).map(c => c.dataset.projId);
        exportCfg.selectedProjects = checked.length === projects.length ? [] : checked;
      });
    });
    $all(".exp-task-cb", $("#page-content")).forEach(cb => {
      cb.addEventListener("change", () => {
        const allTaskBoxes = $all(".exp-task-cb", $("#page-content"));
        const checked = $all(".exp-task-cb:checked", $("#page-content")).map(c => c.dataset.taskId);
        exportCfg.selectedTaskIds = checked.length === allTaskBoxes.length ? [] : checked;
      });
    });
    $all("[data-select-projects]", $("#page-content")).forEach((button) => button.addEventListener("click", () => {
      $all(".exp-proj-cb", $("#page-content")).forEach((box) => { box.checked = button.dataset.selectProjects === "all"; });
      exportCfg.selectedProjects = button.dataset.selectProjects === "all" ? [] : projects.map((project) => project.id);
    }));
    $all("[data-select-tasks]", $("#page-content")).forEach((button) => button.addEventListener("click", () => {
      const boxes = $all(".exp-task-cb", $("#page-content"));
      boxes.forEach((box) => { box.checked = button.dataset.selectTasks === "all"; });
      exportCfg.selectedTaskIds = button.dataset.selectTasks === "all" ? [] : boxes.map((box) => box.dataset.taskId);
    }));
    const genBtn = $("#btn-gen-pptx", $("#page-content"));
    if (genBtn) {
      genBtn.addEventListener("click", async () => {
        const statusEl = $("#pgroup-exp-status", $("#page-content"));
        const api = window.AEWTTR.ProjectPptxExport;
        if (!api || typeof api.exportGroupStatusPptx !== "function") {
          toast("PowerPoint group export is unavailable in this package.", "error");
          return;
        }
        const ids = exportCfg.selectedProjects && exportCfg.selectedProjects.length ? exportCfg.selectedProjects : projects.map(p => p.id);
        const sel = projects.filter(p => ids.includes(p.id));
        if (!sel.length) { toast("Select at least one project.", "warn"); return; }
        genBtn.disabled = true;
        genBtn.innerHTML = `<i class="bx bx-loader-alt bx-spin"></i> Generating…`;
        if (statusEl) { statusEl.textContent = "Building…"; statusEl.dataset.state = "saving"; }
        const openTarget = reserveStatusExportOpen("PULSE-status-briefing.pptx");
        try {
          const reportCfg = Object.assign({}, exportCfg, { groupName, groupType, popup: openTarget });
          if (reportCfg.reportLayout === "onePage") { reportCfg.includeSummaries = false; reportCfg.includeTasks = false; reportCfg.includeRisks = false; }
          if (reportCfg.reportLayout === "projectPack") { reportCfg.includeTasks = false; reportCfg.includeRisks = false; }
          const result = await api.exportGroupStatusPptx(sel, reportCfg);
          if (result.mode === "sharepoint" && result.fileUrl) {
            toastStatusPptxSaved(result.fileUrl, result.fileName);
            if (statusEl) { statusEl.textContent = "Saved to SharePoint"; statusEl.dataset.state = "saved"; }
          } else {
            if (statusEl) { statusEl.textContent = "Downloaded"; statusEl.dataset.state = "saved"; }
            toast("PowerPoint downloaded", "success");
          }
        } catch (e) {
          closeStatusExportOpen(openTarget);
          console.warn("PULSE: group PPTX export failed", e);
          if (statusEl) { statusEl.textContent = "Export failed"; statusEl.dataset.state = "error"; }
          toast(`Export failed: ${(e && e.message) || "unknown error"}`, "error");
        }
        genBtn.disabled = false;
        genBtn.innerHTML = `<i class="bx bxs-slideshow"></i> Generate PowerPoint`;
        setTimeout(() => { if (statusEl && statusEl.isConnected) { statusEl.textContent = ""; delete statusEl.dataset.state; } }, 5000);
      });
    }
  }

  draw();
}

/* =====================================================================
   End of Portfolio & EIC aggregation views
   ===================================================================== */

function tagPickerHtml(selected, idPrefix, opts) {
  opts = Object.assign({
    emptyText: "None selected yet.",
    placeholder: "Search or add…",
    hint: ""
  }, opts || {});
  const selectedList = (selected || []).map((n) => String(n || "").trim()).filter(Boolean);
  return `
    <div class="tag-picker" data-tag-picker="${escapeHtml(idPrefix)}">
      <div class="tag-picker-selected" id="${escapeHtml(idPrefix)}-selected">
        ${selectedList.length ? selectedList.map((name) => `
          <span class="tag-chip" data-tag="${escapeHtml(name)}">
            <em>${escapeHtml(name)}</em>
            <button type="button" class="tag-chip-remove" data-remove-tag="${escapeHtml(name)}" aria-label="Remove ${escapeHtml(name)}"><i class="bx bx-x"></i></button>
          </span>`).join("") : `<span class="tag-picker-empty">${escapeHtml(opts.emptyText)}</span>`}
      </div>
      <div class="tag-picker-add-row">
        <input class="input-aewttr" id="${escapeHtml(idPrefix)}-input" placeholder="${escapeHtml(opts.placeholder)}" autocomplete="off">
        <button type="button" class="tag-picker-add-btn" id="${escapeHtml(idPrefix)}-add" aria-label="Add"${tip("Add")}><i class="bx bx-plus"></i></button>
      </div>
      <div class="tag-picker-suggestions" id="${escapeHtml(idPrefix)}-suggestions" hidden></div>
      ${opts.hint ? `<p class="tag-picker-hint">${opts.hint}</p>` : ""}
    </div>`;
}

function wireTagPicker(root, selectedSet, idPrefix, opts) {
  opts = opts || {};
  const normalize = typeof opts.normalize === "function" ? opts.normalize : ((v) => String(v || "").trim());
  const getKnown = typeof opts.getKnown === "function" ? opts.getKnown : (() => []);
  const remember = typeof opts.remember === "function" ? opts.remember : (() => {});
  const onChange = typeof opts.onChange === "function" ? opts.onChange : null;
  const selectedEl = $(`#${idPrefix}-selected`, root);
  const input = $(`#${idPrefix}-input`, root);
  const addBtn = $(`#${idPrefix}-add`, root);
  const suggestions = $(`#${idPrefix}-suggestions`, root);

  function syncSelectedSetCasing(name) {
    if (opts.singleSelect) selectedSet.clear();
    const lower = name.toLowerCase();
    for (const existing of Array.from(selectedSet)) {
      if (String(existing).toLowerCase() === lower) {
        selectedSet.delete(existing);
        break;
      }
    }
    selectedSet.add(name);
  }

  function isSelected(name) {
    const lower = String(name).toLowerCase();
    return Array.from(selectedSet).some((n) => String(n).toLowerCase() === lower);
  }

  function hideSuggestions() {
    suggestions.hidden = true;
    suggestions.innerHTML = "";
  }

  function renderSelected() {
    const list = Array.from(selectedSet);
    selectedEl.innerHTML = list.length
      ? list.map((name) => `
          <span class="tag-chip" data-tag="${escapeHtml(name)}">
            <em>${escapeHtml(name)}</em>
            <button type="button" class="tag-chip-remove" data-remove-tag="${escapeHtml(name)}" aria-label="Remove ${escapeHtml(name)}"><i class="bx bx-x"></i></button>
          </span>`).join("")
      : `<span class="tag-picker-empty">None selected yet.</span>`;
    $all("[data-remove-tag]", selectedEl).forEach((btn) => btn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const name = btn.dataset.removeTag;
      Array.from(selectedSet).forEach((n) => {
        if (String(n).toLowerCase() === String(name).toLowerCase()) selectedSet.delete(n);
      });
      renderSelected();
      hideSuggestions();
      if (onChange) onChange(selectedSet);
    }));
  }

  function renderSuggestions(showAll) {
    const query = normalize(input.value).toLowerCase();
    if (!query && !showAll) {
      hideSuggestions();
      return;
    }
    const known = getKnown().filter((n) => !isSelected(n));
    const matches = query
      ? known.filter((n) => String(n).toLowerCase().includes(query)).slice(0, 8)
      : known.slice(0, 12);
    const exactExists = getKnown().some((n) => String(n).toLowerCase() === query);
    if (!matches.length && !query) {
      hideSuggestions();
      return;
    }
    suggestions.hidden = false;
    suggestions.innerHTML = [
      ...matches.map((name, i) => `
        <button type="button" class="tag-picker-suggestion" data-pick-index="${i}">
          <strong>${escapeHtml(name)}</strong>
        </button>`),
      (!exactExists && query.length >= 1
        ? `<button type="button" class="tag-picker-suggestion tag-picker-suggestion--create" data-create="1">
             <i class="bx bx-plus"></i> Create "${escapeHtml(normalize(input.value))}"
           </button>`
        : "")
    ].filter(Boolean).join("") || `<div class="tag-picker-suggestion-empty">No matches</div>`;

    $all("[data-pick-index]", suggestions).forEach((btn) => btn.addEventListener("click", () => {
      const name = matches[Number(btn.dataset.pickIndex)];
      if (!name) return;
      remember([name]);
      syncSelectedSetCasing(name);
      input.value = "";
      renderSelected();
      hideSuggestions();
      if (onChange) onChange(selectedSet);
      input.focus();
    }));
    const createBtn = $("[data-create]", suggestions);
    if (createBtn) createBtn.addEventListener("click", () => addFromInput());
  }

  function addFromInput() {
    const name = normalize(input.value);
    if (!name) return;
    remember([name]);
    // Prefer canonical casing from catalog when available.
    const known = getKnown().find((n) => String(n).toLowerCase() === name.toLowerCase()) || name;
    syncSelectedSetCasing(known);
    input.value = "";
    renderSelected();
    hideSuggestions();
    if (onChange) onChange(selectedSet);
    input.focus();
  }

  renderSelected();
  addBtn.addEventListener("click", addFromInput);
  input.addEventListener("input", () => renderSuggestions(false));
  input.addEventListener("focus", () => renderSuggestions(true));
  input.addEventListener("blur", () => {
    setTimeout(() => {
      if (!suggestions.contains(document.activeElement)) hideSuggestions();
    }, 150);
  });
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      const first = $("[data-pick-index], [data-create]", suggestions);
      if (first && !suggestions.hidden) first.click();
      else addFromInput();
    } else if (e.key === "Escape") {
      hideSuggestions();
    }
  });
  return selectedSet;
}

function portfolioPickerHtml(selected, idPrefix) {
  return tagPickerHtml(selected, idPrefix, {
    emptyText: "No portfolios selected.",
    placeholder: "Search or add portfolio…",
    hint: "Select one or more. New names are remembered for future projects."
  });
}

function wirePortfolioPicker(root, selectedSet, idPrefix, onChange) {
  return wireTagPicker(root, selectedSet, idPrefix, {
    normalize: normalizePortfolioName,
    getKnown: getKnownPortfolioNames,
    remember: rememberPortfolioNames,
    onChange
  });
}

function locationPickerHtml(selected, idPrefix) {
  return tagPickerHtml(selected, idPrefix, {
    emptyText: "No locations selected.",
    placeholder: "Search or add location…",
    hint: "Select one or more. New names are remembered for future projects."
  });
}

function wireLocationPicker(root, selectedSet, idPrefix, onChange) {
  return wireTagPicker(root, selectedSet, idPrefix, {
    normalize: normalizeLocationName,
    getKnown: getKnownLocationNames,
    remember: rememberLocationNames,
    onChange
  });
}

function contractorPickerHtml(selected, idPrefix) {
  return tagPickerHtml(selected, idPrefix, {
    emptyText: "No contractors selected.",
    placeholder: "Search or add contractor…",
    hint: "Select one or more. New names are remembered like portfolios and locations."
  });
}

function wireContractorPicker(root, selectedSet, idPrefix, onChange) {
  return wireTagPicker(root, selectedSet, idPrefix, {
    normalize: normalizeContractorName,
    getKnown: getKnownContractorNames,
    remember: rememberContractorNames,
    onChange
  });
}

function configEndItemPickerHtml(selected, idPrefix) {
  return tagPickerHtml(selected, idPrefix, {
    emptyText: "No config end item selected.",
    placeholder: "Search or add config end item…",
    hint: "Pick one. New names are remembered for future projects and dividers."
  });
}

function wireConfigEndItemPicker(root, selectedSet, idPrefix, onChange) {
  return wireTagPicker(root, selectedSet, idPrefix, {
    normalize: normalizeConfigEndItemName,
    getKnown: getKnownConfigEndItemNames,
    remember: rememberConfigEndItemNames,
    singleSelect: true,
    onChange
  });
}

function openNewProjectModal() {
  const db = window.AEWTTR.db;
  const pendingMembers = [];
  const pendingPortfolios = new Set();
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>New Project</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body">
      <div class="form-row"><label>Project name</label><input class="input-aewttr" id="np-name"></div>
      <div class="form-row"><label>Portfolios</label>
        ${portfolioPickerHtml([], "np-portfolios")}
      </div>
      <div class="form-row"><label>Cover image <small style="font-weight:400;color:var(--aewttr-muted);">(optional)</small></label>
        <div class="cover-upload-row">
          <div class="cover-upload-preview" id="np-cover-preview"><i class="bx bx-image cover-upload-empty-icon"></i></div>
          <div>
            <input type="file" accept="image/*" id="np-cover-file" style="display:none;">
            <button type="button" class="btn-aewttr-outline btn-aewttr-sm" id="np-cover-pick"${tip("Upload a cover image for this project")}><i class="bx bx-image-add"></i> Upload Image</button>
            <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="np-cover-clear" style="display:none;"${tip("Remove the selected image")}>Remove</button>
            <p style="font-size:11px;color:var(--aewttr-muted);margin:6px 0 0;">JPG or PNG, shown as the project's banner background.</p>
          </div>
        </div>
      </div>
      <div class="form-row"><label>Description</label><textarea class="textarea-aewttr" id="np-desc" placeholder="One or two sentences describing this project."></textarea></div>
      <div class="form-row"><label>Members</label>
        <div class="traveler-picker">
          <div id="np-members-selected" class="traveler-chip-list"></div>
          <input class="input-aewttr" id="np-members-input" placeholder="Search people or groups…">
          <div id="np-members-suggestions" class="traveler-suggestions"></div>
        </div>
        <p style="font-size:11.5px;color:var(--aewttr-muted);margin:6px 0 0;">Added to the project roster — automatically available in the Meeting tab.</p>
      </div>
    </div>
    <div class="aewttr-modal-foot">
      <button class="btn-aewttr-ghost" id="np-cancel">Cancel</button>
      <button class="btn-aewttr" id="np-save">Create Project</button>
    </div>
  `, { wide: true });
  wirePeoplePicker(modal, pendingMembers, { mount: "np-members-selected", input: "np-members-input", suggestions: "np-members-suggestions" }, { allowManualEmail: false, includeGroups: true, expandGroups: true });
  wirePortfolioPicker(modal, pendingPortfolios, "np-portfolios");
  let newCoverImage = "";
  let newCoverFile = null;
  function setCoverPreview(url) {
    const preview = $("#np-cover-preview", modal);
    const clearBtn = $("#np-cover-clear", modal);
    if (url) {
      preview.style.background = `url('${url}') center/cover`;
      preview.classList.add("has-image");
      clearBtn.style.display = "";
    } else {
      preview.style.background = "";
      preview.classList.remove("has-image");
      clearBtn.style.display = "none";
    }
  }
  $("#np-cover-pick", modal).addEventListener("click", () => $("#np-cover-file", modal).click());
  $("#np-cover-clear", modal).addEventListener("click", () => {
    newCoverFile = null;
    newCoverImage = "";
    $("#np-cover-file", modal).value = "";
    setCoverPreview("");
  });
  $("#np-cover-file", modal).addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;
    newCoverFile = file;
    const reader = new FileReader();
    reader.onload = () => {
      newCoverImage = reader.result;
      setCoverPreview(newCoverImage);
    };
    reader.onerror = () => toast("Couldn't read that image file.", "error");
    reader.readAsDataURL(file);
  });
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#np-cancel", modal).addEventListener("click", closeModal);
  $("#np-save", modal).addEventListener("click", async () => {
    const name = $("#np-name", modal).value.trim();
    if (!name) { toast("Project name is required", "error"); return; }
    // Prefer the live db — a background refresh may have replaced
    // window.AEWTTR.db while this create modal was open.
    const liveDb = window.AEWTTR.db;
    const nextNum = (liveDb.projects || []).length + 1;
    const portfolios = Array.from(pendingPortfolios);
    rememberPortfolioNames(portfolios);
    const newProj = {
      id: "P" + String(nextNum).padStart(2, "0") + "x",
      name, team: "", priority: "", effort: "", rag: "Green", lifecycleStatus: "Active",
      updated: new Date().toISOString().slice(0, 10), tools: "TBD",
      stakeholders: "",
      sharepointFolderUrl: "",
      coverImage: "", description: $("#np-desc", modal).value.trim(),
      pm: "", engineer: "", isso: "", rangePoc: "", contractor: "", taskOrder: "",
      fundingType: "", fiscalYear: "", fundingStatus: "", changeRequestRequired: false,
      configEndItem: "", locations: [], portfolios, startDate: "", dueDate: "", completionDate: ""
    };
    if (newCoverFile || newCoverImage) {
      try {
        newProj.coverImage = await persistProjectCoverImage(newProj, newCoverImage, newCoverFile);
      } catch (e) {
        toast((e && e.friendly) || e.message || "Cover image upload failed — project will save without it.", "error");
      }
    }
    if (!liveDb.projects) liveDb.projects = [];
    liveDb.projects.unshift(newProj);
    await Repo.save("project", newProj);

    if (!liveDb.projectPeople) liveDb.projectPeople = {};
    const peopleEntries = pendingMembers.map((person) => {
      const known = (liveDb.members || []).find((m) => m.id === person.id || (m.email && person.email && m.email.toLowerCase() === person.email.toLowerCase()));
      return known
        ? { id: uid("ppj"), type: "member", memberId: known.id, label: known.name, role: known.role || "", company: "", email: known.email || "" }
        : { id: uid("ppj"), type: "person", memberId: "", label: person.name, role: "", company: "", email: person.email || "" };
    });
    // Always include the creator — otherwise their own new project doesn't
    // show up under "My Projects" until they separately add themselves.
    const me = currentUserAsProjectPersonEntry();
    if (me && !peopleEntries.some((p) => (me.memberId && p.memberId === me.memberId) || (!me.memberId && p.label === me.label))) {
      peopleEntries.push(me);
    }
    liveDb.projectPeople[newProj.id] = peopleEntries;
    if (liveDb.projectPeople[newProj.id].length) await Repo.save("project", newProj);
    if (typeof syncProjectPulseGroup === "function") {
      try { await syncProjectPulseGroup(newProj); } catch (e) { console.warn("project group sync", e); }
    }

    closeModal();
    toast("Project created", "success");
    if (typeof notifyLocalDataChanged === "function") notifyLocalDataChanged("project-create");
    navigate("projects/" + newProj.id);
  });
}

const PROJECT_BANNER_COLORS = ["#2F5FE0", "#546B2F", "#7A5C10", "#A4262C", "#31405B", "#8764B8"];
function projectBannerColor(proj) {
  let hash = 0;
  for (let i = 0; i < proj.id.length; i++) hash = (hash * 31 + proj.id.charCodeAt(i)) >>> 0;
  return PROJECT_BANNER_COLORS[hash % PROJECT_BANNER_COLORS.length];
}
function projectBannerHtml(proj, options) {
  options = options || {};
  const bg = proj.coverImage
    ? `background-image:linear-gradient(180deg, rgba(10,14,30,.15), rgba(10,14,30,.75)), url('${proj.coverImage}');`
    : `background-image:linear-gradient(135deg, ${projectBannerColor(proj)}, rgba(10,14,30,.85));`;
  return `
    <div class="project-banner" style="${bg}">
      <div class="project-banner-inner">
        <div class="project-banner-copy">
          <div class="project-banner-meta">
            ${priorityTag(proj.priority)}
          </div>
          <h1>${escapeHtml(proj.name)}</h1>
          ${proj.description ? `<p>${escapeHtml(proj.description)}</p>` : ""}
        </div>
        ${options.actionsHtml ? `<div class="project-banner-actions">${options.actionsHtml}</div>` : ""}
      </div>
    </div>`;
}

function renderProjectDetail(id, tab, boardId) {
  const db = window.AEWTTR.db;
  const proj = (db.projects || []).find(p => p.id === id);
  if (!proj) { navigate("projects"); return; }

  setTopbar("", "");

  // The project banner (cover image, name, description, Export/All Projects)
  // now lives inside the Home tab's own content (see drawWorkspace) instead
  // of sitting above every tab — it was eating vertical space on tabs like
  // Tracker/Boards where you want to see content, not a re-statement of
  // which project you're already in (the left nav already shows the ID).
  const navItems = [
    { key: "workspace", label: "Home", icon: "bx-home-alt" },
    { key: "notes", label: "Notes", icon: "bx-notepad" },
    { key: "people", label: "People", icon: "bx-group" },
    { key: "documents", label: "Documents", icon: "bx-folder-open" },
    { key: "photos", label: "Photos", icon: "bx-image" },
    { key: "risks", label: "Risks", icon: "bx-shield-quarter" },
    { key: "meeting", label: "Meeting", icon: "bx-conversation" },
    { key: "tracker", label: "Tracker", icon: "bx-table" },
    { key: "reporting", label: "Reporting", icon: "bxs-slideshow" },
    { key: "settings", label: "Settings", icon: "bx-cog" }
  ];
  const activeTab = tab;
  const trackerCount = ((db.ganttTasks && db.ganttTasks[proj.id]) || []).filter(t => t && !isTrackerDivider(t)).length;
  const openTrackerCount = ((db.ganttTasks && db.ganttTasks[proj.id]) || []).filter((task) => task && !isTrackerDivider(task) && task.status !== "Done" && task.status !== "Cancelled").length;
  const riskCount = projectRisks(proj.id).filter((risk) => risk.status !== "Closed").length;
  const activeNavItem = navItems.find((item) => item.key === activeTab) || navItems[0];
  const projectHealth = proj.technicalStatus || computeProjectTechStatus(proj) || "On Track";
  const projectHealthTone = /off track|blocked|red/i.test(projectHealth) ? "red" : /at risk|watch|amber/i.test(projectHealth) ? "amber" : "green";
  const projectHeadings = {
    workspace: { kicker: "Delivery overview", title: "Project delivery at a glance" },
    notes: { kicker: "Project record", title: "Notes, decisions, and risk context" },
    people: { kicker: "Team capacity", title: "People and project roles" },
    documents: { kicker: "Working files", title: "Project documents" },
    photos: { kicker: "Visual record", title: "Project photos" },
    risks: { kicker: "Risk register", title: "Risks and mitigation" },
    meeting: { kicker: "Team coordination", title: "Project meeting" },
    tracker: { kicker: "Workload", title: "Tasks and milestones" },
    reporting: { kicker: "Status reporting", title: "Project briefing and exports" },
    settings: { kicker: "Project administration", title: "Project settings" }
  };
  const activeHeading = projectHeadings[activeTab] || { kicker: "Project workspace", title: activeNavItem.label };

  $("#page-content").innerHTML = `
    <div class="project-spo-layout pgroup-workspace project-spo-layout--scroll">
      <aside class="project-spo-nav pgroup-workspace-sidebar">
        <button type="button" class="project-spo-back pgroup-back-btn" id="project-sidebar-back"><i class="bx bx-chevron-left"></i> All projects</button>
        <div class="pgroup-workspace-identity project-spo-identity">
          <span class="pgroup-kicker">Project</span>
          <h2 class="pgroup-detail-name">${escapeHtml(proj.name || proj.id)}</h2>
          <span class="pgroup-workspace-health pgroup-workspace-health--${projectHealthTone}">${pgroupHealthDotHtml(projectHealthTone, projectHealth)}${escapeHtml(projectHealth)}</span>
        </div>
        <div class="pgroup-workspace-summary project-spo-summary">
          <span><b>${escapeHtml(proj.id)}</b> project ID</span>
          <span><b>${openTrackerCount}</b> open task${openTrackerCount === 1 ? "" : "s"}</span>
          <span><b>${riskCount}</b> open risk${riskCount === 1 ? "" : "s"}</span>
        </div>
        <nav class="project-spo-menu pgroup-workspace-nav">
          ${navItems.map((item) => `
            <button type="button" class="project-spo-link pgroup-workspace-nav-btn ${activeTab === item.key ? "active" : ""}" data-tab="${item.key}">
              <i class="bx ${item.icon}"></i>
              <span>${item.label}</span>
              ${item.key === "tracker" && trackerCount ? `<em class="project-spo-badge">${trackerCount}</em>` : ""}
              ${item.key === "risks" && riskCount ? `<em class="project-spo-badge">${riskCount}</em>` : ""}
            </button>`).join("")}
        </nav>
      </aside>
      <main class="project-spo-main pgroup-workspace-main">
        <header class="project-spo-main-head pgroup-workspace-head"><div><span class="pgroup-kicker">${escapeHtml(activeHeading.kicker)}</span><h3>${escapeHtml(activeHeading.title)}</h3></div><div class="project-spo-main-actions"><button type="button" class="btn-aewttr-outline btn-aewttr-sm" id="project-workspace-back"><i class="bx bx-grid-alt"></i> All Projects</button><button type="button" class="btn-aewttr btn-aewttr-sm" id="project-workspace-settings"><i class="bx bx-cog"></i> Project Settings</button></div></header>
        <div id="proj-tab-body" class="project-spo-tab-body"></div>
      </main>
    </div>
  `;
  $all(".project-spo-link", $("#page-content")).forEach((link) => link.addEventListener("click", () => navigate(`projects/${id}/${link.dataset.tab}`)));
  $("#project-sidebar-back", $("#page-content")).addEventListener("click", () => navigate("projects"));
  $("#project-workspace-back", $("#page-content")).addEventListener("click", () => navigate("projects"));
  $("#project-workspace-settings", $("#page-content")).addEventListener("click", () => navigate(`projects/${id}/settings`));

  const body = $("#proj-tab-body");
  if (tab === "workspace") return drawWorkspace(body, proj);
  if (tab === "notes") return drawProjectNotes(body, proj);
  if (tab === "people") return drawPeople(body, proj);
  if (tab === "documents") return drawProjectDocuments(body, proj);
  if (tab === "photos") return drawProjectPhotos(body, proj);
  if (tab === "risks") return drawProjectRisks(body, proj);
  if (tab === "meeting") return drawProjectMeeting(body, proj);
  if (tab === "tracker") return drawTracker(body, proj);
  if (tab === "reporting") return drawProjectReporting(body, proj);
  if (tab === "import") return drawImport(body, proj);
  if (tab === "finance") return drawFinance(body, proj);
  if (tab === "tickets") return drawProjectTickets(body, proj);
  if (tab === "settings") return drawProjectSettings(body, proj);
  navigate(`projects/${id}/workspace`);
}

async function exportProjectXlsx(proj) {
  const all = window.AEWTTR.db.ganttTasks[proj.id] || [];
  const tasks = trackerPlainTasks(all);
  const columns = ["Project ID", "Project Name", "Status", "Title", "Assignee", "Health", "Progress", "Start", "End", "Notes"];
  const dataRows = tasks.map((task) => [
    proj.id,
    proj.name,
    computeProjectStatus(proj),
    task.title || "",
    task.assignee || "",
    task.health || "",
    `${taskProgressPct(task)}%`,
    task.start || "",
    task.end || "",
    (Array.isArray(task.notes)
      ? task.notes.map((n) => n.text || "").filter(Boolean).join(" | ")
      : String(task.notes || "")).replace(/\n/g, " ")
  ]);
  const service = window.AEWTTR && window.AEWTTR.ExportService;
  if (!service || typeof service.exportXlsx !== "function") {
    toast("Excel export is unavailable in this package.", "error");
    return;
  }
  const fileName = `project-${proj.id}-tracker.xlsx`;
  const popup = typeof service.reserveOpen === "function"
    ? service.reserveOpen(fileName, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    : null;
  try {
    const result = await service.exportXlsx(fileName, columns, dataRows, {
      popup,
      folderName: "Project Exports",
      title: "PULSE Project Tracker",
      subtitle: `${proj.id || "Project"} | ${proj.name || "Project"} | ${computeProjectStatus(proj)}`,
      sectionLabel: "PROJECT TASK INVENTORY",
      footerLabel: "PULSE PROJECT TRACKER",
      sheetName: "Project Tracker"
    });
    if (result.mode === "sharepoint") toast("Project workbook saved and opened in SharePoint.", "success");
    else if (result.uploadError) toast("SharePoint upload failed — downloaded a local copy instead.", "warn");
    else toast("Project workbook downloaded.", "success");
  } catch (error) {
    closeStatusExportOpen(popup);
    toast((error && error.message) || "Excel export failed.", "error");
  }
}

function reserveStatusExportOpen(fileName) {
  const api = window.AEWTTR && window.AEWTTR.OfficeDesktop;
  if (typeof isSharePointMode !== "function" || !isSharePointMode()
    || !api || typeof api.reserveSharePointFileWindow !== "function") {
    return null;
  }
  return api.reserveSharePointFileWindow(
    fileName || "PULSE-status-briefing.pptx",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  );
}

function closeStatusExportOpen(popup) {
  const api = window.AEWTTR && window.AEWTTR.OfficeDesktop;
  if (api && typeof api.closeReservedSharePointFileWindow === "function") {
    api.closeReservedSharePointFileWindow(popup);
  } else if (popup && !popup.closed) {
    popup.close();
  }
}

function toastStatusPptxSaved(fileUrl, fileName) {
  let stack = $(".aewttr-toast-stack");
  if (!stack) {
    stack = el(`<div class="aewttr-toast-stack"></div>`);
    document.body.appendChild(stack);
  }
  const openUrl = `${fileUrl}${fileUrl.includes("?") ? "&" : "?"}web=1`;
  const t = el(`<div class="aewttr-toast success">PowerPoint saved and opened in SharePoint. <a href="${escapeHtml(openUrl)}" target="_blank" rel="noopener" style="color:#9ed0ff;text-decoration:underline;">Open again</a></div>`);
  stack.appendChild(t);
  setTimeout(() => { t.style.opacity = "0"; t.style.transition = "opacity .25s"; setTimeout(() => t.remove(), 250); }, 10000);
}

async function exportProjectPptx(proj) {
  const api = window.AEWTTR.ProjectPptxExport;
  if (!api || typeof api.exportProjectStatusPptxToSharePoint !== "function") {
    toast("PowerPoint export is unavailable in this package.", "error");
    return;
  }
  const openTarget = reserveStatusExportOpen(`${proj.id || "project"}-project-update.pptx`);
  try {
    toast("Building status PowerPoint…", "info");
    const result = await api.exportProjectStatusPptxToSharePoint(proj, { popup: openTarget });
    if (result.mode === "sharepoint" && result.fileUrl) {
      toastStatusPptxSaved(result.fileUrl, result.fileName);
    } else if (result.uploadError) {
      const detail = (result.uploadError && (result.uploadError.friendly || result.uploadError.message)) || "upload failed";
      toast(`SharePoint upload failed (${detail}) — downloaded a local copy instead.`, "warn");
    } else {
      toast("Status PowerPoint downloaded", "success");
    }
  } catch (e) {
    closeStatusExportOpen(openTarget);
    console.warn("PULSE: PPTX export failed", e);
    toast(`Export failed: ${(e && e.message) || "unknown error"}`, "error");
  }
}

function openProjectExportMenu(proj, anchor) {
  const existing = document.querySelector(".project-export-menu");
  if (existing) existing.remove();
  const menu = el(`<div class="project-export-menu" role="menu">
    <button type="button" role="menuitem" data-export="pptx"><i class="bx bxs-slideshow"></i> Status deck (.pptx)</button>
    <button type="button" role="menuitem" data-export="xlsx"><i class="bx bx-spreadsheet"></i> Tracker workbook (.xlsx)</button>
  </div>`);
  document.body.appendChild(menu);
  const rect = anchor.getBoundingClientRect();
  menu.style.top = `${Math.round(rect.bottom + 6)}px`;
  menu.style.left = `${Math.round(Math.min(rect.left, window.innerWidth - 220))}px`;
  function close() {
    menu.remove();
    document.removeEventListener("mousedown", onDoc);
  }
  function onDoc(ev) {
    if (!menu.contains(ev.target) && ev.target !== anchor) close();
  }
  document.addEventListener("mousedown", onDoc);
  $all("[data-export]", menu).forEach((btn) => btn.addEventListener("click", () => {
    const kind = btn.dataset.export;
    close();
    if (kind === "pptx") exportProjectPptx(proj);
    else exportProjectXlsx(proj);
  }));
}

function ensureWorkspaceExtra(extra, proj) {
  if (!extra.workspaceLinks) {
    extra.workspaceLinks = [
      { id: uid("wl"), label: "Tracker", route: `projects/${proj.id}/tracker`, icon: "bx-table" },
      { id: uid("wl"), label: "Meeting", route: `projects/${proj.id}/meeting`, icon: "bx-conversation" },
      { id: uid("wl"), label: "Documents", route: `projects/${proj.id}/documents`, icon: "bx-folder-open" }
    ];
  }
  return extra;
}

function projectRoleHomeChipHtml(proj, field) {
  if (field.key === "contractor") {
    const company = proj.contractor || "";
    return `
      <div class="project-home-role-chip${company ? "" : " is-empty"}">
        <span class="project-home-role-label">${escapeHtml(field.label)}</span>
        <span class="project-home-role-name">${escapeHtml(company || "Unassigned")}</span>
      </div>`;
  }
  const entry = resolvedAssignedPerson(proj, field.key);
  return `
    <div class="project-home-role-chip${entry ? "" : " is-empty"}">
      <span class="project-home-role-label">${escapeHtml(field.label)}</span>
      <span class="project-home-role-name">${escapeHtml(entry ? entry.label : "Unassigned")}</span>
    </div>`;
}

function projectRoleCandidateList(proj) {
  const db = window.AEWTTR.db;
  const roster = (db.projectPeople && db.projectPeople[proj.id]) || [];
  const rosterMemberIds = new Set(roster.filter((p) => p.memberId).map((p) => p.memberId));
  const rosterCandidates = roster.map((p) => ({ source: "roster", personId: p.id, name: p.label, email: p.email || "", sub: p.company || (p.type === "member" ? "Team member" : "Project contact") }));
  const siteUsers = getMemberDirectory()
    .filter((m) => !rosterMemberIds.has(m.id))
    .map((m) => ({ source: "member", memberId: m.id, name: m.name, email: m.email || "", sub: "Site user" }));
  return rosterCandidates.concat(siteUsers);
}

function projectRolePickerHtml(proj, field) {
  const entry = resolvedAssignedPerson(proj, field.key);
  const pickerId = `role-picker-${field.key}`;
  return `
    <div class="project-role-picker" id="${pickerId}" data-role-key="${field.key}" data-role-label="${escapeHtml(field.label)}">
      <div class="project-role-picker-head">
        <label class="project-role-picker-label">${escapeHtml(field.label)}</label>
        ${entry ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-sm project-role-clear"${tip(`Clear ${field.label}`)}>Clear</button>` : ""}
      </div>
      ${entry ? `
        <div class="project-role-current">
          ${userAvatarHtml(entry.label, entry.email || memberEmailForPerson(entry.label), 28)}
          <div class="project-role-current-copy">
            <strong>${escapeHtml(entry.label)}</strong>
            <span>${escapeHtml(entry.email || entry.company || "No email on file")}</span>
          </div>
        </div>` : `<div class="project-role-current is-empty">No one assigned yet</div>`}
      <div class="project-role-search-wrap">
        <i class="bx bx-search"></i>
        <input type="search" class="input-aewttr project-role-search" placeholder="Type a name to replace…" autocomplete="off">
      </div>
      <div class="project-role-suggestions traveler-suggestions" hidden></div>
      <button type="button" class="btn-aewttr-outline btn-aewttr-sm project-role-new-btn"><i class="bx bx-user-plus"></i> Add new person</button>
      <div class="project-role-new-form" hidden>
        <div class="form-row"><label>Name</label><input class="input-aewttr project-role-new-name" placeholder="Full name"></div>
        <div class="form-grid-2">
          <div class="form-row"><label>Email</label><input class="input-aewttr project-role-new-email" placeholder="name@example.com"></div>
          <div class="form-row"><label>Company</label><input class="input-aewttr project-role-new-company" placeholder="Optional"></div>
        </div>
        <div class="project-role-new-actions">
          <button type="button" class="btn-aewttr-ghost btn-aewttr-sm project-role-new-cancel">Cancel</button>
          <button type="button" class="btn-aewttr btn-aewttr-sm project-role-new-save"><i class="bx bx-check"></i> Add &amp; assign</button>
        </div>
      </div>
    </div>`;
}

function projectContractorCompanyPickerHtml(proj) {
  const company = proj.contractor || "";
  return `
    <div class="project-role-picker" id="role-picker-contractor-company" data-contractor-company-picker>
      <div class="project-role-picker-head">
        <label class="project-role-picker-label">Contractor</label>
        ${company ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-sm project-contractor-clear">Clear</button>` : ""}
      </div>
      ${company ? `
        <div class="project-role-current">
          <i class="bx bx-building" style="font-size:20px;flex-shrink:0;color:var(--aewttr-muted);"></i>
          <div class="project-role-current-copy"><strong>${escapeHtml(company)}</strong><span>Contractor company</span></div>
        </div>` : `<div class="project-role-current is-empty">No company assigned</div>`}
      <div class="project-role-search-wrap">
        <i class="bx bx-search"></i>
        <input type="search" class="input-aewttr project-contractor-company-search" placeholder="Search contractor companies…" autocomplete="off">
      </div>
      <div class="project-role-suggestions traveler-suggestions project-contractor-company-suggestions" hidden></div>
    </div>`;
}

function wireProjectContractorCompanyPicker(scope, proj, onUpdate) {
  const db = window.AEWTTR.db;
  const picker = $("[data-contractor-company-picker]", scope);
  if (!picker) return;

  function getAllCompanies() {
    const seen = new Set();
    const list = [];
    Object.values(db.projectContractors || {}).forEach((arr) => {
      (arr || []).forEach((c) => {
        if (c.company && !seen.has(c.company.toLowerCase())) {
          seen.add(c.company.toLowerCase());
          list.push(c.company);
        }
      });
    });
    return list.sort();
  }

  function redraw() {
    const parent = picker.parentElement;
    const fresh = el(projectContractorCompanyPickerHtml(proj));
    picker.replaceWith(fresh);
    wireProjectContractorCompanyPicker(scope, proj, onUpdate);
  }

  async function selectCompany(name) {
    proj.contractor = name;
    try {
      await Repo.save("project", proj);
      toast("Contractor assigned", "success");
    } catch (e) { /* toast already shown */ }
    redraw();
    if (typeof onUpdate === "function") onUpdate();
  }

  const clearBtn = $(".project-contractor-clear", picker);
  if (clearBtn) clearBtn.addEventListener("click", async () => {
    proj.contractor = "";
    try { await Repo.save("project", proj); } catch (e) {}
    redraw();
    if (typeof onUpdate === "function") onUpdate();
  });

  const searchInput = $(".project-contractor-company-search", picker);
  const suggestions = $(".project-contractor-company-suggestions", picker);
  if (!searchInput) return;

  searchInput.addEventListener("input", () => {
    const q = (searchInput.value || "").trim().toLowerCase();
    if (!q) { suggestions.hidden = true; suggestions.innerHTML = ""; return; }
    const matches = getAllCompanies().filter((c) => c.toLowerCase().includes(q)).slice(0, 8);
    if (!matches.length) {
      suggestions.hidden = false;
      suggestions.innerHTML = `<button type="button" class="traveler-suggestion" data-company="${escapeHtml(searchInput.value.trim())}"><strong>${escapeHtml(searchInput.value.trim())}</strong><small>Add as new company</small></button>`;
    } else {
      suggestions.hidden = false;
      suggestions.innerHTML = matches.map((c) => `<button type="button" class="traveler-suggestion" data-company="${escapeHtml(c)}"><strong>${escapeHtml(c)}</strong><small>Contractor company</small></button>`).join("");
    }
    $all("[data-company]", suggestions).forEach((btn) => btn.addEventListener("click", () => selectCompany(btn.dataset.company)));
  });

  document.addEventListener("click", function hide(e) {
    if (!picker.contains(e.target)) { suggestions.hidden = true; document.removeEventListener("click", hide); }
  }, true);
}

async function assignProjectRole(proj, fieldKey, entry, fieldLabel) {
  if (typeof reanchorProject === "function") reanchorProject(proj);
  const previous = proj[fieldKey];
  proj[fieldKey] = entry.id;
  try {
    await Repo.save("project", proj);
    if (typeof syncProjectPulseGroup === "function") {
      try { await syncProjectPulseGroup(proj); } catch (e) { console.warn("project group sync", e); }
    }
    toast(`${fieldLabel} assigned`, "success");
  } catch (e) {
    proj[fieldKey] = previous;
    throw e;
  }
}

function wireProjectRolePickers(scope, proj, onUpdate) {
  const db = window.AEWTTR.db;
  if (!db.projectPeople) db.projectPeople = {};
  if (!db.projectPeople[proj.id]) db.projectPeople[proj.id] = [];
  const roster = db.projectPeople[proj.id];

  $all(".project-role-picker", scope).forEach((picker) => {
    if (picker.dataset.contractorCompanyPicker !== undefined) return;
    const fieldKey = picker.dataset.roleKey;
    const fieldLabel = picker.dataset.roleLabel;
    const searchInput = $(".project-role-search", picker);
    if (!searchInput) return;
    const suggestions = $(".project-role-suggestions", picker);
    const newForm = $(".project-role-new-form", picker);
    const newBtn = $(".project-role-new-btn", picker);

    function redrawPicker() {
      const parent = picker.parentElement;
      const fresh = el(projectRolePickerHtml(proj, { key: fieldKey, label: fieldLabel }));
      picker.replaceWith(fresh);
      wireProjectRolePickers(parent, proj, onUpdate);
    }

    async function afterAssign() {
      redrawPicker();
      if (typeof onUpdate === "function") onUpdate();
    }

    async function pickCandidate(cand) {
      let entry = null;
      if (cand.source === "roster") entry = roster.find((p) => p.id === cand.personId);
      else entry = findOrCreateProjectPersonEntry(proj, { type: "member", memberId: cand.memberId, label: cand.name, email: cand.email, role: fieldLabel });
      if (!entry) return;
      try {
        await assignProjectRole(proj, fieldKey, entry, fieldLabel);
        await afterAssign();
      } catch (e) { /* toast already shown by Repo */ }
    }

    function hideSuggestions() {
      suggestions.hidden = true;
      suggestions.innerHTML = "";
    }

    function drawSuggestions() {
      const query = (searchInput.value || "").trim().toLowerCase();
      if (query.length < 2) {
        hideSuggestions();
        return;
      }
      const current = resolvedAssignedPerson(proj, fieldKey);
      const matches = projectRoleCandidateList(proj)
        .filter((c) => c.name.toLowerCase().includes(query) || (c.email || "").toLowerCase().includes(query))
        .filter((c) => {
          if (!current) return true;
          if (c.source === "roster" && c.personId === current.id) return false;
          if (c.source === "member" && current.memberId && c.memberId === current.memberId) return false;
          if (c.source === "member" && current.email && c.email && c.email.toLowerCase() === current.email.toLowerCase()) return false;
          return true;
        })
        .slice(0, 8);
      suggestions.hidden = false;
      suggestions.innerHTML = matches.length ? matches.map((c, i) => `
        <button type="button" class="traveler-suggestion" data-cand-index="${i}">
          <strong>${escapeHtml(c.name)}</strong>
          <span>${escapeHtml(c.email || c.sub || "No email on file")}</span>
        </button>`).join("") : `<div class="traveler-suggestion-empty">No matches — add them as a new person below.</div>`;
      $all("[data-cand-index]", suggestions).forEach((btn) => btn.addEventListener("click", () => pickCandidate(matches[Number(btn.dataset.candIndex)])));
    }

    searchInput.addEventListener("input", drawSuggestions);
    searchInput.addEventListener("focus", () => {
      if ((searchInput.value || "").trim().length >= 2) drawSuggestions();
      else hideSuggestions();
    });
    searchInput.addEventListener("blur", () => {
      // Delay so a mousedown on a suggestion still registers.
      setTimeout(() => {
        if (!picker.contains(document.activeElement)) hideSuggestions();
      }, 150);
    });

    const clearBtn = $(".project-role-clear", picker);
    if (clearBtn) clearBtn.addEventListener("click", async () => {
      if (typeof reanchorProject === "function") reanchorProject(proj);
      const previous = proj[fieldKey];
      proj[fieldKey] = "";
      try {
        await Repo.save("project", proj);
        toast(`${fieldLabel} assignment cleared`, "success");
        await afterAssign();
      } catch (e) {
        proj[fieldKey] = previous;
      }
    });

    newBtn.addEventListener("click", () => {
      newForm.hidden = false;
      newBtn.hidden = true;
      hideSuggestions();
      $(".project-role-new-name", picker).focus();
    });
    $(".project-role-new-cancel", picker).addEventListener("click", () => {
      newForm.hidden = true;
      newBtn.hidden = false;
    });
    $(".project-role-new-save", picker).addEventListener("click", async () => {
      const name = $(".project-role-new-name", picker).value.trim();
      if (!name) { toast("Enter a name.", "error"); return; }
      const entry = findOrCreateProjectPersonEntry(proj, {
        type: "person",
        memberId: "",
        label: name,
        email: $(".project-role-new-email", picker).value.trim(),
        company: $(".project-role-new-company", picker).value.trim(),
        role: fieldLabel
      });
      try {
        await assignProjectRole(proj, fieldKey, entry, fieldLabel);
        await afterAssign();
      } catch (e) { /* toast already shown by Repo */ }
    });
  });
}

/* Home is a read-only snapshot — every field/panel edits via the Settings
   tab, reached here through the header "Edit Details" button or a per-panel
   "Edit" button, never inline on Home itself. */
function drawWorkspace(body, proj) {
  const db = window.AEWTTR.db;
  const extra = (db.projectExtra && db.projectExtra[proj.id]) || { history: [], risks: [], handoff: "", meetingNotes: "", notes: [] };
  if (!extra.notes) extra.notes = [];
  const trackerTasks = (db.ganttTasks && db.ganttTasks[proj.id]) || [];
  const people = (db.projectPeople && db.projectPeople[proj.id]) || [];
  const doneStatus = "Complete";
  const doneTasks = trackerTasks.filter((t) => (t.status || "") === doneStatus).length;
  const openTasks = Math.max(0, trackerTasks.length - doneTasks);
  const homeRisks = projectRisks(proj.id);

  const headerActions = `
    ${proj.sharepointFolderUrl ? `<a href="${escapeHtml(proj.sharepointFolderUrl)}" target="_blank" class="btn-aewttr-outline btn-aewttr-sm"${tip("Open SharePoint Folder")}><i class="bx bxl-microsoft"></i> Open Folder</a>` : ""}
    <button type="button" class="btn-aewttr-outline btn-aewttr-sm" id="btn-export-project"${tip("Export a status PowerPoint or themed tracker workbook")}><i class="bx bx-download"></i> Export</button>
    <button type="button" class="btn-aewttr btn-aewttr-sm" id="btn-edit-details" data-route="projects/${proj.id}/settings"${tip("Edit project details")}><i class="bx bx-edit-alt"></i> Edit Details</button>
  `;

  const openRisks = homeRisks.filter((r) => r.status !== "Closed").length;
  const dueLabel = proj.dueDate ? fmtDate(proj.dueDate) : (proj.completionDate ? fmtDate(proj.completionDate) : "No due date");
  const locLabel = (proj.locations || []).length ? escapeHtml(proj.locations.join(", ")) : "No locations set";

  body.innerHTML = `
    <div class="project-home">
      <header class="project-home-header">
        <div class="project-home-header-top">
          <div class="project-home-header-id">
            ${proj.lifecycleStatus ? lifecyclePill(proj.lifecycleStatus) : ""}
            ${priorityTag(proj.priority)}
            ${(proj.technicalStatus || computeProjectTechStatus(proj)) ? ragPill(proj.technicalStatus || computeProjectTechStatus(proj)) : ""}
          </div>
          <div class="project-home-header-actions">${headerActions}</div>
        </div>
        <h1>${escapeHtml(proj.name)}</h1>
        ${proj.description ? `<p class="project-home-desc">${escapeHtml(proj.description)}</p>` : ""}
        <div class="project-home-header-meta">
          <span><i class="bx bx-calendar" aria-hidden="true"></i> ${proj.startDate ? fmtDate(proj.startDate) : "No start"} → ${dueLabel}</span>
          <span><i class="bx bx-map" aria-hidden="true"></i> ${locLabel}</span>
          ${openRisks ? `<span class="project-home-header-meta-risk"><i class="bx bx-shield-quarter" aria-hidden="true"></i> ${openRisks} open risk${openRisks === 1 ? "" : "s"}</span>` : ""}
        </div>
      </header>

      <div class="project-home-metrics" role="list" aria-label="Project snapshot">
        <div class="project-home-metric" role="listitem">
          <strong>${openTasks}</strong>
          <span class="k">Open tasks</span>
        </div>
        <div class="project-home-metric" role="listitem">
          <strong>${doneTasks}</strong>
          <span class="k">Complete</span>
        </div>
        <div class="project-home-metric" role="listitem">
          <strong>${people.length}</strong>
          <span class="k">People</span>
        </div>
      </div>

      <div class="project-home-panels">
        <section class="project-home-panel">
          <div class="project-home-panel-head">
            <h2>Key roles</h2>
            <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-route="projects/${proj.id}/settings">Edit</button>
          </div>
          <div class="project-home-roles">
            ${ASSIGNABLE_PROJECT_ROLE_FIELDS.map((field) => projectRoleHomeChipHtml(proj, field)).join("")}
          </div>
        </section>

        <section class="project-home-panel">
          <div class="project-home-panel-head">
            <h2>Program &amp; funding</h2>
            <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-route="projects/${proj.id}/settings">Edit</button>
          </div>
          <dl class="project-home-facts">
            <div><dt>Portfolio</dt><dd>${(projectPortfolios(proj) || []).length ? escapeHtml((projectPortfolios(proj)).join(", ")) : "—"}</dd></div>
            <div><dt>Program</dt><dd>${escapeHtml(proj.program) || "—"}</dd></div>
            <div><dt>Contract</dt><dd>${escapeHtml(proj.contract) || "—"}</dd></div>
            <div><dt>Task order</dt><dd>${escapeHtml(proj.taskOrder) || "—"}</dd></div>
            <div><dt>Funding type</dt><dd>${escapeHtml(proj.fundingType) || "—"}</dd></div>
            <div><dt>Fiscal year</dt><dd>${escapeHtml(proj.fiscalYear) || "—"}</dd></div>
            <div><dt>Funding status</dt><dd>${escapeHtml(proj.fundingStatus) || "—"}</dd></div>
            <div><dt>Funded on contract</dt><dd>${proj.fundedOnContractAmount ? `$${Number(proj.fundedOnContractAmount).toLocaleString()}` : "—"}</dd></div>
            <div><dt>Reimbursable</dt><dd>${proj.reimbursableAmount ? `$${Number(proj.reimbursableAmount).toLocaleString()}` : "—"}</dd></div>
            <div><dt>Config end item</dt><dd>${escapeHtml(proj.configEndItem) || "—"}</dd></div>
            <div><dt>Change request</dt><dd>${proj.changeRequestRequired ? "Required" : "Not required"}</dd></div>
            ${proj.fundingNotes ? `<div class="project-home-facts-span"><dt>Funding notes</dt><dd>${escapeHtml(proj.fundingNotes)}</dd></div>` : ""}
          </dl>
        </section>

        <section class="project-home-panel">
          <div class="project-home-panel-head">
            <h2>Schedule &amp; location</h2>
            <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-route="projects/${proj.id}/settings">Edit</button>
          </div>
          <dl class="project-home-facts project-home-facts--compact">
            <div><dt>Lifecycle</dt><dd>${escapeHtml(proj.lifecycleStatus) || "—"}</dd></div>
            <div><dt>Technical status</dt><dd>${escapeHtml(proj.technicalStatus) || "—"}</dd></div>
            <div><dt>Priority</dt><dd>${escapeHtml(proj.priority) || "—"}</dd></div>
            <div><dt>Start date</dt><dd>${proj.startDate ? fmtDate(proj.startDate) : "—"}</dd></div>
            <div><dt>Due date</dt><dd>${proj.dueDate ? fmtDate(proj.dueDate) : "—"}</dd></div>
            <div><dt>Completion date</dt><dd>${proj.completionDate ? fmtDate(proj.completionDate) : "—"}</dd></div>
            <div class="project-home-facts-span"><dt>Location</dt><dd>${(proj.locations || []).length ? escapeHtml(proj.locations.join(", ")) : "—"}</dd></div>
          </dl>
        </section>

        <section class="project-home-panel project-home-panel--scope">
          <div class="project-home-panel-head">
            <h2>Scope &amp; objectives</h2>
            <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-route="projects/${proj.id}/settings">Edit</button>
          </div>
          <div class="project-home-text-grid">
            <div class="project-home-text-block${proj.scope ? "" : " is-empty"}">
              <h3>Scope</h3>
              <p>${proj.scope ? escapeHtml(proj.scope) : "No scope defined yet."}</p>
            </div>
            <div class="project-home-text-block${proj.objectives ? "" : " is-empty"}">
              <h3>Objectives</h3>
              <p>${proj.objectives ? escapeHtml(proj.objectives) : "No objectives defined yet."}</p>
            </div>
          </div>
        </section>

        <section class="project-home-panel project-home-panel--risk" id="project-home-risk-panel">
          <div class="project-home-panel-head">
            <h2>Risks</h2>
            <div class="project-home-risk-tools">
              <div class="project-home-view-tabs" role="tablist" aria-label="Risk view">
                <button type="button" class="project-home-view-tab active" data-home-risk-view="matrix" role="tab" aria-selected="true">Matrix</button>
                <button type="button" class="project-home-view-tab" data-home-risk-view="burndown" role="tab" aria-selected="false">Risk Burn Down</button>
              </div>
              <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-route="projects/${proj.id}/risks">View all</button>
            </div>
          </div>
          <div class="project-home-risk-view" data-risk-view-panel="matrix">
            ${homeRisks.length ? riskMatrixHtml(homeRisks) : `<div class="project-home-empty">No risks logged yet.</div>`}
          </div>
          <div class="project-home-risk-view" data-risk-view-panel="burndown" hidden>
            ${homeRisks.length ? riskBurnDownHomeHtml(homeRisks) : `<div class="project-home-empty">No risks logged yet.</div>`}
          </div>
        </section>

        <section class="project-home-panel">
          <div class="project-home-panel-head">
            <h2>Quick links</h2>
          </div>
          <div class="project-home-quick-nav" aria-label="Quick links">
            <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-route="projects/${proj.id}/tracker"><i class="bx bx-task" aria-hidden="true"></i> Tracker</button>
            <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-route="projects/${proj.id}/people"><i class="bx bx-group" aria-hidden="true"></i> People</button>
            <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-route="projects/${proj.id}/meeting"><i class="bx bx-calendar-event" aria-hidden="true"></i> Meeting</button>
            <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-route="projects/${proj.id}/risks"><i class="bx bx-shield-quarter" aria-hidden="true"></i> Risks</button>
          </div>
        </section>

        <section class="project-home-panel project-home-panel--notes">
          <div class="project-home-panel-head">
            <h2>Notes</h2>
            <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-route="projects/${proj.id}/notes">View all</button>
          </div>
          <div class="proj-notes-widget" id="proj-home-notes-feed">
            ${(extra.notes || []).length === 0 ? `<div class="project-home-empty">No notes yet. Add one below.</div>` : (extra.notes || []).slice().reverse().slice(0, 4).map((n) => `
              <div class="proj-notes-widget-item">
                <div class="proj-notes-widget-meta"><strong>${escapeHtml(n.author || "Unknown")}</strong><span>${escapeHtml(n.date || "")}${n.time ? ` · ${escapeHtml(n.time)}` : ""}</span></div>
                <div class="proj-notes-widget-text">${escapeHtml(n.text || "")}</div>
              </div>`).join("")}
          </div>
          <div class="proj-notes-widget-compose">
            <input class="input-aewttr" id="proj-home-note-input" placeholder="Quick note…" autocomplete="off" aria-label="Quick note">
            <button type="button" class="btn-aewttr btn-aewttr-sm" id="proj-home-note-send" aria-label="Send note"${tip("Post note")}><i class="bx bx-send" aria-hidden="true"></i></button>
          </div>
        </section>
      </div>
    </div>
  `;

  const exportBtn = $("#btn-export-project", body);
  if (exportBtn) exportBtn.addEventListener("click", () => openProjectExportMenu(proj, exportBtn));

  $all("[data-home-risk-view]", body).forEach((btn) => btn.addEventListener("click", () => {
    const view = btn.dataset.homeRiskView;
    $all("[data-home-risk-view]", body).forEach((tab) => {
      const on = tab.dataset.homeRiskView === view;
      tab.classList.toggle("active", on);
      tab.setAttribute("aria-selected", on ? "true" : "false");
    });
    $all("[data-risk-view-panel]", body).forEach((panel) => {
      panel.hidden = panel.dataset.riskViewPanel !== view;
    });
  }));

  $all("[data-route]", body).forEach((node) => node.addEventListener("click", (e) => {
    if (!node.dataset.route) return;
    e.preventDefault();
    navigate(node.dataset.route);
  }));
  $all("[data-risk-open]", body).forEach((btn) => btn.addEventListener("click", () => openRiskModal(proj, btn.dataset.riskOpen, () => drawWorkspace(body, proj))));

  // Quick note from homepage widget
  const homeNoteInput = $("#proj-home-note-input", body);
  const homeNoteSend = $("#proj-home-note-send", body);
  if (homeNoteInput && homeNoteSend) {
    function postHomeNote() {
      const text = homeNoteInput.value.trim();
      if (!text) return;
      const me = (window.AEWTTR && window.AEWTTR.db && window.AEWTTR.db.user && window.AEWTTR.db.user.name) || "Unknown";
      const now = new Date();
      const note = {
        id: uid("pn"),
        author: me,
        date: now.toISOString().slice(0, 10),
        time: now.toTimeString().slice(0, 5),
        text
      };
      if (!extra.notes) extra.notes = [];
      extra.notes.unshift(note);
      Repo.save("project", proj);
      homeNoteInput.value = "";
      drawWorkspace(body, proj);
    }
    homeNoteSend.addEventListener("click", postHomeNote);
    homeNoteInput.addEventListener("keydown", (e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); postHomeNote(); } });
  }

  // Background refresh updates window.AEWTTR.db in place but never
  // re-renders on its own — the Home tab's "Key roles" chips (PM/Engineer/
  // ISSO/etc.) were static until you left and came back, so someone else's
  // role reassignment never showed up live. Re-registers on every draw
  // (removing the previous listener first) so it always calls the current
  // draw with the current body/proj — a Set-based register-once guard would
  // leave a stale closure pointing at a detached #proj-tab-body once you
  // navigate away and back to Home.
  if (window.AEWTTR._projectHomeLiveRefreshHandler) {
    window.removeEventListener("pulse:data-refreshed", window.AEWTTR._projectHomeLiveRefreshHandler);
  }
  window.AEWTTR._projectHomeLiveRefreshHandler = () => {
    if (!document.querySelector(".project-home")) return;
    if (document.querySelector(".aewttr-modal-backdrop")) return;
    const active = document.activeElement;
    if (active && (active.tagName === "INPUT" || active.tagName === "TEXTAREA" || active.tagName === "SELECT" || active.isContentEditable)) return;
    const current = (window.AEWTTR.db.projects || []).find((p) => p.id === proj.id);
    if (!current) return;
    const scroller = document.querySelector(".aewttr-content");
    const scrollTop = scroller ? scroller.scrollTop : 0;
    drawWorkspace(body, current);
    if (scroller) scroller.scrollTop = scrollTop;
  };
  window.addEventListener("pulse:data-refreshed", window.AEWTTR._projectHomeLiveRefreshHandler);
}

const RISK_RESPONSE_OPTIONS = ["Accept", "Avoid", "Mitigate", "Transfer", "Escalate"];
const RISK_MITIGATION_TYPES = ["Monitor and track", "Develop contingency plan", "Implement control measure", "Allocate additional resources", "Add schedule buffer", "Increase testing / validation", "Identify backup vendor", "Escalate to leadership", "Review contract terms", "Other / custom"];
const RISK_STATUS_OPTIONS = ["Open", "Monitoring", "Mitigating", "Escalated", "Closed"];

/* Derive a RAG status from the highest-severity active risk when no manual
   technicalStatus is set. Returns "Red", "Amber", "Green", or "". */
function computeProjectTechStatus(proj) {
  const risks = projectRisks(proj.id).filter(r => r.status !== "Closed");
  if (!risks.length) return "";
  const ratings = risks.map(r => r.rating || riskRating(r.likelihood, r.impact));
  if (ratings.some(r => r === "Red")) return "Red";
  if (ratings.some(r => r === "Amber")) return "Amber";
  return "Green";
}
const RISK_CATEGORY_OPTIONS = ["Schedule", "Cost", "Technical", "Cyber", "Contract", "Staffing", "Operational", "External"];

function riskScore(risk) {
  return Number(risk.likelihood || 0) * Number(risk.impact || 0);
}

function riskRating(likelihood, impact) {
  const score = Number(likelihood || 0) * Number(impact || 0);
  if (score >= 15) return "Red";
  if (score >= 6) return "Amber";
  return "Green";
}

function riskRatingRank(value) {
  return { Red: 3, Amber: 2, Green: 1 }[value] || 0;
}

function riskRatingPill(value) {
  const rating = value || "Green";
  return `<span class="risk-rating-pill risk-${escapeHtml(rating)}">${escapeHtml(rating)}</span>`;
}

/* Home / export burn-down: H/M/L counts over time from review dates,
   last-reviewed, and due dates. Thin history is synthesized as open→mid→now. */
function riskBurnDownHomeHtml(risks) {
  const exportApi = window.AEWTTR.ProjectPptxExport;
  const series = exportApi && typeof exportApi.riskBurnDownSeries === "function"
    ? exportApi.riskBurnDownSeries(risks)
    : [];
  const maxCount = Math.max(1, ...series.map((p) => (p.H || 0) + (p.M || 0) + (p.L || 0)));
  const w = 560;
  const h = 200;
  const padL = 36;
  const padR = 12;
  const padT = 16;
  const padB = 28;
  const plotW = w - padL - padR;
  const plotH = h - padT - padB;
  const bandH = plotH / 3;
  const pointsFor = (key) => series.map((p, i) => {
    const x = padL + (series.length <= 1 ? plotW / 2 : (i / (series.length - 1)) * plotW);
    const count = Number(p[key] || 0);
    const band = key === "H" ? 0 : key === "M" ? 1 : 2;
    const y = padT + band * bandH + bandH * (1 - Math.min(count / maxCount, 0.92));
    return `${x},${y}`;
  }).join(" ");

  const reviewRich = risks.some((r) => (r.reviewNotes || []).length > 1 || r.lastReviewedDate);
  const note = reviewRich
    ? "Series built from risk review dates and current H/M/L ratings."
    : "Limited review history — chart is a best-effort path from open risks to current ratings. Add review notes for a tighter burn-down.";

  return `
    <div class="risk-burndown-home">
      <svg class="risk-burndown-svg" viewBox="0 0 ${w} ${h}" role="img" aria-label="Risk burn down chart">
        <rect x="${padL}" y="${padT}" width="${plotW}" height="${bandH}" fill="rgba(163,43,34,.12)"></rect>
        <rect x="${padL}" y="${padT + bandH}" width="${plotW}" height="${bandH}" fill="rgba(201,162,39,.14)"></rect>
        <rect x="${padL}" y="${padT + bandH * 2}" width="${plotW}" height="${bandH}" fill="rgba(47,111,72,.12)"></rect>
        <text x="8" y="${padT + bandH * 0.55}" class="risk-burndown-axis">H</text>
        <text x="8" y="${padT + bandH * 1.55}" class="risk-burndown-axis">M</text>
        <text x="8" y="${padT + bandH * 2.55}" class="risk-burndown-axis">L</text>
        ${series.length ? `
          <polyline class="risk-burndown-line risk-burndown-line--h" fill="none" points="${pointsFor("H")}"></polyline>
          <polyline class="risk-burndown-line risk-burndown-line--m" fill="none" points="${pointsFor("M")}"></polyline>
          <polyline class="risk-burndown-line risk-burndown-line--l" fill="none" points="${pointsFor("L")}"></polyline>
        ` : ""}
        ${series.map((p, i) => {
          const x = padL + (series.length <= 1 ? plotW / 2 : (i / (series.length - 1)) * plotW);
          const label = p.date ? fmtDate(p.date) : "";
          return `<text x="${x}" y="${h - 8}" text-anchor="middle" class="risk-burndown-tick">${escapeHtml(label)}</text>`;
        }).join("")}
      </svg>
      <div class="risk-burndown-legend" aria-hidden="true">
        <span><i class="risk-burndown-swatch risk-burndown-swatch--h"></i> High</span>
        <span><i class="risk-burndown-swatch risk-burndown-swatch--m"></i> Medium</span>
        <span><i class="risk-burndown-swatch risk-burndown-swatch--l"></i> Low</span>
      </div>
      <p class="risk-burndown-note">${escapeHtml(note)}</p>
    </div>`;
}

function normalizeRiskRecord(risk, projectId) {
  const normalized = {
    id: risk.id || uid("risk"),
    projectId: risk.projectId || projectId || "",
    portfolio: risk.portfolio || "",
    name: risk.name || risk.title || "Untitled risk",
    title: risk.name || risk.title || "Untitled risk",
    description: risk.description || risk.text || "",
    text: risk.description || risk.text || risk.name || risk.title || "",
    owner: risk.owner || risk.ownerName || "",
    ownerName: risk.owner || risk.ownerName || "",
    ownerEmail: risk.ownerEmail || "",
    likelihood: Number(risk.likelihood || 1),
    impact: Number(risk.impact || 1),
    category: risk.category || "Operational",
    mitigationPlan: risk.mitigationPlan || "",
    responseStrategy: risk.responseStrategy || "Mitigate",
    due: risk.due || "",
    status: risk.status || "Open",
    lastReviewedDate: risk.lastReviewedDate || "",
    reviewNotes: Array.isArray(risk.reviewNotes) ? risk.reviewNotes : [],
    ratingHistory: Array.isArray(risk.ratingHistory) ? risk.ratingHistory : [],
    mitigations: Array.isArray(risk.mitigations) ? risk.mitigations : [],
    residualLikelihood: Number(risk.residualLikelihood || 0),
    residualImpact: Number(risk.residualImpact || 0)
  };
  normalized.rating = risk.rating || risk.rag || risk.level || riskRating(normalized.likelihood, normalized.impact);
  normalized.level = normalized.rating;
  if (risk._spId) normalized._spId = risk._spId;
  return normalized;
}

function projectRisks(projectId) {
  const db = window.AEWTTR.db;
  const extra = (db.projectExtra && db.projectExtra[projectId]) || {};
  const seen = new Set();
  return (extra.risks || []).map((risk) => normalizeRiskRecord(risk, projectId)).filter((risk) => {
    const key = risk._spId
      ? `sp:${risk._spId}`
      : `sig:${String(risk.name || "").trim().toLowerCase()}|${String(risk.description || "").trim().toLowerCase()}|${risk.due || ""}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

/* Keep the project mirror and the dedicated PULSE Risks record in lockstep.
   The previous delete paths removed a normalized display copy and then saved
   the project only in the background. A refresh could therefore restore the
   risk before its list-item delete completed. */
async function deleteProjectRisk(proj, risk) {
  const extra = ensureProjectExtra(proj.id);
  if (!Array.isArray(extra.risks)) extra.risks = [];
  const signatureFor = (item) => [
    String(item.name || item.title || "").trim().toLowerCase(),
    String(item.description || item.text || "").trim().toLowerCase(),
    String(item.due || ""),
    String(item.owner || item.ownerName || "").trim().toLowerCase()
  ].join("|");
  const riskSignature = signatureFor(risk);
  const index = extra.risks.findIndex((item) =>
    item === risk ||
    (item.id && risk.id && String(item.id) === String(risk.id)) ||
    (item._spId && risk._spId && String(item._spId) === String(risk._spId)) ||
    // Some older project mirrors predate risk IDs. Match their stable
    // operational fields so those risks remain deletable too.
    (riskSignature !== "|||" && signatureFor(item) === riskSignature)
  );
  if (index < 0) return null;
  const removed = extra.risks.splice(index, 1)[0];
  try {
    if (typeof reanchorProject === "function") reanchorProject(proj);
    // Commit the project-side RisksJson first so a reload cannot resurrect it.
    await Repo.save("project", proj);
    await Repo.remove("risk", removed);
    return { removed, index };
  } catch (error) {
    extra.risks.splice(index, 0, removed);
    // Best-effort repair of the project mirror if its save was the operation
    // that failed after the optimistic UI update.
    Repo.save("project", proj).catch(() => {});
    throw error;
  }
}

function allProjectRisks() {
  const db = window.AEWTTR.db;
  return (db.projects || []).flatMap((project) => projectRisks(project.id).map((risk) => ({ risk, project })));
}

function portfolioForRisk(risk, project) {
  if (risk.portfolio) return risk.portfolio;
  const portfolios = typeof projectPortfolios === "function" ? projectPortfolios(project) : (project.portfolios || []);
  return (portfolios && portfolios[0]) || project.program || project.product || "Unassigned";
}

function riskOwnersForProject(projectId) {
  const db = window.AEWTTR.db;
  const extra = (db.projectExtra && db.projectExtra[projectId]) || {};
  const names = new Set();
  (extra.risks || []).forEach((risk) => {
    const owner = risk.owner || risk.ownerName;
    if (owner) names.add(owner);
  });
  ((db.projectPeople && db.projectPeople[projectId]) || []).forEach((person) => {
    if (person.label) names.add(person.label);
  });
  (db.members || []).forEach((member) => { if (member.name) names.add(member.name); });
  return Array.from(names).sort((a, b) => a.localeCompare(b));
}

function sortRisksHighestFirst(risks) {
  return risks.slice().sort((a, b) => {
    const aRating = riskRatingRank(a.rating || riskRating(a.likelihood, a.impact));
    const bRating = riskRatingRank(b.rating || riskRating(b.likelihood, b.impact));
    return bRating - aRating || riskScore(b) - riskScore(a) || String(a.due || "9999").localeCompare(String(b.due || "9999"));
  });
}

/* Every risk name inside a cell is its own clickable button (data-risk-open)
   so the matrix — on both the Risks tab and the Project Home widget — opens
   straight to that risk's details instead of just being a static count.
   Callers must wire `[data-risk-open]` after inserting this HTML. */
function riskMatrixHtml(risks) {
  const byCell = new Map();
  risks.forEach((risk) => {
    const likelihood = Math.max(1, Math.min(5, Number(risk.likelihood || 1)));
    const impact = Math.max(1, Math.min(5, Number(risk.impact || 1)));
    const key = `${likelihood}-${impact}`;
    if (!byCell.has(key)) byCell.set(key, []);
    byCell.get(key).push(risk);
  });
  const rows = [];
  for (let impact = 5; impact >= 1; impact--) {
    rows.push(`
      <div class="risk-matrix-axis risk-matrix-impact">${impact}</div>
      ${[1, 2, 3, 4, 5].map((likelihood) => {
        const cellRisks = byCell.get(`${likelihood}-${impact}`) || [];
        const rating = riskRating(likelihood, impact);
        const shown = cellRisks.slice(0, 3);
        const overflow = cellRisks.length - shown.length;
        return `
          <div class="risk-matrix-cell risk-matrix-cell--${rating.toLowerCase()}${cellRisks.length ? "" : " is-empty"}"${tip(`${cellRisks.length} risk${cellRisks.length === 1 ? "" : "s"} at likelihood ${likelihood}, impact ${impact}`)}>
            <strong>${cellRisks.length || ""}</strong>
            <div class="risk-matrix-items">
              ${shown.map((risk) => `<button type="button" class="risk-matrix-item" data-risk-open="${escapeHtml(risk.id)}" title="${escapeHtml(risk.name)}">${escapeHtml(risk.name)}</button>`).join("")}
              ${overflow > 0 ? `<span class="risk-matrix-item-overflow">+${overflow} more</span>` : ""}
            </div>
          </div>`;
      }).join("")}`);
  }
  return `
    <div class="risk-matrix-wrap">
      <div class="risk-matrix-title-row"><span>Impact</span><em>Likelihood →</em></div>
      <div class="risk-matrix-grid">
        ${rows.join("")}
        <div class="risk-matrix-axis"></div>
        ${[1, 2, 3, 4, 5].map((n) => `<div class="risk-matrix-axis">${n}</div>`).join("")}
      </div>
    </div>`;
}

/* Inline register — prioritized columns: title, severity (RAG), status,
   owner, due / last reviewed, then L/I for scoring. Category + delete live
   in the details modal / context menu so the table stays scannable. */
function riskInlineTableHtml(risks, owners) {
  const ownerListId = "risk-owner-datalist";
  const rowsHtml = risks.map((risk) => {
    const rating = risk.rating || riskRating(risk.likelihood, risk.impact);
    const residualL = Number(risk.residualLikelihood || 0);
    const residualI = Number(risk.residualImpact || 0);
    const residualRating = (residualL && residualI) ? riskRating(residualL, residualI) : null;
    const hasNotes = !!(risk.description || risk.mitigationPlan || (risk.reviewNotes || []).length || (risk.mitigations || []).length);
    const statusClass = `risk-status-${String(risk.status || "Open").replace(/\s+/g, "-")}`;
    return `
      <tr class="monday-row monday-row--risk" data-risk-id="${escapeHtml(risk.id)}">
        <td class="risk-expand-td"><button type="button" class="risk-expand-btn" data-risk-expand="${escapeHtml(risk.id)}" title="Expand details">›</button></td>
        <td class="monday-item-cell risk-title-td"><input type="text" class="monday-inline-text" data-risk-field="name" data-risk-id="${escapeHtml(risk.id)}" value="${escapeHtml(risk.name)}" placeholder="Risk title"></td>
        <td class="risk-rating-td" data-risk-rating-cell="${escapeHtml(risk.id)}">${riskRatingPill(rating)}${residualRating ? `<span class="risk-residual-arrow" aria-hidden="true">→</span>${riskRatingPill(residualRating)}` : ""}</td>
        <td class="monday-select-cell">
          <select class="monday-status-select ${statusClass}" data-risk-field="status" data-risk-id="${escapeHtml(risk.id)}" aria-label="Status">
            ${RISK_STATUS_OPTIONS.map((option) => `<option ${risk.status === option ? "selected" : ""}>${escapeHtml(option)}</option>`).join("")}
          </select>
        </td>
        <td class="monday-owner-td"><input type="text" class="monday-inline-text" list="${ownerListId}" data-risk-field="owner" data-risk-id="${escapeHtml(risk.id)}" value="${escapeHtml(risk.owner)}" placeholder="Unassigned"></td>
        <td class="monday-date-td"><input type="date" class="monday-date-input" data-risk-field="due" data-risk-id="${escapeHtml(risk.id)}" value="${escapeHtml(risk.due || "")}" aria-label="Due date"></td>
        <td class="risk-reviewed-td" data-risk-reviewed-cell="${escapeHtml(risk.id)}">${risk.lastReviewedDate ? escapeHtml(fmtDate(risk.lastReviewedDate)) : "—"}</td>
        <td class="monday-select-cell">
          <select class="monday-status-select risk-li-select" data-risk-field="likelihood" data-risk-id="${escapeHtml(risk.id)}" aria-label="Likelihood">
            ${[1, 2, 3, 4, 5].map((n) => `<option value="${n}" ${Number(risk.likelihood) === n ? "selected" : ""}>${n}</option>`).join("")}
          </select>
        </td>
        <td class="monday-select-cell">
          <select class="monday-status-select risk-li-select" data-risk-field="impact" data-risk-id="${escapeHtml(risk.id)}" aria-label="Impact">
            ${[1, 2, 3, 4, 5].map((n) => `<option value="${n}" ${Number(risk.impact) === n ? "selected" : ""}>${n}</option>`).join("")}
          </select>
        </td>
        <td class="monday-notes-td">
          <button type="button" class="monday-notes-btn" data-risk-details="${escapeHtml(risk.id)}"${tip("Description, category, mitigation & review notes")}>
            <i class="bx bx-note"></i>${hasNotes ? `<span class="monday-notes-btn-count">•</span>` : ""}
          </button>
        </td>
      </tr>
      <tr class="risk-detail-row risk-detail-row--hidden" data-detail-for="${escapeHtml(risk.id)}">
        <td colspan="10" class="risk-detail-td"></td>
      </tr>`;
  }).join("");
  return `
    <div class="monday-table-wrap">
      <datalist id="${ownerListId}">${(owners || []).map((owner) => `<option value="${escapeHtml(owner)}"></option>`).join("")}</datalist>
      <table class="monday-table monday-table--risk">
        <thead><tr>
          <th class="risk-expand-td"></th>
          <th class="risk-th-title"><button class="risk-col-filter-btn" data-col-filter="name">Risk <i class="bx bx-filter-alt"></i></button></th>
          <th class="risk-th-severity"><button class="risk-col-filter-btn" data-col-filter="rating">Severity <i class="bx bx-filter-alt"></i></button></th>
          <th class="risk-th-status"><button class="risk-col-filter-btn" data-col-filter="status">Status <i class="bx bx-filter-alt"></i></button></th>
          <th class="risk-th-owner"><button class="risk-col-filter-btn" data-col-filter="owner">Owner <i class="bx bx-filter-alt"></i></button></th>
          <th class="risk-th-due"><button class="risk-col-filter-btn" data-col-filter="due">Due <i class="bx bx-filter-alt"></i></button></th>
          <th class="risk-th-reviewed"><button class="risk-col-filter-btn" data-col-filter="lastReviewedDate">Reviewed <i class="bx bx-filter-alt"></i></button></th>
          <th class="risk-th-li"><button class="risk-col-filter-btn" data-col-filter="likelihood">Likelihood <i class="bx bx-filter-alt"></i></button></th>
          <th class="risk-th-li"><button class="risk-col-filter-btn" data-col-filter="impact">Impact <i class="bx bx-filter-alt"></i></button></th>
          <th class="risk-th-notes">Notes</th>
        </tr></thead>
        <tbody>
          ${rowsHtml}
          <tr class="monday-inline-add-row"><td colspan="10"><button type="button" class="monday-inline-add-btn" id="risk-inline-add"><i class="bx bx-plus"></i> Add risk</button></td></tr>
        </tbody>
      </table>
      ${risks.length ? "" : `<div class="monday-table-empty">No risks match this view. Click "Add risk" to log one.</div>`}
    </div>`;
}

function riskReviewBoardHtml(risks) {
  const lanes = RISK_STATUS_OPTIONS;
  const byStatus = new Map(lanes.map((lane) => [lane, []]));
  risks.forEach((risk) => {
    const lane = byStatus.has(risk.status) ? risk.status : "Open";
    byStatus.get(lane).push(risk);
  });
  return `
    <div class="risk-review-board">
      ${lanes.map((lane) => {
        const laneRisks = byStatus.get(lane) || [];
        return `
          <section class="risk-review-lane">
            <div class="risk-review-lane-head"><strong>${lane}</strong><span>${laneRisks.length}</span></div>
            <div class="risk-review-cards">
              ${laneRisks.length ? laneRisks.map((risk) => `
                <button type="button" class="risk-review-card" data-review-risk="${escapeHtml(risk.id)}">
                  <div class="risk-review-card-top">${riskRatingPill(risk.rating)}<span>${escapeHtml(risk.category || "Operational")}</span></div>
                  <strong>${escapeHtml(risk.name)}</strong>
                  <em>${escapeHtml(risk.owner || "Unassigned")} · Due ${risk.due ? fmtDate(risk.due) : "—"}</em>
                </button>`).join("") : `<div class="risk-review-empty">No risks</div>`}
            </div>
          </section>`;
      }).join("")}
    </div>`;
}

function riskPassesFilters(risk, state) {
  if (state.filter === "Red / Amber" && !["Red", "Amber"].includes(risk.rating)) return false;
  if (state.filter === "Open" && risk.status !== "Open") return false;
  if (state.filter === "Active" && risk.status === "Closed") return false;
  if (state.owner !== "All" && risk.owner !== state.owner) return false;
  if (state.search) {
    const q = state.search.toLowerCase();
    const haystack = `${risk.name} ${risk.description} ${risk.owner} ${risk.category} ${risk.mitigationPlan} ${risk.status}`.toLowerCase();
    if (!haystack.includes(q)) return false;
  }
  return true;
}

function riskSummaryFromRaw(rawRisks) {
  let red = 0, amber = 0, active = 0;
  rawRisks.forEach((risk) => {
    if (risk.rating === "Red") red += 1;
    else if (risk.rating === "Amber") amber += 1;
    if (risk.status !== "Closed") active += 1;
  });
  return { red, amber, active, posture: red ? "Red" : amber ? "Amber" : "Green" };
}

function updateRiskSummaryStrip(body, summary) {
  const strip = $(".risk-summary-strip", body);
  if (!strip) return;
  strip.innerHTML = `
    <div class="risk-summary-card risk-summary-card--posture risk-summary-card--${summary.posture.toLowerCase()}"><span>Posture</span><strong>${summary.posture}</strong></div>
    <div class="risk-summary-card"><span>Active</span><strong>${summary.active}</strong></div>
    <div class="risk-summary-card risk-summary-card--tone-red"><span>Red</span><strong>${summary.red}</strong></div>
    <div class="risk-summary-card risk-summary-card--tone-amber"><span>Amber</span><strong>${summary.amber}</strong></div>`;
}

function drawProjectRisks(body, proj, opts) {
  opts = opts || {};
  if (!window.AEWTTR.state.projectRisks) window.AEWTTR.state.projectRisks = {};
  if (!window.AEWTTR.state.projectRisks[proj.id]) {
    window.AEWTTR.state.projectRisks[proj.id] = { view: "register", filter: "Active", owner: "All", search: "", colSort: null, colSortDir: "asc", colFilters: {}, expanded: {} };
  }
  const state = window.AEWTTR.state.projectRisks[proj.id];
  if (!state.colFilters) state.colFilters = {};
  if (!state.colSort) state.colSort = null;
  if (!state.colSortDir) state.colSortDir = "asc";
  const rawRisks = projectRisks(proj.id);
  let risks = rawRisks.filter((risk) => riskPassesFilters(risk, state));
  // Apply column value filters
  Object.keys(state.colFilters).forEach((col) => {
    const allowed = state.colFilters[col];
    if (allowed && allowed.size) risks = risks.filter((r) => allowed.has(riskColValue(r, col)));
  });
  // Apply column sort or default severity sort
  risks = state.colSort
    ? sortRisksByCol(risks, state.colSort, state.colSortDir)
    : sortRisksHighestFirst(risks);
  const summary = riskSummaryFromRaw(rawRisks);
  const owners = riskOwnersForProject(proj.id);
  const extra = ensureProjectExtra(proj.id);
  if (!Array.isArray(extra.risks)) extra.risks = [];

  const workspace = $(".risk-workspace", body);
  const viewOnly = opts.viewOnly && workspace;

  const isEmbed = opts.trackerEmbed || !!(body._riskEmbed);
  if (!viewOnly) {
    if (isEmbed) {
      body._riskEmbed = true;
      state.view = "register";
      body.innerHTML = `
        <div class="risk-workspace risk-workspace--embed">
          <div class="risk-embed-toolbar">
            <button type="button" class="monday-tracker-new" id="risk-new"><i class="bx bx-plus"></i> New risk</button>
            <span class="risk-save-status" id="risk-save-status" data-state="idle"></span>
          </div>
          <div class="risk-view-body"></div>
        </div>
      `;
    } else {
      const searchEl = $("#risk-search", body);
      const hadSearchFocus = searchEl && document.activeElement === searchEl;
      const caret = hadSearchFocus ? searchEl.selectionStart : null;
      body.innerHTML = `
        <div class="risk-workspace">
          <div class="monday-tracker-head risk-head">
            <div class="monday-tracker-tabs" id="risk-view-tabs">
              ${["register", "matrix", "review"].map((v) => `
                <button type="button" data-risk-view="${v}" class="monday-tracker-tab ${state.view === v ? "active" : ""}">
                  ${v === "register" ? "Register" : v === "matrix" ? "Matrix" : "Review board"}
                </button>`).join("")}
            </div>
            <div class="monday-tracker-toolbar">
              <div class="monday-tracker-toolbar-left">
                <button type="button" class="monday-tracker-new" id="risk-new"><i class="bx bx-plus"></i> New risk</button>
                <select class="select-aewttr" id="risk-filter" style="max-width:140px;">
                  ${["Active", "Red / Amber", "Open", "All"].map((filter) => `<option value="${filter}" ${state.filter === filter ? "selected" : ""}>${filter}</option>`).join("")}
                </select>
                <select class="select-aewttr" id="risk-owner-filter" style="max-width:170px;">
                  <option value="All">All owners</option>
                  ${owners.map((owner) => `<option value="${escapeHtml(owner)}" ${state.owner === owner ? "selected" : ""}>${escapeHtml(owner)}</option>`).join("")}
                </select>
                <div class="monday-tracker-search">
                  <i class="bx bx-search"></i>
                  <input type="text" id="risk-search" placeholder="Search risks" value="${escapeHtml(state.search)}" autocomplete="off">
                  <button type="button" class="gantt-search-clear" id="risk-search-clear" style="${state.search ? "" : "display:none;"}">&times;</button>
                </div>
              </div>
              <span class="risk-save-status" id="risk-save-status" data-state="idle"></span>
            </div>
          </div>
          <div class="risk-view-body"></div>
        </div>
      `;
      if (hadSearchFocus) {
        const nextSearch = $("#risk-search", body);
        if (nextSearch) {
          nextSearch.focus();
          if (caret != null) nextSearch.setSelectionRange(caret, caret);
        }
      }
    }
    wireRiskWorkspaceShell(body, proj);
  } else {
    updateRiskSummaryStrip(body, summary);
    $all("[data-risk-view]", body).forEach((btn) => btn.classList.toggle("active", btn.dataset.riskView === state.view));
    const clearBtn = $("#risk-search-clear", body);
    if (clearBtn) clearBtn.style.display = state.search ? "" : "none";
  }

  if (isEmbed) state.view = "register";
  const viewBody = $(".risk-view-body", body);
  if (!viewBody) return;
  viewBody.innerHTML = state.view === "matrix"
    ? riskMatrixHtml(risks)
    : state.view === "register"
      ? riskInlineTableHtml(risks, owners)
      : riskReviewBoardHtml(risks);
  wireRiskViewBody(body, proj, { rawRisks, risks, owners, summary });
}

function wireRiskWorkspaceShell(body, proj) {
  const state = window.AEWTTR.state.projectRisks[proj.id];
  const extra = ensureProjectExtra(proj.id);
  if (!Array.isArray(extra.risks)) extra.risks = [];

  let _riskSaveStatusTimer = null;
  let _riskPendingTimer = null;
  function setRiskSaveStatus(state) {
    const el = $("#risk-save-status", body);
    if (!el) return;
    if (_riskSaveStatusTimer) { clearTimeout(_riskSaveStatusTimer); _riskSaveStatusTimer = null; }
    el.dataset.state = state;
    el.textContent = state === "saving" ? "Saving…" : state === "saved" ? "Saved" : state === "error" ? "Save failed" : state === "pending" ? "Saving soon…" : "";
    if (state === "saved") _riskSaveStatusTimer = setTimeout(() => { el.dataset.state = "idle"; el.textContent = ""; }, 2500);
  }

  function scheduleRiskSave(risk) {
    setRiskSaveStatus("pending");
    if (_riskPendingTimer) clearTimeout(_riskPendingTimer);
    _riskPendingTimer = setTimeout(() => {
      _riskPendingTimer = null;
      persistRisk(risk, false).catch(() => {});
    }, 400);
  }

  async function persistRisk(risk, mirrorProject) {
    if (typeof reanchorProject === "function") reanchorProject(proj);
    if (isSharePointMode() && !risk._spId) {
      try { await ensureRisksList(window.AEWTTR.siteUrl); } catch (e) {}
    }
    setRiskSaveStatus("saving");
    try {
      await Repo.save("risk", risk);
      // Field edits hit PULSE Risks only. Mirror into project RisksJson on
      // create/delete (and when explicitly requested) so we don't rewrite the
      // entire Projects list item (BoardsJson/PeopleJson/…) on every cell commit.
      if (mirrorProject) await Repo.save("project", proj);
      setRiskSaveStatus("saved");
    } catch (e) {
      setRiskSaveStatus("error");
      throw e;
    }
  }

  async function createInlineRisk() {
    const risk = normalizeRiskRecord({ name: "", owner: currentUserName() }, proj.id);
    extra.risks.unshift(risk);
    state.view = "register";
    state.filter = "All";
    // Paint the row first; SharePoint writes continue in the background.
    const hasShell = !!$(".risk-workspace", body);
    drawProjectRisks(body, proj, hasShell ? { viewOnly: true } : {});
    const nameInput = $(`input[data-risk-field="name"][data-risk-id="${risk.id}"]`, body);
    if (nameInput) { nameInput.focus(); nameInput.select(); }
    try {
      await persistRisk(risk, true);
    } catch (e) {
      const idx = extra.risks.findIndex((r) => r.id === risk.id);
      if (idx >= 0) extra.risks.splice(idx, 1);
      toast((e && e.friendly) || "Couldn’t create risk", "error");
      drawProjectRisks(body, proj, { viewOnly: true });
    }
  }

  async function deleteRiskOptimistic(risk, { closeUi } = {}) {
    try {
      const result = await deleteProjectRisk(proj, risk);
      if (!result) return false;
      if (typeof closeUi === "function") closeUi();
      drawProjectRisks(body, proj, { viewOnly: !!$(".risk-workspace", body) });
      toast("Risk deleted", "success");
      return true;
    } catch (e) {
      drawProjectRisks(body, proj, { viewOnly: true });
      toast((e && e.friendly) || "Couldn’t delete risk", "error");
      return false;
    }
  }

  body._riskPersist = persistRisk;
  body._riskScheduleSave = scheduleRiskSave;
  body._riskCreate = createInlineRisk;
  body._riskDelete = deleteRiskOptimistic;
  body._riskRefresh = (flags) => drawProjectRisks(body, proj, flags || {});
  body._riskRefreshView = () => drawProjectRisks(body, proj, { viewOnly: true });

  const newRiskBtn = $("#risk-new", body);
  if (newRiskBtn) newRiskBtn.addEventListener("click", createInlineRisk);
  $all("[data-risk-view]", body).forEach((btn) => btn.addEventListener("click", () => {
    state.view = btn.dataset.riskView;
    drawProjectRisks(body, proj, { viewOnly: true });
  }));
  const filterEl = $("#risk-filter", body);
  if (filterEl) filterEl.addEventListener("change", (e) => {
    state.filter = e.target.value;
    drawProjectRisks(body, proj, { viewOnly: true });
  });
  const ownerFilterEl = $("#risk-owner-filter", body);
  if (ownerFilterEl) ownerFilterEl.addEventListener("change", (e) => {
    state.owner = e.target.value;
    drawProjectRisks(body, proj, { viewOnly: true });
  });

  let searchTimer = null;
  const searchEl2 = $("#risk-search", body);
  if (searchEl2) searchEl2.addEventListener("input", (e) => {
    state.search = e.target.value;
    const clearBtn = $("#risk-search-clear", body);
    if (clearBtn) clearBtn.style.display = state.search ? "" : "none";
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => drawProjectRisks(body, proj, { viewOnly: true }), 180);
  });
  const searchClearEl = $("#risk-search-clear", body);
  if (searchClearEl) searchClearEl.addEventListener("click", () => {
    clearTimeout(searchTimer);
    state.search = "";
    const input = $("#risk-search", body);
    if (input) input.value = "";
    drawProjectRisks(body, proj, { viewOnly: true });
  });

  // One delegated context menu on the workspace — survives view-only table
  // refreshes without rebinding per row.
  const workspace = $(".risk-workspace", body);
  if (workspace && !workspace._riskContextWired) {
    workspace.addEventListener("contextmenu", (e) => {
      const row = e.target.closest("tr.monday-row--risk[data-risk-id], .risk-review-card[data-review-risk], .risk-matrix-item[data-risk-open]");
      if (!row) return;
      e.preventDefault();
      const id = row.dataset.riskId || row.dataset.reviewRisk || row.dataset.riskOpen;
      const risk = extra.risks.find((r) => String(r.id) === String(id));
      if (!risk || typeof showContextMenu !== "function") return;
      const refresh = body._riskRefresh || (() => drawProjectRisks(body, proj));
      const refreshView = body._riskRefreshView || (() => drawProjectRisks(body, proj, { viewOnly: true }));
      const items = [
        { label: "Open details", icon: "bx-note", action: () => openRiskModal(proj, risk.id, () => refresh()) },
        { label: "Mark reviewed", icon: "bx-check-circle", action: async () => {
          const today = new Date().toISOString().slice(0, 10);
          risk.lastReviewedDate = today;
          risk.reviewNotes = risk.reviewNotes || [];
          risk.reviewNotes.unshift({ id: uid("rrn"), author: currentUserName(), date: today, note: "Marked reviewed" });
          scheduleRiskSave(risk);
          const cell = $(`[data-risk-reviewed-cell="${risk.id}"]`, body);
          if (cell) cell.textContent = fmtDate(today);
          else refreshView();
          toast("Marked reviewed", "success");
        }},
        { separator: true },
        ...RISK_STATUS_OPTIONS.filter((status) => status !== risk.status).map((status) => ({
          label: `Status: ${status}`,
          icon: "bx-flag",
          action: async () => {
            risk.status = status;
            scheduleRiskSave(risk);
            refreshView();
          }
        })),
        { separator: true },
        { label: "Delete risk", icon: "bx-trash", danger: true, action: async () => {
          const ok = await confirmDialog({ title: "Delete risk", message: `Delete "${risk.name}"?`, confirmLabel: "Delete", danger: true });
          if (!ok) return;
          await deleteRiskOptimistic(risk);
        }}
      ];
      showContextMenu(e.clientX, e.clientY, items);
    });
    workspace._riskContextWired = true;
  }
}

function wireRiskViewBody(body, proj) {
  const state = window.AEWTTR.state.projectRisks[proj.id];
  const extra = ensureProjectExtra(proj.id);
  const persistRisk = body._riskPersist || (async (risk, mirror) => {
    if (typeof reanchorProject === "function") reanchorProject(proj);
    await Repo.save("risk", risk);
    if (mirror) await Repo.save("project", proj);
  });
  const refresh = body._riskRefresh || (() => drawProjectRisks(body, proj));
  const refreshView = body._riskRefreshView || (() => drawProjectRisks(body, proj, { viewOnly: true }));
  const createInlineRisk = body._riskCreate;

  const inlineAddBtn = $("#risk-inline-add", body);
  if (inlineAddBtn && createInlineRisk) inlineAddBtn.addEventListener("click", createInlineRisk);

  // Tab or Enter on the last risk row's name field creates a new risk
  $all('[data-risk-field="name"]', body).forEach((input) => {
    input.addEventListener("keydown", (e) => {
      if ((e.key === "Tab" && !e.shiftKey) || e.key === "Enter") {
        const rows = $all("tr.monday-row--risk", body);
        const myRow = input.closest("tr");
        if (myRow && rows[rows.length - 1] === myRow) {
          e.preventDefault();
          if (createInlineRisk) createInlineRisk();
        }
      }
    });
  });

  function riskById(id) { return extra.risks.find((r) => String(r.id) === String(id)); }

  function riskStillVisible(risk) {
    return riskPassesFilters(normalizeRiskRecord(risk, proj.id), state);
  }

  async function commitRiskField(riskId, field, value, immediate) {
    const risk = riskById(riskId);
    if (!risk) return { risk: null, needsViewRefresh: false };
    let needsViewRefresh = state.view !== "register";
    if (field === "name") {
      const trimmed = String(value || "").trim();
      if (!trimmed) {
        const input = $(`[data-risk-field="name"][data-risk-id="${riskId}"]`, body);
        if (input) input.value = risk.name || "";
        return { risk: null, needsViewRefresh: false };
      }
      risk.name = trimmed;
      risk.title = risk.name;
    } else if (field === "owner") {
      risk.owner = String(value || "").trim();
      risk.ownerName = risk.owner;
    } else if (field === "likelihood" || field === "impact") {
      risk[field] = Number(value) || 1;
      risk.rating = riskRating(risk.likelihood, risk.impact);
      risk.level = risk.rating;
      risk.ratingHistory = risk.ratingHistory || [];
      risk.ratingHistory.push({ date: new Date().toISOString().slice(0, 10), likelihood: risk.likelihood, impact: risk.impact, rating: risk.rating, status: risk.status });
      const cell = $(`[data-risk-rating-cell="${riskId}"]`, body);
      if (cell) cell.innerHTML = riskRatingPill(risk.rating);
    } else if (field === "status") {
      risk.status = value;
      risk.ratingHistory = risk.ratingHistory || [];
      risk.ratingHistory.push({ date: new Date().toISOString().slice(0, 10), likelihood: risk.likelihood, impact: risk.impact, rating: risk.rating, status: risk.status });
      const select = $(`select[data-risk-field="status"][data-risk-id="${riskId}"]`, body);
      if (select) {
        select.className = `monday-status-select risk-status-${String(value || "Open").replace(/\s+/g, "-")}`;
      }
    } else {
      risk[field] = value;
    }
    // Don't block the cell UI on SharePoint debounce + network. The scheduler
    // belongs to the persistent risk workspace shell, so obtain it from the
    // shell rather than reaching for an out-of-scope local function.
    const scheduleSave = body._riskScheduleSave;
    if (immediate && typeof body._riskPersist === "function") {
      body._riskPersist(risk, false).catch(() => {});
    } else if (typeof scheduleSave === "function") {
      scheduleSave(risk);
    } else if (typeof body._riskPersist === "function") {
      body._riskPersist(risk, false).catch(() => {});
    }
    updateRiskSummaryStrip(body, riskSummaryFromRaw(projectRisks(proj.id)));
    if (!riskStillVisible(risk)) needsViewRefresh = true;
    return { risk, needsViewRefresh };
  }

  $all('[data-risk-field="name"], [data-risk-field="owner"]', body).forEach((input) => {
    const commit = (immediate) => commitRiskField(input.dataset.riskId, input.dataset.riskField, input.value, immediate).then((result) => {
      if (result.needsViewRefresh) refreshView();
    });
    input.addEventListener("blur", () => commit(input.dataset.riskField === "name"));
  });
  $all('[data-risk-field="due"]', body).forEach((input) => {
    input.addEventListener("change", () => commitRiskField(input.dataset.riskId, "due", input.value));
  });
  $all("select[data-risk-field]", body).forEach((select) => {
    select.addEventListener("change", async () => {
      const result = await commitRiskField(select.dataset.riskId, select.dataset.riskField, select.value);
      if (result.needsViewRefresh) refreshView();
    });
  });

  $all("[data-risk-details]", body).forEach((btn) => btn.addEventListener("click", () => openRiskModal(proj, btn.dataset.riskDetails, () => refresh())));
  $all("[data-review-risk]", body).forEach((btn) => btn.addEventListener("click", () => openRiskModal(proj, btn.dataset.reviewRisk, () => refresh())));
  $all("[data-risk-open]", body).forEach((btn) => btn.addEventListener("click", () => openRiskModal(proj, btn.dataset.riskOpen, () => refresh())));

  wireRiskColumnFilters(body, proj, refreshView);
  wireRiskExpand(body, proj);
}

function wireRiskExpand(body, proj) {
  const state = window.AEWTTR.state.projectRisks[proj.id];
  if (!state.expanded) state.expanded = {};
  const extra = ensureProjectExtra(proj.id);

  function riskById(id) { return extra.risks.find((r) => String(r.id) === String(id)); }

  function renderDetailContent(detailTd, risk) {
    const mits = risk.mitigations || [];
    const mitListHtml = mits.length ? mits.map((m, idx) => {
      const isDone = m.status === "Complete";
      const isCustom = !RISK_MITIGATION_TYPES.slice(0, -1).includes(m.actionType || "");
      const selectVal = (m.actionType && !isCustom) ? m.actionType : "Other / custom";
      return `<div class="risk-detail-mit-item${isDone ? " risk-detail-mit-item--done" : ""}" data-midx="${idx}">
        <button type="button" class="risk-detail-mit-check${isDone ? " done" : ""}" data-midx="${idx}" title="${isDone ? "Mark open" : "Mark complete"}">${isDone ? "✓" : "○"}</button>
        <div class="risk-detail-mit-body">
          <select class="select-aewttr risk-detail-mit-type" data-midx="${idx}">
            ${RISK_MITIGATION_TYPES.map((t) => `<option value="${escapeHtml(t)}"${selectVal === t ? " selected" : ""}>${escapeHtml(t)}</option>`).join("")}
          </select>
          <input class="input-aewttr risk-detail-mit-desc" data-midx="${idx}" value="${escapeHtml(m.description || "")}" placeholder="Detail…">
          <input class="input-aewttr risk-detail-mit-owner" data-midx="${idx}" value="${escapeHtml(m.owner || "")}" placeholder="Owner">
          <input class="input-aewttr risk-detail-mit-due" type="date" data-midx="${idx}" value="${escapeHtml(m.dueDate || "")}">
        </div>
        <button type="button" class="risk-detail-mit-del" data-midx="${idx}" title="Remove">✕</button>
      </div>`;
    }).join("") : `<div style="font-size:12px;color:var(--aewttr-muted);padding:4px 0;">No mitigation steps yet.</div>`;

    detailTd.innerHTML = `<div class="risk-detail-panel">
      <div class="risk-detail-left">
        <div class="risk-detail-section-lbl">Description</div>
        <textarea class="textarea-aewttr risk-detail-desc" rows="3" placeholder="Risk description…">${escapeHtml(risk.description || "")}</textarea>
        <div class="risk-detail-section-lbl" style="margin-top:4px;">Category</div>
        <select class="select-aewttr risk-detail-category">
          ${RISK_CATEGORY_OPTIONS.map((opt) => `<option${risk.category === opt ? " selected" : ""}>${escapeHtml(opt)}</option>`).join("")}
        </select>
        <div class="risk-detail-section-lbl" style="margin-top:4px;">Response Strategy</div>
        <select class="select-aewttr risk-detail-response">
          ${RISK_RESPONSE_OPTIONS.map((opt) => `<option${risk.responseStrategy === opt ? " selected" : ""}>${escapeHtml(opt)}</option>`).join("")}
        </select>
      </div>
      <div class="risk-detail-right">
        <div class="risk-detail-section-lbl">Mitigation Steps</div>
        <div class="risk-detail-mit-list">${mitListHtml}</div>
        <button type="button" class="risk-detail-add-mit">+ Add step</button>
      </div>
      <div class="risk-detail-foot">
        <a href="#" class="risk-detail-more-link" data-risk-more="${escapeHtml(risk.id)}">More › (review notes &amp; residual risk)</a>
      </div>
    </div>`;

    // Wire mitigation toggles
    $all(".risk-detail-mit-check", detailTd).forEach((btn) => {
      btn.addEventListener("click", () => {
        const idx = Number(btn.dataset.midx);
        if (!risk.mitigations) risk.mitigations = [];
        risk.mitigations[idx].status = risk.mitigations[idx].status === "Complete" ? "Open" : "Complete";
        if (typeof body._riskScheduleSave === "function") body._riskScheduleSave(risk);
        renderDetailContent(detailTd, risk);
      });
    });

    $all(".risk-detail-mit-type", detailTd).forEach((sel) => {
      sel.addEventListener("change", () => {
        const idx = Number(sel.dataset.midx);
        if (!risk.mitigations) risk.mitigations = [];
        risk.mitigations[idx].actionType = sel.value;
        if (typeof body._riskScheduleSave === "function") body._riskScheduleSave(risk);
      });
    });

    $all(".risk-detail-mit-desc", detailTd).forEach((inp) => {
      inp.addEventListener("change", () => {
        const idx = Number(inp.dataset.midx);
        if (!risk.mitigations) risk.mitigations = [];
        risk.mitigations[idx].description = inp.value.trim();
        if (typeof body._riskScheduleSave === "function") body._riskScheduleSave(risk);
      });
    });

    $all(".risk-detail-mit-owner", detailTd).forEach((inp) => {
      inp.addEventListener("change", () => {
        const idx = Number(inp.dataset.midx);
        if (!risk.mitigations) risk.mitigations = [];
        risk.mitigations[idx].owner = inp.value.trim();
        if (typeof body._riskScheduleSave === "function") body._riskScheduleSave(risk);
      });
    });

    $all(".risk-detail-mit-due", detailTd).forEach((inp) => {
      inp.addEventListener("change", () => {
        const idx = Number(inp.dataset.midx);
        if (!risk.mitigations) risk.mitigations = [];
        risk.mitigations[idx].dueDate = inp.value;
        if (typeof body._riskScheduleSave === "function") body._riskScheduleSave(risk);
      });
    });

    $all(".risk-detail-mit-del", detailTd).forEach((btn) => {
      btn.addEventListener("click", () => {
        const idx = Number(btn.dataset.midx);
        if (!risk.mitigations) risk.mitigations = [];
        risk.mitigations.splice(idx, 1);
        if (typeof body._riskScheduleSave === "function") body._riskScheduleSave(risk);
        renderDetailContent(detailTd, risk);
      });
    });

    const addBtn = $(".risk-detail-add-mit", detailTd);
    if (addBtn) {
      addBtn.addEventListener("click", () => {
        if (!risk.mitigations) risk.mitigations = [];
        risk.mitigations.push({ id: uid("mit"), actionType: "", description: "", owner: "", dueDate: "", status: "Open", notes: "" });
        if (typeof body._riskScheduleSave === "function") body._riskScheduleSave(risk);
        renderDetailContent(detailTd, risk);
      });
    }

    const descTa = $(".risk-detail-desc", detailTd);
    if (descTa) {
      let descTimer = null;
      descTa.addEventListener("input", () => {
        risk.description = descTa.value;
        clearTimeout(descTimer);
        descTimer = setTimeout(() => { if (typeof body._riskScheduleSave === "function") body._riskScheduleSave(risk); }, 600);
      });
    }

    const catSel = $(".risk-detail-category", detailTd);
    if (catSel) {
      catSel.addEventListener("change", () => {
        risk.category = catSel.value;
        if (typeof body._riskScheduleSave === "function") body._riskScheduleSave(risk);
      });
    }

    const respSel = $(".risk-detail-response", detailTd);
    if (respSel) {
      respSel.addEventListener("change", () => {
        risk.responseStrategy = respSel.value;
        if (typeof body._riskScheduleSave === "function") body._riskScheduleSave(risk);
      });
    }

    const moreLink = $("[data-risk-more]", detailTd);
    if (moreLink) {
      moreLink.addEventListener("click", (e) => {
        e.preventDefault();
        const refresh = body._riskRefresh || (() => drawProjectRisks(body, proj));
        openRiskModal(proj, moreLink.dataset.riskMore, () => refresh());
      });
    }
  }

  $all("[data-risk-expand]", body).forEach((btn) => {
    btn.addEventListener("click", () => {
      const riskId = btn.dataset.riskExpand;
      const detailRow = $(`[data-detail-for="${riskId}"]`, body);
      if (!detailRow) return;
      const isExpanded = !!state.expanded[riskId];
      state.expanded[riskId] = !isExpanded;
      btn.classList.toggle("expanded", !isExpanded);
      detailRow.classList.toggle("risk-detail-row--hidden", isExpanded);
      if (!isExpanded) {
        const detailTd = detailRow.querySelector("td");
        if (detailTd && !detailTd._rendered) {
          const risk = riskById(riskId);
          if (risk) renderDetailContent(detailTd, risk);
          detailTd._rendered = true;
        }
      }
    });
  });
}

function riskColValue(risk, col) {
  switch (col) {
    case "name": return risk.name || "";
    case "rating": return risk.rating || riskRating(risk.likelihood, risk.impact) || "";
    case "status": return risk.status || "Open";
    case "owner": return risk.owner || "";
    case "due": return risk.due ? fmtDate(risk.due) : "—";
    case "lastReviewedDate": return risk.lastReviewedDate ? fmtDate(risk.lastReviewedDate) : "—";
    case "likelihood": return String(risk.likelihood || 1);
    case "impact": return String(risk.impact || 1);
    default: return "";
  }
}

function sortRisksByCol(risks, col, dir) {
  return [...risks].sort((a, b) => {
    const av = riskColValue(a, col);
    const bv = riskColValue(b, col);
    const cmp = av.localeCompare(bv, undefined, { numeric: true, sensitivity: "base" });
    return dir === "desc" ? -cmp : cmp;
  });
}

function wireRiskColumnFilters(body, proj, redrawFn) {
  const state = window.AEWTTR.state.projectRisks && window.AEWTTR.state.projectRisks[proj.id];
  if (!state) return;
  const extra = ensureProjectExtra(proj.id);
  const rawRisks = projectRisks(proj.id);

  function closeDropdown() {
    const existing = document.getElementById("risk-col-filter-dropdown");
    if (existing) existing.remove();
  }

  $all("[data-col-filter]", body).forEach((btn) => {
    const col = btn.dataset.colFilter;
    const isActive = !!(state.colSort === col || (state.colFilters[col] && state.colFilters[col].size));
    if (isActive) btn.classList.add("is-active");
    if (state.colSort === col) btn.classList.add("is-sorted");

    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      closeDropdown();

      // Collect unique values for this column from all raw risks
      const allVals = [...new Set(rawRisks.map((r) => riskColValue(r, col)).filter(Boolean))].sort((a, b) =>
        a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));
      const activeFilter = state.colFilters[col];

      const dropdown = document.createElement("div");
      dropdown.id = "risk-col-filter-dropdown";
      dropdown.className = "risk-col-filter-dropdown";

      let valSearch = "";
      function renderDropdown() {
        const filtered = valSearch ? allVals.filter((v) => v.toLowerCase().includes(valSearch.toLowerCase())) : allVals;
        dropdown.innerHTML = `
          <button class="rcfd-sort-btn" data-rcfd-sort="asc"><i class="bx bx-sort-a-z"></i> Sort A → Z</button>
          <button class="rcfd-sort-btn" data-rcfd-sort="desc"><i class="bx bx-sort-z-a"></i> Sort Z → A</button>
          <div class="rcfd-sep"></div>
          <input type="text" class="rcfd-search-input" placeholder="Search values…" value="${escapeHtml(valSearch)}">
          <div class="rcfd-values">
            ${filtered.length ? filtered.map((v) => `
              <label class="rcfd-val-label">
                <input type="checkbox" value="${escapeHtml(v)}" ${!activeFilter || !activeFilter.size || activeFilter.has(v) ? "checked" : ""}>
                <span>${escapeHtml(v)}</span>
              </label>`).join("") : `<div style="padding:8px 12px;font-size:12px;color:var(--aewttr-muted);">No values</div>`}
          </div>
          <div class="rcfd-actions">
            <button class="rcfd-clear-btn" id="rcfd-clear">Clear filter</button>
            <button class="btn-aewttr btn-aewttr-sm" id="rcfd-apply">Apply</button>
          </div>
        `;
        const searchInput = dropdown.querySelector(".rcfd-search-input");
        if (searchInput) {
          searchInput.focus();
          searchInput.addEventListener("input", () => { valSearch = searchInput.value; renderDropdown(); });
        }
        $all("[data-rcfd-sort]", dropdown).forEach((sortBtn) => sortBtn.addEventListener("click", () => {
          state.colSort = col;
          state.colSortDir = sortBtn.dataset.rcfdSort;
          closeDropdown();
          if (typeof redrawFn === "function") redrawFn();
        }));
        const clearBtn = dropdown.querySelector("#rcfd-clear");
        if (clearBtn) clearBtn.addEventListener("click", () => {
          delete state.colFilters[col];
          if (state.colSort === col) { state.colSort = null; state.colSortDir = "asc"; }
          closeDropdown();
          if (typeof redrawFn === "function") redrawFn();
        });
        const applyBtn = dropdown.querySelector("#rcfd-apply");
        if (applyBtn) applyBtn.addEventListener("click", () => {
          const checked = new Set($all("input[type=checkbox]:checked", dropdown).map((cb) => cb.value));
          const allChecked = allVals.every((v) => checked.has(v));
          if (allChecked) delete state.colFilters[col];
          else state.colFilters[col] = checked;
          closeDropdown();
          if (typeof redrawFn === "function") redrawFn();
        });
      }
      renderDropdown();

      const rect = btn.getBoundingClientRect();
      dropdown.style.top = `${rect.bottom + 4}px`;
      dropdown.style.left = `${Math.min(rect.left, window.innerWidth - 270)}px`;
      document.body.appendChild(dropdown);

      setTimeout(() => {
        document.addEventListener("click", closeDropdown, { once: true });
        document.addEventListener("keydown", (ev) => { if (ev.key === "Escape") closeDropdown(); }, { once: true });
      }, 0);
    });
  });
}

/* Risk Details — long-form fields (description, category, mitigation,
   response, portfolio, review history). Title/owner/severity/status/due
   stay inline in the register. */
function openRiskModal(proj, riskId, onDone) {
  const extra = ensureProjectExtra(proj.id);
  if (!Array.isArray(extra.risks)) extra.risks = [];
  const existing = riskId ? extra.risks.find((risk) => String(risk.id) === String(riskId)) : null;
  const risk = existing ? normalizeRiskRecord(existing, proj.id) : normalizeRiskRecord({ projectId: proj.id }, proj.id);
  const portfolios = typeof projectPortfolios === "function" ? projectPortfolios(proj) : (proj.portfolios || []);
  const rating = risk.rating || riskRating(risk.likelihood, risk.impact);
  const modal = openModal(`
    <div class="aewttr-modal-head">
      <div>
        <div class="risk-details-subline">${riskRatingPill(rating)} <span>L${Number(risk.likelihood || 0)} / I${Number(risk.impact || 0)}</span> <span>${escapeHtml(risk.owner || "Unassigned")}</span></div>
      </div>
      <button class="aewttr-modal-close">&times;</button>
    </div>
    <div class="aewttr-modal-body">
      <div class="form-row"><label>Risk title <span class="required-star">*</span></label><input class="input-aewttr" id="risk-name" value="${escapeHtml(risk.name)}" placeholder="Describe this risk in a few words…"></div>
      <div class="form-row"><label>Description</label><textarea class="textarea-aewttr" id="risk-description" placeholder="More detail on the risk and its context…">${escapeHtml(risk.description)}</textarea></div>
      <div class="form-grid-2">
        <div class="form-row"><label>Category</label><select class="select-aewttr" id="risk-category">${RISK_CATEGORY_OPTIONS.map((option) => `<option ${risk.category === option ? "selected" : ""}>${escapeHtml(option)}</option>`).join("")}</select></div>
        <div class="form-row"><label>Response strategy</label><select class="select-aewttr" id="risk-response">${RISK_RESPONSE_OPTIONS.map((option) => `<option ${risk.responseStrategy === option ? "selected" : ""}>${escapeHtml(option)}</option>`).join("")}</select></div>
      </div>
      <div class="form-row"><label>Portfolio / project link</label><input class="input-aewttr" id="risk-portfolio" list="risk-portfolio-options" value="${escapeHtml(risk.portfolio || portfolios[0] || "")}" placeholder="${escapeHtml(proj.id + " — " + (proj.name || ""))}"><datalist id="risk-portfolio-options">${portfolios.map((portfolio) => `<option value="${escapeHtml(portfolio)}"></option>`).join("")}</datalist></div>
      <div class="risk-mitigations-section">
        <div class="risk-mitigations-head">
          <span>Mitigation actions <span class="risk-mit-progress" id="risk-mit-progress" style="display:none;"></span></span>
          <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="risk-add-mitigation"><i class="bx bx-plus"></i> Add action</button>
        </div>
        <div id="risk-mitigations-list" class="risk-mitigations-list"></div>
      </div>

      <div class="risk-residual-section">
        <div class="risk-residual-head">
          <span>Residual risk</span>
          <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="risk-residual-suggest"${tip("Calculate residual from mitigation completion progress")}>Suggest from mitigations</button>
        </div>
        <div class="risk-residual-body">
          <div class="risk-residual-controls">
            <div class="risk-residual-field">
              <label>Residual likelihood</label>
              <select class="select-aewttr" id="risk-residual-l">
                <option value="0"${!risk.residualLikelihood ? " selected" : ""}>Not set</option>
                ${[1,2,3,4,5].map((n) => `<option value="${n}"${Number(risk.residualLikelihood) === n ? " selected" : ""}>${n}</option>`).join("")}
              </select>
            </div>
            <div class="risk-residual-field">
              <label>Residual impact</label>
              <select class="select-aewttr" id="risk-residual-i">
                <option value="0"${!risk.residualImpact ? " selected" : ""}>Not set</option>
                ${[1,2,3,4,5].map((n) => `<option value="${n}"${Number(risk.residualImpact) === n ? " selected" : ""}>${n}</option>`).join("")}
              </select>
            </div>
            <div class="risk-residual-rating-wrap">
              <div class="risk-residual-rating-display" id="risk-residual-rating-display" style="display:none;"></div>
              <div class="risk-residual-hint">Residual = inherent risk remaining after mitigations are applied.</div>
            </div>
          </div>
        </div>
      </div>

      <div class="risk-review-history-head">Review history</div>
      <div class="risk-review-history-list">
        ${(risk.reviewNotes || []).length ? risk.reviewNotes.map((entry) => `
          <div class="risk-review-history-item">
            <div class="risk-review-history-meta">${fmtDate(entry.date)} · ${escapeHtml(entry.author || "Unknown")}</div>
            <div class="risk-review-history-text">${escapeHtml(entry.note)}</div>
          </div>`).join("") : `<div class="empty-state" style="padding:8px 0;font-size:12px;">No review notes yet.</div>`}
      </div>
      <div class="form-row" style="margin-bottom:0;"><label>Add review note <small style="font-weight:400;color:var(--aewttr-muted);">(optional)</small></label><textarea class="textarea-aewttr" id="risk-review-note" placeholder="Add a dated review note..." rows="2"></textarea></div>
    </div>
    <div class="aewttr-modal-foot">
      <button class="btn-aewttr-ghost" id="risk-delete">Delete</button>
      <button class="btn-aewttr-ghost" id="risk-cancel">Cancel</button>
      <button class="btn-aewttr" id="risk-save">Save</button>
    </div>
  `);
  let mitigations = (risk.mitigations || []).map((m) => Object.assign({}, m));
  const notesOpen = new Set();

  function updateResidualDisplay() {
    const lSel = $("#risk-residual-l", modal);
    const iSel = $("#risk-residual-i", modal);
    const display = $("#risk-residual-rating-display", modal);
    const completedCount = mitigations.filter((m) => m.status === "Complete").length;
    const totalCount = mitigations.length;
    const pct = $("#risk-mit-progress", modal);
    if (pct) {
      if (totalCount) {
        pct.textContent = `${completedCount}/${totalCount} complete`;
        pct.style.display = "";
      } else {
        pct.style.display = "none";
      }
    }
    if (!lSel || !iSel || !display) return;
    const rl = Number(lSel.value);
    const ri = Number(iSel.value);
    if (rl && ri) {
      display.innerHTML = riskRatingPill(riskRating(rl, ri));
      display.style.display = "";
    } else {
      display.style.display = "none";
    }
  }

  function renderMitigations() {
    const list = $("#risk-mitigations-list", modal);
    if (!list) return;
    if (!mitigations.length) {
      list.innerHTML = `<div class="risk-mitigations-empty">No actions yet — click "Add action" to track mitigation steps.</div>`;
      updateResidualDisplay();
      return;
    }
    list.innerHTML = mitigations.map((m, idx) => {
      const isCustom = !RISK_MITIGATION_TYPES.slice(0, -1).includes(m.actionType || "");
      const selectVal = (m.actionType && !isCustom) ? m.actionType : "Other / custom";
      return `
      <div class="risk-mitigation-item ${m.status === "Complete" ? "risk-mitigation-item--done" : ""}" data-midx="${idx}">
        <div class="risk-mitigation-main">
          <button type="button" class="risk-mitigation-toggle" data-midx="${idx}" title="${m.status === "Complete" ? "Mark open" : "Mark complete"}" aria-label="${m.status === "Complete" ? "Mark open" : "Mark complete"}">
            <i class="bx ${m.status === "Complete" ? "bxs-check-circle" : "bx-circle"}"></i>
          </button>
          <div class="risk-mitigation-fields">
            <select class="select-aewttr risk-mitigation-type" data-midx="${idx}" data-field="actionType">
              ${RISK_MITIGATION_TYPES.map((t) => `<option value="${escapeHtml(t)}"${selectVal === t ? " selected" : ""}>${escapeHtml(t)}</option>`).join("")}
            </select>
            <input class="input-aewttr risk-mitigation-desc" data-midx="${idx}" data-field="description" value="${escapeHtml(m.description || "")}" placeholder="Additional detail…"${selectVal !== "Other / custom" && !m.description ? ' style="display:none;"' : ""}>
            <div class="risk-mitigation-meta">
              <input class="input-aewttr risk-mitigation-owner" data-midx="${idx}" data-field="owner" value="${escapeHtml(m.owner || "")}" placeholder="Owner">
              <input class="input-aewttr risk-mitigation-due" type="date" data-midx="${idx}" data-field="dueDate" value="${escapeHtml(m.dueDate || "")}">
            </div>
          </div>
          <button type="button" class="risk-mitigation-notes-btn ${notesOpen.has(idx) ? "active" : ""}" data-midx="${idx}" title="${notesOpen.has(idx) ? "Hide notes" : "Add notes"}" aria-label="Toggle notes">
            <i class="bx bx-note"></i>${m.notes ? `<span class="risk-mit-notes-dot"></span>` : ""}
          </button>
          <button type="button" class="risk-mitigation-delete" data-midx="${idx}" title="Remove" aria-label="Remove mitigation"><i class="bx bx-trash"></i></button>
        </div>
        ${notesOpen.has(idx) ? `
        <div class="risk-mitigation-notes-area">
          <textarea class="textarea-aewttr risk-mitigation-notes-ta" data-midx="${idx}" data-field="notes" rows="3" placeholder="Additional context, blockers, or progress notes…">${escapeHtml(m.notes || "")}</textarea>
        </div>` : ""}
      </div>`;
    }).join("");

    $all(".risk-mitigation-toggle", list).forEach((btn) => btn.addEventListener("click", () => {
      const idx = Number(btn.dataset.midx);
      mitigations[idx].status = mitigations[idx].status === "Complete" ? "Open" : "Complete";
      renderMitigations();
    }));
    $all(".risk-mitigation-type", list).forEach((sel) => sel.addEventListener("change", () => {
      const idx = Number(sel.dataset.midx);
      mitigations[idx].actionType = sel.value;
      const descInput = list.querySelector(`.risk-mitigation-desc[data-midx="${idx}"]`);
      if (descInput) descInput.style.display = sel.value === "Other / custom" ? "" : "";
    }));
    $all(".risk-mitigation-desc, .risk-mitigation-owner, .risk-mitigation-due", list).forEach((input) => input.addEventListener("change", () => {
      const idx = Number(input.dataset.midx);
      mitigations[idx][input.dataset.field] = input.value.trim();
    }));
    $all(".risk-mitigation-notes-ta", list).forEach((ta) => {
      ta.addEventListener("input", () => {
        const idx = Number(ta.dataset.midx);
        mitigations[idx].notes = ta.value;
      });
    });
    $all(".risk-mitigation-notes-btn", list).forEach((btn) => btn.addEventListener("click", () => {
      const idx = Number(btn.dataset.midx);
      if (notesOpen.has(idx)) notesOpen.delete(idx);
      else notesOpen.add(idx);
      renderMitigations();
      const ta = list.querySelector(`.risk-mitigation-notes-ta[data-midx="${idx}"]`);
      if (ta) ta.focus();
    }));
    $all(".risk-mitigation-delete", list).forEach((btn) => btn.addEventListener("click", () => {
      const idx = Number(btn.dataset.midx);
      mitigations.splice(idx, 1);
      notesOpen.delete(idx);
      renderMitigations();
    }));
    updateResidualDisplay();
  }

  renderMitigations();

  $("#risk-add-mitigation", modal).addEventListener("click", () => {
    mitigations.push({ id: uid("mit"), actionType: "", description: "", owner: "", dueDate: "", status: "Open", notes: "" });
    renderMitigations();
    const selects = $all(".risk-mitigation-type", modal);
    if (selects.length) selects[selects.length - 1].focus();
  });

  const residualLSel = $("#risk-residual-l", modal);
  const residualISel = $("#risk-residual-i", modal);
  if (residualLSel) residualLSel.addEventListener("change", updateResidualDisplay);
  if (residualISel) residualISel.addEventListener("change", updateResidualDisplay);
  updateResidualDisplay();

  const suggestBtn = $("#risk-residual-suggest", modal);
  if (suggestBtn) {
    suggestBtn.addEventListener("click", () => {
      const total = mitigations.length;
      if (!total) { toast("Add mitigation actions first.", "error"); return; }
      const done = mitigations.filter((m) => m.status === "Complete").length;
      const pct = done / total;
      const sugL = Math.max(1, Math.round(risk.likelihood * (1 - pct * 0.5)));
      const sugI = Math.max(1, Math.round(risk.impact * (1 - pct * 0.4)));
      if (residualLSel) residualLSel.value = String(sugL);
      if (residualISel) residualISel.value = String(sugI);
      updateResidualDisplay();
      toast(`Suggested residual: L${sugL} / I${sugI} (${Math.round(pct * 100)}% of mitigations complete)`, "success");
    });
  }

  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#risk-cancel", modal).addEventListener("click", closeModal);
  $("#risk-save", modal).addEventListener("click", async () => {
    $all(".risk-mitigation-type", modal).forEach((sel) => {
      const idx = Number(sel.dataset.midx);
      if (mitigations[idx]) mitigations[idx].actionType = sel.value;
    });
    $all(".risk-mitigation-desc, .risk-mitigation-owner, .risk-mitigation-due", modal).forEach((input) => {
      const idx = Number(input.dataset.midx);
      if (mitigations[idx]) mitigations[idx][input.dataset.field] = input.value.trim();
    });
    $all(".risk-mitigation-notes-ta", modal).forEach((ta) => {
      const idx = Number(ta.dataset.midx);
      if (mitigations[idx]) mitigations[idx].notes = ta.value;
    });
    risk.mitigations = mitigations.filter((m) => m.actionType || m.description);
    risk.residualLikelihood = Number((residualLSel && residualLSel.value) || 0) || 0;
    risk.residualImpact = Number((residualISel && residualISel.value) || 0) || 0;
    const nameVal = $("#risk-name", modal).value.trim();
    if (!nameVal) { toast("Risk title is required", "error"); return; }
    risk.name = nameVal;
    risk.title = nameVal;
    risk.description = $("#risk-description", modal).value.trim();
    risk.text = risk.description || risk.name;
    risk.category = $("#risk-category", modal).value;
    risk.responseStrategy = $("#risk-response", modal).value;
    risk.portfolio = $("#risk-portfolio", modal).value.trim();
    const reviewNote = $("#risk-review-note", modal).value.trim();
    if (reviewNote) {
      risk.reviewNotes = risk.reviewNotes || [];
      risk.reviewNotes.unshift({ id: uid("rrn"), author: currentUserName(), date: new Date().toISOString().slice(0, 10), note: reviewNote });
      risk.lastReviewedDate = new Date().toISOString().slice(0, 10);
    }
    const isNew = !existing;
    if (existing) Object.assign(existing, risk);
    else extra.risks.unshift(risk);
    closeModal();
    toast("Risk updated", "success");
    if (onDone) onDone();
    try {
      if (typeof reanchorProject === "function") reanchorProject(proj);
      if (isSharePointMode() && isNew) {
        try { await ensureRisksList(window.AEWTTR.siteUrl); } catch (e) {}
      }
      await Repo.save("risk", existing || risk);
      await Repo.save("project", proj);
    } catch (e) {
      toast((e && e.friendly) || "Couldn’t save risk details", "error");
      if (onDone) onDone();
    }
  });
  $("#risk-delete", modal).addEventListener("click", async () => {
    const ok = await confirmDialog({ title: "Delete risk", message: `Delete "${risk.name}"?`, confirmLabel: "Delete", danger: true });
    if (!ok) return;
    try {
      const result = await deleteProjectRisk(proj, risk);
      if (!result) throw new Error("The risk could not be found in this project.");
      closeModal();
      toast("Risk deleted", "success");
      if (onDone) onDone();
    } catch (e) {
      if (onDone) onDone();
      toast((e && e.friendly) || "Couldn’t delete risk", "error");
    }
  });
}

/* The Notes tab is a folder-style hub: sidebar tree = Project → Tasks →
   nested subitems (unlimited depth); Overview mirrors that hierarchy.
   Thread views reuse the same chat composer and author-only edit rules. */
function drawProjectNotes(body, proj) {
  const db = window.AEWTTR.db;
  const extra = ensureProjectExtra(proj.id);
  if (!Array.isArray(extra.notes)) extra.notes = [];
  const stateKey = `projNotesFolder:${proj.id}`;
  const collapseKey = `projNotesTreeCollapse:${proj.id}`;
  if (!window.AEWTTR.state.projNotesFolder) window.AEWTTR.state.projNotesFolder = {};
  if (!window.AEWTTR.state.projNotesTreeCollapse) window.AEWTTR.state.projNotesTreeCollapse = {};
  let activeFolder = window.AEWTTR.state.projNotesFolder[stateKey] || "overview";
  let collapsedMap = window.AEWTTR.state.projNotesTreeCollapse[collapseKey] || {};
  let editingId = null;
  let folderFilter = "";

  function projectTasks() {
    return (db.ganttTasks && db.ganttTasks[proj.id]) || [];
  }

  function folderKeyForTask(taskId) { return `task:${taskId}`; }
  function folderKeyForSub(taskId, path) { return `sub:${taskId}:${path}`; }

  function noteSortKey(note) {
    return `${note.date || ""}T${note.time || "00:00"}`;
  }

  function sortNotesDesc(list) {
    return (list || []).slice().sort((a, b) => String(noteSortKey(b)).localeCompare(String(noteSortKey(a))));
  }

  function subLabelTrail(task, pathStr) {
    const parts = String(pathStr || "").split(".").filter((p) => p !== "");
    const labels = [];
    let list = task.subtasks || [];
    for (let i = 0; i < parts.length; i++) {
      const node = list[Number(parts[i])];
      if (!node) break;
      labels.push(node.text || "Untitled subitem");
      list = node.subtasks || [];
    }
    return labels;
  }

  function assignedLabelForSub(task, pathStr) {
    const trail = subLabelTrail(task, pathStr);
    const depth = trail.length;
    const leaf = trail[trail.length - 1] || "Untitled subitem";
    if (depth <= 1) return `Subitem · ${leaf}`;
    return `Subitem · ${trail.join(" › ")}`;
  }

  function resolveFolder(folderId) {
    if (folderId === "overview") return { id: "overview", kind: "overview", label: "Overview", notes: null, parentTask: null };
    if (folderId === "project") return { id: "project", kind: "project", label: "Project Notes", notes: extra.notes, parentTask: null, assignedTo: "Project" };
    if (folderId === "risks") return { id: "risks", kind: "risks", label: "Risks", notes: null, parentTask: null };
    if (String(folderId || "").startsWith("task:")) {
      const taskId = folderId.slice(5);
      const task = projectTasks().find((t) => t.id === taskId);
      if (!task) return null;
      if (!Array.isArray(task.notes)) task.notes = [];
      return { id: folderId, kind: "task", label: task.title || "Untitled task", notes: task.notes, parentTask: task, assignedTo: `Task · ${task.title || "Untitled"}` };
    }
    if (String(folderId || "").startsWith("sub:")) {
      const rest = folderId.slice(4);
      const colonIdx = rest.indexOf(":");
      if (colonIdx < 0) return null;
      const taskId = rest.slice(0, colonIdx);
      const path = rest.slice(colonIdx + 1);
      const task = projectTasks().find((t) => t.id === taskId);
      const sub = typeof getSubtaskAtPath === "function"
        ? getSubtaskAtPath(task, path)
        : (task && task.subtasks ? task.subtasks[Number(path)] : null);
      if (!task || !sub) return null;
      if (!Array.isArray(sub.notes)) sub.notes = [];
      return {
        id: folderId,
        kind: "sub",
        label: sub.text || "Untitled subitem",
        notes: sub.notes,
        parentTask: task,
        subPath: path,
        subIndex: Number(String(path).split(".")[0]),
        assignedTo: assignedLabelForSub(task, path),
        breadcrumb: [`Task · ${task.title || "Untitled"}`].concat(subLabelTrail(task, path))
      };
    }
    return null;
  }

  function collectOverviewNotes() {
    const rows = [];
    (extra.notes || []).forEach((note) => {
      rows.push({ note, assignedTo: "Project", folderId: "project", sortKey: noteSortKey(note) });
    });
    projectTasks().forEach((task) => {
      (task.notes || []).forEach((note) => {
        rows.push({
          note,
          assignedTo: `Task · ${task.title || "Untitled"}`,
          folderId: folderKeyForTask(task.id),
          sortKey: noteSortKey(note)
        });
      });
      walkNestedSubtasks(task.subtasks || [], (sub, _pathArr, pathStr) => {
        (sub.notes || []).forEach((note) => {
          rows.push({
            note,
            assignedTo: assignedLabelForSub(task, pathStr),
            folderId: folderKeyForSub(task.id, pathStr),
            sortKey: noteSortKey(note)
          });
        });
      });
    });
    rows.sort((a, b) => String(b.sortKey).localeCompare(String(a.sortKey)));
    return rows;
  }

  function buildSubOverviewBranch(task, list, pathPrefix) {
    const branches = [];
    (list || []).forEach((sub, si) => {
      const path = pathPrefix ? `${pathPrefix}.${si}` : String(si);
      const children = buildSubOverviewBranch(task, sub.subtasks || [], path);
      const notes = sortNotesDesc(sub.notes);
      if (!notes.length && !children.length) return;
      branches.push({
        sub,
        path,
        folderId: folderKeyForSub(task.id, path),
        label: sub.text || "Untitled subitem",
        assignedTo: assignedLabelForSub(task, path),
        notes,
        children,
        depth: String(path).split(".").length
      });
    });
    return branches;
  }

  function buildOverviewTree() {
    const projectNotes = sortNotesDesc(extra.notes);
    const tasks = [];
    projectTasks().forEach((task) => {
      const children = buildSubOverviewBranch(task, task.subtasks || [], "");
      const notes = sortNotesDesc(task.notes);
      if (!notes.length && !children.length) return;
      tasks.push({
        task,
        folderId: folderKeyForTask(task.id),
        label: task.title || "Untitled task",
        notes,
        children
      });
    });
    return { projectNotes, tasks, total: collectOverviewNotes().length };
  }

  function countNotes(list) { return (list || []).length; }

  function subtreeHasText(sub, q) {
    if ((sub.text || "").toLowerCase().includes(q)) return true;
    return (sub.subtasks || []).some((child) => subtreeHasText(child, q));
  }

  function taskMatchesFilter(task, q) {
    if (!q) return true;
    if ((task.title || "").toLowerCase().includes(q)) return true;
    return (task.subtasks || []).some((s) => subtreeHasText(s, q));
  }

  function isTreeCollapsed(nodeId) {
    return !!collapsedMap[nodeId];
  }

  function setTreeCollapsed(nodeId, collapsed) {
    if (collapsed) collapsedMap[nodeId] = true;
    else delete collapsedMap[nodeId];
    window.AEWTTR.state.projNotesTreeCollapse[collapseKey] = collapsedMap;
  }

  function expandAncestorsOfActive() {
    if (!String(activeFolder || "").startsWith("sub:")) return;
    const rest = activeFolder.slice(4);
    const colonIdx = rest.indexOf(":");
    if (colonIdx < 0) return;
    const taskId = rest.slice(0, colonIdx);
    const path = rest.slice(colonIdx + 1);
    setTreeCollapsed(folderKeyForTask(taskId), false);
    const parts = String(path).split(".");
    for (let i = 0; i < parts.length - 1; i++) {
      setTreeCollapsed(folderKeyForSub(taskId, parts.slice(0, i + 1).join(".")), false);
    }
  }

  function persistThread(folder) {
    if (!folder) return;
    if (folder.kind === "project") Repo.save("project", proj);
    else if (folder.parentTask) Repo.save("actionItem", folder.parentTask, { projectCode: proj.id, source: "Tracker" });
  }

  function setActiveFolder(folderId) {
    activeFolder = folderId;
    window.AEWTTR.state.projNotesFolder[stateKey] = folderId;
    editingId = null;
    expandAncestorsOfActive();
    renderAll();
  }

  function noteBubbleHtml(n, { showAssigned } = {}) {
    const isMine = typeof isNoteAuthor === "function" ? isNoteAuthor(n) : false;
    const isEditing = editingId === n.id;
    const stamp = typeof formatNoteTimestamp === "function" ? formatNoteTimestamp(n) : `${n.date || ""}${n.time ? ` · ${n.time}` : ""}`;
    return `
      <article class="proj-notes-card ${isMine ? "mine" : ""}" data-note-id="${escapeHtml(n.id)}">
        <header class="proj-notes-card-head">
          ${showAssigned ? `<div class="proj-notes-card-assigned"><i class="bx bx-link-alt"></i> ${escapeHtml(showAssigned)}</div>` : ""}
          <div class="proj-notes-card-meta">
            <strong>${escapeHtml(n.author || "Unknown")}</strong>
            <span>${escapeHtml(stamp)}</span>
            ${isMine && !isEditing ? `
              <span class="task-notes-bubble-actions">
                <button type="button" data-edit-note="${n.id}" aria-label="Edit"><i class="bx bx-pencil"></i></button>
                <button type="button" data-del-note="${n.id}" aria-label="Delete"><i class="bx bx-trash"></i></button>
              </span>` : ""}
          </div>
        </header>
        ${isEditing
          ? `<textarea class="task-notes-input" id="proj-note-edit-${n.id}" rows="3">${escapeHtml(n.text || "")}</textarea>
             <div class="task-notes-edit-actions">
               <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-cancel-edit>Cancel</button>
               <button type="button" class="btn-aewttr btn-aewttr-sm" data-save-note="${n.id}">Save</button>
             </div>`
          : `<div class="proj-notes-card-text">${escapeHtml(n.text || "")}</div>`}
      </article>`;
  }

  function wireNoteActions(container, folder) {
    $all("[data-edit-note]", container).forEach((btn) => btn.addEventListener("click", () => {
      const note = (folder.notes || []).find((n) => n.id === btn.dataset.editNote);
      if (typeof isNoteAuthor === "function" && !isNoteAuthor(note)) {
        toast("You can only edit your own notes.", "error");
        return;
      }
      editingId = btn.dataset.editNote;
      renderMain();
    }));
    $all("[data-del-note]", container).forEach((btn) => btn.addEventListener("click", async () => {
      const note = (folder.notes || []).find((n) => n.id === btn.dataset.delNote);
      if (typeof isNoteAuthor === "function" && !isNoteAuthor(note)) {
        toast("You can only delete your own notes.", "error");
        return;
      }
      const ok = await confirmDialog({ title: "Delete note", message: "Delete this note? This cannot be undone.", confirmLabel: "Delete", danger: true });
      if (!ok) return;
      folder.notes = (folder.notes || []).filter((n) => n.id !== btn.dataset.delNote);
      if (folder.kind === "project") extra.notes = folder.notes;
      persistThread(folder);
      editingId = null;
      renderAll();
    }));
    $all("[data-save-note]", container).forEach((btn) => btn.addEventListener("click", () => {
      const ta = $(`#proj-note-edit-${btn.dataset.saveNote}`, container);
      const text = ta ? ta.value.trim() : "";
      if (!text) { toast("Note can't be empty.", "error"); return; }
      const note = (folder.notes || []).find((n) => n.id === btn.dataset.saveNote);
      if (!note || (typeof isNoteAuthor === "function" && !isNoteAuthor(note))) {
        toast("You can only edit your own notes.", "error");
        editingId = null;
        renderMain();
        return;
      }
      note.text = text;
      if (typeof touchNoteTimestamp === "function") touchNoteTimestamp(note);
      persistThread(folder);
      editingId = null;
      renderAll();
    }));
    $all("[data-cancel-edit]", container).forEach((btn) => btn.addEventListener("click", () => {
      editingId = null;
      renderMain();
    }));
  }

  function renderTreeRow({ folderId, label, icon, depth, noteCount, hasChildren, treeId, kind }) {
    const filtering = !!folderFilter.trim();
    const open = hasChildren && (filtering || !isTreeCollapsed(treeId));
    const caret = hasChildren
      ? `<button type="button" class="proj-notes-tree-caret" data-toggle-tree="${escapeHtml(treeId)}" aria-expanded="${open ? "true" : "false"}" aria-label="${open ? "Collapse" : "Expand"}">
           <i class="bx bx-chevron-${open ? "down" : "right"}"></i>
         </button>`
      : `<span class="proj-notes-tree-caret-spacer" aria-hidden="true"></span>`;
    return `
      <div class="proj-notes-folder-row proj-notes-folder-row--${escapeHtml(kind || "item")}" style="--notes-depth:${depth}">
        ${caret}
        <button type="button" class="proj-notes-folder ${activeFolder === folderId ? "is-active" : ""}" data-folder="${escapeHtml(folderId)}">
          <i class="bx ${icon}"></i>
          <span>${escapeHtml(label)}</span>
          ${noteCount ? `<em>${noteCount}</em>` : ""}
        </button>
      </div>`;
  }

  function renderSubTree(task, list, pathPrefix, depth) {
    let html = "";
    const filtering = !!folderFilter.trim();
    const q = folderFilter.trim().toLowerCase();
    const taskTitleHit = q && (task.title || "").toLowerCase().includes(q);
    (list || []).forEach((sub, si) => {
      if (q && !taskTitleHit && !subtreeHasText(sub, q)) return;
      const path = pathPrefix ? `${pathPrefix}.${si}` : String(si);
      const folderId = folderKeyForSub(task.id, path);
      const treeId = folderId;
      const children = sub.subtasks || [];
      const hasChildren = children.length > 0;
      const open = hasChildren && (filtering || !isTreeCollapsed(treeId));
      html += renderTreeRow({
        folderId,
        label: sub.text || "Untitled subitem",
        icon: depth > 1 ? "bx-git-merge" : "bx-subdirectory-right",
        depth,
        noteCount: countNotes(sub.notes),
        hasChildren,
        treeId,
        kind: "sub"
      });
      if (open) html += renderSubTree(task, children, path, depth + 1);
    });
    return html;
  }

  function renderSidebar() {
    const treeEl = $("#proj-notes-tree", body);
    if (!treeEl) return;
    const q = folderFilter.trim().toLowerCase();
    const filtering = !!q;
    const allItems = projectTasks();
    const dividers = allItems.filter((t) => isTrackerDivider(t));
    const tasks = allItems.filter((t) => !isTrackerDivider(t) && taskMatchesFilter(t, q));
    const filteredDividers = dividers.filter((d) => !q || (d.title || "").toLowerCase().includes(q));
    const overviewCount = collectOverviewNotes().length;
    const projectCount = countNotes(extra.notes);
    const riskCount = projectRisks(proj.id).filter((risk) => risk.status !== "Closed").length;

    treeEl.innerHTML = `
      <div class="proj-notes-folder-row" style="--notes-depth:0">
        <span class="proj-notes-tree-caret-spacer" aria-hidden="true"></span>
        <button type="button" class="proj-notes-folder ${activeFolder === "overview" ? "is-active" : ""}" data-folder="overview">
          <i class="bx bx-grid-alt"></i>
          <span>Overview</span>
          <em>${overviewCount}</em>
        </button>
      </div>
      <div class="proj-notes-folder-row proj-notes-folder-row--project" style="--notes-depth:0">
        <span class="proj-notes-tree-caret-spacer" aria-hidden="true"></span>
        <button type="button" class="proj-notes-folder ${activeFolder === "project" ? "is-active" : ""}" data-folder="project">
          <i class="bx bx-folder"></i>
          <span>Project Notes</span>
          <em>${projectCount}</em>
        </button>
      </div>
      <div class="proj-notes-folder-section">Operational records</div>
      <div class="proj-notes-folder-row proj-notes-folder-row--risks" style="--notes-depth:0">
        <span class="proj-notes-tree-caret-spacer" aria-hidden="true"></span>
        <button type="button" class="proj-notes-folder ${activeFolder === "risks" ? "is-active" : ""}" data-folder="risks">
          <i class="bx bx-shield-quarter"></i>
          <span>Risks</span>
          <em>${riskCount}</em>
        </button>
      </div>
      <div class="proj-notes-folder-section">Tasks</div>
      ${tasks.length ? tasks.map((task) => {
        const taskKey = folderKeyForTask(task.id);
        const treeId = taskKey;
        const subs = task.subtasks || [];
        const hasChildren = subs.length > 0;
        const open = hasChildren && (filtering || !isTreeCollapsed(treeId));
        return `
          <div class="proj-notes-folder-group">
            ${renderTreeRow({
              folderId: taskKey,
              label: task.title || "Untitled task",
              icon: "bx-task",
              depth: 0,
              noteCount: countNotes(task.notes),
              hasChildren,
              treeId,
              kind: "task"
            })}
            ${open ? renderSubTree(task, subs, "", 1) : ""}
          </div>`;
      }).join("") : `<div class="proj-notes-folder-empty">${q ? "No matching tasks." : "No tasks yet."}</div>`}
    `;

    $all("[data-folder]", treeEl).forEach((btn) => btn.addEventListener("click", () => setActiveFolder(btn.dataset.folder)));
    $all("[data-toggle-tree]", treeEl).forEach((btn) => btn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const id = btn.dataset.toggleTree;
      setTreeCollapsed(id, !isTreeCollapsed(id));
      renderSidebar();
    }));
  }

  function renderMain() {
    const mainEl = $("#proj-notes-main", body);
    if (!mainEl) return;
    const folder = resolveFolder(activeFolder) || resolveFolder("overview");
    if (!folder) {
      activeFolder = "overview";
      window.AEWTTR.state.projNotesFolder[stateKey] = "overview";
      return renderMain();
    }

    if (folder.kind === "overview") {
      const tree = buildOverviewTree();
      const rows = tree.total;

      function overviewBranchHtml(branch) {
        return `
          <div class="proj-notes-overview-branch" style="--notes-depth:${branch.depth}">
            <button type="button" class="proj-notes-overview-heading proj-notes-overview-heading--sub" data-jump-folder="${escapeHtml(branch.folderId)}">
              <i class="bx ${branch.depth > 1 ? "bx-git-merge" : "bx-subdirectory-right"}"></i>
              <span>${escapeHtml(branch.label)}</span>
              <em>${branch.notes.length}</em>
            </button>
            ${branch.notes.map((note) => `
              <div class="proj-notes-overview-item" data-jump-folder="${escapeHtml(branch.folderId)}">
                ${noteBubbleHtml(note, { showAssigned: branch.assignedTo })}
              </div>`).join("")}
            ${(branch.children || []).map((child) => overviewBranchHtml(child)).join("")}
          </div>`;
      }

      mainEl.innerHTML = `
        <div class="proj-notes-main-head">
          <div>
            <h2><i class="bx bx-grid-alt"></i> Overview</h2>
            <p>Project → tasks → subitems. New notes here are saved as project notes.</p>
          </div>
          <span class="aewttr-muted">${rows} note${rows === 1 ? "" : "s"}</span>
        </div>
        <div class="proj-notes-main-body" id="proj-notes-main-body">
          ${rows
            ? `<div class="proj-notes-overview-tree">
                ${tree.projectNotes.length ? `
                  <section class="proj-notes-overview-group">
                    <button type="button" class="proj-notes-overview-heading proj-notes-overview-heading--project" data-jump-folder="project">
                      <i class="bx bx-folder"></i>
                      <span>Project Notes</span>
                      <em>${tree.projectNotes.length}</em>
                    </button>
                    ${tree.projectNotes.map((note) => `
                      <div class="proj-notes-overview-item" data-jump-folder="project">
                        ${noteBubbleHtml(note, { showAssigned: "Project" })}
                      </div>`).join("")}
                  </section>` : ""}
                ${tree.tasks.map((taskNode) => `
                  <section class="proj-notes-overview-group">
                    <button type="button" class="proj-notes-overview-heading proj-notes-overview-heading--task" data-jump-folder="${escapeHtml(taskNode.folderId)}">
                      <i class="bx bx-task"></i>
                      <span>${escapeHtml(taskNode.label)}</span>
                      <em>${taskNode.notes.length}</em>
                    </button>
                    ${taskNode.notes.map((note) => `
                      <div class="proj-notes-overview-item" data-jump-folder="${escapeHtml(taskNode.folderId)}">
                        ${noteBubbleHtml(note, { showAssigned: `Task · ${taskNode.label}` })}
                      </div>`).join("")}
                    ${(taskNode.children || []).map((child) => overviewBranchHtml(child)).join("")}
                  </section>`).join("")}
              </div>`
            : `<div class="empty-state" style="padding:40px;text-align:center;">No notes yet — write one below to start the project thread.</div>`}
        </div>
        <div class="task-notes-input-row proj-notes-compose">
          <textarea class="task-notes-input" id="proj-notes-input" placeholder="Write a project note — Enter to post…" rows="1"></textarea>
          <button type="button" class="btn-aewttr btn-aewttr-sm task-notes-send" id="proj-notes-send"><i class="bx bx-send"></i></button>
        </div>`;

      // Overview is read-only for editing through the flattened list; jump into
      // the owning folder so author-only edit/delete still has a clear save target.
      $all("[data-jump-folder]", mainEl).forEach((el) => {
        el.addEventListener("click", (e) => {
          if (e.target.closest("button:not([data-jump-folder]), textarea, .proj-notes-compose")) return;
          setActiveFolder(el.dataset.jumpFolder);
        });
      });
      // Allow edit/delete from overview by resolving owning folder per note
      $all("[data-edit-note], [data-del-note], [data-save-note], [data-cancel-edit]", mainEl).forEach((btn) => {
        btn.addEventListener("click", (e) => {
          e.stopPropagation();
          const card = btn.closest("[data-jump-folder]");
          const owning = resolveFolder(card && card.dataset.jumpFolder);
          if (!owning || !owning.notes) return;
          if (btn.hasAttribute("data-edit-note")) {
            const note = owning.notes.find((n) => n.id === btn.dataset.editNote);
            if (typeof isNoteAuthor === "function" && !isNoteAuthor(note)) {
              toast("You can only edit your own notes.", "error");
              return;
            }
            setActiveFolder(owning.id);
            editingId = btn.dataset.editNote;
            renderMain();
            return;
          }
          if (btn.hasAttribute("data-del-note")) {
            (async () => {
              const note = owning.notes.find((n) => n.id === btn.dataset.delNote);
              if (typeof isNoteAuthor === "function" && !isNoteAuthor(note)) {
                toast("You can only delete your own notes.", "error");
                return;
              }
              const ok = await confirmDialog({ title: "Delete note", message: "Delete this note? This cannot be undone.", confirmLabel: "Delete", danger: true });
              if (!ok) return;
              owning.notes = owning.notes.filter((n) => n.id !== btn.dataset.delNote);
              if (owning.kind === "project") extra.notes = owning.notes;
              persistThread(owning);
              renderAll();
            })();
          }
        });
      });

      const overviewInput = $("#proj-notes-input", mainEl);
      const overviewSend = $("#proj-notes-send", mainEl);
      function postOverviewNote() {
        const text = (overviewInput.value || "").trim();
        if (!text) return;
        const note = typeof stampNoteAuthor === "function"
          ? stampNoteAuthor({ id: uid("pn"), text })
          : { id: uid("pn"), author: (db.user && db.user.name) || "Unknown", text };
        if (typeof touchNoteTimestamp === "function") touchNoteTimestamp(note);
        delete note.editedAt;
        if (!Array.isArray(extra.notes)) extra.notes = [];
        extra.notes.unshift(note);
        Repo.save("project", proj);
        overviewInput.value = "";
        overviewInput.style.height = "auto";
        renderAll();
      }
      if (overviewSend) overviewSend.addEventListener("click", postOverviewNote);
      if (overviewInput) {
        overviewInput.addEventListener("keydown", (e) => {
          if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); postOverviewNote(); }
        });
        overviewInput.addEventListener("input", () => {
          overviewInput.style.height = "auto";
          overviewInput.style.height = Math.min(overviewInput.scrollHeight, 120) + "px";
        });
      }
      return;
    }

    if (folder.kind === "risks") {
      const risks = projectRisks(proj.id).slice().sort((a, b) => {
        const order = { Red: 0, Amber: 1, Green: 2 };
        return (order[a.rating] ?? 3) - (order[b.rating] ?? 3);
      });
      mainEl.innerHTML = `
        <div class="proj-notes-main-head">
          <div><h2><i class="bx bx-shield-quarter"></i> Risks</h2><p>Current project risk context, mitigation, and review history.</p></div>
          <button type="button" class="btn-aewttr-outline btn-aewttr-sm" id="proj-notes-open-risks">Open risk register</button>
        </div>
        <div class="proj-notes-main-body proj-notes-risk-list">
          ${risks.length ? risks.map((risk) => `
            <article class="proj-notes-risk-card">
              <div class="proj-notes-risk-head">
                <div>${riskRatingPill(risk.rating)} <strong>${escapeHtml(risk.name)}</strong></div>
                <span class="proj-notes-risk-actions">
                  <button type="button" data-open-risk="${escapeHtml(risk.id)}" title="Edit risk"><i class="bx bx-pencil"></i></button>
                  <button type="button" data-del-risk="${escapeHtml(risk.id)}" title="Delete risk"><i class="bx bx-trash"></i></button>
                </span>
              </div>
              <p>${escapeHtml(risk.description || risk.mitigationPlan || "No additional risk context recorded.")}</p>
              <div class="proj-notes-risk-meta"><span><b>Owner</b> ${escapeHtml(risk.owner || "Unassigned")}</span><span><b>Status</b> ${escapeHtml(risk.status || "Open")}</span><span><b>Due</b> ${risk.due ? escapeHtml(fmtDate(risk.due)) : "Not set"}</span></div>
              ${risk.mitigationPlan ? `<div class="proj-notes-risk-mitigation"><b>Mitigation</b> ${escapeHtml(risk.mitigationPlan)}</div>` : ""}
            </article>`).join("") : `<div class="empty-state" style="padding:32px;text-align:center;">No risks are recorded for this project.</div>`}
        </div>`;
      const registerBtn = $("#proj-notes-open-risks", mainEl);
      if (registerBtn) registerBtn.addEventListener("click", () => navigate(`projects/${proj.id}/risks`));
      $all("[data-open-risk]", mainEl).forEach((btn) => btn.addEventListener("click", () => openRiskModal(proj, btn.dataset.openRisk, () => renderAll())));
      $all("[data-del-risk]", mainEl).forEach((btn) => btn.addEventListener("click", async () => {
        const riskId = btn.dataset.delRisk;
        const allRisks = projectRisks(proj.id);
        const risk = allRisks.find((r) => r.id === riskId);
        if (!risk) return;
        if (!confirm(`Delete "${risk.name || "this risk"}"? This cannot be undone.`)) return;
        try {
          await deleteProjectRisk(proj, risk);
          toast("Risk deleted.", "success");
          renderAll();
        } catch (e) {
          toast((e && e.friendly) || "Could not delete risk.", "error");
        }
      }));
      return;
    }

    const notes = (folder.notes || []).slice();
    const crumb = (folder.breadcrumb && folder.breadcrumb.length)
      ? folder.breadcrumb.join(" › ")
      : (folder.assignedTo || folder.label);
    mainEl.innerHTML = `
      <div class="proj-notes-main-head">
        <div>
          <h2><i class="bx ${folder.kind === "project" ? "bx-folder" : folder.kind === "sub" ? "bx-subdirectory-right" : "bx-task"}"></i> ${escapeHtml(folder.label)}</h2>
          <p>${escapeHtml(crumb)}</p>
        </div>
        <span class="aewttr-muted">${notes.length} note${notes.length === 1 ? "" : "s"}</span>
      </div>
      <div class="proj-notes-main-body" id="proj-notes-main-body">
        ${notes.length
          ? notes.map((n) => noteBubbleHtml(n, { showAssigned: folder.assignedTo })).join("")
          : `<div class="empty-state" style="padding:32px;text-align:center;">No notes in this folder yet — add the first one below.</div>`}
      </div>
      <div class="task-notes-input-row proj-notes-compose">
        <textarea class="task-notes-input" id="proj-notes-input" placeholder="Type a note — Enter to post…" rows="1"></textarea>
        <button type="button" class="btn-aewttr btn-aewttr-sm task-notes-send" id="proj-notes-send"><i class="bx bx-send"></i></button>
      </div>`;

    const bodyEl = $("#proj-notes-main-body", mainEl);
    wireNoteActions(bodyEl || mainEl, folder);

    const input = $("#proj-notes-input", mainEl);
    const sendBtn = $("#proj-notes-send", mainEl);
    function postNote() {
      const text = (input.value || "").trim();
      if (!text) return;
      const note = typeof stampNoteAuthor === "function"
        ? stampNoteAuthor({ id: uid("pn"), text })
        : { id: uid("pn"), author: (db.user && db.user.name) || "Unknown", text };
      if (typeof touchNoteTimestamp === "function") touchNoteTimestamp(note);
      delete note.editedAt;
      folder.notes.unshift(note);
      if (folder.kind === "project") extra.notes = folder.notes;
      persistThread(folder);
      input.value = "";
      input.style.height = "auto";
      renderAll();
    }
    if (sendBtn) sendBtn.addEventListener("click", postNote);
    if (input) {
      input.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); postNote(); }
      });
      input.addEventListener("input", () => {
        input.style.height = "auto";
        input.style.height = Math.min(input.scrollHeight, 120) + "px";
      });
    }
  }

  function renderAll() {
    renderSidebar();
    renderMain();
  }

  body.innerHTML = `
    <div class="proj-notes-shell" id="proj-notes-shell">
      <aside class="proj-notes-sidebar">
        <div class="proj-notes-sidebar-head">
          <h2>Notes</h2>
          <div class="proj-notes-task-search">
            <i class="bx bx-search"></i>
            <input type="search" id="proj-notes-folder-filter" placeholder="Filter folders…" aria-label="Filter note folders">
          </div>
        </div>
        <div class="proj-notes-tree" id="proj-notes-tree"></div>
      </aside>
      <div class="proj-notes-resizer" id="proj-notes-resizer" aria-hidden="true" title="Drag to resize sidebar"></div>
      <section class="proj-notes-main" id="proj-notes-main"></section>
    </div>
  `;

  const filterInput = $("#proj-notes-folder-filter", body);
  if (filterInput) {
    filterInput.addEventListener("input", (e) => {
      folderFilter = e.target.value;
      renderSidebar();
    });
  }

  // Resizable sidebar drag
  (function wireNotesResizer() {
    const shell = $("#proj-notes-shell", body);
    const resizer = $("#proj-notes-resizer", body);
    if (!shell || !resizer) return;
    const SIDEBAR_KEY = "projNotesSidebarW";
    const saved = parseInt(window.AEWTTR.state[SIDEBAR_KEY] || 0, 10);
    if (saved >= 160 && saved <= 480) shell.style.setProperty("--sidebar-w", saved + "px");
    let startX = 0;
    let startW = 0;
    resizer.addEventListener("mousedown", (e) => {
      e.preventDefault();
      startX = e.clientX;
      startW = shell.getBoundingClientRect().width;
      const sidebar = shell.querySelector(".proj-notes-sidebar");
      startW = sidebar ? sidebar.getBoundingClientRect().width : (parseInt(getComputedStyle(shell).getPropertyValue("--sidebar-w")) || 240);
      resizer.classList.add("is-dragging");
      document.body.style.cursor = "col-resize";
      document.body.style.userSelect = "none";
      const onMove = (me) => {
        const delta = me.clientX - startX;
        const newW = Math.min(480, Math.max(160, startW + delta));
        shell.style.setProperty("--sidebar-w", newW + "px");
      };
      const onUp = () => {
        resizer.classList.remove("is-dragging");
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        window.removeEventListener("mousemove", onMove);
        window.removeEventListener("mouseup", onUp);
        const finalW = parseInt(shell.style.getPropertyValue("--sidebar-w")) || 240;
        window.AEWTTR.state[SIDEBAR_KEY] = finalW;
      };
      window.addEventListener("mousemove", onMove);
      window.addEventListener("mouseup", onUp);
    });
  })();

  expandAncestorsOfActive();
  renderAll();
}

function drawProjectMeeting(body, proj) {
  if (window.AEWTTR && typeof window.AEWTTR.renderProjectMeetingApp === "function") {
    return window.AEWTTR.renderProjectMeetingApp(body, proj);
  }
  body.innerHTML = `<div class="empty-state" style="padding:40px;">Project meeting app is loading.</div>`;
}

function openProjectContractorModal(proj, existing, onDone) {
  const db = window.AEWTTR.db;
  if (!db.projectContractors) db.projectContractors = {};
  if (!db.projectContractors[proj.id]) db.projectContractors[proj.id] = [];
  const isEdit = !!existing;
  const c = existing || {};

  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>${isEdit ? "Edit Contractor / Vendor" : "Add Contractor / Vendor"}</h3><button class="aewttr-modal-close" type="button">&times;</button></div>
    <div class="aewttr-modal-body">
      <div class="form-row"><label>Company name <span class="required-star">*</span></label><input class="input-aewttr" id="pc-company" placeholder="e.g., Booz Allen Hamilton" value="${escapeHtml(c.company || "")}"></div>
      <div class="form-row"><label>Relationship / role on project</label><input class="input-aewttr" id="pc-role" placeholder="e.g., Prime contractor, Subcontractor, Vendor" value="${escapeHtml(c.role || "")}"></div>
      <div class="form-row"><label>Point of contact name</label><input class="input-aewttr" id="pc-poc-name" placeholder="Full name" value="${escapeHtml(c.pocName || "")}"></div>
      <div class="form-row form-row-2col">
        <div><label>POC email</label><input class="input-aewttr" id="pc-poc-email" type="email" placeholder="email@company.com" value="${escapeHtml(c.pocEmail || "")}"></div>
        <div><label>POC phone</label><input class="input-aewttr" id="pc-poc-phone" placeholder="DSN or cell" value="${escapeHtml(c.pocPhone || "")}"></div>
      </div>
      <div class="form-row" style="margin-bottom:0;"><label>Notes</label><textarea class="textarea-aewttr" id="pc-notes" rows="2" placeholder="Contract number, task order, period of performance…">${escapeHtml(c.notes || "")}</textarea></div>
    </div>
    <div class="aewttr-modal-foot">
      ${isEdit ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-danger" id="pc-delete"><i class="bx bx-trash"></i> Remove</button>` : ""}
      <button type="button" class="btn-aewttr-ghost" id="pc-cancel">Cancel</button>
      <button type="button" class="btn-aewttr" id="pc-save">${isEdit ? "Save" : "Add Company"}</button>
    </div>
  `);

  function close() { modal.closest(".aewttr-modal-backdrop")?.remove(); }
  const cancelBtn = $("#pc-cancel", modal);
  if (cancelBtn) cancelBtn.addEventListener("click", close);

  const deleteBtn = $("#pc-delete", modal);
  if (deleteBtn) deleteBtn.addEventListener("click", async () => {
    if (!confirm("Remove this contractor from the project?")) return;
    db.projectContractors[proj.id] = (db.projectContractors[proj.id] || []).filter(x => x.id !== existing.id);
    try { if (typeof aewttrSaveStore === "function") aewttrSaveStore(); } catch (_) {}
    close();
    if (typeof onDone === "function") onDone();
  });

  const saveBtn = $("#pc-save", modal);
  if (saveBtn) saveBtn.addEventListener("click", async () => {
    const company = ($("#pc-company", modal).value || "").trim();
    if (!company) { toast("Company name is required.", "error"); return; }
    const payload = {
      id: existing ? existing.id : uid("pc"),
      company,
      role: ($("#pc-role", modal).value || "").trim(),
      pocName: ($("#pc-poc-name", modal).value || "").trim(),
      pocEmail: ($("#pc-poc-email", modal).value || "").trim(),
      pocPhone: ($("#pc-poc-phone", modal).value || "").trim(),
      notes: ($("#pc-notes", modal).value || "").trim()
    };
    if (isEdit) {
      const idx = db.projectContractors[proj.id].findIndex(x => x.id === existing.id);
      if (idx !== -1) db.projectContractors[proj.id][idx] = payload;
    } else {
      db.projectContractors[proj.id].push(payload);
    }
    try { if (typeof aewttrSaveStore === "function") aewttrSaveStore(); } catch (_) {}
    close();
    if (typeof onDone === "function") onDone();
  });
}

function drawPeople(body, proj) {
  const db = window.AEWTTR.db;
  if (!db.projectPeople) db.projectPeople = {};
  if (!db.projectPeople[proj.id]) db.projectPeople[proj.id] = [];
  if (!db.projectContractors) db.projectContractors = {};
  if (!db.projectContractors[proj.id]) db.projectContractors[proj.id] = [];
  const people = db.projectPeople[proj.id];
  const contractors = db.projectContractors[proj.id];

  const pmRoles = new Set(["pm", "project manager", "program manager", "lead pm", "deputy pm", "asst pm", "assistant pm"]);
  function isPmRole(role) { return pmRoles.has((role || "").trim().toLowerCase()); }

  const PEOPLE_PRESET_ROLES = ["Project Manager","Deputy PM","Program Manager","Lead Engineer","Systems Engineer","Software Engineer","ISSO","Range POC","Contracting Officer","Technical Advisor","SME","Reviewer","Contractor","Subcontractor","Stakeholder"];
  function roleComboHtml(person) {
    const cur = person.role || "";
    const isPreset = PEOPLE_PRESET_ROLES.includes(cur);
    return `<div class="role-combo-wrap" data-person-role="${person.id}" style="min-width:130px;">
      <select class="select-aewttr role-combo-sel" style="width:100%;">
        <option value="">Set role…</option>
        ${PEOPLE_PRESET_ROLES.map(r => `<option value="${escapeHtml(r)}"${cur === r ? " selected" : ""}>${escapeHtml(r)}</option>`).join("")}
        <option value="__custom__"${!isPreset && cur ? " selected" : ""}>Custom…</option>
      </select>
      <input class="input-aewttr role-combo-inp" style="${!isPreset && cur ? "margin-top:4px;" : "display:none;margin-top:4px;"}" placeholder="Custom role…" value="${escapeHtml(!isPreset ? cur : "")}">
    </div>`;
  }

  body.innerHTML = `
    <div class="project-people-page">
      <div class="project-people-head">
         <div>
            <div class="side-panel-title">Project People</div>
            <p data-help>This is the project roster. Type a role directly in the list — start typing for suggestions.</p>
         </div>
          <button class="btn-aewttr" id="btn-add-project-person"${tip("Search and add a person to this project roster")}><i class="bx bx-user-plus"></i> Add Person</button>
      </div>
      <div class="project-people-table-wrap">
       ${people.length ? `
         <table class="aewttr-table">
           <thead><tr><th>Name</th><th>Role</th><th>Email</th><th>Type</th><th>Project Admin</th><th></th></tr></thead>
           <tbody>
             ${people.map(person => `
               <tr style="cursor:default;">
                 <td>
                   <div class="proj-person-name-cell">
                     <strong>${escapeHtml(person.label)}</strong>
                     ${isPmRole(person.role) ? `<span class="proj-person-pm-badge"${tip("Project Manager")}>PM</span>` : ""}
                   </div>
                 </td>
                 <td>${roleComboHtml(person)}</td>
                 <td>${escapeHtml(person.email || "—")}</td>
                 <td><span class="proj-person-type-chip">${escapeHtml(person.type === "member" ? "Team" : person.type === "person" ? "Contractor" : "Company")}</span></td>
                 <td style="text-align:center;">
                   <button type="button" class="proj-person-admin-toggle ${person.isProjectAdmin ? "active" : ""}" data-toggle-admin="${person.id}"${tip(person.isProjectAdmin ? "Project Admin — click to remove" : "Grant Project Admin access")} aria-pressed="${!!person.isProjectAdmin}">
                     <i class="bx ${person.isProjectAdmin ? "bxs-shield-alt-2" : "bx-shield"}"></i>
                   </button>
                 </td>
                 <td class="travel-actions-cell">
                    <button class="btn-aewttr-outline btn-aewttr-sm" data-edit-person="${person.id}"${tip(`Edit ${person.label}'s contact details`)}>Details</button>
                   <button class="btn-aewttr-outline btn-aewttr-sm" data-remove-person="${person.id}"${tip(`Remove ${person.label} from this project`)}>Remove</button>
                 </td>
               </tr>`).join("")}
           </tbody>
         </table>` : `<div class="empty-state">No one is attached to this project yet.</div>`}
       </div>
     </div>
   </div>
 `;

  $("#btn-add-project-person", body).addEventListener("click", () => openProjectRosterSearchModal(proj, () => drawPeople(body, proj)));
  $all("[data-person-role]", body).forEach((wrap) => {
    const person = people.find((entry) => entry.id === wrap.dataset.personRole);
    if (!person) return;
    const sel = wrap.querySelector(".role-combo-sel");
    const inp = wrap.querySelector(".role-combo-inp");
    if (!sel) return;
    async function saveRole(role) {
      if (!role) return;
      person.role = role;
      await Repo.save("project", proj);
      toast("Project role saved", "success");
    }
    sel.addEventListener("change", async () => {
      if (sel.value === "__custom__") {
        inp.style.display = "block";
        inp.focus();
      } else {
        inp.style.display = "none";
        await saveRole(sel.value);
      }
    });
    if (inp) inp.addEventListener("change", () => saveRole(inp.value.trim()));
  });
  $all("[data-edit-person]", body).forEach(btn => btn.addEventListener("click", () => {
    const person = people.find(p => p.id === btn.dataset.editPerson);
    if (person) openProjectPersonModal(proj, person, () => drawPeople(body, proj));
  }));
  $all("[data-toggle-admin]", body).forEach(btn => btn.addEventListener("click", async () => {
    const person = people.find(p => p.id === btn.dataset.toggleAdmin);
    if (!person) return;
    person.isProjectAdmin = !person.isProjectAdmin;
    await Repo.save("project", proj);
    drawPeople(body, proj);
    toast(person.isProjectAdmin ? `${person.label} granted Project Admin` : `${person.label} removed as Project Admin`, "success");
  }));
  $all("[data-remove-person]", body).forEach(btn => btn.addEventListener("click", async () => {
    db.projectPeople[proj.id] = (db.projectPeople[proj.id] || []).filter(person => person.id !== btn.dataset.removePerson);
    if (db.weeklyMeeting && db.weeklyMeeting.projectMeetings && db.weeklyMeeting.projectMeetings[proj.id]) {
      db.weeklyMeeting.projectMeetings[proj.id].participants = db.projectPeople[proj.id].map(person => ({ ...person }));
    }
    await Repo.save("project", proj);
    if (typeof syncProjectPulseGroup === "function") {
      try { await syncProjectPulseGroup(proj); } catch (e) { console.warn("project group sync", e); }
    }
    drawPeople(body, proj);
    toast("Project person removed", "success");
  }));

}

/* `editing` is null to add a new person, or an existing db.projectPeople
   entry to edit it in place (used both by the People tab's Edit button and
   by openAssignPersonModal's "New Person" flow). */
function openProjectRosterSearchModal(proj, onDone) {
  const db = window.AEWTTR.db;
  if (!db.projectPeople) db.projectPeople = {};
  if (!db.projectPeople[proj.id]) db.projectPeople[proj.id] = [];
  if (!db.projectContractors) db.projectContractors = {};
  const roster = db.projectPeople[proj.id];

  // Build a deduplicated list of known contractor companies across all projects
  function getAllContractorCompanies() {
    const seen = new Set();
    const list = [];
    Object.values(db.projectContractors || {}).forEach((arr) => {
      (arr || []).forEach((c) => {
        if (c.company && !seen.has(c.company.toLowerCase())) {
          seen.add(c.company.toLowerCase());
          list.push(c.company);
        }
      });
    });
    (db.projects || []).forEach((p) => {
      if (p.contractor && !seen.has(p.contractor.toLowerCase())) {
        seen.add(p.contractor.toLowerCase());
        list.push(p.contractor);
      }
    });
    return list.sort();
  }

  let activeTab = "team";
  let selectedMember = null;

  const ROLE_DATALIST = `
    <datalist id="roster-role-suggestions">
      <option value="Project Manager"><option value="Deputy PM"><option value="Program Manager">
      <option value="Lead Engineer"><option value="Systems Engineer"><option value="Software Engineer">
      <option value="ISSO"><option value="Range POC"><option value="Contracting Officer">
      <option value="Technical Advisor"><option value="SME"><option value="Reviewer">
      <option value="Contractor"><option value="Subcontractor"><option value="Stakeholder">
    </datalist>`;

  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>Add Person to Project</h3><button class="aewttr-modal-close" type="button">&times;</button></div>
    <div class="aewttr-modal-body" style="padding-top:0;">
      ${ROLE_DATALIST}
      <div class="add-person-tabs" style="display:flex;gap:0;border-bottom:1px solid var(--aewttr-border);margin-bottom:16px;">
        <button type="button" class="add-person-tab active" data-tab="team" style="flex:1;padding:10px 6px;font-size:12px;font-weight:600;background:none;border:none;border-bottom:2px solid var(--aewttr-accent);color:var(--aewttr-accent);cursor:pointer;"><i class="bx bx-group"></i> Team</button>
        <button type="button" class="add-person-tab" data-tab="contractor" style="flex:1;padding:10px 6px;font-size:12px;font-weight:600;background:none;border:none;border-bottom:2px solid transparent;color:var(--aewttr-muted);cursor:pointer;"><i class="bx bx-hard-hat"></i> Contractor</button>
        <button type="button" class="add-person-tab" data-tab="gov" style="flex:1;padding:10px 6px;font-size:12px;font-weight:600;background:none;border:none;border-bottom:2px solid transparent;color:var(--aewttr-muted);cursor:pointer;"><i class="bx bx-building-house"></i> Gov</button>
      </div>

      <!-- TEAM TAB -->
      <div id="tab-team">
        <div class="form-row" style="position:relative;">
          <label>Search app users</label>
          <div id="project-person-selected" class="traveler-chip-list"></div>
          <input class="input-aewttr" id="project-person-search" placeholder="Search name or email…" autocomplete="off">
          <div id="project-person-results" class="traveler-suggestions"></div>
        </div>
        <div id="project-person-role-row">
          <div class="form-row" style="margin-bottom:0;"><label>Role on this project</label><input class="input-aewttr" id="project-person-role" list="roster-role-suggestions" placeholder="e.g., Lead Engineer, PM, ISSO…"></div>
        </div>
      </div>

      <!-- CONTRACTOR TAB -->
      <div id="tab-contractor" hidden>
        <div class="form-row"><label>Name <span class="required-star">*</span></label><input class="input-aewttr" id="pc-name" placeholder="Full name"></div>
        <div class="form-row"><label>Email</label><input class="input-aewttr" id="pc-email" type="email" placeholder="name@contractor.com (optional)"></div>
        <div class="form-row">
          <label>Contractor company</label>
          <div style="position:relative;">
            <div id="pc-company-display" style="cursor:pointer;padding:8px 10px;background:var(--aewttr-surface-2);border-radius:6px;border:1px solid var(--aewttr-border);font-size:13px;display:flex;align-items:center;justify-content:space-between;gap:6px;min-height:38px;">
              <span id="pc-company-label" style="color:var(--aewttr-muted);">Select company…</span>
              <i class="bx bx-chevron-down" style="color:var(--aewttr-muted);flex-shrink:0;"></i>
            </div>
            <input type="hidden" id="pc-company" value="">
            <input class="input-aewttr" id="pc-company-search" placeholder="Search company…" autocomplete="off" style="display:none;margin-top:4px;">
            <div id="pc-company-suggestions" style="display:none;position:absolute;top:calc(100% + 2px);left:0;right:0;z-index:9999;background:var(--aewttr-surface);border:1px solid var(--aewttr-border);border-radius:6px;box-shadow:0 4px 16px rgba(0,0,0,.18);max-height:180px;overflow-y:auto;"></div>
          </div>
        </div>
        <div class="form-row" style="margin-bottom:0;"><label>Role on this project</label><input class="input-aewttr" id="pc-role" list="roster-role-suggestions" placeholder="e.g., Systems Engineer, Subcontractor…"></div>
      </div>

      <!-- GOV TAB -->
      <div id="tab-gov" hidden>
        <p style="font-size:12px;color:var(--aewttr-muted);margin:0 0 12px;">Use for government contacts, CORs, TRs, or other gov stakeholders not in the app.</p>
        <div class="form-row"><label>Name <span class="required-star">*</span></label><input class="input-aewttr" id="pg-name" placeholder="Full name"></div>
        <div class="form-row"><label>Email</label><input class="input-aewttr" id="pg-email" type="email" placeholder="name@mail.mil (optional)"></div>
        <div class="form-row"><label>Organization</label><input class="input-aewttr" id="pg-org" placeholder="e.g., Army, DEVCOM, PEO EW&S…"></div>
        <div class="form-row" style="margin-bottom:0;"><label>Role / title</label><input class="input-aewttr" id="pg-role" list="roster-role-suggestions" placeholder="e.g., COR, Technical Representative, PM…"></div>
      </div>
    </div>
    <div class="aewttr-modal-foot">
      <button type="button" class="btn-aewttr-ghost" id="project-person-cancel">Cancel</button>
      <button type="button" class="btn-aewttr" id="project-person-add" disabled>Add to Project</button>
    </div>
  `);

  const addButton = $("#project-person-add", modal);

  // Wire custom contractor company picker
  (function() {
    const display = $("#pc-company-display", modal);
    const label = $("#pc-company-label", modal);
    const hidden = $("#pc-company", modal);
    const search = $("#pc-company-search", modal);
    const sugg = $("#pc-company-suggestions", modal);
    if (!display) return;
    const companies = getAllContractorCompanies();

    function showSuggestions(q) {
      const lq = (q || "").toLowerCase();
      const trimmed = (q || "").trim();
      const matches = companies.filter(c => !lq || c.toLowerCase().includes(lq)).slice(0, 10);
      const btnStyle = "display:block;width:100%;text-align:left;padding:7px 12px;background:none;border:none;font-size:13px;cursor:pointer;color:var(--aewttr-text);";
      let html = matches.map(c => `<button type="button" data-co="${escapeHtml(c)}" style="${btnStyle}">${escapeHtml(c)}</button>`).join("");
      if (trimmed && !matches.find(c => c.toLowerCase() === trimmed.toLowerCase())) {
        html += `<button type="button" data-co="${escapeHtml(trimmed)}" style="${btnStyle}font-weight:600;color:var(--aewttr-accent);border-top:${matches.length ? "1px solid var(--aewttr-border)" : "none"};"><i class="bx bx-plus-circle"></i> Use "${escapeHtml(trimmed)}"</button>`;
      }
      if (!html) {
        html = `<div style="padding:8px 12px;color:var(--aewttr-muted);font-size:12px;">No companies saved yet — type a name above.</div>`;
      }
      sugg.innerHTML = html;
      $all("[data-co]", sugg).forEach(btn => {
        btn.addEventListener("mousedown", (e) => {
          e.preventDefault();
          const val = btn.dataset.co;
          hidden.value = val;
          label.textContent = val;
          label.style.color = "var(--aewttr-text)";
          sugg.style.display = "none";
          search.style.display = "none";
          display.style.display = "flex";
        });
      });
      sugg.style.display = "block";
    }

    display.addEventListener("click", () => {
      display.style.display = "none";
      search.style.display = "block";
      search.value = hidden.value;
      search.focus();
      showSuggestions(hidden.value);
    });

    search.addEventListener("input", () => showSuggestions(search.value));

    search.addEventListener("blur", () => {
      setTimeout(() => {
        if (search.value.trim()) {
          hidden.value = search.value.trim();
          label.textContent = search.value.trim();
          label.style.color = "var(--aewttr-text)";
        }
        sugg.style.display = "none";
        search.style.display = "none";
        display.style.display = "flex";
      }, 160);
    });
  }());

  // Tab switching
  $all(".add-person-tab", modal).forEach((tab) => {
    tab.addEventListener("click", () => {
      activeTab = tab.dataset.tab;
      selectedMember = null;
      addButton.disabled = true;
      $all(".add-person-tab", modal).forEach((t) => {
        const isActive = t.dataset.tab === activeTab;
        t.classList.toggle("active", isActive);
        t.style.borderBottomColor = isActive ? "var(--aewttr-accent)" : "transparent";
        t.style.color = isActive ? "var(--aewttr-accent)" : "var(--aewttr-muted)";
      });
      $("#tab-team", modal).hidden = activeTab !== "team";
      $("#tab-contractor", modal).hidden = activeTab !== "contractor";
      $("#tab-gov", modal).hidden = activeTab !== "gov";
      if (activeTab === "contractor") {
        addButton.disabled = false;
        $("#pc-name", modal).focus();
      }
      if (activeTab === "gov") {
        addButton.disabled = false;
        $("#pg-name", modal).focus();
      }
      if (activeTab === "team") {
        addButton.disabled = false;
        const search = $("#project-person-search", modal);
        if (search) { search.value = ""; search.focus(); }
      }
    });
  });

  // Team search
  const selectedMembers = [];
  wirePeoplePicker(modal, selectedMembers, { mount: "project-person-selected", input: "project-person-search", suggestions: "project-person-results" }, { allowManualEmail: false, includeGroups: false, expandGroups: false });

  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#project-person-cancel", modal).addEventListener("click", closeModal);

  addButton.addEventListener("click", async () => {
    let entry;
    if (activeTab === "team") {
      if (!selectedMembers.length) { toast("Select at least one team member.", "error"); return; }
      const roleVal = ($("#project-person-role", modal).value || "").trim();
      const entries = selectedMembers.map(selectedMember => ({
        type: "member", memberId: selectedMember.id, label: selectedMember.name, email: selectedMember.email || "", company: "", role: roleVal || selectedMember.role || ""
      }));
      
      entries.forEach(newEntry => {
        const existing = roster.find((p) =>
          (newEntry.memberId && p.memberId === newEntry.memberId) ||
          (!newEntry.memberId && newEntry.email && p.email && p.email.toLowerCase() === newEntry.email.toLowerCase()) ||
          (!newEntry.memberId && p.label.toLowerCase() === newEntry.label.toLowerCase())
        );
        if (existing) { Object.assign(existing, newEntry); }
        else roster.push(Object.assign({ id: uid("ppj") }, newEntry));
      });
      // Skip the default entry push at the end since we handled it in the loop
      entry = null;
    } else if (activeTab === "contractor") {
      const label = ($("#pc-name", modal).value || "").trim();
      if (!label) { toast("Contractor name is required.", "error"); return; }
      entry = {
        type: "person", memberId: "", label,
        email: ($("#pc-email", modal).value || "").trim(),
        company: ($("#pc-company", modal).value || "").trim(),
        role: ($("#pc-role", modal).value || "").trim() || "Contractor"
      };
    } else {
      const label = ($("#pg-name", modal).value || "").trim();
      if (!label) { toast("Name is required.", "error"); return; }
      entry = {
        type: "person", memberId: "", label,
        email: ($("#pg-email", modal).value || "").trim(),
        company: ($("#pg-org", modal).value || "").trim(),
        role: ($("#pg-role", modal).value || "").trim() || "Government",
        isGov: true
      };
    }
    if (entry) {
      const existing = roster.find((p) =>
        (entry.memberId && p.memberId === entry.memberId) ||
        (!entry.memberId && entry.email && p.email && p.email.toLowerCase() === entry.email.toLowerCase()) ||
        (!entry.memberId && p.label.toLowerCase() === entry.label.toLowerCase())
      );
      if (existing) { Object.assign(existing, entry); }
      else roster.push(Object.assign({ id: uid("ppj") }, entry));
    }
    
    if (db.weeklyMeeting && db.weeklyMeeting.projectMeetings && db.weeklyMeeting.projectMeetings[proj.id]) {
      db.weeklyMeeting.projectMeetings[proj.id].participants = roster.map((p) => Object.assign({}, p));
    }
    await Repo.save("project", proj);
    if (typeof syncProjectPulseGroup === "function") { try { await syncProjectPulseGroup(proj); } catch (e) { console.warn("project group sync", e); } }
    closeModal();
    toast("Project team updated", "success");
    if (onDone) onDone();
  });
}

function openProjectPersonModal(proj, editing, onDone) {
  const db = window.AEWTTR.db;
  if (!db.projectPeople) db.projectPeople = {};
  if (!db.projectPeople[proj.id]) db.projectPeople[proj.id] = [];
  const isEdit = !!editing;
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>${isEdit ? "Edit Project Person" : "Add Project Person"}</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body">
      <div class="form-row"><label>Person Type</label><select class="select-aewttr" id="pp-type" ${isEdit && editing.type === "member" ? "disabled" : ""}><option value="member">Internal Team Member</option><option value="person">Contractor Name</option><option value="company">Company</option></select></div>
      <div class="form-row" id="pp-member-row"><label>Team Member</label><select class="select-aewttr" id="pp-member">${db.members.map(member => `<option value="${member.id}" ${isEdit && editing.memberId === member.id ? "selected" : ""}>${escapeHtml(member.name)} — ${escapeHtml(member.role)}</option>`).join("")}</select></div>
      <div class="form-row" id="pp-label-row" style="display:none;"><label>Name / Company</label><input class="input-aewttr" id="pp-label" value="${escapeHtml((isEdit && editing.type !== "member" && editing.label) || "")}"></div>
      <div class="form-row" id="pp-email-row" style="display:none;"><label>Email <small style="font-weight:400;color:var(--aewttr-muted);">(optional)</small></label><input class="input-aewttr" id="pp-email" placeholder="name@example.com" value="${escapeHtml((isEdit && editing.email) || "")}"></div>
      <div class="form-row" id="pp-company-row" style="display:none;"><label>Company</label><input class="input-aewttr" id="pp-company" value="${escapeHtml((isEdit && editing.company) || "")}"></div>
      <div class="form-row"><label>Role</label>
        <select class="select-aewttr" id="pp-role">
          ${(isEdit && editing.role && !["Employee", "Contractor", "Guest", "Product Manager", "Engineer", "Developer", "ISSO", "Range POC"].includes(editing.role)) ? `<option value="${escapeHtml(editing.role)}">${escapeHtml(editing.role)}</option>` : (!isEdit ? `<option value="Employee" selected>Employee</option>` : "")}
          ${["Employee", "Contractor", "Guest", "Product Manager", "Engineer", "Developer", "ISSO", "Range POC"].map(r => `<option ${isEdit && editing.role === r ? "selected" : ""}>${r}</option>`).join("")}
        </select>
      </div>
    </div>
    <div class="aewttr-modal-foot">
      <button class="btn-aewttr-ghost" id="pp-cancel">Cancel</button>
      <button class="btn-aewttr" id="pp-save">${isEdit ? "Save Changes" : "Add To Project"}</button>
    </div>
  `);

  function syncForm() {
    const type = $("#pp-type", modal).value;
    $("#pp-member-row", modal).style.display = type === "member" ? "" : "none";
    $("#pp-label-row", modal).style.display = type === "member" ? "none" : "";
    $("#pp-email-row", modal).style.display = type === "member" ? "none" : "";
    $("#pp-company-row", modal).style.display = type === "person" ? "" : "none";
  }
  if (isEdit) $("#pp-type", modal).value = editing.type;
  syncForm();
  $("#pp-type", modal).addEventListener("change", syncForm);
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#pp-cancel", modal).addEventListener("click", closeModal);
  $("#pp-save", modal).addEventListener("click", async () => {
    const type = $("#pp-type", modal).value;
    let updates;
    if (type === "member") {
      const memberId = $("#pp-member", modal).value;
      const member = db.members.find(item => item.id === memberId);
      if (!member) return;
      updates = { type, memberId, label: member.name, role: $("#pp-role", modal).value.trim() || member.role, company: "", email: member.email || "" };
    } else {
      const label = $("#pp-label", modal).value.trim();
      if (!label) { toast("Name or company is required", "error"); return; }
      updates = {
        type, memberId: "", label,
        role: $("#pp-role", modal).value.trim() || (type === "company" ? "Company" : "Contractor"),
        company: type === "company" ? label : $("#pp-company", modal).value.trim(),
        email: $("#pp-email", modal).value.trim()
      };
    }
    if (isEdit) {
      Object.assign(editing, updates);
    } else {
      db.projectPeople[proj.id].push({ id: uid("ppj"), ...updates });
    }
    if (db.weeklyMeeting && db.weeklyMeeting.projectMeetings && db.weeklyMeeting.projectMeetings[proj.id]) {
      db.weeklyMeeting.projectMeetings[proj.id].participants = db.projectPeople[proj.id].map(item => ({ ...item }));
    }
    await Repo.save("project", proj);
    if (typeof syncProjectPulseGroup === "function") {
      try { await syncProjectPulseGroup(proj); } catch (e) { console.warn("project group sync", e); }
    }
    closeModal();
    toast(isEdit ? "Project person updated" : "Project person added", "success");
    if (onDone) onDone();
  });
}

/* PM/Engineer/ISSO/Range POC/Contractor are stored as a projectPeople entry
   id (proj.pm etc.) rather than free text, so "who is the PM" always points
   at one roster row with a real role/email — editable from the People tab
   the same way for a site user or a manually-added contact. */
const ASSIGNABLE_PROJECT_ROLE_FIELDS = [
  { key: "pm", label: "PM" },
  { key: "engineer", label: "Engineer" },
  { key: "isso", label: "ISSO" },
  { key: "rangePoc", label: "Range POC" },
  { key: "contractor", label: "Contractor" }
];

function findOrCreateProjectPersonEntry(proj, { type, memberId, label, email, company, role }) {
  const db = window.AEWTTR.db;
  if (!db.projectPeople) db.projectPeople = {};
  if (!db.projectPeople[proj.id]) db.projectPeople[proj.id] = [];
  const roster = db.projectPeople[proj.id];
  let existing = null;
  if (type === "member" && memberId) existing = roster.find((p) => p.memberId === memberId);
  if (!existing && email) existing = roster.find((p) => p.email && p.email.toLowerCase() === email.toLowerCase());
  if (!existing && label) {
    const key = String(label).toLowerCase();
    existing = roster.find((p) => p.label && p.label.toLowerCase() === key)
      || roster.find((p) => !p.memberId && p.label && p.label.toLowerCase() === key);
  }
  if (existing) {
    if (role && !existing.role) existing.role = role;
    return existing;
  }
  const entry = { id: uid("ppj"), type, memberId: memberId || "", label, role: role || "", company: company || "", email: email || "" };
  roster.push(entry);
  return entry;
}

/* When someone is assigned as a key/divider role or task/subtask owner,
   also land them on the project people roster (idempotent). New roster
   rows trigger the existing project auto-group sync. */
function resolveProjectRef(projOrId) {
  if (!projOrId) return null;
  if (typeof projOrId === "object" && projOrId.id) return projOrId;
  const db = window.AEWTTR.db;
  return (db.projects || []).find((p) => p.id === projOrId) || null;
}

async function ensureAssigneeOnProjectPeople(projOrId, assignee, opts) {
  opts = opts || {};
  const name = String(
    typeof assignee === "string"
      ? assignee
      : (assignee && (assignee.name || assignee.label)) || ""
  ).trim();
  if (!name || /^unassigned$/i.test(name)) return null;
  const proj = resolveProjectRef(projOrId);
  if (!proj) return null;
  const db = window.AEWTTR.db;
  let member = null;
  if (typeof assignee === "object" && assignee) {
    if (assignee.memberId) member = (db.members || []).find((m) => m.id === assignee.memberId) || null;
    if (!member && assignee.email) {
      const email = String(assignee.email).toLowerCase();
      member = (db.members || []).find((m) => m.email && String(m.email).toLowerCase() === email) || null;
    }
  }
  if (!member && typeof memberMatchesAssignee === "function") {
    member = (db.members || []).find((m) => memberMatchesAssignee(m, name)) || null;
  }
  if (!member) {
    member = (db.members || []).find((m) => m.name && m.name.toLowerCase() === name.toLowerCase()) || null;
  }
  const rosterBefore = (db.projectPeople && db.projectPeople[proj.id]) || [];
  const beforeIds = new Set(rosterBefore.map((p) => p.id));
  const entry = findOrCreateProjectPersonEntry(proj, {
    type: member ? "member" : "person",
    memberId: member ? member.id : "",
    label: member ? member.name : name,
    email: member ? (member.email || "") : ((typeof assignee === "object" && assignee && assignee.email) || ""),
    company: (typeof assignee === "object" && assignee && assignee.company) || "",
    role: opts.role || ""
  });
  const added = entry && !beforeIds.has(entry.id);
  if ((added || opts.forceSync) && typeof syncProjectPulseGroup === "function") {
    try { await syncProjectPulseGroup(proj); } catch (e) { console.warn("project group sync", e); }
  }
  return entry;
}

function ensureAssigneesFromTask(projOrId, task) {
  if (!task || isTrackerDivider(task)) return Promise.resolve();
  const jobs = [];
  if (task.assignee) jobs.push(ensureAssigneeOnProjectPeople(projOrId, task.assignee));
  if (typeof walkNestedSubtasks === "function") {
    walkNestedSubtasks(task.subtasks || [], (sub) => {
      if (sub && sub.assignee) jobs.push(ensureAssigneeOnProjectPeople(projOrId, sub.assignee));
    });
  } else {
    (task.subtasks || []).forEach((sub) => {
      if (sub && sub.assignee) jobs.push(ensureAssigneeOnProjectPeople(projOrId, sub.assignee));
    });
  }
  return Promise.all(jobs);
}

function ensureDividerRolesOnProjectPeople(projOrId, divider) {
  if (!divider) return Promise.resolve();
  const jobs = [];
  ASSIGNABLE_PROJECT_ROLE_FIELDS.forEach((f) => {
    const value = divider[f.key] || (divider.metadata && divider.metadata[f.key]) || "";
    if (typeof value === "string" && value.trim()) {
      jobs.push(ensureAssigneeOnProjectPeople(projOrId, value.trim(), { role: f.label }));
    }
  });
  const contractorRaw = divider.contractor || (divider.metadata && divider.metadata.contractor) || "";
  const contractors = typeof parseContractorList === "function"
    ? parseContractorList(contractorRaw)
    : String(contractorRaw).split(/[;,]/).map((s) => s.trim()).filter(Boolean);
  contractors.forEach((name) => {
    jobs.push(ensureAssigneeOnProjectPeople(projOrId, name, { role: "Contractor" }));
  });
  return Promise.all(jobs);
}

function resolvedAssignedPerson(proj, fieldKey) {
  const id = proj[fieldKey];
  if (!id) return null;
  const roster = (window.AEWTTR.db.projectPeople && window.AEWTTR.db.projectPeople[proj.id]) || [];
  return roster.find((p) => p.id === id) || null;
}

async function notifyTaskAssignee(task, projectCode) {
  if (typeof isSharePointMode !== "function" || !isSharePointMode()) return;
  if (!task || !task.assignee || task.assignee === "Unassigned") return;
  const db = window.AEWTTR.db;
  const member = (db.members || []).find((m) => memberMatchesAssignee(m, task.assignee));
  if (!member || !member.email) return;
  if (String(member.email).toLowerCase() === String(db.user.email || "").toLowerCase()) return;
  const proj = (db.projects || []).find((p) => p.id === projectCode);
  try {
    await notifyUsers({
      to: [member.email],
      subject: `PULSE Projects: "${task.title}" assigned to you${proj ? ` — ${proj.id}` : ""}`,
      area: "Projects",
      kind: "action",
      preview: `${db.user.name} assigned you a task${proj ? ` on ${proj.id} — ${proj.name}` : ""}.`,
      facts: [
        { title: "Task", value: task.title },
        { title: "Project", value: proj ? `${proj.id} — ${proj.name}` : (projectCode || "—") },
        { title: "Due", value: task.end ? fmtDate(task.end) : "—" },
        { title: "Assigned by", value: db.user.name }
      ],
      actionUrl: projectTaskDeepLinkActionUrl(projectCode, task),
      actionTitle: "Open Task"
    });
  } catch (e) {
    console.warn("PULSE: task assignment notification failed.", e);
  }
}

function projectTaskDeepLinkActionUrl(projectCode, task, subIndex) {
  if (!projectCode || typeof pulseAppRouteUrl !== "function") return pulseAppUrl();
  const query = {};
  if (task && task.id) query.task = task.id;
  if (subIndex != null) query.sub = subIndex;
  return pulseAppRouteUrl(`projects/${projectCode}/tracker`, query);
}

async function notifyPersonAssignedToProjectRole(proj, fieldLabel, entry) {
  if (typeof isSharePointMode !== "function" || !isSharePointMode() || !entry || !entry.email) return;
  const db = window.AEWTTR.db;
  // Don't notify people of their own assignment action.
  if (String(entry.email).toLowerCase() === String(db.user.email || "").toLowerCase()) return;
  try {
    await notifyUsers({
      to: [entry.email],
      subject: `PULSE Projects: You're the ${fieldLabel} on ${proj.id} — ${proj.name}`,
      area: "Projects",
      kind: "action",
      preview: `${db.user.name} assigned you as ${fieldLabel} on ${proj.id} — ${proj.name}.`,
      facts: [
        { title: "Project", value: `${proj.id} — ${proj.name}` },
        { title: "Role", value: fieldLabel },
        { title: "Portfolios", value: (typeof projectPortfolios === "function" ? projectPortfolios(proj) : (proj.portfolios || [])).join(", ") || "—" },
        { title: "Assigned by", value: db.user.name }
      ],
      actionUrl: typeof pulseAppRouteUrl === "function" ? pulseAppRouteUrl(`projects/${proj.id}`) : pulseAppUrl(),
      actionTitle: "Open Project"
    });
  } catch (e) {
    console.warn("PULSE: project role assignment notification failed.", e);
  }
}

function openAssignPersonModal(proj, fieldKey, fieldLabel, onDone) {
  const db = window.AEWTTR.db;
  if (!db.projectPeople) db.projectPeople = {};
  if (!db.projectPeople[proj.id]) db.projectPeople[proj.id] = [];
  const roster = db.projectPeople[proj.id];
  const currentEntry = resolvedAssignedPerson(proj, fieldKey);
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>Assign ${escapeHtml(fieldLabel)}</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body">
      <div id="ap-search-view">
        <div class="form-row">
          <label>Search site users or this project's people</label>
          <input class="input-aewttr" id="ap-search-input" placeholder="Type a name or email...">
        </div>
        <div id="ap-suggestions" class="traveler-suggestions" style="position:static;box-shadow:none;max-height:260px;"></div>
        <button type="button" class="btn-aewttr-outline btn-aewttr-sm" id="ap-new-person-btn" style="margin-top:10px;"><i class="bx bx-user-plus"></i> New Person (not in the site)</button>
      </div>
      <div id="ap-new-view" style="display:none;">
        <div class="form-row"><label>Name</label><input class="input-aewttr" id="ap-new-name" placeholder="Full name"></div>
        <div class="form-row"><label>Email <small style="font-weight:400;color:var(--aewttr-muted);">(optional)</small></label><input class="input-aewttr" id="ap-new-email" placeholder="name@example.com"></div>
        <div class="form-row"><label>Company <small style="font-weight:400;color:var(--aewttr-muted);">(optional)</small></label><input class="input-aewttr" id="ap-new-company" placeholder="e.g. Northrop Grumman"></div>
        <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="ap-back-to-search"><i class="bx bx-chevron-left"></i> Back to search</button>
      </div>
    </div>
    <div class="aewttr-modal-foot">
      ${currentEntry ? `<button class="btn-aewttr-ghost" id="ap-clear" style="margin-right:auto;">Clear assignment</button>` : ""}
      <button class="btn-aewttr-ghost" id="ap-cancel">Cancel</button>
      <button class="btn-aewttr" id="ap-save-new" style="display:none;"><i class="bx bx-check"></i> Add &amp; Assign</button>
    </div>
  `);

  async function finishAssign(entry) {
    proj[fieldKey] = entry.id;
    await Repo.save("project", proj);
    if (typeof syncProjectPulseGroup === "function") {
      try { await syncProjectPulseGroup(proj); } catch (e) { console.warn("project group sync", e); }
    }
    notifyPersonAssignedToProjectRole(proj, fieldLabel, entry);
    closeModal();
    toast(`${fieldLabel} assigned`, "success");
    if (onDone) onDone();
  }

  function candidateList() {
    const rosterMemberIds = new Set(roster.filter((p) => p.memberId).map((p) => p.memberId));
    const rosterCandidates = roster.map((p) => ({ source: "roster", personId: p.id, name: p.label, email: p.email || "", sub: p.company || (p.type === "member" ? "Team member" : "Project contact") }));
    const siteUsers = getMemberDirectory()
      .filter((m) => !rosterMemberIds.has(m.id))
      .map((m) => ({ source: "member", memberId: m.id, name: m.name, email: m.email || "", sub: "Site user" }));
    return rosterCandidates.concat(siteUsers);
  }

  function drawSuggestions() {
    const query = ($("#ap-search-input", modal).value || "").trim().toLowerCase();
    const suggestions = $("#ap-suggestions", modal);
    const all = candidateList();
    const matches = (query ? all.filter((c) => c.name.toLowerCase().includes(query) || c.email.toLowerCase().includes(query)) : all).slice(0, 20);
    suggestions.innerHTML = matches.length ? matches.map((c, i) => `
      <button type="button" class="traveler-suggestion" data-cand-index="${i}">
        <strong>${escapeHtml(c.name)}</strong>
        <span>${escapeHtml(c.email || c.sub || "No email on file")}</span>
      </button>`).join("") : `<div class="traveler-suggestion-empty">No matches. Use "New Person" below if they're not a site user.</div>`;
    $all("[data-cand-index]", suggestions).forEach((btn) => btn.addEventListener("click", async () => {
      const cand = matches[Number(btn.dataset.candIndex)];
      if (cand.source === "roster") {
        const entry = roster.find((p) => p.id === cand.personId);
        if (entry) await finishAssign(entry);
      } else {
        const entry = findOrCreateProjectPersonEntry(proj, { type: "member", memberId: cand.memberId, label: cand.name, email: cand.email, role: fieldLabel });
        await finishAssign(entry);
      }
    }));
  }
  drawSuggestions();
  $("#ap-search-input", modal).addEventListener("input", drawSuggestions);

  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#ap-cancel", modal).addEventListener("click", closeModal);
  const clearBtn = $("#ap-clear", modal);
  if (clearBtn) clearBtn.addEventListener("click", async () => {
    proj[fieldKey] = "";
    Repo.save("project", proj);
    closeModal();
    toast(`${fieldLabel} assignment cleared`, "success");
    if (onDone) onDone();
  });

  const searchView = $("#ap-search-view", modal);
  const newView = $("#ap-new-view", modal);
  const saveNewBtn = $("#ap-save-new", modal);
  $("#ap-new-person-btn", modal).addEventListener("click", () => {
    searchView.style.display = "none";
    newView.style.display = "";
    saveNewBtn.style.display = "";
    $("#ap-new-name", modal).focus();
  });
  $("#ap-back-to-search", modal).addEventListener("click", () => {
    newView.style.display = "none";
    searchView.style.display = "";
    saveNewBtn.style.display = "none";
  });
  saveNewBtn.addEventListener("click", async () => {
    const name = $("#ap-new-name", modal).value.trim();
    if (!name) { toast("Enter a name.", "error"); return; }
    const email = $("#ap-new-email", modal).value.trim();
    const company = $("#ap-new-company", modal).value.trim();
    const entry = findOrCreateProjectPersonEntry(proj, { type: "person", memberId: "", label: name, email, company, role: fieldLabel });
    await finishAssign(entry);
  });
}

/* ---------- Data Import ---------- */
const DOC_IMPORT_TYPES = {
  contracts: {
    label: "Contracts",
    applyLabel: "Finance + contract details",
    parsedKeys: [
      { key: "contractNumber", label: "Contract Number" },
      { key: "contractor", label: "Contractor" },
      { key: "clin", label: "Primary CLIN" },
      { key: "fundedValue", label: "Funded Value" },
      { key: "period", label: "Period Of Performance" }
    ],
    sample(fileName) {
      return {
        contractNumber: "N00019-26-F-2184",
        contractor: "Northrop Grumman",
        clin: "0004AA",
        fundedValue: "$1.85M",
        period: "2026-08-01 to 2027-07-31",
        sourceFile: fileName
      };
    }
  },
  invoices: {
    label: "Invoices",
    applyLabel: "Finance invoice queue",
    parsedKeys: [
      { key: "invoiceNo", label: "Invoice Number" },
      { key: "vendor", label: "Vendor" },
      { key: "amount", label: "Amount" },
      { key: "received", label: "Received Date" },
      { key: "due", label: "Due Date" }
    ],
    sample(fileName) {
      return {
        invoiceNo: "INV-" + String(fileName.length).padStart(4, "0"),
        vendor: "BAE Systems",
        amount: "$84,250",
        received: "2026-06-26",
        due: "2026-07-10",
        sourceFile: fileName
      };
    }
  },
  communication: {
    label: "Communication",
    applyLabel: "Communication log + follow-up fields",
    parsedKeys: [
      { key: "subject", label: "Subject" },
      { key: "counterparty", label: "Counterparty" },
      { key: "date", label: "Date" },
      { key: "action", label: "Action / Follow-Up" }
    ],
    sample(fileName) {
      return {
        subject: "Contractor funding alignment note",
        counterparty: "Northrop Grumman",
        date: "2026-06-26",
        action: `Pulled key action items from ${fileName} for manual entry and follow-up tracking.`,
        sourceFile: fileName
      };
    }
  }
};

function financeFmt(value) {
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 }).format(Number(value || 0));
}

function financeNumber(value) {
  return Number(String(value || "").replace(/[^0-9.-]/g, "")) || 0;
}

function projectImportDb(projectId) {
  const db = window.AEWTTR.db;
  if (!db.projectImports) db.projectImports = {};
  if (!db.projectImports[projectId]) db.projectImports[projectId] = { docs: [] };
  return db.projectImports[projectId];
}

function projectFinanceDb(projectId) {
  const db = window.AEWTTR.db;
  if (!db.projectFinance) db.projectFinance = {};
  if (!db.projectFinance[projectId]) {
    db.projectFinance[projectId] = {
      summary: {},
      snapshot: { totalCeiling: 0, fundedToDate: 0, obligated: 0, accrued: 0, actuals: 0, forecastAtComplete: 0 },
      clins: [],
      invoices: [],
      watchlist: [],
      communications: []
    };
  }
  return db.projectFinance[projectId];
}

function drawImport(body, proj) {
  const importDb = projectImportDb(proj.id);

  body.innerHTML = `
    <div class="aewttr-dash-grid">
      <div>
        <div class="aewttr-card aewttr-card-pad" style="margin-bottom:16px;">
          <div class="toolbar-row" style="margin-bottom:14px;">
            <div>
              <div class="side-panel-title" style="margin:0;">Document Intake</div>
              <p style="font-size:12px;color:var(--aewttr-muted);margin:6px 0 0;">Upload a source document for project access, parse it, then manually review the suggested fields before they touch project data.</p>
            </div>
            <button class="btn-aewttr" id="btn-open-import-flow"${tip("Start document intake for this project")}><i class="bx bx-upload"></i> Start Intake</button>
          </div>
          <div class="timeline-item"><div class="tdate">1</div><div>Pick the document type: contracts, invoices, or communication.</div></div>
          <div class="timeline-item"><div class="tdate">2</div><div>Upload the document so the team has quick access to the source file in the project.</div></div>
          <div class="timeline-item"><div class="tdate">3</div><div>Review the parsed suggestions and manually apply only the fields that make sense for this project.</div></div>
        </div>
        <div class="aewttr-card aewttr-card-pad">
          <div class="side-panel-title">Imported Project Documents</div>
          ${importDb.docs.length ? `
            <table class="aewttr-table">
              <thead><tr><th>Type</th><th>File</th><th>Uploaded</th><th>Suggested Data</th><th>Status</th><th></th></tr></thead>
              <tbody>
                ${importDb.docs.map(doc => `
                  <tr style="cursor:default;">
                    <td>${escapeHtml((DOC_IMPORT_TYPES[doc.type] || {}).label || doc.type)}</td>
                    <td><strong>${escapeHtml(doc.fileName)}</strong><div style="font-size:11.5px;color:var(--aewttr-muted);">Stored for project access</div></td>
                    <td>${fmtDate(doc.uploadedOn)} · ${escapeHtml(doc.uploadedBy)}</td>
                    <td>${Object.entries(doc.parsed || {}).filter(([key]) => key !== "sourceFile").slice(0, 2).map(([, value]) => escapeHtml(String(value))).join(" · ") || "No suggestions"}</td>
                    <td>${doc.applied ? `<span class="kc-badge">Applied</span>` : `<span class="kc-badge">Needs Review</span>`}</td>
                    <td><button class="btn-aewttr-outline btn-aewttr-sm" data-review-import="${doc.id}"${tip("Review imported document fields")}>Review</button></td>
                  </tr>`).join("")}
              </tbody>
            </table>` : `<div class="empty-state">No project documents have been imported yet.</div>`}
        </div>
      </div>
      <div>
        <div class="aewttr-card aewttr-card-pad">
          <div class="side-panel-title">How This Works</div>
          <div class="timeline-item"><div class="tdate">Access</div><div>The uploaded file lives here for easy project access even if no data gets applied.</div></div>
          <div class="timeline-item"><div class="tdate">Parsed</div><div>The system proposes fields based on doc type, but it does not auto-commit project finance or contract data.</div></div>
          <div class="timeline-item"><div class="tdate">Manual</div><div>PMs review the suggestions, then apply them into Finance or the communication log one intake at a time.</div></div>
        </div>
      </div>
    </div>
  `;

  $("#btn-open-import-flow", body).addEventListener("click", () => openProjectImportModal(proj, () => drawImport(body, proj)));
  $all("[data-review-import]", body).forEach(btn => btn.addEventListener("click", () => openProjectImportReviewModal(proj, btn.dataset.reviewImport, () => drawImport(body, proj))));
}

function openProjectImportModal(proj, onDone) {
  const importDb = projectImportDb(proj.id);
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>Start Document Intake</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body">
      <div class="form-row"><label>Document Type</label>
        <select class="select-aewttr" id="pi-type">
          ${Object.entries(DOC_IMPORT_TYPES).map(([key, spec]) => `<option value="${key}">${escapeHtml(spec.label)}</option>`).join("")}
        </select>
      </div>
      <div class="form-row"><label>Upload Document</label>
        <div class="dropzone" id="pi-dropzone">
          <i class="bx bx-cloud-upload"></i>
          Upload the project document for access and parsing
          <input type="file" id="pi-file" accept=".pdf,.doc,.docx,.xlsx,.csv,.msg,.txt" style="display:none;">
        </div>
        <div id="pi-file-label" style="font-size:12px;color:var(--aewttr-muted);margin-top:8px;">No file selected yet.</div>
      </div>
      <div class="aewttr-card" style="display:none;padding:14px;" id="pi-preview"></div>
    </div>
    <div class="aewttr-modal-foot">
      <button class="btn-aewttr-ghost" id="pi-cancel">Cancel</button>
      <button class="btn-aewttr" id="pi-save">Save Intake</button>
    </div>
  `, { wide: true });

  let fileName = "";
  let type = "contracts";
  let parsed = null;

  function drawPreview() {
    const preview = $("#pi-preview", modal);
    if (!parsed) {
      preview.style.display = "none";
      return;
    }
    preview.style.display = "block";
    preview.innerHTML = `
      <div class="side-panel-title">Parsed Suggestions</div>
      <p style="font-size:12px;color:var(--aewttr-muted);margin-top:0;">These are suggestions only. The uploaded file will be saved for quick access, and the fields can be reviewed manually before anything is applied.</p>
      <div class="kv-list">
        ${DOC_IMPORT_TYPES[type].parsedKeys.map(field => `<div class="kv-row"><span class="k">${escapeHtml(field.label)}</span><span class="v">${escapeHtml(String(parsed[field.key] || "—"))}</span></div>`).join("")}
      </div>
    `;
  }

  function readFile(file) {
    if (!file) return;
    fileName = file.name;
    $("#pi-file-label", modal).textContent = fileName;
    parsed = DOC_IMPORT_TYPES[type].sample(fileName);
    drawPreview();
  }

  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#pi-cancel", modal).addEventListener("click", closeModal);
  $("#pi-type", modal).addEventListener("change", e => {
    type = e.target.value;
    if (fileName) parsed = DOC_IMPORT_TYPES[type].sample(fileName);
    drawPreview();
  });
  $("#pi-dropzone", modal).addEventListener("click", () => $("#pi-file", modal).click());
  ["dragover", "dragenter"].forEach(ev => $("#pi-dropzone", modal).addEventListener(ev, e => { e.preventDefault(); $("#pi-dropzone", modal).classList.add("drag-active"); }));
  ["dragleave", "drop"].forEach(ev => $("#pi-dropzone", modal).addEventListener(ev, e => { e.preventDefault(); $("#pi-dropzone", modal).classList.remove("drag-active"); }));
  $("#pi-dropzone", modal).addEventListener("drop", e => readFile(e.dataTransfer.files[0]));
  $("#pi-file", modal).addEventListener("change", e => readFile(e.target.files[0]));
  $("#pi-save", modal).addEventListener("click", () => {
    if (!fileName || !parsed) { toast("Upload a document first", "error"); return; }
    importDb.docs.unshift({
      id: uid("imp"),
      type,
      fileName,
      uploadedOn: new Date().toISOString().slice(0, 10),
      uploadedBy: window.AEWTTR.db.user.name,
      accessOnly: true,
      parsed,
      applied: false
    });
    aewttrSaveStore();
    closeModal();
    toast("Project document intake saved", "success");
    if (onDone) onDone();
  });
}

function openProjectImportReviewModal(proj, docId, onDone) {
  const importDb = projectImportDb(proj.id);
  const finance = projectFinanceDb(proj.id);
  const doc = (importDb.docs || []).find(item => item.id === docId);
  if (!doc) return;
  const spec = DOC_IMPORT_TYPES[doc.type];
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>Review Imported ${escapeHtml(spec.label)}</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body">
      <p style="font-size:12px;color:var(--aewttr-muted);margin-top:0;">The source document stays attached to the project for access. Review and edit the parsed values below before applying them.</p>
      ${spec.parsedKeys.map(field => `<div class="form-row"><label>${escapeHtml(field.label)}</label><input class="input-aewttr" data-import-field="${field.key}" value="${escapeHtml(String(doc.parsed[field.key] || ""))}"></div>`).join("")}
    </div>
    <div class="aewttr-modal-foot">
      <button class="btn-aewttr-ghost" id="pir-close">Close</button>
      <button class="btn-aewttr" id="pir-apply">Apply ${escapeHtml(spec.applyLabel)}</button>
    </div>
  `, { wide: true });
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#pir-close", modal).addEventListener("click", closeModal);
  $("#pir-apply", modal).addEventListener("click", () => {
    const values = {};
    spec.parsedKeys.forEach(field => { values[field.key] = $(`[data-import-field="${field.key}"]`, modal).value.trim(); });
    doc.parsed = { ...doc.parsed, ...values };
    if (doc.type === "contracts") {
      finance.summary.contractNumber = values.contractNumber;
      finance.summary.contractor = values.contractor;
      finance.summary.popStart = (values.period || "").split(" to ")[0] || finance.summary.popStart || "";
      finance.summary.popEnd = (values.period || "").split(" to ")[1] || finance.summary.popEnd || "";
      if (values.fundedValue) finance.snapshot.fundedToDate = financeNumber(values.fundedValue);
      if (values.clin) {
        const existingClin = (finance.clins || []).find(clin => clin.code === values.clin);
        if (!existingClin) (finance.clins || (finance.clins = [])).unshift({ id: uid("clin"), code: values.clin, title: "Imported contract line", type: "Contract", funded: finance.snapshot.fundedToDate || 0, obligated: 0, spent: 0, status: "Imported" });
      }
    }
    if (doc.type === "invoices") {
      finance.invoices.unshift({
        id: uid("inv"),
        invoiceNo: values.invoiceNo,
        vendor: values.vendor,
        amount: financeNumber(values.amount),
        received: values.received,
        due: values.due,
        status: "In Review"
      });
    }
    if (doc.type === "communication") {
      finance.communications.unshift({
        id: uid("fc"),
        subject: values.subject,
        counterparty: values.counterparty,
        date: values.date,
        action: values.action
      });
    }
    doc.applied = true;
    aewttrSaveStore();
    closeModal();
    toast("Parsed data applied after manual review", "success");
    if (onDone) onDone();
  });
}

function drawFinance(body, proj) {
  const finance = projectFinanceDb(proj.id);
  const remaining = (finance.snapshot.fundedToDate || 0) - (finance.snapshot.actuals || 0);
  const burnPct = finance.snapshot.fundedToDate ? Math.round(((finance.snapshot.actuals || 0) / finance.snapshot.fundedToDate) * 100) : 0;
  body.innerHTML = `
    <div class="stat-grid" style="margin-bottom:18px;">
      <div class="stat-card"><div class="label">Funded To Date</div><div class="value">${financeFmt(finance.snapshot.fundedToDate)}</div><div class="desc">Current funding actually available to the project.</div></div>
      <div class="stat-card"><div class="label">Actuals</div><div class="value">${financeFmt(finance.snapshot.actuals)}</div><div class="desc">Current spend recorded against project funding.</div></div>
      <div class="stat-card"><div class="label">Remaining</div><div class="value">${financeFmt(remaining)}</div><div class="desc">Available headroom before current funding is exhausted.</div></div>
      <div class="stat-card"><div class="label">Burn</div><div class="value">${burnPct}%</div><div class="desc">Approximate burn against funded dollars.</div></div>
    </div>
    <div class="aewttr-dash-grid">
      <div>
        <div class="aewttr-card aewttr-card-pad" style="margin-bottom:16px;">
          <div class="side-panel-title">Contract & Funding Summary</div>
          <div class="kv-list">
            <div class="kv-row"><span class="k">Appropriation</span><span class="v">${escapeHtml(finance.summary.appropriation || "—")}</span></div>
            <div class="kv-row"><span class="k">Sponsor</span><span class="v">${escapeHtml(finance.summary.sponsor || "—")}</span></div>
            <div class="kv-row"><span class="k">Charge Object</span><span class="v">${escapeHtml(finance.summary.chargeObject || "—")}</span></div>
            <div class="kv-row"><span class="k">Contract Number</span><span class="v">${escapeHtml(finance.summary.contractNumber || "—")}</span></div>
            <div class="kv-row"><span class="k">Contractor</span><span class="v">${escapeHtml(finance.summary.contractor || "—")}</span></div>
            <div class="kv-row"><span class="k">Funding Type</span><span class="v">${escapeHtml(finance.summary.fundingType || "—")}</span></div>
            <div class="kv-row"><span class="k">POP</span><span class="v">${finance.summary.popStart ? `${fmtDate(finance.summary.popStart)} – ${fmtDate(finance.summary.popEnd)}` : "—"}</span></div>
          </div>
        </div>
        <div class="aewttr-card aewttr-card-pad" style="margin-bottom:16px;">
          <div class="side-panel-title">CLIN / Line Funding</div>
          <table class="aewttr-table">
            <thead><tr><th>CLIN</th><th>Title</th><th>Funded</th><th>Obligated</th><th>Spent</th><th>Status</th></tr></thead>
            <tbody>
              ${finance.clins.length ? finance.clins.map(clin => `<tr style="cursor:default;"><td>${escapeHtml(clin.code)}</td><td>${escapeHtml(clin.title)}</td><td>${financeFmt(clin.funded)}</td><td>${financeFmt(clin.obligated)}</td><td>${financeFmt(clin.spent)}</td><td>${escapeHtml(clin.status)}</td></tr>`).join("") : `<tr style="cursor:default;"><td colspan="6"><div class="empty-state">No CLIN data yet.</div></td></tr>`}
            </tbody>
          </table>
        </div>
        <div class="aewttr-card aewttr-card-pad">
          <div class="side-panel-title">Invoices</div>
          <table class="aewttr-table">
            <thead><tr><th>Invoice</th><th>Vendor</th><th>Amount</th><th>Received</th><th>Due</th><th>Status</th></tr></thead>
            <tbody>
              ${finance.invoices.length ? finance.invoices.map(invoice => `<tr style="cursor:default;"><td>${escapeHtml(invoice.invoiceNo)}</td><td>${escapeHtml(invoice.vendor)}</td><td>${financeFmt(invoice.amount)}</td><td>${fmtDate(invoice.received)}</td><td>${fmtDate(invoice.due)}</td><td>${escapeHtml(invoice.status)}</td></tr>`).join("") : `<tr style="cursor:default;"><td colspan="6"><div class="empty-state">No invoices logged yet.</div></td></tr>`}
            </tbody>
          </table>
        </div>
      </div>
      <div>
        <div class="aewttr-card aewttr-card-pad" style="margin-bottom:16px;">
          <div class="side-panel-title">Financial Snapshot</div>
          <div class="kv-list">
            <div class="kv-row"><span class="k">Total Ceiling</span><span class="v">${financeFmt(finance.snapshot.totalCeiling)}</span></div>
            <div class="kv-row"><span class="k">Funded To Date</span><span class="v">${financeFmt(finance.snapshot.fundedToDate)}</span></div>
            <div class="kv-row"><span class="k">Obligated</span><span class="v">${financeFmt(finance.snapshot.obligated)}</span></div>
            <div class="kv-row"><span class="k">Accrued</span><span class="v">${financeFmt(finance.snapshot.accrued)}</span></div>
            <div class="kv-row"><span class="k">Actuals</span><span class="v">${financeFmt(finance.snapshot.actuals)}</span></div>
            <div class="kv-row"><span class="k">EAC</span><span class="v">${financeFmt(finance.snapshot.forecastAtComplete)}</span></div>
          </div>
        </div>
        <div class="aewttr-card aewttr-card-pad" style="margin-bottom:16px;">
          <div class="side-panel-title">Funding Watchlist</div>
          ${finance.watchlist.length ? finance.watchlist.map(item => `<div class="timeline-item"><div class="tdate">${fmtDate(item.due)}</div><div><strong>${escapeHtml(item.title)}</strong><br>${escapeHtml(item.owner)} · ${escapeHtml(item.status)} · ${escapeHtml(item.note)}</div></div>`).join("") : `<div class="empty-state">No funding watch items yet.</div>`}
        </div>
        <div class="aewttr-card aewttr-card-pad">
          <div class="side-panel-title">Finance Communications</div>
          ${finance.communications.length ? finance.communications.map(item => `<div class="timeline-item"><div class="tdate">${fmtDate(item.date)}</div><div><strong>${escapeHtml(item.subject)}</strong><br>${escapeHtml(item.counterparty)} — ${escapeHtml(item.action)}</div></div>`).join("") : `<div class="empty-state">No finance communications logged yet.</div>`}
        </div>
      </div>
    </div>
  `;
}

/* ---------- Tracker dividers (mini-projects) + unlimited nested subitems ---------- */

function trackerDividerCollapsedStore(stateKey) {
  if (!window.AEWTTR.state.trackerDividerCollapsed) window.AEWTTR.state.trackerDividerCollapsed = {};
  if (!window.AEWTTR.state.trackerDividerCollapsed[stateKey]) window.AEWTTR.state.trackerDividerCollapsed[stateKey] = {};
  return window.AEWTTR.state.trackerDividerCollapsed[stateKey];
}

function trackerSubitemExpandedStore(stateKey) {
  if (!window.AEWTTR.state.trackerSubitemExpanded) window.AEWTTR.state.trackerSubitemExpanded = {};
  if (!window.AEWTTR.state.trackerSubitemExpanded[stateKey]) window.AEWTTR.state.trackerSubitemExpanded[stateKey] = {};
  return window.AEWTTR.state.trackerSubitemExpanded[stateKey];
}

function syncDividerMetadata(divider) {
  if (!divider || !isTrackerDivider(divider)) return divider;
  const meta = typeof blankDividerMetadata === "function"
    ? blankDividerMetadata(Object.assign({}, divider.metadata || {}, divider))
    : (divider.metadata || {});
  divider.metadata = meta;
  divider.itemType = "divider";
  divider.workItemLevel = "Divider";
  divider.startDate = meta.startDate || "";
  divider.dueDate = meta.dueDate || "";
  divider.start = divider.startDate;
  divider.end = divider.dueDate;
  [
    "pm", "engineer", "isso", "rangePoc", "contract", "contractor",
    "taskOrder", "fundingType", "fiscalYear", "fundingStatus",
    "changeRequestRequired", "configEndItem", "program", "product",
    "locations", "portfolios", "priority", "rag", "lifecycleStatus",
    "scope", "objectives", "description", "completionDate", "isMilestone"
  ].forEach((key) => { divider[key] = meta[key]; });
  divider.isMilestone = !!meta.isMilestone;
  return divider;
}

function createTrackerDivider(partial) {
  const today = new Date().toISOString().slice(0, 10);
  const base = {
    id: uid("div"),
    itemType: "divider",
    workItemLevel: "Divider",
    title: (partial && partial.title) || "New project divider",
    assignee: "",
    health: "On Track",
    status: "Not Started",
    subtasks: [],
    notes: [],
    parentDividerId: "",
    startDate: today,
    dueDate: today
  };
  return syncDividerMetadata(Object.assign(base, partial || {}));
}

function trackerPlainTasks(tasks) {
  return (tasks || []).filter((t) => t && !isTrackerDivider(t));
}

function trackerDividers(tasks) {
  return (tasks || []).filter((t) => isTrackerDivider(t));
}

/** Group tasks under dividers while preserving array order of dividers.
 *  Ungrouped tasks (no parentDividerId, or orphaned id) stay in `ungrouped`. */
function groupTrackerItems(tasks) {
  const dividers = trackerDividers(tasks);
  const plain = trackerPlainTasks(tasks);
  const byDivider = new Map(dividers.map((d) => [d.id, []]));
  const ungrouped = [];
  plain.forEach((task) => {
    const pid = task.parentDividerId || "";
    if (pid && byDivider.has(pid)) byDivider.get(pid).push(task);
    else ungrouped.push(task);
  });
  return {
    dividers: dividers.map((d) => ({ divider: d, tasks: byDivider.get(d.id) || [] })),
    ungrouped
  };
}

function rebuildTrackerOrder(tasks, grouped) {
  const next = [];
  const dividers = (grouped && grouped.dividers) || [];
  dividers.forEach(({ divider, tasks: sectionTasks }) => {
    if (divider) next.push(divider);
    (sectionTasks || []).forEach((task) => {
      if (task) next.push(task);
    });
  });
  ((grouped && grouped.ungrouped) || []).forEach((task) => {
    if (task) next.push(task);
  });
  tasks.splice(0, tasks.length, ...next);
  tasks.forEach((item, index) => {
    if (item) item._sortOrder = index;
  });
  return tasks;
}

function reorderTrackerTask(tasks, draggedId, targetId) {
  if (!Array.isArray(tasks) || !draggedId || !targetId || draggedId === targetId) return null;
  const grouped = groupTrackerItems(tasks);
  const allSections = grouped.dividers.map((section) => ({
    key: section.divider && section.divider.id,
    dividerId: section.divider && section.divider.id,
    tasks: section.tasks
  })).concat([{ key: "__ungrouped__", dividerId: "", tasks: grouped.ungrouped }]);

  let moved = null;
  let sourceSection = null;
  let sourceIndex = -1;
  let targetSection = null;
  let targetIndex = -1;

  allSections.forEach((section) => {
    const items = section.tasks || [];
    const movedIdx = items.findIndex((item) => item && item.id === draggedId);
    if (movedIdx >= 0) {
      sourceSection = section;
      sourceIndex = movedIdx;
      moved = items[movedIdx];
    }
    const targetIdx = items.findIndex((item) => item && item.id === targetId);
    if (targetIdx >= 0) {
      targetSection = section;
      targetIndex = targetIdx;
    }
  });

  if (!moved || !sourceSection || !targetSection || targetIndex < 0) return null;

  sourceSection.tasks.splice(sourceIndex, 1);
  const adjustedTargetIndex = sourceSection === targetSection && sourceIndex < targetIndex
    ? targetIndex - 1
    : targetIndex;
  moved.parentDividerId = targetSection.dividerId || "";
  targetSection.tasks.splice(adjustedTargetIndex, 0, moved);
  rebuildTrackerOrder(tasks, grouped);
  return moved;
}

function reorderTrackerDivider(tasks, movedId, targetId) {
  if (!Array.isArray(tasks) || !movedId || !targetId || movedId === targetId) return false;
  const grouped = groupTrackerItems(tasks);
  const dividers = grouped.dividers || [];
  const from = dividers.findIndex((section) => section.divider && section.divider.id === movedId);
  const to = dividers.findIndex((section) => section.divider && section.divider.id === targetId);
  if (from < 0 || to < 0 || from === to) return false;
  const [section] = dividers.splice(from, 1);
  dividers.splice(to, 0, section);
  rebuildTrackerOrder(tasks, grouped);
  return true;
}

function countNestedSubitems(subs) {
  let n = 0;
  (subs || []).forEach((s) => {
    n += 1;
    n += countNestedSubitems(s.subtasks || s.children || []);
  });
  return n;
}

function walkNestedSubtasks(subs, visitor, pathPrefix) {
  const prefix = pathPrefix || [];
  (subs || []).forEach((sub, i) => {
    const path = prefix.concat(i);
    visitor(sub, path, path.join("."));
    walkNestedSubtasks(sub.subtasks || [], visitor, path);
  });
}

function getSubtaskAtPath(task, pathStr) {
  if (!task || pathStr == null || pathStr === "") return null;
  const parts = String(pathStr).split(".").map((p) => Number(p));
  let node = null;
  let list = task.subtasks || [];
  for (let i = 0; i < parts.length; i++) {
    node = list[parts[i]];
    if (!node) return null;
    list = node.subtasks || [];
  }
  return node;
}

function removeSubtaskAtPath(task, pathStr) {
  if (!task || pathStr == null || pathStr === "") return null;
  const parts = String(pathStr).split(".").map((p) => Number(p));
  if (!parts.length) return null;
  if (parts.length === 1) {
    const removed = (task.subtasks || []).splice(parts[0], 1)[0];
    return removed || null;
  }
  const parentPath = parts.slice(0, -1).join(".");
  const parent = getSubtaskAtPath(task, parentPath);
  if (!parent || !parent.subtasks) return null;
  const removed = parent.subtasks.splice(parts[parts.length - 1], 1)[0];
  return removed || null;
}

function addNestedSubtask(parentLike, seed) {
  const child = normalizeTaskSubtask(parentLike, seed || { text: "New subitem" });
  if (!parentLike.subtasks) parentLike.subtasks = [];
  parentLike.subtasks.push(child);
  return child;
}

function flattenSubtasksForProgress(subs) {
  const out = [];
  walkNestedSubtasks(subs, (sub) => out.push(sub));
  return out;
}

function dividerMetaChipHtml(divider) {
  const bits = [];
  if (divider.isMilestone) bits.push("Milestone");
  if (divider.taskOrder) bits.push(`TO ${divider.taskOrder}`);
  if (divider.fundingType) bits.push(divider.fundingType);
  if (divider.fundingStatus) bits.push(divider.fundingStatus);
  if (divider.fiscalYear) bits.push(divider.fiscalYear);
  const roles = [];
  if (divider.pm) roles.push("PM");
  if (divider.engineer) roles.push("Eng");
  if (divider.isso) roles.push("ISSO");
  if (roles.length) bits.push(roles.join(" · "));
  if (divider.startDate || divider.dueDate) {
    bits.push([divider.startDate ? fmtDate(divider.startDate) : "—", divider.dueDate ? fmtDate(divider.dueDate) : "—"].join(" → "));
  }
  return bits.slice(0, 4).map((b) => `<span class="tracker-divider-chip">${escapeHtml(b)}</span>`).join("");
}

function projectTrackerMetaStripHtml(proj) {
  if (!proj) return "";
  const roles = ASSIGNABLE_PROJECT_ROLE_FIELDS.map((f) => {
    const person = typeof resolvedAssignedPerson === "function" ? resolvedAssignedPerson(proj, f.key) : null;
    const label = person ? person.label : (proj[f.key] || "");
    return label ? `<span class="tracker-meta-chip"><em>${escapeHtml(f.label)}</em> ${escapeHtml(label)}</span>` : "";
  }).filter(Boolean).join("");
  const funding = [
    proj.taskOrder ? `TO ${proj.taskOrder}` : "",
    proj.fundingType || "",
    proj.fundingStatus || "",
    proj.fiscalYear || ""
  ].filter(Boolean).join(" · ");
  const schedule = (proj.startDate || proj.dueDate)
    ? `${proj.startDate ? fmtDate(proj.startDate) : "—"} → ${proj.dueDate ? fmtDate(proj.dueDate) : "—"}`
    : "";
  if (!roles && !funding && !schedule && !(proj.portfolios || []).length) return "";
  return `
    <div class="tracker-project-meta" aria-label="Project metadata">
      <div class="tracker-project-meta-main">
        ${roles || `<span class="tracker-meta-quiet">No key roles set</span>`}
        ${funding ? `<span class="tracker-meta-chip tracker-meta-chip--muted">${escapeHtml(funding)}</span>` : ""}
        ${schedule ? `<span class="tracker-meta-chip tracker-meta-chip--muted">${escapeHtml(schedule)}</span>` : ""}
        ${(proj.portfolios || []).slice(0, 3).map((p) => `<span class="tracker-meta-chip tracker-meta-chip--muted">${escapeHtml(p)}</span>`).join("")}
      </div>
      <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-open-project-settings${tip("Edit project settings")}>
        <i class="bx bx-cog"></i> Project fields
      </button>
    </div>`;
}

function openDividerSettingsModal(divider, proj, onDone, opts) {
  opts = opts || {};
  const isNew = !!(opts.isNew || divider._isDraft);
  syncDividerMetadata(divider);
  const meta = divider.metadata || blankDividerMetadata(divider);
  const selectedPortfolios = new Set(meta.portfolios || []);
  const selectedLocations = new Set(meta.locations || []);
  const selectedContractors = new Set(parseContractorList(meta.contractor));
  const selectedConfigEnd = new Set();
  const initialConfig = normalizeConfigEndItemName(meta.configEndItem || "");
  if (initialConfig) selectedConfigEnd.add(initialConfig);
  const roleFields = ASSIGNABLE_PROJECT_ROLE_FIELDS.filter((f) => f.key !== "contractor");
  const roleDraft = {};
  roleFields.forEach((f) => {
    roleDraft[f.key] = meta[f.key] || divider[f.key] || "";
  });

  function readFormState() {
    return {
      title: ($("#div-title", modal).value || "").trim(),
      priority: $("#div-priority", modal).value,
      rag: meta.rag || "Green",
      startDate: $("#div-start", modal).value,
      dueDate: $("#div-due", modal).value,
      completionDate: $("#div-complete", modal).value,
      lifecycleStatus: $("#div-lifecycle", modal).value,
      program: $("#div-program", modal).value.trim(),
      contract: $("#div-contract", modal).value.trim(),
      taskOrder: $("#div-taskorder", modal).value.trim(),
      fundingType: $("#div-fundingtype", modal).value.trim(),
      fiscalYear: $("#div-fiscalyear", modal).value.trim(),
      fundingStatus: $("#div-fundingstatus", modal).value.trim(),
      configEndItem: Array.from(selectedConfigEnd)[0] || "",
      contractor: Array.from(selectedContractors).join("; "),
      changeRequestRequired: (($(`input[name="div-crr"]:checked`, modal) || {}).value === "Yes"),
      portfolios: Array.from(selectedPortfolios),
      locations: Array.from(selectedLocations),
      roles: Object.fromEntries(roleFields.map((f) => [f.key, String(roleDraft[f.key] || "").trim()]))
    };
  }

  const baseline = {
    title: String(divider.title || "").trim() || "New project divider",
    priority: meta.priority || "",
    rag: meta.rag || "Green",
    startDate: meta.startDate || "",
    dueDate: meta.dueDate || "",
    completionDate: meta.completionDate || "",
    lifecycleStatus: meta.lifecycleStatus || "",
    program: meta.program || "",
    contract: meta.contract || "",
    taskOrder: meta.taskOrder || "",
    fundingType: meta.fundingType || "",
    fiscalYear: meta.fiscalYear || "",
    fundingStatus: meta.fundingStatus || "",
    configEndItem: initialConfig,
    contractor: Array.from(selectedContractors).join("; "),
    changeRequestRequired: !!meta.changeRequestRequired,
    portfolios: Array.from(selectedPortfolios),
    locations: Array.from(selectedLocations),
    roles: Object.fromEntries(roleFields.map((f) => [f.key, String(roleDraft[f.key] || "").trim()]))
  };

  function statesEqual(a, b) {
    return JSON.stringify(a) === JSON.stringify(b);
  }

  function hasMeaningfulChanges() {
    return !statesEqual(readFormState(), baseline);
  }

  let settled = false;
  let saveTimer = null;
  let persistedOnce = !isNew && !!divider._spId;

  function applyFormToDivider() {
    const state = readFormState();
    const contractors = Array.from(selectedContractors);
    const next = {
      description: "",
      priority: state.priority,
      rag: state.rag || "Green",
      startDate: state.startDate,
      dueDate: state.dueDate,
      completionDate: state.completionDate,
      lifecycleStatus: state.lifecycleStatus,
      program: state.program,
      contract: state.contract,
      taskOrder: state.taskOrder,
      fundingType: state.fundingType,
      fiscalYear: state.fiscalYear,
      fundingStatus: state.fundingStatus,
      configEndItem: state.configEndItem,
      contractor: contractors.join("; "),
      changeRequestRequired: state.changeRequestRequired,
      scope: "",
      objectives: "",
      portfolios: state.portfolios,
      locations: state.locations
    };
    roleFields.forEach((f) => { next[f.key] = state.roles[f.key] || ""; });
    next.isMilestone = !!divider.isMilestone;
    divider.title = state.title || (isNew ? "Untitled divider" : (divider.title || "Untitled divider"));
    Object.assign(divider, next);
    syncDividerMetadata(divider);
    rememberPortfolioNames(next.portfolios);
    rememberLocationNames(next.locations);
    rememberContractorNames(contractors);
    if (next.configEndItem) rememberConfigEndItemNames([next.configEndItem]);
    return state;
  }

  async function persistDivider() {
    applyFormToDivider();
    delete divider._isDraft;
    try {
      await Repo.save("actionItem", divider, { projectCode: proj.id, source: "Tracker", immediate: true });
      persistedOnce = true;
      ensureDividerRolesOnProjectPeople(proj, divider);
      return true;
    } catch (e) {
      toast((e && e.friendly) || "Couldn’t save divider", "error");
      return false;
    }
  }

  function scheduleAutosave() {
    if (!hasMeaningfulChanges() && isNew && !persistedOnce) return;
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      saveTimer = null;
      persistDivider().then((ok) => {
        if (ok && typeof onDone === "function") onDone();
      });
    }, 400);
  }

  function discardDraftIfNeeded() {
    if (!isNew || persistedOnce) return false;
    if (hasMeaningfulChanges()) return false;
    const list = (window.AEWTTR.db.ganttTasks && window.AEWTTR.db.ganttTasks[proj.id]) || [];
    const idx = list.findIndex((t) => t && t.id === divider.id);
    if (idx >= 0) list.splice(idx, 1);
    return true;
  }

  function dismiss() {
    if (settled) return;
    settled = true;
    clearTimeout(saveTimer);
    const discarded = discardDraftIfNeeded();
    if (!discarded && hasMeaningfulChanges()) {
      // Clicked out after typing — flush one last save, then close.
      persistDivider().finally(() => {
        closeModal();
        if (typeof onDone === "function") onDone();
      });
      return;
    }
    closeModal();
    if (typeof onDone === "function") onDone();
  }

  function dsViewField(label, value) {
    return `<div class="ds-vf"><span class="ds-vf-label">${label}</span><span class="ds-vf-value">${escapeHtml(value || "—")}</span></div>`;
  }
  function dsViewDate(label, value) {
    return dsViewField(label, value ? fmtDate(value) : "");
  }

  function buildViewHtml(sec) {
    if (sec === "overview") {
      return `<div class="ds-view-grid">
        ${dsViewField("Name", divider.title)}
        ${dsViewField("Priority", meta.priority)}
        ${dsViewDate("Start", meta.startDate)}
        ${dsViewDate("Due", meta.dueDate)}
        ${dsViewDate("Completed", meta.completionDate)}
        ${dsViewField("Lifecycle", meta.lifecycleStatus)}
      </div>`;
    }
    if (sec === "roles") {
      return `<div class="ds-view-grid">
        ${roleFields.map((f) => dsViewField(f.label, roleDraft[f.key])).join("")}
      </div>`;
    }
    if (sec === "program") {
      return `<div class="ds-view-grid">
        ${dsViewField("Program", meta.program)}
        ${dsViewField("Contract", meta.contract)}
        ${dsViewField("Task Order", meta.taskOrder)}
        ${dsViewField("Funding Type", meta.fundingType)}
        ${dsViewField("Fiscal Year", meta.fiscalYear)}
        ${dsViewField("Funding Status", meta.fundingStatus)}
        ${dsViewField("Config end item", Array.from(selectedConfigEnd)[0])}
        ${dsViewField("Change request", meta.changeRequestRequired ? "Yes" : "No")}
      </div>`;
    }
    if (sec === "tags") {
      return `<div class="ds-view-grid ds-view-grid--wide">
        ${dsViewField("Contractor", Array.from(selectedContractors).join(", "))}
        ${dsViewField("Portfolios", Array.from(selectedPortfolios).join(", "))}
        ${dsViewField("Locations", Array.from(selectedLocations).join(", "))}
      </div>`;
    }
    return "";
  }

  function updateViewContent(sec) {
    const el = modal && modal.querySelector(`#ds-view-${sec}`);
    if (el) el.innerHTML = buildViewHtml(sec);
  }

  const SECTIONS = [
    { id: "overview", label: "Overview" },
    { id: "roles", label: "Key roles" },
    { id: "program", label: "Program & funding" },
    { id: "tags", label: "Tags & assignments" }
  ];

  function sectionShell(sec, title, editHtml) {
    return `
      <section class="divider-settings-section">
        <div class="ds-section-head">
          <h4 class="divider-settings-section-title">${title}</h4>
          <button type="button" class="ds-edit-btn" data-ds-edit="${sec.id}"><i class="bx bx-pencil"></i> Edit</button>
        </div>
        <div class="ds-view-content" id="ds-view-${sec.id}">${buildViewHtml(sec.id)}</div>
        <div class="ds-edit-content" id="ds-edit-${sec.id}" style="display:none;">${editHtml}</div>
      </section>`;
  }

  const modal = openModal(`
    <div class="aewttr-modal-head">
      <h3>${isNew ? "New divider" : "Divider settings"}</h3>
      <span class="ds-save-status" id="ds-save-status"></span>
      <button type="button" class="aewttr-modal-close" aria-label="Close">&times;</button>
    </div>
    <div class="aewttr-modal-body divider-settings-body">
      ${sectionShell(SECTIONS[0], "Overview", `
        <div class="divider-settings-grid">
          <div class="form-row divider-settings-span-2"><label>Name</label><input class="input-aewttr" id="div-title" value="${escapeHtml(divider.title || "")}" placeholder="Divider name"></div>
          <div class="form-row"><label>Priority</label>
            <select class="select-aewttr" id="div-priority">
              <option value="">None</option>
              ${["High", "Medium", "Low"].map((o) => `<option value="${o}" ${meta.priority === o ? "selected" : ""}>${o}</option>`).join("")}
            </select>
          </div>
          <div class="form-row"><label>Start</label><input type="date" class="input-aewttr" id="div-start" value="${escapeHtml(meta.startDate || "")}"></div>
          <div class="form-row"><label>Due</label><input type="date" class="input-aewttr" id="div-due" value="${escapeHtml(meta.dueDate || "")}"></div>
          <div class="form-row"><label>Completed</label><input type="date" class="input-aewttr" id="div-complete" value="${escapeHtml(meta.completionDate || "")}"></div>
          <div class="form-row"><label>Lifecycle</label>
            <select class="select-aewttr" id="div-lifecycle">
              <option value="">None</option>
              ${["Planned", "Awaiting Funding", "Active", "Paused", "Complete"].map((o) => `<option value="${o}" ${(meta.lifecycleStatus === o || (o === "Planned" && meta.lifecycleStatus === "Planning") || (o === "Paused" && meta.lifecycleStatus === "On Hold") || (o === "Complete" && meta.lifecycleStatus === "Completed")) ? "selected" : ""}>${o}</option>`).join("")}
            </select>
          </div>
        </div>
      `)}
      ${sectionShell(SECTIONS[1], "Key roles", `
        <div class="divider-settings-grid">
          ${roleFields.map((f) => `
            <div class="form-row"><label>${escapeHtml(f.label)}</label>
              <input class="input-aewttr" id="div-role-${f.key}" list="div-member-list" value="${escapeHtml(roleDraft[f.key] || "")}" placeholder="Name">
            </div>`).join("")}
        </div>
        <datalist id="div-member-list">${(window.AEWTTR.db.members || []).map((m) => `<option value="${escapeHtml(m.name)}">`).join("")}</datalist>
      `)}
      ${sectionShell(SECTIONS[2], "Program &amp; funding", `
        <div class="divider-settings-grid">
          <div class="form-row"><label>Program</label><input class="input-aewttr" id="div-program" value="${escapeHtml(meta.program || "")}"></div>
          <div class="form-row"><label>Contract</label><input class="input-aewttr" id="div-contract" value="${escapeHtml(meta.contract || "")}"></div>
          <div class="form-row"><label>Task Order</label><input class="input-aewttr" id="div-taskorder" value="${escapeHtml(meta.taskOrder || "")}"></div>
          <div class="form-row"><label>Funding Type</label><input class="input-aewttr" id="div-fundingtype" value="${escapeHtml(meta.fundingType || "")}"></div>
          <div class="form-row"><label>Fiscal Year</label><input class="input-aewttr" id="div-fiscalyear" value="${escapeHtml(meta.fiscalYear || "")}"></div>
          <div class="form-row"><label>Funding Status</label><input class="input-aewttr" id="div-fundingstatus" value="${escapeHtml(meta.fundingStatus || "")}"></div>
          <div class="form-row divider-settings-span-2"><label>Config end item</label>${configEndItemPickerHtml(Array.from(selectedConfigEnd), "div-config")}</div>
          <div class="form-row">
            <label>Change request</label>
            <div class="travel-choice-row">${travelChoiceGroup("div-crr", ["Yes", "No"], meta.changeRequestRequired ? "Yes" : "No")}</div>
          </div>
        </div>
      `)}
      ${sectionShell(SECTIONS[3], "Tags &amp; assignments", `
        <div class="divider-settings-grid">
          <div class="form-row divider-settings-span-2"><label>Contractor</label>${contractorPickerHtml(Array.from(selectedContractors), "div-contractors")}</div>
          <div class="form-row divider-settings-span-2"><label>Portfolios</label>${portfolioPickerHtml(Array.from(selectedPortfolios), "div-portfolios")}</div>
          <div class="form-row divider-settings-span-2"><label>Locations</label>${locationPickerHtml(Array.from(selectedLocations), "div-locations")}</div>
        </div>
      `)}
    </div>
    <div class="aewttr-modal-foot">
      <button type="button" class="btn-aewttr-ghost" id="div-cancel">${isNew ? "Discard" : "Close"}</button>
    </div>
  `, { wide: true, className: "divider-settings-modal", onDismiss: dismiss });

  wirePortfolioPicker(modal, selectedPortfolios, "div-portfolios", scheduleAutosave);
  wireLocationPicker(modal, selectedLocations, "div-locations", scheduleAutosave);
  wireContractorPicker(modal, selectedContractors, "div-contractors", scheduleAutosave);
  wireConfigEndItemPicker(modal, selectedConfigEnd, "div-config", scheduleAutosave);

  roleFields.forEach((f) => {
    const input = $(`#div-role-${f.key}`, modal);
    if (!input) return;
    input.addEventListener("input", () => { roleDraft[f.key] = input.value; scheduleAutosave(); });
    input.addEventListener("change", () => { roleDraft[f.key] = input.value; scheduleAutosave(); });
  });

  $all("input, select", modal).forEach((el) => {
    if (el.closest(".tag-picker")) return;
    if (el.id && String(el.id).indexOf("div-role-") === 0) return;
    el.addEventListener("input", scheduleAutosave);
    el.addEventListener("change", scheduleAutosave);
  });
  $all(`input[name="div-crr"]`, modal).forEach((el) => el.addEventListener("change", scheduleAutosave));

  // Per-section edit toggles
  $all("[data-ds-edit]", modal).forEach((btn) => {
    btn.addEventListener("click", () => {
      const secId = btn.dataset.dsEdit;
      const viewEl = modal.querySelector(`#ds-view-${secId}`);
      const editEl = modal.querySelector(`#ds-edit-${secId}`);
      const isCurrentlyEditing = editEl && editEl.style.display !== "none";
      if (isCurrentlyEditing) {
        // Switching back to view mode: update view content with current values
        if (secId === "roles") {
          roleFields.forEach((f) => { const inp = $(`#div-role-${f.key}`, modal); if (inp) roleDraft[f.key] = inp.value; });
        }
        updateViewContent(secId);
        if (viewEl) viewEl.style.display = "";
        if (editEl) editEl.style.display = "none";
        btn.innerHTML = '<i class="bx bx-pencil"></i> Edit';
      } else {
        if (viewEl) viewEl.style.display = "none";
        if (editEl) editEl.style.display = "";
        btn.innerHTML = '<i class="bx bx-check"></i> Done';
        // Focus first input
        const first = editEl && editEl.querySelector("input, select");
        if (first) requestAnimationFrame(() => first.focus());
      }
    });
  });

  $(".aewttr-modal-close", modal).addEventListener("click", dismiss);
  $("#div-cancel", modal).addEventListener("click", dismiss);

  if (isNew) {
    // For new dividers, open overview for editing immediately
    requestAnimationFrame(() => {
      const editBtn = modal.querySelector('[data-ds-edit="overview"]');
      if (editBtn) editBtn.click();
    });
  }
}

/* ---------- Project Tracker: Gantt + linked Checklist ---------- */
function taskMatchesAssigneeQuery(task, tokens) {
  if (!tokens.length) return true;
  if (isTrackerDivider(task)) {
    const hay = [task.title || "", task.taskOrder || "", task.fundingType || ""].join(" ").toLowerCase();
    return tokens.some((tok) => hay.includes(tok));
  }
  const haystacks = [task.assignee || ""].concat([]);
  walkNestedSubtasks(task.subtasks || [], (s) => haystacks.push(s.assignee || ""));
  return haystacks.some(h => {
    const hl = h.toLowerCase();
    return tokens.some(tok => hl.includes(tok));
  });
}

function ensureTrackerUiState(stateKey, defaultView) {
  const s = window.AEWTTR.state;
  if (!s.trackerView) s.trackerView = {};
  if (!s.ganttExpanded) s.ganttExpanded = {};
  if (!s.trackerFilters) s.trackerFilters = {};
  if (!s.trackerView[stateKey]) s.trackerView[stateKey] = defaultView || "table";
  if (!s.ganttExpanded[stateKey]) s.ganttExpanded[stateKey] = {};
  if (!s.trackerFilters[stateKey]) s.trackerFilters[stateKey] = { assignee: "" };
  return {
    view: s.trackerView[stateKey],
    expanded: s.ganttExpanded[stateKey],
    filterState: s.trackerFilters[stateKey]
  };
}

function renderTrackerWorkspace(mount, config) {
  const db = window.AEWTTR.db;
  const proj = config.proj;
  const stateKey = config.stateKey || proj.id;
  const ui = ensureTrackerUiState(stateKey, config.defaultView);
  const view = ui.view;
  const expanded = ui.expanded;
  const filterState = ui.filterState;
  if (config.initialAssigneeFilter && !filterState.assignee) filterState.assignee = config.initialAssigneeFilter;

  if (!db.ganttTasks[proj.id]) db.ganttTasks[proj.id] = [];
  const allTasks = config.tasks || db.ganttTasks[proj.id];
  const saveSource = config.saveSource || "Tracker";
  const saveProjectCode = config.projectCode || proj.id;
  const memberOptions = (db.members || []).map((m) => `<option value="${escapeHtml(m.name)}"></option>`).join("");
  const mountId = `tracker-mount-${stateKey.replace(/[^a-zA-Z0-9_-]/g, "-")}`;
  const showHint = config.showHint !== false;
  const showAdd = config.showAddButton !== false;
  const meetingMode = !!config.meetingMode;
  const noToolbar = !!config.noToolbar;
  const hideViewTabs = !!config.hideViewTabs;
  const shellClass = meetingMode ? "tracker-shell tracker-shell--meeting" : "tracker-shell";
  const hideGroupHeader = !!config.hideGroupHeader;
  const collapseGanttGroupsByDefault = !!config.collapseGanttGroupsByDefault;
  const collapseDividersByDefault = config.collapseDividersByDefault != null
    ? !!config.collapseDividersByDefault
    : true;
  const expandTasksByDefault = !!config.expandTasksByDefault;
  const dividerCollapsed = trackerDividerCollapsedStore(stateKey);
  const subitemExpanded = trackerSubitemExpandedStore(stateKey);
  if (expandTasksByDefault) {
    if (!window.AEWTTR.state._trackerExpandInit) window.AEWTTR.state._trackerExpandInit = {};
    if (!window.AEWTTR.state._trackerExpandInit[stateKey]) {
      trackerPlainTasks(allTasks).forEach((task) => { expanded[task.id] = true; });
      window.AEWTTR.state._trackerExpandInit[stateKey] = true;
    }
  }

  const redraw = () => {
    if (typeof config.onRedraw === "function") config.onRedraw();
    else renderTrackerWorkspace(mount, config);
  };

  const showRisksTab = config.showRisksTab !== false;
  const metaStrip = "";

  mount.innerHTML = `
    <div class="${shellClass}">
      ${metaStrip}
      ${noToolbar && hideViewTabs ? "" : `
      <div class="monday-tracker-head">
        ${hideViewTabs ? "" : `
        <div class="monday-tracker-tabs" id="tracker-view-toggle-${mountId}">
          <button type="button" data-view="table" class="monday-tracker-tab ${view === "table" ? "active" : ""}">Main table</button>
          <button type="button" data-view="timeline" class="monday-tracker-tab ${view === "timeline" ? "active" : ""}">Timeline</button>
          ${showRisksTab ? `<button type="button" data-view="risks" class="monday-tracker-tab ${view === "risks" ? "active" : ""}">Risks</button>` : ""}
        </div>`}
        ${noToolbar || view === "risks" ? "" : `
        <div class="monday-tracker-toolbar">
          <div class="monday-tracker-toolbar-left">
            <button type="button" class="monday-tracker-ghost monday-tracker-ghost--label" id="btn-expand-all-${mountId}"${tip("Expand all tasks to show subitems")}><i class="bx bx-expand-vertical"></i> Expand all</button>
            <button type="button" class="monday-tracker-ghost monday-tracker-ghost--label" id="btn-collapse-all-${mountId}"${tip("Collapse all tasks")}><i class="bx bx-collapse-vertical"></i> Collapse all</button>
            <div class="monday-tracker-search">
              <i class="bx bx-search"></i>
              <input type="text" id="tracker-assignee-search-${mountId}" list="tracker-assignee-options-${mountId}" placeholder="Filter by owner" value="${escapeHtml(filterState.assignee)}" autocomplete="off">
              <datalist id="tracker-assignee-options-${mountId}">${memberOptions}</datalist>
              <button type="button" class="gantt-search-clear" id="tracker-assignee-clear-${mountId}"${tip("Clear assignee filter")} style="${filterState.assignee ? "" : "display:none;"}">&times;</button>
            </div>
          </div>
          <span class="tracker-save-status" id="tracker-save-status-${mountId}" data-state="idle"></span>
        </div>`}
      </div>`}
      <div class="tracker-view-mount tracker-view-mount--monday ${meetingMode ? "tracker-view-mount--meeting" : ""} ${view === "timeline" ? "tracker-view-mount--gantt" : ""} ${view === "risks" ? "tracker-view-mount--risks" : ""}" id="${mountId}"></div>
    </div>
  `;

  $all("[data-view]", $(`#tracker-view-toggle-${mountId}`, mount)).forEach((b) => b.addEventListener("click", () => {
    window.AEWTTR.state.trackerView[stateKey] = b.dataset.view;
    redraw();
  }));
  const settingsBtn = $("[data-open-project-settings]", mount);
  if (settingsBtn) {
    settingsBtn.addEventListener("click", () => navigate(`projects/${proj.id}/settings`));
  }

  // Live-update the tracker badge count in the project sidebar without re-rendering the whole nav.
  function updateTrackerNavBadge() {
    const taskCount = allTasks.filter((t) => t && !isTrackerDivider(t)).length;
    const btn = document.querySelector(".project-spo-link[data-tab='tracker']");
    if (!btn) return;
    let badge = btn.querySelector(".project-spo-badge");
    if (taskCount) {
      if (!badge) { badge = document.createElement("em"); badge.className = "project-spo-badge"; btn.appendChild(badge); }
      badge.textContent = taskCount;
    } else if (badge) {
      badge.remove();
    }
  }

  // Save status indicator — shows "Saving…" then "Saved" after each write.
  let _saveStatusTimer = null;
  function showSaveStatus(state) {
    const el = $(`#tracker-save-status-${mountId}`, mount);
    if (!el) return;
    clearTimeout(_saveStatusTimer);
    el.dataset.state = state;
    el.textContent = state === "saving" ? "Saving…" : state === "saved" ? "Saved" : state === "error" ? "Save failed" : "";
    if (state === "saved") _saveStatusTimer = setTimeout(() => { el.dataset.state = "idle"; el.textContent = ""; }, 2500);
  }

  function createInlineTrackerTask(parentDividerId, afterTaskId) {
    const today = new Date().toISOString().slice(0, 10);
    const id = uid("g");
    const newTask = {
      id,
      itemType: "task",
      workItemLevel: "Task",
      title: "New task",
      assignee: config.defaultAssignee || "Unassigned",
      start: today,
      end: "",
      health: "On Track",
      status: "Not Started",
      parentDividerId: parentDividerId || "",
      subtasks: [],
      notes: []
    };
    if (meetingMode) newTask.reviewStatus = "Not Reviewed";
    if (afterTaskId) {
      const afterIdx = allTasks.findIndex((t) => t.id === afterTaskId);
      allTasks.splice(afterIdx >= 0 ? afterIdx + 1 : allTasks.length, 0, newTask);
    } else {
      allTasks.push(newTask);
    }
    allTasks.forEach((item, idx) => { if (item) item._sortOrder = idx; });
    expanded[id] = true;
    // Ensure the section holding the new task is visible
    if (parentDividerId) dividerCollapsed[parentDividerId] = false;
    else dividerCollapsed.__ungrouped__ = false;
    window.AEWTTR.state.trackerFocusTaskId = id;
    if (meetingMode && typeof findActiveMeetingScopeForProject === "function" && typeof logMeetingActivity === "function") {
      const scope = findActiveMeetingScopeForProject(saveProjectCode);
      if (scope) {
        const actor = (window.AEWTTR.db.user && window.AEWTTR.db.user.name) || "Someone";
        logMeetingActivity(scope, `${actor} added a new task to project ${saveProjectCode}: ${newTask.title}.`, {
          type: "create",
          projectId: saveProjectCode,
          taskId: newTask.id,
          taskTitle: newTask.title
        });
      }
    }
    showSaveStatus("saving");
    Repo.save("actionItem", newTask, { projectCode: saveProjectCode, source: saveSource })
      .then(() => showSaveStatus("saved"))
      .catch(() => { showSaveStatus("error"); });
    ensureAssigneesFromTask(saveProjectCode, newTask);
    afterSave();
    // In-place content refresh — avoid remounting the tracker shell (tabs/search).
    // Meeting person filters still need a shell remount to recompute visibility.
    if (meetingMode) redraw();
    else refreshContent();
    updateTrackerNavBadge();
    if (view === "timeline" && !meetingMode) openTaskSidePanel(newTask, allTasks, refreshContent);
    return newTask;
  }

  function createInlineDivider() {
    const divider = createTrackerDivider({ title: "New project divider" });
    if (!divider.metadata) divider.metadata = {};
    divider.metadata.clientLocalId = divider.id;
    divider.metadata.projectCode = saveProjectCode;
    divider._isDraft = true;
    allTasks.push(divider);
    dividerCollapsed[divider.id] = false;
    // Do not persist until the user enters a name or other fields.
    openDividerSettingsModal(divider, proj, refreshContent, { isNew: true });
    return divider;
  }

  const addBtn = $(`#btn-add-gantt-task-${mountId}`, mount);
  if (addBtn) {
    addBtn.addEventListener("click", () => {
      if (typeof config.onAddTask === "function") return config.onAddTask(() => createInlineTrackerTask(""));
      createInlineTrackerTask("");
    });
  }
  const addDivBtn = $(`#btn-add-divider-${mountId}`, mount);
  if (addDivBtn) addDivBtn.addEventListener("click", () => createInlineDivider());

  const expandAllBtn = $(`#btn-expand-all-${mountId}`, mount);
  if (expandAllBtn) expandAllBtn.addEventListener("click", () => {
    trackerPlainTasks(allTasks).forEach((t) => { expanded[t.id] = true; });
    refreshContent();
  });
  const collapseAllBtn = $(`#btn-collapse-all-${mountId}`, mount);
  if (collapseAllBtn) collapseAllBtn.addEventListener("click", () => {
    trackerPlainTasks(allTasks).forEach((t) => { expanded[t.id] = false; });
    refreshContent();
  });

  const searchInput = $(`#tracker-assignee-search-${mountId}`, mount);
  if (searchInput) {
    searchInput.addEventListener("input", () => {
      filterState.assignee = searchInput.value;
      $(`#tracker-assignee-clear-${mountId}`, mount).style.display = filterState.assignee ? "" : "none";
      renderView();
    });
    $(`#tracker-assignee-clear-${mountId}`, mount).addEventListener("click", () => { filterState.assignee = ""; refreshContent(); });
  }

  const openEditor = (taskId, subIndex) => {
    if (typeof config.onOpenEditor === "function") return config.onOpenEditor(taskId, subIndex);
    const task = allTasks.find((t) => t.id === taskId);
    if (!task) return;
    expanded[taskId] = true;
    if (subIndex != null && subIndex !== "" && !Number.isNaN(subIndex)) {
      // Path strings ("0.1") or legacy numeric indices.
      const path = String(subIndex);
      const sub = getSubtaskAtPath(task, path.includes(".") ? path : path);
      if (sub && path.includes(".")) {
        // Open nearest flat index side panel for deeply nested? Prefer path-aware editor.
        openSubtaskSidePanel(task, allTasks, refreshContent, path);
        return;
      }
      openSubtaskSidePanel(task, allTasks, refreshContent, typeof subIndex === "number" ? subIndex : Number(path));
      return;
    }
    if (isTrackerDivider(task)) {
      openDividerSettingsModal(task, proj, refreshContent);
      return;
    }
    openTaskSidePanel(task, allTasks, refreshContent);
  };
  const toggleExpand = (taskId) => { expanded[taskId] = !expanded[taskId]; renderView(); };
  const viewMount = $(`#${mountId}`, mount);
  const ganttOpts = {
    mondayStyle: true,
    saveSource,
    projectCode: saveProjectCode,
    startDate: proj.startDate || "",
    dueDate: proj.dueDate || "",
    dividerCollapsed,
    collapseDividersByDefault
  };

  function afterSave() {
    showSaveStatus("saved");
    if (typeof config.onAfterSave === "function") config.onAfterSave();
  }

  function renderView() {
    const tokens = filterState.assignee.split(",").map((t) => t.trim().toLowerCase()).filter(Boolean);
    let visibleTasks = allTasks.filter((t) => taskMatchesAssigneeQuery(t, tokens));
    if (typeof config.taskFilter === "function") visibleTasks = visibleTasks.filter(config.taskFilter);
    const viewOpts = { ...ganttOpts, meetingMode, collapseGanttGroupsByDefault };
    if (view === "risks") { drawProjectRisks(viewMount, proj, { trackerEmbed: true }); return; }
    if (view === "timeline") { const ganttTasks = allTasks.filter((t) => !isTrackerDivider(t)); renderGanttChart(viewMount, ganttTasks, expanded, toggleExpand, openEditor, refreshContent, visibleTasks.filter((t) => !isTrackerDivider(t)), viewOpts); }
    else renderTrackerTableView(viewMount, allTasks, expanded, toggleExpand, openEditor, refreshContent, visibleTasks, proj, {
      saveSource,
      projectCode: saveProjectCode,
      onAfterSave: afterSave,
      onBeforeSave: () => showSaveStatus("saving"),
      onSaveError: () => showSaveStatus("error"),
      meetingMode,
      hideGroupHeader,
      hideHeader: !!config.hideHeader,
      defaultAssignee: config.defaultAssignee || "",
      onCreateTask: createInlineTrackerTask,
      dividerCollapsed,
      subitemExpanded,
      stateKey
    });
  }

  /* Refresh table/gantt content without remounting the tracker shell (tabs,
     toolbar, search). Use full redraw() only when the shell itself must change. */
  function refreshContent() {
    renderView();
  }

  if (!window.AEWTTR._trackerLiveRefreshKeys) window.AEWTTR._trackerLiveRefreshKeys = new Set();
  if (!window.AEWTTR._trackerLiveRefreshKeys.has(stateKey)) {
    window.AEWTTR._trackerLiveRefreshKeys.add(stateKey);
    window.addEventListener("pulse:data-refreshed", () => {
      if (!document.getElementById(mountId)) return;
      if (typeof isBackgroundDataSwapBlocked === "function" && isBackgroundDataSwapBlocked()) return;
      if (document.querySelector(".aewttr-modal-backdrop")) return;
      const active = document.activeElement;
      if (active && (active.tagName === "INPUT" || active.tagName === "TEXTAREA" || active.tagName === "SELECT" || active.isContentEditable)) return;
      const scroller = document.querySelector(".aewttr-content");
      const scrollTop = scroller ? scroller.scrollTop : 0;
      // Remount (don't just renderView) so allTasks is re-read from the
      // live db after a background refresh replaced window.AEWTTR.db.
      redraw();
      if (scroller) scroller.scrollTop = scrollTop;
    });
  }

  renderView();
}

function drawTracker(body, proj) {
  renderTrackerWorkspace(body, { proj, stateKey: proj.id });
}

function normalizeTaskSubtask(task, subtask) {
  const nested = (subtask && (subtask.subtasks || subtask.children)) || [];
  return {
    id: (subtask && subtask.id) || "",
    text: subtask && subtask.text ? subtask.text : "",
    assignee: subtask && subtask.assignee ? subtask.assignee : "",
    done: !!(subtask && subtask.done),
    isMilestone: !!(subtask && subtask.isMilestone),
    health: (subtask && subtask.health) || (task && task.health) || "On Track",
    start: subtask && subtask.start ? subtask.start : (task && task.start ? task.start : ""),
    end: subtask && subtask.end ? subtask.end : "",
    linked: !(subtask && subtask.linked === false),
    notes: Array.isArray(subtask && subtask.notes) ? subtask.notes : [],
    relatedDocs: (subtask && subtask.relatedDocs || []).map(doc => ({
      fileName: doc && doc.fileName ? doc.fileName : "",
      location: doc && doc.location ? doc.location : "Windchill"
    })).filter(doc => doc.fileName),
    subtasks: nested.map((child) => normalizeTaskSubtask({
      start: (subtask && subtask.start) || (task && task.start) || "",
      end: (subtask && subtask.end) || (task && task.end) || "",
      health: (subtask && subtask.health) || (task && task.health) || "On Track"
    }, child))
  };
}

/* ---------- task ↔ subtask date linking ---------- */
/* Project-level master switch (Settings > "Link subtasks to task dates").
   Defaults ON when the flag is unset. */
function subtaskLinkEnabled(tasks) {
  const db = window.AEWTTR.db;
  const pid = projectCodeForTaskList(tasks);
  const proj = (db.projects || []).find(p => p.id === pid);
  return !proj || proj.linkSubtaskDates !== false;
}
function addDaysIso(iso, delta) {
  if (!iso || !delta) return iso;
  return ganttIsoDate(ganttAddDays(ganttParseDate(iso), delta));
}
/* Re-flow a task's linked subtasks after its dates change.
   - Whole-task move (start and end shift by the same delta): subtasks travel with it.
   - End-only change: subtasks pinned to the old end follow to the new end; all
     subtasks are clamped so they never fall outside the [start, end] window.
   `snapshot` holds each subtask's pre-change {start,end} so repeated live drag
   updates stay stable rather than accumulating. */
function cascadeSubtaskDates(tasks, task, prevStart, prevEnd, snapshot) {
  if (!subtaskLinkEnabled(tasks) || !task.subtasks || !task.subtasks.length) return;
  const startDelta = ganttDaysBetween(prevStart, task.start);
  const endDelta = ganttDaysBetween(prevEnd, task.end);
  if (startDelta === 0 && endDelta === 0) return;
  const wholeMove = startDelta === endDelta;
  task.subtasks.forEach((s, i) => {
    if (s.linked === false || !s.start || !s.end) return;
    const base = (snapshot && snapshot[i]) ? snapshot[i] : { start: s.start, end: s.end };
    if (wholeMove) {
      s.start = addDaysIso(base.start, startDelta);
      s.end = addDaysIso(base.end, endDelta);
    } else {
      s.start = base.start;
      s.end = base.end === prevEnd ? task.end : base.end;
    }
    // clamp inside the task window
    if (task.end && s.end > task.end) s.end = task.end;
    if (task.start && s.start < task.start) s.start = task.start;
    if (task.start && s.end < task.start) s.end = task.start;
    if (s.start > s.end) s.start = s.end;
    if (s.subtasks && s.subtasks.length) {
      cascadeSubtaskDates(tasks, { start: s.start, end: s.end, subtasks: s.subtasks }, base.start, base.end, null);
    }
  });
}

function subtaskStateClass(subtask) {
  if (subtask.done) return "done";
  return "";
}

function ganttTaskTipText(task) {
  const parts = [task.title || "Untitled task"];
  if (task.assignee) parts.push(`Owner: ${task.assignee}`);
  if (task.start && task.end) parts.push(`${fmtDate(task.start)} – ${fmtDate(task.end)}`);
  if (task.status) parts.push(task.status);
  parts.push(`${taskProgressPct(task)}% progress`);
  return parts.join(" · ");
}

function ganttSubtaskTipText(task, subtask) {
  const parts = [subtask.text || "Untitled subtask"];
  if (subtask.assignee) parts.push(`Owner: ${subtask.assignee}`);
  if (subtask.start && subtask.end) parts.push(`${fmtDate(subtask.start)} – ${fmtDate(subtask.end)}`);
  parts.push(subtask.done ? "Done" : "Open");
  if (task && task.title) parts.push(`Parent: ${task.title}`);
  return parts.join(" · ");
}

function renderSubtaskPreview(task, rawSubtask, index) {
  const subtask = normalizeTaskSubtask(task, rawSubtask);
  const stateClass = subtaskStateClass(subtask);
  const parentStart = task.start || subtask.start;
  const parentEnd = task.end || subtask.end || parentStart;
  let miniGantt = `<div class="mini-gantt-row"><div class="mini-gantt-dates">No subtask timeline set yet.</div></div>`;
  if (parentStart && parentEnd && subtask.start && subtask.end) {
    const parentDays = Math.max(ganttDaysBetween(parentStart, parentEnd) + 1, 1);
    const offset = Math.max(ganttDaysBetween(parentStart, subtask.start), 0);
    const span = Math.max(ganttDaysBetween(subtask.start, subtask.end) + 1, 1);
    const left = Math.min((offset / parentDays) * 100, 100);
    const width = Math.min((span / parentDays) * 100, 100 - left);
    miniGantt = `
      <div class="mini-gantt-row">
        <div class="mini-gantt-track">
          <div class="mini-gantt-bar ${stateClass}" style="left:${left}%;width:${Math.max(width, 6)}%;"></div>
        </div>
        <div class="mini-gantt-dates">${fmtDate(subtask.start)} – ${fmtDate(subtask.end)}</div>
      </div>`;
  }
  return `
    <div class="gantt-sub-item">
      <div class="gantt-sub-top">
        <input type="checkbox" data-task="${task.id}" data-sub="${index}" ${subtask.done ? "checked" : ""}>
        <div class="gantt-sub-text" style="${subtask.done ? "text-decoration:line-through;color:var(--aewttr-muted);" : ""}">${escapeHtml(subtask.text || "Untitled subtask")}</div>
        ${subtask.assignee ? `<span class="kc-badge">${escapeHtml(subtask.assignee)}</span>` : `<span class="kc-badge">Unassigned</span>`}
        <button class="btn-aewttr-outline btn-aewttr-sm gantt-sub-edit" data-open-sub="${task.id}:${index}"${tip("Edit this subtask")}><i class="bx bx-edit"></i> Edit</button>
      </div>
      ${miniGantt}
      ${subtask.relatedDocs.length ? `
        <div class="subtask-files-row">
          ${subtask.relatedDocs.map(doc => `<span class="subtask-file-chip"><b>${escapeHtml(doc.location)}</b> ${escapeHtml(doc.fileName)}</span>`).join("")}
        </div>` : `
        <div class="mini-gantt-row"><div class="mini-gantt-dates">No file references added yet.</div></div>`}
    </div>`;
}

function renderInlineSubtaskBars(task, rangeStart, dayWidth) {
  const subtasks = (task.subtasks || []).filter(s => s && (s.text || s.start || s.end));
  if (!subtasks.length) return "";
  return `
    <div class="gantt-subtrack-wrap">
      ${subtasks.map((rawSubtask, index) => {
        const subtask = normalizeTaskSubtask(task, rawSubtask);
        const stateClass = subtaskStateClass(subtask);
        if (!subtask.start || !subtask.end) {
          return `
            <button class="gantt-subbar gantt-subbar-untimed ${stateClass}" data-open-sub="${task.id}:${index}" data-drag-sub="${task.id}:${index}" data-day-width="${dayWidth}" style="top:${index * 16}px;"${typeof ganttTip === "function" ? ganttTip(ganttSubtaskTipText(task, subtask)) : ""}>
              <span class="gantt-subbar-label">${escapeHtml(subtask.text || "Untitled subtask")}</span>
              <span class="gantt-resize-handle left" data-drag-sub="${task.id}:${index}" data-mode="start"></span>
              <span class="gantt-resize-handle right" data-drag-sub="${task.id}:${index}" data-mode="end"></span>
            </button>`;
        }
        const left = ganttDaysBetween(ganttIsoDate(rangeStart), subtask.start) * dayWidth;
        const width = Math.max((ganttDaysBetween(subtask.start, subtask.end) + 1) * dayWidth, 10);
        return `
          <button class="gantt-subbar ${stateClass}" data-open-sub="${task.id}:${index}" data-drag-sub="${task.id}:${index}" data-day-width="${dayWidth}" style="left:${left}px;width:${width}px;top:${index * 16}px;"${typeof ganttTip === "function" ? ganttTip(ganttSubtaskTipText(task, subtask)) : ""}>
            <span class="gantt-subbar-label">${escapeHtml(subtask.text || "Untitled subtask")}</span>
            <span class="gantt-resize-handle left" data-drag-sub="${task.id}:${index}" data-mode="start"></span>
            <span class="gantt-resize-handle right" data-drag-sub="${task.id}:${index}" data-mode="end"></span>
          </button>`;
      }).join("")}
    </div>`;
}

function clampDateOrder(start, end) {
  if (start > end) return [end, start];
  return [start, end];
}

function shiftRange(start, end, deltaDays) {
  return [
    ganttIsoDate(ganttAddDays(ganttParseDate(start), deltaDays)),
    ganttIsoDate(ganttAddDays(ganttParseDate(end), deltaDays))
  ];
}

function resizeRange(start, end, deltaDays, mode) {
  let nextStart = start;
  let nextEnd = end;
  if (mode === "start") nextStart = ganttIsoDate(ganttAddDays(ganttParseDate(start), deltaDays));
  if (mode === "end") nextEnd = ganttIsoDate(ganttAddDays(ganttParseDate(end), deltaDays));
  [nextStart, nextEnd] = clampDateOrder(nextStart, nextEnd);
  return [nextStart, nextEnd];
}

function wireGanttBarDragging(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, dragOpts) {
  dragOpts = dragOpts || {};
  const saveSource = dragOpts.saveSource || "Tracker";
  const saveProjectCode = dragOpts.projectCode || projectCodeForTaskList(tasks);
  let dragState = null;

  function beginDrag(e, payload) {
    e.preventDefault();
    e.stopPropagation();
    // A truthy dragState here means a gesture is already in progress (e.g. a
    // second touch point during multi-touch) — the window listeners from
    // that gesture are already live, so don't attach a second set of them.
    const alreadyDragging = !!dragState;
    // lastDelta starts at 0 (not undefined/null) so the first pointermove —
    // which is almost always a sub-pixel jitter that rounds to a 0-day
    // delta — is recognized as "no actual change" and skipped below, rather
    // than being treated as a completed drag that then blocks the
    // subsequent click from opening the task/subtask editor.
    // Capture the gantt range start from the board's CSS variable for in-place updates
    const boardEl = mount.querySelector(".monday-gantt-board, .gantt-board");
    let rangeStartMs = null;
    if (boardEl) {
      const style = getComputedStyle(boardEl);
      const todayEl = mount.querySelector(".monday-gantt-today, .gantt-today");
      if (todayEl && payload.dayWidth) {
        const todayLeft = parseFloat(todayEl.style.left) || 0;
        const todayDays = todayLeft / payload.dayWidth;
        const now = new Date();
        now.setHours(0, 0, 0, 0);
        rangeStartMs = now.getTime() - todayDays * 86400000;
      }
    }
    dragState = { ...payload, startX: e.clientX, lastDelta: 0, moved: false, rangeStartMs };
    document.body.classList.add("gantt-dragging");
    if (!alreadyDragging) {
      window.addEventListener("pointermove", onPointerMove);
      window.addEventListener("pointerup", onPointerUp);
    }
  }

  function updateBarInPlace(barEl, start, end) {
    if (!barEl || !dragState) return;
    const rangeStartEl = mount.querySelector(".monday-gantt-board, .gantt-board");
    const dayW = dragState.dayWidth;
    // Calculate range start from the gantt's CSS variable or re-derive from the first dated task
    const rangeStartMs = dragState.rangeStartMs;
    if (!rangeStartMs) return;
    const rangeStart = new Date(rangeStartMs);
    const left = ganttDaysBetween(ganttIsoDate(rangeStart), start) * dayW;
    const width = Math.max((ganttDaysBetween(start, end) + 1) * dayW, dayW * 2);
    barEl.style.left = left + "px";
    barEl.style.width = width + "px";
  }

  function onPointerMove(e) {
    if (!dragState) return;
    const deltaDays = Math.round((e.clientX - dragState.startX) / dragState.dayWidth);
    if (deltaDays === dragState.lastDelta) return;
    dragState.lastDelta = deltaDays;
    dragState.moved = true;
    const task = tasks.find(t => t.id === dragState.taskId);
    if (!task) return;
    if (dragState.kind === "task") {
      const next = dragState.mode === "move"
        ? shiftRange(dragState.originalStart, dragState.originalEnd, deltaDays)
        : resizeRange(dragState.originalStart, dragState.originalEnd, deltaDays, dragState.mode);
      [task.start, task.end] = next;
      // In-place DOM update — no full redraw during drag for smoothness
      const barEl = mount.querySelector(`[data-drag-task="${CSS.escape(task.id)}"]`);
      updateBarInPlace(barEl, task.start, task.end);
      return;
    }
    const subtask = task.subtasks && task.subtasks[dragState.subIndex];
    if (!subtask) return;
    const next = dragState.mode === "move"
      ? shiftRange(dragState.originalStart, dragState.originalEnd, deltaDays)
      : resizeRange(dragState.originalStart, dragState.originalEnd, deltaDays, dragState.mode);
    subtask.start = next[0];
    subtask.end = next[1];
    // In-place DOM update for subtask bar
    const subBarEl = mount.querySelector(`[data-drag-sub="${CSS.escape(dragState.taskId + ":" + dragState.subIndex)}"]`);
    updateBarInPlace(subBarEl, subtask.start, subtask.end);
  }

  function onPointerUp() {
    // Calling preventDefault() on pointerdown (above, in beginDrag — needed
    // so a drag doesn't also select page text) makes the browser skip the
    // compatibility "click" event it would otherwise synthesize for a mouse
    // pointer. Since the bar/subbar elements are also the app's click-to-
    // open targets (data-open/data-open-sub live on the same node as
    // data-drag-task/data-drag-sub), that native click never reliably
    // arrives — this was the actual "popups don't open" bug, not just an
    // edge case. So a plain press-and-release (no real movement, and not a
    // resize-handle grab) opens the editor directly here instead of relying
    // on a click event at all.
    if (dragState && !dragState.moved && dragState.mode === "move") {
      if (dragState.kind === "task") onOpenEditor(dragState.taskId);
      else onOpenEditor(dragState.taskId, dragState.subIndex);
    } else if (dragState && dragState.moved) {
      // Save and redraw only once on drop (not on every pixel during drag)
      const task = tasks.find(t => t.id === dragState.taskId);
      if (task) {
        if (dragState.kind === "subtask") {
          // If subtask now extends beyond parent task, expand the parent
          const sub = task.subtasks && task.subtasks[dragState.subIndex];
          if (sub && sub.start && sub.end && task.start && task.end) {
            let parentChanged = false;
            if (sub.start < task.start) { task.start = sub.start; parentChanged = true; }
            if (sub.end > task.end) { task.end = sub.end; parentChanged = true; }
            if (parentChanged) {
              const parentBarEl = mount.querySelector(`[data-drag-task="${CSS.escape(task.id)}"]`);
              updateBarInPlace(parentBarEl, task.start, task.end);
            }
          }
        }
        Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource });
      }
      redraw();
      mount.dataset.dragJustFinished = "1";
      setTimeout(() => { delete mount.dataset.dragJustFinished; }, 120);
    }
    dragState = null;
    document.body.classList.remove("gantt-dragging");
    window.removeEventListener("pointermove", onPointerMove);
    window.removeEventListener("pointerup", onPointerUp);
  }

  $all(".gantt-bar[data-drag-task]", mount).forEach(node => node.addEventListener("pointerdown", (e) => {
    const task = tasks.find(t => t.id === node.dataset.dragTask);
    if (!task || !task.start || !task.end) return;
    const handle = e.target.closest(".gantt-resize-handle");
    beginDrag(e, {
      kind: "task",
      taskId: task.id,
      mode: handle ? handle.dataset.mode : "move",
      originalStart: task.start,
      originalEnd: task.end,
      originalSubtasks: (task.subtasks || []).map(s => ({ start: s.start, end: s.end })),
      dayWidth: +node.dataset.dayWidth
    });
  }));

  $all(".gantt-subbar[data-drag-sub]", mount).forEach(node => node.addEventListener("pointerdown", (e) => {
    const [taskId, subIndexText] = node.dataset.dragSub.split(":");
    const task = tasks.find(t => t.id === taskId);
    const subIndex = +subIndexText;
    const subtask = task && task.subtasks && task.subtasks[subIndex];
    if (!task || !subtask || !subtask.start || !subtask.end) return;
    const handle = e.target.closest(".gantt-resize-handle");
    beginDrag(e, {
      kind: "subtask",
      taskId,
      subIndex,
      mode: handle ? handle.dataset.mode : "move",
      originalStart: subtask.start,
      originalEnd: subtask.end,
      dayWidth: +node.dataset.dayWidth
    });
  }));
}

function ganttDateRangeLabel(start, end) {
  if (!start && !end) return "No dates";
  if (!start || !end) return fmtDate(start || end);
  const s = new Date(start + "T00:00:00");
  const e = new Date(end + "T00:00:00");
  if (isNaN(s) || isNaN(e)) return `${fmtDate(start)} – ${fmtDate(end)}`;
  const startLabel = s.toLocaleDateString("en-US", { month: "short", day: "numeric" });
  const endLabel = e.toLocaleDateString("en-US", { month: "short", day: "numeric" });
  return `${startLabel} – ${endLabel}`;
}

const MONDAY_GANTT_GROUPS = [
  { key: "planned", label: "Not started", dot: "#8590A2", bar: "monday-bar--planned" },
  { key: "progress", label: "In progress", dot: "#2451C4", bar: "monday-bar--progress" },
  { key: "blocked", label: "Blocked", dot: "#A32B22", bar: "monday-bar--blocked" },
  { key: "done", label: "Complete", dot: "#1A7A52", bar: "monday-bar--done" }
];

function ganttGroupKeyForTask(task) {
  const pct = taskProgressPct(task);
  if (pct === 100) return "done";
  if (task.health === "Off Track" || task.status === "Blocked") return "blocked";
  if (pct > 0) return "progress";
  return "planned";
}

function ganttGroupForTask(task) {
  const key = ganttGroupKeyForTask(task);
  return MONDAY_GANTT_GROUPS.find((g) => g.key === key) || MONDAY_GANTT_GROUPS[0];
}

function ganttZoomDayWidth(zoom) {
  if (zoom === "months") return 10;
  if (zoom === "days") return 28;
  return 20;
}

/* Range mode ("day"/"week"/"month") constrains which window of the
   timeline is shown, independent of the density zoom above — "whole"
   falls back to the existing task-derived min/max range. */
function ganttWeekStart(date) {
  const d = new Date(date);
  d.setDate(d.getDate() - d.getDay());
  d.setHours(0, 0, 0, 0);
  return d;
}
function ganttMonthStart(date) {
  const d = new Date(date.getFullYear(), date.getMonth(), 1);
  d.setHours(0, 0, 0, 0);
  return d;
}
function ganttRangeWindowFor(mode, anchorIso) {
  const anchor = anchorIso ? ganttParseDate(anchorIso) : new Date();
  if (mode === "day") return { start: anchor, end: anchor };
  if (mode === "week") { const start = ganttWeekStart(anchor); return { start, end: ganttAddDays(start, 6) }; }
  if (mode === "month") { const start = ganttMonthStart(anchor); return { start, end: new Date(start.getFullYear(), start.getMonth() + 1, 0) }; }
  return null;
}
function ganttRangeDayWidth(mode) {
  // Preferred density for each range — used as a floor for Whole Project
  // scroll density. Day/week/month windows scale up from these via
  // ganttResolveDayWidth so the timeline always fills the mount width.
  if (mode === "day") return 720;
  if (mode === "week") return 64;
  if (mode === "month") return 22;
  return null;
}

function ganttInfoColumnWidth(projectCode) {
  const stored = window.AEWTTR.state
    && window.AEWTTR.state.ganttInfoWidth
    && window.AEWTTR.state.ganttInfoWidth[projectCode];
  return Math.max(200, Number(stored) || 300);
}

/** Usable px for the timeline column (mount width minus the sticky info lane). */
function ganttMeasureTimelineAvail(mount, infoWidth) {
  if (!mount) return 0;
  const rectW = mount.getBoundingClientRect ? mount.getBoundingClientRect().width : 0;
  const width = Math.max(mount.clientWidth || 0, Math.floor(rectW) || 0);
  // 2px buffer for border/subpixel so the board doesn't force a 1px scrollbar.
  return Math.max(0, width - infoWidth - 2);
}

/**
 * Scale day columns to the mount: fixed Day/Week/Month windows fill the
 * section; Whole Project stretches when it would be a stub, else keeps
 * preferred density and scrolls horizontally.
 */
function ganttResolveDayWidth(preferred, availWidth, totalDays, fillExact) {
  if (!totalDays || totalDays < 1) return preferred;
  if (!(availWidth > 0)) return preferred;
  const fill = availWidth / totalDays;
  if (fillExact) {
    // Keep hour/day ticks readable on very narrow Firepit panes.
    const floor = preferred >= 100 ? 56 : preferred >= 40 ? 32 : 14;
    return Math.max(floor, fill);
  }
  return Math.max(preferred, fill);
}
function ganttRangeLabel(mode, window) {
  if (mode === "day") return fmtDate(ganttIsoDate(window.start));
  if (mode === "week") return `${fmtDate(ganttIsoDate(window.start))} – ${fmtDate(ganttIsoDate(window.end))}`;
  if (mode === "month") return window.start.toLocaleDateString("en-US", { month: "long", year: "numeric" });
  return "Whole project";
}

function renderMondayGanttChart(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts) {
  const saveSource = opts.saveSource || "Tracker";
  const saveProjectCode = opts.projectCode || projectCodeForTaskList(tasks);
  const meetingMode = !!opts.meetingMode;
  visibleTasks = visibleTasks || tasks;
  if (!tasks.length) {
    mount.innerHTML = `<div class="monday-gantt-empty">No tasks yet. Click <strong>New task</strong> to start building your plan.</div>`;
    return;
  }
  if (!visibleTasks.length) {
    mount.innerHTML = `<div class="monday-gantt-empty">No tasks match this filter.</div>`;
    return;
  }

  if (!window.AEWTTR.state.ganttZoom) window.AEWTTR.state.ganttZoom = {};
  if (!window.AEWTTR.state.ganttGroupCollapsed) window.AEWTTR.state.ganttGroupCollapsed = {};
  if (!window.AEWTTR.state.ganttRange) window.AEWTTR.state.ganttRange = {};
  const zoomKey = saveProjectCode;
  // Keep meeting collapse state separate from the project Tracker page so
  // "closed by default" in Weekly Meeting doesn't fight the project Gantt.
  const groupKey = opts.collapseGanttGroupsByDefault ? `meeting:${saveProjectCode}` : saveProjectCode;
  // Defaults to "week" (today's week) per-project the first time its Gantt
  // is opened; switching to Whole Project falls back to the old
  // task-derived min/max range and density dropdown.
  if (!window.AEWTTR.state.ganttRange[zoomKey]) window.AEWTTR.state.ganttRange[zoomKey] = { mode: "week", anchor: null };
  const rangeState = window.AEWTTR.state.ganttRange[zoomKey];
  const rangeMode = rangeState.mode || "week";
  const explicitWindow = ganttRangeWindowFor(rangeMode, rangeState.anchor);
  // Meeting / multi-project views start with every status group collapsed so
  // the list stays scannable until the facilitator opens a section.
  if (opts.collapseGanttGroupsByDefault && window.AEWTTR.state.ganttGroupCollapsed[groupKey] == null) {
    const init = {};
    (typeof MONDAY_GANTT_GROUPS !== "undefined" ? MONDAY_GANTT_GROUPS : []).forEach((group) => {
      init[group.key] = true;
    });
    window.AEWTTR.state.ganttGroupCollapsed[groupKey] = init;
  }
  const collapsedGroups = window.AEWTTR.state.ganttGroupCollapsed[groupKey] || {};

  const datedForRange = visibleTasks.filter((t) => !isTrackerDivider(t) && t.start && t.end);
  const rangeSource = datedForRange.length ? datedForRange : visibleTasks.filter((t) => !isTrackerDivider(t));
  let starts = rangeSource.map((t) => ganttParseDate(t.start || new Date().toISOString().slice(0, 10)).getTime());
  let ends = rangeSource.map((t) => ganttParseDate(t.end || t.start || new Date().toISOString().slice(0, 10)).getTime());
  if (!starts.length) {
    const now = Date.now();
    starts = [now];
    ends = [now];
  }
  const taskRangeStart = ganttAddDays(new Date(Math.min(...starts)), -5);
  const taskRangeEnd = ganttAddDays(new Date(Math.max(...ends)), 10);
  const rangeStart = explicitWindow ? explicitWindow.start : taskRangeStart;
  const rangeEnd = explicitWindow ? ganttAddDays(explicitWindow.end, 1) : taskRangeEnd;
  const totalDays = explicitWindow
    ? Math.max(ganttDaysBetween(ganttIsoDate(rangeStart), ganttIsoDate(rangeEnd)), 1) + 1
    : Math.max(ganttDaysBetween(ganttIsoDate(rangeStart), ganttIsoDate(rangeEnd)), 21) + 1;
  const zoom = window.AEWTTR.state.ganttZoom[zoomKey] || "weeks";
  const preferredDayWidth = ganttRangeDayWidth(rangeMode) || ganttZoomDayWidth(zoom);
  const infoWidth = ganttInfoColumnWidth(saveProjectCode);
  const timelineAvail = ganttMeasureTimelineAvail(mount, infoWidth);
  const dayWidth = ganttResolveDayWidth(preferredDayWidth, timelineAvail, totalDays, !!explicitWindow);
  const timelineWidth = totalDays * dayWidth;
  const monthPills = buildMondayGanttMonthPills(rangeStart, totalDays + 1, dayWidth);
  const tickMode = explicitWindow ? rangeMode : "whole";
  const ticksHtml = buildMondayGanttTicks(tickMode, rangeStart, dayWidth, totalDays + 1);
  const quarterLabel = explicitWindow ? ganttRangeLabel(rangeMode, explicitWindow) : buildMondayGanttQuarterLabel(rangeStart, rangeEnd);
  const todayOffset = ganttDaysBetween(ganttIsoDate(rangeStart), ganttIsoDate(new Date()));
  const todayLine = (todayOffset >= 0 && todayOffset <= totalDays) ? `<div class="monday-gantt-today" style="left:${todayOffset * dayWidth}px;"></div>` : "";
  const startDateOffset = opts.startDate ? ganttDaysBetween(ganttIsoDate(rangeStart), opts.startDate) : null;
  const startDateLine = (startDateOffset != null && startDateOffset >= 0 && startDateOffset <= totalDays)
    ? `<div class="monday-gantt-due" style="left:${startDateOffset * dayWidth}px;background:var(--aewttr-green);" title="Project start date: ${escapeHtml(fmtDate(opts.startDate))}"></div>` : "";
  const dueDateOffset = opts.dueDate ? ganttDaysBetween(ganttIsoDate(rangeStart), opts.dueDate) : null;
  const dueDateLine = (dueDateOffset != null && dueDateOffset >= 0 && dueDateOffset <= totalDays)
    ? `<div class="monday-gantt-due" style="left:${dueDateOffset * dayWidth}px;" title="Project due date: ${escapeHtml(fmtDate(opts.dueDate))}"></div>` : "";

  const groupedByDivider = groupTrackerItems(tasks);
  const visibleIdSet = new Set(visibleTasks.map((t) => t.id));
  const hasDividers = groupedByDivider.dividers.length > 0;
  let bodyRows = "";

  function renderGanttTaskRows(taskList) {
    let html = "";
    taskList.forEach((t) => {
      if (isTrackerDivider(t)) return;
      const subitems = t.subtasks || [];
      const isOpen = !!expanded[t.id];
      const left = ganttDaysBetween(ganttIsoDate(rangeStart), t.start) * dayWidth;
      const width = Math.max((ganttDaysBetween(t.start, t.end) + 1) * dayWidth, dayWidth * 2);
      const taskDone = taskIsFullyDone(t);
      html += `
        <div class="monday-gantt-task-label ${taskDone ? "is-complete" : ""}" data-row-id="${t.id}">
          <button type="button" class="cl-drag-handle monday-gantt-drag"${tip("Drag to reorder")}><i class="bx bx-grid-vertical"></i></button>
          <button type="button" class="monday-gantt-complete" data-task-complete="${t.id}" aria-pressed="${taskDone ? "true" : "false"}"${tip(taskDone ? "Mark as not done" : "Mark done (includes all subitems)")} aria-label="Mark task done"><i class="bx bx-check"></i></button>
          ${subitems.length ? `<button type="button" class="monday-gantt-chevron" data-toggle="${t.id}"${tip(isOpen ? "Hide subitems" : "Show subitems")}><i class="bx bx-chevron-${isOpen ? "down" : "right"}"></i></button>` : `<span class="monday-gantt-chevron-spacer"></span>`}
          <div class="monday-gantt-task-copy">
            <span class="monday-gantt-task-name" contenteditable="true" spellcheck="false" data-gantt-field="title" data-task-id="${t.id}">${escapeHtml(t.title || "Untitled")}</span>
            <button type="button" class="monday-gantt-task-meta" data-open="${t.id}"${tip("Edit task dates and details")}>${ganttDateRangeLabel(t.start, t.end)} · ${escapeHtml(t.assignee || "Unassigned")}</button>
          </div>
          <button type="button" class="monday-gantt-edit" data-add-subtask="${t.id}"${tip("Add a subitem")} aria-label="Add subitem"><i class="bx bx-list-plus"></i></button>
          <button type="button" class="monday-gantt-edit" data-open="${t.id}"${tip("Edit task")} aria-label="Edit task"><i class="bx bx-edit"></i></button>
          <button type="button" class="monday-gantt-edit monday-gantt-edit--danger" data-delete-task="${t.id}"${tip("Delete this task")} aria-label="Delete task"><i class="bx bx-trash"></i></button>
        </div>
        <div class="monday-gantt-task-rail monday-gantt-rail" style="--gantt-day:${dayWidth}px;">
          ${todayLine}
          ${startDateLine}
          ${dueDateLine}
          <button type="button" class="monday-gantt-bar ${ganttBarClassForTask(t)}" data-open="${t.id}" data-drag-task="${t.id}" data-day-width="${dayWidth}" style="left:${left}px;width:${width}px;"${typeof ganttTip === "function" ? ganttTip(ganttTaskTipText(t)) : ""}>
            <span class="monday-gantt-bar-progress" style="width:${taskProgressPct(t)}%;"></span>
            <span class="gantt-resize-handle left" data-drag-task="${t.id}" data-mode="start"></span>
            <span class="gantt-resize-handle right" data-drag-task="${t.id}" data-mode="end"></span>
          </button>
          ${isOpen ? renderInlineSubtaskBars(t, rangeStart, dayWidth) : ""}
        </div>`;
      if (!isOpen) return;
      if (!subitems.length) {
        html += `
          <div class="monday-gantt-subtask-label monday-gantt-subtask-label--empty">
            <button type="button" class="monday-gantt-inline-add" data-add-subtask="${t.id}"${tip("Add a subitem")}><i class="bx bx-plus"></i> Add subitem</button>
          </div>
          <div class="monday-gantt-subtask-rail monday-gantt-rail monday-gantt-subtask-rail--empty" style="--gantt-day:${dayWidth}px;"></div>`;
        return;
      }
      subitems.forEach((rawSubtask, si) => {
        const subtask = normalizeTaskSubtask(t, rawSubtask);
        const subLeft = subtask.start && subtask.end ? ganttDaysBetween(ganttIsoDate(rangeStart), subtask.start) * dayWidth : null;
        const subWidth = subtask.start && subtask.end ? Math.max((ganttDaysBetween(subtask.start, subtask.end) + 1) * dayWidth, dayWidth * 1.5) : 120;
        const stateClass = subtaskStateClass(subtask);
        html += `
          <div class="monday-gantt-subtask-label ${subtask.done ? "is-complete" : ""}">
            <button type="button" class="monday-gantt-complete monday-gantt-complete--sub" data-subtask-complete="${t.id}:${si}" aria-pressed="${subtask.done ? "true" : "false"}"${tip(subtask.done ? "Mark as not done" : "Mark done")} aria-label="Mark subitem done"><i class="bx bx-check"></i></button>
            <button type="button" class="monday-gantt-subtask-copy" data-open-sub="${t.id}:${si}">
              <span class="monday-gantt-subtask-name">${escapeHtml(subtask.text || "Untitled subitem")}</span>
              <span class="monday-gantt-subtask-meta">${subtask.start && subtask.end ? ganttDateRangeLabel(subtask.start, subtask.end) : "No dates"}</span>
            </button>
            <button type="button" class="monday-gantt-edit monday-gantt-edit--sub" data-open-sub="${t.id}:${si}"${tip("Edit subitem")} aria-label="Edit subitem"><i class="bx bx-edit"></i></button>
            <button type="button" class="monday-gantt-subtask-delete" data-delete-subtask="${t.id}:${si}"${tip("Delete this subitem")} aria-label="Delete subitem"><i class="bx bx-trash"></i></button>
          </div>
          <div class="monday-gantt-subtask-rail monday-gantt-rail" style="--gantt-day:${dayWidth}px;">
            ${subLeft == null ? `
              <button type="button" class="gantt-subbar gantt-subbar-untimed ${stateClass} monday-gantt-subbar" data-open-sub="${t.id}:${si}" style="left:12px;width:${subWidth}px;"${typeof ganttTip === "function" ? ganttTip(ganttSubtaskTipText(t, subtask)) : ""}></button>` : `
              <button type="button" class="gantt-subbar ${stateClass} monday-gantt-subbar" data-open-sub="${t.id}:${si}" data-drag-sub="${t.id}:${si}" data-day-width="${dayWidth}" style="left:${subLeft}px;width:${subWidth}px;"${typeof ganttTip === "function" ? ganttTip(ganttSubtaskTipText(t, subtask)) : ""}>
                <span class="gantt-resize-handle left" data-drag-sub="${t.id}:${si}" data-mode="start"></span>
                <span class="gantt-resize-handle right" data-drag-sub="${t.id}:${si}" data-mode="end"></span>
              </button>`}
          </div>`;
      });
    });
    return html;
  }

  if (hasDividers) {
    const dividerState = opts.dividerCollapsed || {};
    if (opts.collapseDividersByDefault && window.AEWTTR.state._ganttDividerInit !== groupKey) {
      groupedByDivider.dividers.forEach(({ divider }) => {
        if (dividerState[divider.id] == null) dividerState[divider.id] = true;
      });
      window.AEWTTR.state._ganttDividerInit = groupKey;
    }
    groupedByDivider.dividers.forEach(({ divider, tasks: sectionTasks }) => {
      const sectionVisible = sectionTasks.filter((t) => visibleIdSet.has(t.id));
      if (!sectionVisible.length && visibleTasks.length !== trackerPlainTasks(tasks).length) return;
      const collapsed = !!dividerState[divider.id];
      const rag = String((divider.rag || (divider.metadata && divider.metadata.rag) || "Green")).replace(/\s+/g, "");
      const milestoneClass = divider.isMilestone ? " is-milestone" : "";
      bodyRows += `
        <div class="monday-gantt-group-label monday-gantt-divider-label${milestoneClass}">
          <button type="button" class="monday-gantt-group-toggle" data-toggle-divider="${divider.id}"${tip(collapsed ? "Expand divider" : "Collapse divider")}>
            <i class="bx bx-chevron-${collapsed ? "right" : "down"}"></i>
            <span class="tracker-divider-rag tracker-divider-rag--${escapeHtml(rag)}" aria-hidden="true"></span>
            <span class="monday-gantt-group-name">${escapeHtml(divider.title || "Untitled divider")}</span>
            ${divider.isMilestone ? `<span class="tracker-divider-milestone tracker-divider-milestone--compact"><i class="bx bxs-flag" aria-hidden="true"></i></span>` : ""}
            <span class="monday-gantt-group-count">${sectionVisible.length}</span>
          </button>
        </div>
        <div class="monday-gantt-group-rail monday-gantt-rail" style="--gantt-day:${dayWidth}px;"></div>`;
      if (!collapsed) bodyRows += renderGanttTaskRows(sectionVisible);
    });
    const ungroupedVisible = groupedByDivider.ungrouped.filter((t) => visibleIdSet.has(t.id));
    if (ungroupedVisible.length) {
      const collapsed = !!dividerState.__ungrouped__;
      bodyRows += `
        <div class="monday-gantt-group-label">
          <button type="button" class="monday-gantt-group-toggle" data-toggle-divider="__ungrouped__"${tip(collapsed ? "Expand" : "Collapse")}>
            <i class="bx bx-chevron-${collapsed ? "right" : "down"}"></i>
            <span class="monday-gantt-group-dot" style="background:var(--aewttr-border-strong);"></span>
            <span class="monday-gantt-group-name">Ungrouped</span>
            <span class="monday-gantt-group-count">${ungroupedVisible.length}</span>
          </button>
        </div>
        <div class="monday-gantt-group-rail monday-gantt-rail" style="--gantt-day:${dayWidth}px;"></div>`;
      if (!collapsed) bodyRows += renderGanttTaskRows(ungroupedVisible);
    }
  } else {
  const grouped = MONDAY_GANTT_GROUPS.map((group) => ({
    ...group,
    tasks: visibleTasks.filter((t) => !isTrackerDivider(t) && ganttGroupKeyForTask(t) === group.key)
  })).filter((group) => group.tasks.length);

  grouped.forEach((group) => {
    const isGroupCollapsed = !!collapsedGroups[group.key];
    bodyRows += `
      <div class="monday-gantt-group-label">
        <button type="button" class="monday-gantt-group-toggle" data-status-group="${group.key}"${tip(isGroupCollapsed ? "Expand group" : "Collapse group")}>
          <i class="bx bx-chevron-${isGroupCollapsed ? "right" : "down"}"></i>
          <span class="monday-gantt-group-dot" style="background:${group.dot};"></span>
          <span class="monday-gantt-group-name">${group.label}</span>
          <span class="monday-gantt-group-count">${group.tasks.length}</span>
        </button>
      </div>
      <div class="monday-gantt-group-rail monday-gantt-rail" style="--gantt-day:${dayWidth}px;"></div>`;

    if (isGroupCollapsed) return;
    bodyRows += renderGanttTaskRows(group.tasks);
  });
  }

  // Only real tasks drive the timeline window — dividers may lack dates.
  mount.innerHTML = `
    <div class="monday-gantt${meetingMode ? " monday-gantt--meeting" : ""}">
      <div class="monday-gantt-chartbar">
        <div class="monday-gantt-chartbar-left">
          <span class="monday-gantt-quarter">${quarterLabel}</span>
        </div>
        <div class="monday-gantt-chartbar-right">
          ${explicitWindow ? `
            <button type="button" class="monday-gantt-nav-btn" data-gantt-nav="prev" aria-label="Previous ${escapeHtml(rangeMode)}"><i class="bx bx-chevron-left"></i></button>
            <button type="button" class="monday-gantt-chartbtn" data-gantt-nav="today">Today</button>
            <button type="button" class="monday-gantt-nav-btn" data-gantt-nav="next" aria-label="Next ${escapeHtml(rangeMode)}"><i class="bx bx-chevron-right"></i></button>
          ` : `
            <button type="button" class="monday-gantt-chartbtn" data-gantt-zoom="fit"${tip("Fit timeline to tasks")}>Auto fit</button>
            <select class="monday-gantt-zoom-select" id="gantt-zoom-select-${mount.id || "main"}" aria-label="Zoom level">
              <option value="months" ${zoom === "months" ? "selected" : ""}>Months</option>
              <option value="weeks" ${zoom === "weeks" ? "selected" : ""}>Weeks</option>
              <option value="days" ${zoom === "days" ? "selected" : ""}>Days</option>
            </select>
          `}
          <div class="monday-gantt-range-tabs" role="tablist">
            ${[["day", "Day"], ["week", "Week"], ["month", "Month"], ["whole", "Whole Project"]].map(([key, label]) => `
              <button type="button" class="monday-gantt-range-tab ${rangeMode === key ? "active" : ""}" data-gantt-range="${key}" role="tab" aria-selected="${rangeMode === key}">${label}</button>
            `).join("")}
          </div>
        </div>
      </div>
      <div class="monday-gantt-scroll">
        <div class="monday-gantt-board" style="--gantt-day:${dayWidth}px;--gantt-info-width:${infoWidth}px;--gantt-timeline-width:${timelineWidth}px;">
          <div class="monday-gantt-head">
            <div class="monday-gantt-head-left"><span class="monday-gantt-head-title">Task</span><div class="monday-gantt-info-resizer"></div></div>
            <div class="monday-gantt-head-timeline">
              <div class="monday-gantt-months">${monthPills}</div>
              <div class="monday-gantt-ticks">${ticksHtml}</div>
            </div>
          </div>
          <div class="monday-gantt-body">
            ${bodyRows}
          </div>
        </div>
      </div>
    </div>`;

  const zoomSelect = $(`#gantt-zoom-select-${mount.id || "main"}`, mount);
  if (zoomSelect) {
    zoomSelect.addEventListener("change", () => {
      window.AEWTTR.state.ganttZoom[zoomKey] = zoomSelect.value;
      renderMondayGanttChart(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts);
    });
  }
  $all("[data-gantt-range]", mount).forEach((btn) => btn.addEventListener("click", () => {
    rangeState.mode = btn.dataset.ganttRange;
    if (rangeState.mode === "whole") rangeState.anchor = null;
    renderMondayGanttChart(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts);
  }));
  $all("[data-gantt-nav]", mount).forEach((btn) => btn.addEventListener("click", () => {
    if (btn.dataset.ganttNav === "today") {
      rangeState.anchor = null;
    } else {
      const delta = btn.dataset.ganttNav === "prev" ? -1 : 1;
      const currentAnchor = rangeState.anchor ? ganttParseDate(rangeState.anchor) : new Date();
      let nextAnchor;
      if (rangeMode === "month") {
        nextAnchor = new Date(currentAnchor.getFullYear(), currentAnchor.getMonth() + delta, 1);
      } else {
        const stepDays = rangeMode === "day" ? 1 : 7;
        nextAnchor = ganttAddDays(currentAnchor, delta * stepDays);
      }
      rangeState.anchor = ganttIsoDate(nextAnchor);
    }
    renderMondayGanttChart(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts);
  }));
  const fitBtn = $("[data-gantt-zoom=fit]", mount);
  if (fitBtn) fitBtn.addEventListener("click", () => {
    // Whole-project + weeks density, then resolve-day-width stretches columns
    // to fill the mount when the natural timeline would be a stub.
    rangeState.mode = "whole";
    rangeState.anchor = null;
    window.AEWTTR.state.ganttZoom[zoomKey] = "weeks";
    renderMondayGanttChart(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts);
  });
  $all("[data-status-group]", mount).forEach((btn) => btn.addEventListener("click", (e) => {
    e.stopPropagation();
    const key = btn.dataset.statusGroup;
    if (!window.AEWTTR.state.ganttGroupCollapsed[groupKey]) window.AEWTTR.state.ganttGroupCollapsed[groupKey] = {};
    window.AEWTTR.state.ganttGroupCollapsed[groupKey][key] = !window.AEWTTR.state.ganttGroupCollapsed[groupKey][key];
    renderMondayGanttChart(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts);
  }));
  $all("[data-toggle-divider]", mount).forEach((btn) => btn.addEventListener("click", (e) => {
    e.stopPropagation();
    const id = btn.dataset.toggleDivider;
    if (!opts.dividerCollapsed) opts.dividerCollapsed = {};
    opts.dividerCollapsed[id] = !opts.dividerCollapsed[id];
    renderMondayGanttChart(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts);
  }));
  $all("[data-toggle]", mount).forEach((elx) => elx.addEventListener("click", (e) => { e.stopPropagation(); onToggleExpand(elx.dataset.toggle); }));
  $all("[data-open]", mount).forEach((elx) => elx.addEventListener("click", (e) => {
    // .monday-gantt-edit is a shared style class for the edit/add-subitem/
    // delete icon buttons alike — it used to be in this exclusion list too,
    // which meant the Edit button (itself styled with that class) matched
    // its own exclusion and could never fire. Exclude by the OTHER buttons'
    // own data-attributes instead, which actually identifies them.
    if (e.target.closest(".gantt-resize-handle, .monday-gantt-task-name, .cl-drag-handle, .monday-gantt-chevron, .monday-gantt-complete, [data-delete-task], [data-delete-subtask], [data-add-subtask]")) return;
    if (mount.dataset.dragJustFinished) return;
    e.stopPropagation();
    onOpenEditor(elx.dataset.open);
  }));
  $all("[data-open-sub]", mount).forEach((elx) => elx.addEventListener("click", (e) => {
    if (mount.dataset.dragJustFinished) return;
    e.stopPropagation();
    const [taskId, subIndex] = elx.dataset.openSub.split(":");
    onOpenEditor(taskId, +subIndex);
  }));
  $all("[data-task-complete]", mount).forEach((btn) => btn.addEventListener("click", (e) => {
    e.stopPropagation();
    const task = tasks.find((t) => t.id === btn.dataset.taskComplete);
    if (!task) return;
    setTaskCompletionState(task, !taskIsFullyDone(task));
    Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource });
    redraw();
  }));
  $all("[data-subtask-complete]", mount).forEach((btn) => btn.addEventListener("click", (e) => {
    e.stopPropagation();
    const [taskId, subIndexText] = btn.dataset.subtaskComplete.split(":");
    const task = tasks.find((t) => t.id === taskId);
    const subtask = task && task.subtasks && task.subtasks[+subIndexText];
    if (!task || !subtask) return;
    setSubtaskCompletionState(subtask, !subtask.done);
    syncTaskStatusFromSubtasks(task);
    Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource });
    redraw();
  }));
  $all("[data-add-subtask]", mount).forEach((btn) => btn.addEventListener("click", async (e) => {
    e.stopPropagation();
    const task = tasks.find((t) => t.id === btn.dataset.addSubtask);
    if (!task) return;
    if (!task.subtasks) task.subtasks = [];
    task.subtasks.push(normalizeTaskSubtask(task, { text: "", start: task.start, end: "", relatedDocs: [] }));
    expanded[task.id] = true;
    Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource }).catch(() => {});
    onOpenEditor(task.id, task.subtasks.length - 1);
  }));
  $all("[data-delete-task]", mount).forEach((btn) => btn.addEventListener("click", async (e) => {
    e.stopPropagation();
    const task = tasks.find((t) => t.id === btn.dataset.deleteTask);
    if (!task) return;
    const ok = await confirmDialog({ title: "Delete task", message: `Delete "${task.title}"? This cannot be undone.`, confirmLabel: "Delete", danger: true });
    if (!ok) return;
    const idx = tasks.findIndex((t) => t.id === task.id);
    if (idx < 0) return;
    tasks.splice(idx, 1);
    toast("Task deleted", "success");
    redraw();
    Repo.remove("actionItem", task).catch(() => {
      tasks.splice(idx, 0, task);
      redraw();
    });
  }));
  $all("[data-delete-subtask]", mount).forEach((btn) => btn.addEventListener("click", async (e) => {
    e.stopPropagation();
    const [taskId, subIndexText] = btn.dataset.deleteSubtask.split(":");
    const task = tasks.find((t) => t.id === taskId);
    if (!task || !task.subtasks || !task.subtasks[+subIndexText]) return;
    const ok = await confirmDialog({ title: "Delete subitem", message: "Delete this subitem? This cannot be undone.", confirmLabel: "Delete", danger: true });
    if (!ok) return;
    const removed = task.subtasks.splice(+subIndexText, 1)[0];
    syncTaskStatusFromSubtasks(task);
    toast("Subitem deleted", "success");
    redraw();
    Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource }).catch(() => {
      task.subtasks.splice(+subIndexText, 0, removed);
      syncTaskStatusFromSubtasks(task);
      redraw();
    });
  }));

  wireMondayGanttTitleEdits(mount, tasks, redraw, onOpenEditor, opts);
  wireGanttBarDragging(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, { saveSource, projectCode: saveProjectCode });
  wireGanttRowReordering(mount, tasks, redraw);
  wireTrackerTableDragDrop(mount, tasks, redraw, { saveSource, projectCode: saveProjectCode });
  wireGanttContextMenu(mount, tasks, expanded, onOpenEditor, redraw, { saveSource, projectCode: saveProjectCode });

  const resizer = $(".monday-gantt-info-resizer", mount);
  if (resizer) {
    const boardEl = $(".monday-gantt-board", mount);
    const startWidthState = window.AEWTTR.state.ganttInfoWidth || {};
    window.AEWTTR.state.ganttInfoWidth = startWidthState;
    if (startWidthState[saveProjectCode]) {
      boardEl.style.setProperty("--gantt-info-width", `${startWidthState[saveProjectCode]}px`);
    }
    resizer.addEventListener("mousedown", (e) => {
      e.preventDefault();
      const startX = e.clientX;
      const computed = getComputedStyle(boardEl).getPropertyValue("--gantt-info-width");
      const startW = parseInt(computed) || 300;
      const onMove = (moveEv) => {
        const nextW = Math.max(200, startW + (moveEv.clientX - startX));
        boardEl.style.setProperty("--gantt-info-width", `${nextW}px`);
        startWidthState[saveProjectCode] = nextW;
      };
      const onUp = () => {
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
        // Re-render so day columns reflow into the new leftover timeline width.
        renderMondayGanttChart(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts);
      };
      document.addEventListener("mousemove", onMove);
      document.addEventListener("mouseup", onUp);
    });
  }

  wireMondayGanttFillResize(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts);
}

function wireMondayGanttFillResize(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts) {
  if (typeof ResizeObserver === "undefined" || !mount) return;
  if (mount._ganttFillResizeObserver) {
    mount._ganttFillResizeObserver.disconnect();
    mount._ganttFillResizeObserver = null;
  }
  let lastW = mount.clientWidth || 0;
  const ro = new ResizeObserver(() => {
    const w = mount.clientWidth || 0;
    if (Math.abs(w - lastW) < 4) return;
    lastW = w;
    if (mount._ganttFillResizeRaf) cancelAnimationFrame(mount._ganttFillResizeRaf);
    mount._ganttFillResizeRaf = requestAnimationFrame(() => {
      if (!mount.isConnected) return;
      renderMondayGanttChart(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts);
    });
  });
  ro.observe(mount);
  mount._ganttFillResizeObserver = ro;
}

function wireTrackerContextMenu(mount, tasks, onOpenEditor, onChange, opts) {
  opts = opts || {};
  const saveSource = opts.saveSource || "Tracker";
  const saveProjectCode = opts.projectCode || projectCodeForTaskList(tasks);
  const proj = opts.proj;
  const subitemExpanded = opts.subitemExpanded || {};

  function addSubtask(task, parentPath) {
    task.subtasks = task.subtasks || [];
    if (parentPath) {
      const parent = getSubtaskAtPath ? getSubtaskAtPath(task, parentPath) : null;
      if (!parent) return;
      const childPath = `${parentPath}.${(parent.subtasks || []).length}`;
      addNestedSubtask(parent, {
        text: "New subtask",
        start: parent.start || task.start,
        end: ""
      });
      subitemExpanded[`${task.id}:${parentPath}`] = true;
      window.AEWTTR.state.trackerFocusSubtask = { taskId: task.id, path: childPath };
    } else {
      const childPath = String(task.subtasks.length);
      task.subtasks.push(normalizeTaskSubtask(task, {
        text: "New subtask",
        start: task.start,
        end: ""
      }));
      window.AEWTTR.state.trackerFocusSubtask = { taskId: task.id, path: childPath };
    }
    if (opts.expanded) opts.expanded[task.id] = true;
    syncTaskStatusFromSubtasks(task);
    Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource }).catch(() => {});
    onChange();
  }

  mount.addEventListener("contextmenu", (e) => {
    const subRow = e.target.closest("tr.monday-sub-row-item[data-task-id][data-sub-path]");
    const taskRow = subRow ? null : e.target.closest("tr.monday-row--tracker[data-id]");
    if (!taskRow && !subRow) return;
    e.preventDefault();

    const taskId = subRow ? subRow.dataset.taskId : taskRow.dataset.id;
    const subPath = subRow ? subRow.dataset.subPath : null;
    const task = tasks.find((t) => t.id === taskId);
    if (!task) return;

    const items = [];
    if (subPath != null) {
      const sub = getSubtaskAtPath ? getSubtaskAtPath(task, subPath) : null;
      items.push({ label: "Expand view", icon: "bx-expand-alt", action: () => openTaskExpandModal(task, tasks, proj, onChange, { subPath }) });
      items.push({ separator: true });
      items.push({ label: "Edit subitem", icon: "bx-pencil", action: () => onOpenEditor(taskId, subPath) });
      items.push({ label: "Add nested subtask", icon: "bx-list-plus", action: () => addSubtask(task, subPath) });
      items.push({ label: sub && sub.done ? "Mark subitem not done" : "Mark subitem done", icon: "bx-check-circle", action: () => {
        if (!sub) return;
        setSubtaskCompletionState(sub, !sub.done);
        syncTaskStatusFromSubtasks(task);
        Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource }).catch(() => {});
        onChange();
      }});
      items.push({ separator: true });
      items.push({ label: "Delete subtask", icon: "bx-trash", danger: true, action: async () => {
        if (!sub) return;
        const ok = await confirmDialog({ title: "Delete subtask", message: `Delete "${sub.text || "this subtask"}"${(sub.subtasks || []).length ? " and its nested subtasks" : ""}?`, confirmLabel: "Delete", danger: true });
        if (!ok) return;
        removeSubtaskAtPath(task, subPath);
        syncTaskStatusFromSubtasks(task);
        Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource }).catch(() => {});
        toast("Subtask deleted", "success");
        onChange();
      }});
    } else {
      items.push({ label: "Expand view", icon: "bx-expand-alt", action: () => openTaskExpandModal(task, tasks, proj, onChange) });
      items.push({ separator: true });
      items.push({ label: "Edit task", icon: "bx-pencil", action: () => onOpenEditor(taskId) });
      items.push({ label: "Add subtask", icon: "bx-list-plus", action: () => addSubtask(task, "") });
      items.push({ label: taskIsFullyDone(task) ? "Mark not done" : "Mark done", icon: "bx-check-circle", action: () => {
        setTaskCompletionState(task, !taskIsFullyDone(task));
        Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource }).catch(() => {});
        onChange();
      }});
      items.push({ separator: true });
      items.push({ label: "Delete task", icon: "bx-trash", danger: true, action: async () => {
        const ok = await confirmDialog({ title: "Delete task", message: `Delete "${task.title}"?`, confirmLabel: "Delete", danger: true });
        if (!ok) return;
        const idx = tasks.findIndex((t) => t.id === task.id);
        if (idx >= 0) tasks.splice(idx, 1);
        toast("Task deleted", "success");
        onChange();
        Repo.remove("actionItem", task).catch(() => { tasks.splice(idx, 0, task); onChange(); });
      }});
    }
    if (typeof showContextMenu === "function") showContextMenu(e.clientX, e.clientY, items);
  });
}

function openTaskExpandModal(task, tasks, proj, onSave, opts) {
  opts = opts || {};
  const pct = typeof taskProgressPct === "function" ? taskProgressPct(task) : 0;
  const subs = (typeof flattenTaskSubitems === "function" ? flattenTaskSubitems(task.subtasks || []) : (task.subtasks || []));
  const doneCount = subs.filter((s) => s.done).length;
  const progressLabel = subs.length ? `${doneCount}/${subs.length} done` : "No subitems";

  const infoHtml = `
    <h2 class="task-expand-title">${escapeHtml(task.title || "Untitled task")}</h2>
    <div class="task-expand-badges">
      ${typeof statusPill === "function" ? statusPill(task.status || "Not Started") : ""}
      ${task.health ? `<span class="monday-health-badge monday-health-${String(task.health).replace(/\s+/g, "-")}">${escapeHtml(task.health)}</span>` : ""}
    </div>
    <div class="task-expand-meta">
      <div class="task-expand-meta-item"><label>Assignee</label><span>${escapeHtml(task.assignee || "—")}</span></div>
      <div class="task-expand-meta-item"><label>Start</label><span>${escapeHtml(typeof fmtDate === "function" ? fmtDate(task.start) : (task.start || "—"))}</span></div>
      <div class="task-expand-meta-item"><label>End</label><span>${escapeHtml(typeof fmtDate === "function" ? fmtDate(task.end) : (task.end || "—"))}</span></div>
      <div class="task-expand-meta-item"><label>Est. Effort</label><span>${task.estimatedEffort ? `${escapeHtml(String(task.estimatedEffort))} hrs` : "—"}</span></div>
    </div>
    ${subs.length ? `
      <div class="task-expand-progress-row">
        <div class="monday-progress-wrap">
          <div class="monday-progress-bar"><span style="width:${pct}%"></span></div>
          <span class="monday-progress-label">${pct}% · ${progressLabel}</span>
        </div>
      </div>` : ""}
    ${subs.length ? `
      <div class="task-expand-section-label">Subitems</div>
      <div class="task-expand-subtask-list">
        ${subs.map((s) => `
          <div class="task-expand-subtask-item${s.done ? " is-done" : ""}">
            <i class="bx ${s.done ? "bx-check-circle" : "bx-circle"}"></i>
            <span>${escapeHtml(s.text || "Untitled")}</span>
            ${s.assignee ? `<span style="margin-left:auto;font-size:11px;color:var(--aewttr-muted)">${escapeHtml(s.assignee)}</span>` : ""}
          </div>`).join("")}
      </div>` : ""}
  `;

  const modal = openModal(`
    <div class="aewttr-modal-head">
      <h3><i class="bx bx-task" style="margin-right:6px;"></i> Task details</h3>
      <button class="aewttr-modal-close" type="button" aria-label="Close">&times;</button>
    </div>
    <div class="task-expand-body">
      <div class="task-expand-info">${infoHtml}</div>
      <div class="task-expand-notes">
        <div class="task-expand-notes-head"><i class="bx bx-chat"></i> Notes &amp; updates</div>
        <div class="task-notes-chat-body" id="tex-notes-body"></div>
        <div class="task-notes-input-row">
          <textarea class="task-notes-input" id="tex-notes-input" placeholder="Post an update — Enter to send…" rows="1"></textarea>
          <button type="button" class="btn-aewttr btn-aewttr-sm task-notes-send" id="tex-notes-send"><i class="bx bx-send"></i></button>
        </div>
      </div>
    </div>
  `, { className: "task-expand-modal" });

  const chatBody = $("#tex-notes-body", modal);
  const input = $("#tex-notes-input", modal);
  let editingId = null;

  function renderNotes() {
    if (!chatBody) return;
    const notes = ((task.notes || []).slice()).reverse();
    chatBody.innerHTML = notes.length
      ? notes.map((n) => {
          const isMine = typeof isNoteAuthor === "function" ? isNoteAuthor(n) : false;
          const isEditing = editingId === n.id;
          return `
            <div class="task-notes-bubble-row ${isMine ? "mine" : ""}">
              <div class="task-notes-bubble">
                <div class="task-notes-bubble-meta">
                  <strong>${escapeHtml(n.author || "Unknown")}</strong>
                  <span>${typeof formatNoteTimestamp === "function" ? escapeHtml(formatNoteTimestamp(n)) : escapeHtml(n.date || "")}</span>
                  ${isMine && !isEditing ? `<span class="task-notes-bubble-actions">
                    <button type="button" data-tex-edit-note="${n.id}"><i class="bx bx-pencil"></i></button>
                    <button type="button" data-tex-del-note="${n.id}"><i class="bx bx-trash"></i></button>
                  </span>` : ""}
                </div>
                ${isEditing
                  ? `<textarea class="task-notes-edit-input" id="tex-edit-${n.id}">${escapeHtml(n.text)}</textarea>
                     <div class="task-notes-edit-actions">
                       <button class="btn-aewttr-ghost btn-aewttr-sm" data-tex-cancel-edit="${n.id}">Cancel</button>
                       <button class="btn-aewttr btn-aewttr-sm" data-tex-save-edit="${n.id}">Save</button>
                     </div>`
                  : `<div class="task-notes-bubble-text">${escapeHtml(n.text)}</div>`}
              </div>
            </div>`;
        }).join("")
      : `<div class="task-notes-empty">No updates yet.</div>`;
    if (!editingId && chatBody) chatBody.scrollTop = chatBody.scrollHeight;
    wireNoteActions();
  }

  function wireNoteActions() {
    $all("[data-tex-edit-note]", chatBody).forEach((btn) => btn.addEventListener("click", () => {
      editingId = btn.dataset.texEditNote;
      renderNotes();
      const el = $(`#tex-edit-${editingId}`, chatBody);
      if (el) { el.focus(); el.setSelectionRange(el.value.length, el.value.length); }
    }));
    $all("[data-tex-cancel-edit]", chatBody).forEach((btn) => btn.addEventListener("click", () => { editingId = null; renderNotes(); }));
    $all("[data-tex-save-edit]", chatBody).forEach((btn) => btn.addEventListener("click", () => {
      const id = btn.dataset.texSaveEdit;
      const el = $(`#tex-edit-${id}`, chatBody);
      const text = el ? el.value.trim() : "";
      if (!text) return;
      const note = (task.notes || []).find((n) => n.id === id);
      if (!note) return;
      note.text = text;
      note.editedAt = new Date().toISOString();
      editingId = null;
      saveNotes();
      renderNotes();
    }));
    $all("[data-tex-del-note]", chatBody).forEach((btn) => btn.addEventListener("click", () => {
      const id = btn.dataset.texDelNote;
      task.notes = (task.notes || []).filter((n) => n.id !== id);
      saveNotes();
      renderNotes();
    }));
  }

  function postNote(text) {
    if (!text.trim()) return;
    if (!task.notes) task.notes = [];
    const note = { id: typeof uid === "function" ? uid("tn") : String(Date.now()), author: typeof currentUserName === "function" ? currentUserName() : "Me", date: new Date().toISOString().slice(0, 10), time: new Date().toTimeString().slice(0, 5), text: text.trim() };
    task.notes.unshift(note);
    saveNotes();
    if (input) input.value = "";
    renderNotes();
  }

  function saveNotes() {
    const saveProjectCode = proj ? proj.id : (typeof projectCodeForTaskList === "function" ? projectCodeForTaskList(tasks) : "");
    Repo.save("actionItem", task, { projectCode: saveProjectCode, source: "Tracker" }).catch(() => {});
    if (typeof onSave === "function") onSave();
  }

  if (input) {
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); postNote(input.value); }
    });
  }
  const sendBtn = $("#tex-notes-send", modal);
  if (sendBtn) sendBtn.addEventListener("click", () => postNote(input ? input.value : ""));

  renderNotes();
}

function closeGanttContextMenu() {
  const menu = document.getElementById("gantt-context-menu");
  if (menu) menu.remove();
}

function wireGanttContextMenu(mount, tasks, expanded, onOpenEditor, redraw, opts) {
  const saveSource = opts.saveSource || "Tracker";
  const saveProjectCode = opts.projectCode || projectCodeForTaskList(tasks);

  mount.addEventListener("contextmenu", (e) => {
    const taskRow = e.target.closest(".monday-gantt-task-label[data-row-id]");
    const subRow = e.target.closest(".monday-gantt-subtask-label");
    if (!taskRow && !subRow) return;
    e.preventDefault();
    closeGanttContextMenu();

    let taskId = taskRow ? taskRow.dataset.rowId : null;
    let subIndex = null;
    if (subRow) {
      const ref = subRow.querySelector("[data-delete-subtask], [data-open-sub], [data-subtask-complete]");
      const key = ref && (ref.dataset.deleteSubtask || ref.dataset.openSub || ref.dataset.subtaskComplete);
      if (key && key.includes(":")) {
        const parts = key.split(":");
        taskId = parts[0];
        subIndex = +parts[1];
      }
    }
    const task = tasks.find((t) => t.id === taskId);
    if (!task) return;

    const menu = document.createElement("div");
    menu.id = "gantt-context-menu";
    menu.className = "gantt-context-menu";
    menu.style.left = `${Math.min(e.clientX, window.innerWidth - 200)}px`;
    menu.style.top = `${Math.min(e.clientY, window.innerHeight - 180)}px`;

    function addItem(label, onClick, danger) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = `gantt-context-menu-item${danger ? " is-danger" : ""}`;
      btn.textContent = label;
      btn.addEventListener("click", (ev) => {
        ev.stopPropagation();
        closeGanttContextMenu();
        onClick();
      });
      menu.appendChild(btn);
    }

    if (subIndex != null && !Number.isNaN(subIndex)) {
      const sub = task.subtasks && task.subtasks[subIndex];
      addItem("Edit subitem", () => onOpenEditor(taskId, subIndex));
      addItem(sub && sub.done ? "Mark subitem not done" : "Mark subitem done", () => {
        const sub = task.subtasks && task.subtasks[subIndex];
        if (!sub) return;
        setSubtaskCompletionState(sub, !sub.done);
        syncTaskStatusFromSubtasks(task);
        Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource });
        redraw();
      });
      addItem("Delete subitem", async () => {
        const ok = await confirmDialog({ title: "Delete subitem", message: "Delete this subitem?", confirmLabel: "Delete", danger: true });
        if (!ok) return;
        task.subtasks.splice(subIndex, 1);
        syncTaskStatusFromSubtasks(task);
        Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource });
        redraw();
      }, true);
    } else {
      addItem("Edit task", () => onOpenEditor(taskId));
      addItem("Add subitem", async () => {
        if (!task.subtasks) task.subtasks = [];
        task.subtasks.push(normalizeTaskSubtask(task, { text: "", start: task.start, end: "", relatedDocs: [] }));
        expanded[task.id] = true;
        Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource });
        onOpenEditor(task.id, task.subtasks.length - 1);
      });
      addItem(taskIsFullyDone(task) ? "Mark task not done" : "Mark task done", () => {
        setTaskCompletionState(task, !taskIsFullyDone(task));
        Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource });
        redraw();
      });
      addItem("Delete task", async () => {
        const ok = await confirmDialog({ title: "Delete task", message: `Delete "${task.title}"?`, confirmLabel: "Delete", danger: true });
        if (!ok) return;
        const idx = tasks.findIndex((t) => t.id === task.id);
        if (idx >= 0) tasks.splice(idx, 1);
        redraw();
        Repo.remove("actionItem", task).catch(() => {
          tasks.splice(idx, 0, task);
          redraw();
        });
      }, true);
    }

    document.body.appendChild(menu);
    const dismiss = () => closeGanttContextMenu();
    setTimeout(() => {
      document.addEventListener("click", dismiss, { once: true });
      document.addEventListener("contextmenu", dismiss, { once: true });
    }, 0);
  });
}

/* Second header row under the month pills — the actual day/week/hour
   indicators the month row alone doesn't provide. "Day" mode gets an hour
   ruler across the single selected day (tasks are date-only, but the ruler
   still orients you within that day the way the other range modes orient
   you within a week/month). Week and Month modes get one label per day,
   spaced by a stride so labels never collide at narrow zoom levels. */
function buildMondayGanttTicks(rangeMode, rangeStart, dayWidth, totalDays) {
  if (rangeMode === "day") {
    const hours = [0, 3, 6, 9, 12, 15, 18, 21];
    const hourWidth = dayWidth / 8;
    return hours.map((h) => {
      const left = (h / 24) * dayWidth;
      const label = h === 0 ? "12 AM" : h < 12 ? `${h} AM` : h === 12 ? "12 PM" : `${h - 12} PM`;
      return `<div class="monday-gantt-tick" style="left:${left}px;width:${hourWidth}px;"><span>${label}</span></div>`;
    }).join("");
  }
  const stride = dayWidth >= 18 ? 1 : dayWidth >= 8 ? 7 : 30;
  const todayIso = ganttIsoDate(new Date());
  let html = "";
  for (let i = 0; i < totalDays; i += stride) {
    const d = ganttAddDays(rangeStart, i);
    const iso = ganttIsoDate(d);
    const isToday = iso === todayIso;
    const isWeekend = stride === 1 && (d.getDay() === 0 || d.getDay() === 6);
    const label = stride !== 1
      ? d.toLocaleDateString("en-US", { month: "short", day: "numeric" })
      : dayWidth >= 40
        ? `${d.toLocaleDateString("en-US", { weekday: "short" })} ${d.getDate()}`
        : `${d.getDate()}`;
    const width = Math.min(stride, totalDays - i) * dayWidth;
    html += `<div class="monday-gantt-tick${isToday ? " is-today" : ""}${isWeekend ? " is-weekend" : ""}" style="left:${i * dayWidth}px;width:${width}px;"><span>${label}</span></div>`;
  }
  return html;
}

function buildMondayGanttMonthPills(rangeStart, totalDays, dayWidth) {
  const now = new Date();
  let html = "";
  let i = 0;
  while (i < totalDays) {
    const d = ganttAddDays(rangeStart, i);
    const monthEnd = new Date(d.getFullYear(), d.getMonth() + 1, 0);
    const endIndex = Math.min(ganttDaysBetween(ganttIsoDate(rangeStart), ganttIsoDate(monthEnd)) + 1, totalDays) - 1;
    const spanDays = endIndex - i + 1;
    const isCurrent = d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
    const label = d.toLocaleDateString("en-US", { month: "long" });
    html += `<div class="monday-gantt-month${isCurrent ? " is-current" : ""}" style="left:${i * dayWidth}px;width:${spanDays * dayWidth}px;"><span>${label}</span></div>`;
    i = endIndex + 1;
  }
  return html;
}

function buildMondayGanttDayGrid(totalDays, dayWidth) {
  let html = "";
  for (let i = 0; i <= totalDays; i++) {
    const showLine = i % 7 === 0;
    html += `<div class="monday-gantt-dayline${showLine ? " monday-gantt-dayline--week" : ""}" style="left:${i * dayWidth}px;"></div>`;
  }
  return html;
}

function buildMondayGanttQuarterLabel(rangeStart, rangeEnd) {
  const start = ganttParseDate(ganttIsoDate(rangeStart));
  const q = Math.floor(start.getMonth() / 3) + 1;
  const year = start.getFullYear();
  const endYear = ganttParseDate(ganttIsoDate(rangeEnd)).getFullYear();
  return endYear !== year ? `Q${q} ${year} – ${endYear}` : `Q${q} ${year}`;
}

function renderSimpleTrackerGantt(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts) {
  opts = opts || {};
  const projectKey = opts.projectCode || projectCodeForTaskList(tasks) || "tracker";
  const state = window.AEWTTR.state;
  state.simpleGanttMonth = state.simpleGanttMonth || {};
  const today = new Date();
  const selected = state.simpleGanttMonth[projectKey] ? new Date(`${state.simpleGanttMonth[projectKey]}-01T12:00:00`) : today;
  const monthStart = new Date(selected.getFullYear(), selected.getMonth(), 1);
  const gridStart = new Date(monthStart);
  gridStart.setDate(gridStart.getDate() - gridStart.getDay());
  const totalDays = 42;
  const dayWidth = 34;
  const gridEnd = ganttAddDays(gridStart, totalDays - 1);
  const rows = [];
  const allowed = new Set((visibleTasks || tasks).filter((t) => !isTrackerDivider(t)).map((t) => t.id));

  function addRows(list, task, depth, prefix) {
    (list || []).forEach((subtask, index) => {
      const path = prefix ? `${prefix}.${index}` : String(index);
      rows.push({ task, item: subtask, path, depth, isSubtask: true });
      if (expanded[`${task.id}:${path}`]) addRows(subtask.subtasks || [], task, depth + 1, path);
    });
  }
  tasks.filter((task) => !isTrackerDivider(task) && allowed.has(task.id)).forEach((task) => {
    rows.push({ task, item: task, path: "", depth: 0, isSubtask: false });
    if (expanded[task.id]) addRows(task.subtasks || [], task, 1, "");
  });

  const days = Array.from({ length: totalDays }, (_, index) => ganttAddDays(gridStart, index));
  const todayKey = ganttIsoDate(today);
  const dayHead = days.map((day) => `<div class="tracker-gantt-day${ganttIsoDate(day) === todayKey ? " is-today" : ""}"><span>${day.toLocaleDateString("en-US", { weekday: "narrow" })}</span><strong>${day.getDate()}</strong></div>`).join("");
  const gridLines = days.map((day) => `<i class="tracker-gantt-gridline${ganttIsoDate(day) === todayKey ? " is-today" : ""}"></i>`).join("");

  function rowHtml(row) {
    const item = row.item;
    const start = item.start || ganttIsoDate(today);
    const end = item.end || "";
    const startOffset = ganttDaysBetween(ganttIsoDate(gridStart), start);
    const rawEndOffset = end ? ganttDaysBetween(ganttIsoDate(gridStart), end) : startOffset;
    const visible = rawEndOffset >= 0 && startOffset < totalDays;
    const left = Math.max(startOffset, 0) * dayWidth;
    const width = end ? Math.max((Math.min(rawEndOffset, totalDays - 1) - Math.max(startOffset, 0) + 1) * dayWidth, dayWidth) : 12;
    const completed = row.isSubtask ? !!item.done : taskIsFullyDone(item);
    const hasChildren = row.isSubtask ? (item.subtasks || []).length > 0 : (item.subtasks || []).length > 0;
    const isOpen = row.isSubtask ? !!expanded[`${row.task.id}:${row.path}`] : !!expanded[row.task.id];
    const label = item.text || item.title || "Untitled task";
    const dateLabel = end ? `${fmtDate(start)} – ${fmtDate(end)}` : `${fmtDate(start)} · Open`;
    const openAttrs = row.isSubtask ? `data-simple-open-sub="${row.task.id}:${row.path}"` : `data-simple-open="${row.task.id}"`;
    const toggleAttrs = row.isSubtask ? `data-simple-toggle-sub="${row.task.id}:${row.path}"` : `data-simple-toggle="${row.task.id}"`;
    return `<div class="tracker-gantt-row${row.isSubtask ? " is-subtask" : ""}${completed ? " is-complete" : ""}" style="--gantt-depth:${row.depth}">
      <div class="tracker-gantt-label">
        ${hasChildren ? `<button type="button" class="tracker-gantt-toggle" ${toggleAttrs} aria-label="${isOpen ? "Collapse" : "Expand"}"><i class="bx bx-chevron-${isOpen ? "down" : "right"}"></i></button>` : `<span class="tracker-gantt-toggle-spacer"></span>`}
        <button type="button" class="tracker-gantt-name" ${openAttrs}><span>${escapeHtml(label)}</span><small>${escapeHtml(dateLabel)}</small></button>
      </div>
      <div class="tracker-gantt-track">${gridLines}${visible ? `<button type="button" class="tracker-gantt-bar${end ? "" : " is-open"}${completed ? " is-complete" : ""}" ${openAttrs} style="left:${left}px;width:${width}px;" title="${escapeHtml(`${label} — ${dateLabel}`)}">${end ? `<span>${escapeHtml(label)}</span>` : ""}</button>` : ""}</div>
    </div>`;
  }

  mount.innerHTML = `<section class="tracker-gantt-shell">
    <header class="tracker-gantt-toolbar">
      <div><strong>${monthStart.toLocaleDateString("en-US", { month: "long", year: "numeric" })}</strong><span>${rows.length} visible item${rows.length === 1 ? "" : "s"}</span></div>
      <div class="tracker-gantt-nav"><button type="button" data-simple-gantt-nav="prev" aria-label="Previous month"><i class="bx bx-chevron-left"></i></button><button type="button" data-simple-gantt-nav="today">Today</button><button type="button" data-simple-gantt-nav="next" aria-label="Next month"><i class="bx bx-chevron-right"></i></button></div>
    </header>
    <div class="tracker-gantt-scroll"><div class="tracker-gantt-board" style="--gantt-days:${totalDays};--gantt-day-width:${dayWidth}px">
      <div class="tracker-gantt-head"><div class="tracker-gantt-head-label">Work item</div><div class="tracker-gantt-days">${dayHead}</div></div>
      <div class="tracker-gantt-body">${rows.length ? rows.map(rowHtml).join("") : `<div class="tracker-gantt-empty">No tasks match the current filter.</div>`}</div>
    </div></div>
  </section>`;

  $all("[data-simple-gantt-nav]", mount).forEach((button) => button.addEventListener("click", () => {
    const date = new Date(monthStart);
    if (button.dataset.simpleGanttNav === "today") date.setTime(today.getTime());
    else date.setMonth(date.getMonth() + (button.dataset.simpleGanttNav === "prev" ? -1 : 1));
    state.simpleGanttMonth[projectKey] = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    renderSimpleTrackerGantt(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts);
  }));
  $all("[data-simple-toggle]", mount).forEach((button) => button.addEventListener("click", () => onToggleExpand(button.dataset.simpleToggle)));
  $all("[data-simple-toggle-sub]", mount).forEach((button) => button.addEventListener("click", () => {
    const raw = button.dataset.simpleToggleSub;
    const split = raw.indexOf(":");
    const key = `${raw.slice(0, split)}:${raw.slice(split + 1)}`;
    expanded[key] = !expanded[key];
    redraw();
  }));
  $all("[data-simple-open]", mount).forEach((button) => button.addEventListener("click", () => onOpenEditor(button.dataset.simpleOpen)));
  $all("[data-simple-open-sub]", mount).forEach((button) => button.addEventListener("click", () => {
    const raw = button.dataset.simpleOpenSub;
    const split = raw.indexOf(":");
    onOpenEditor(raw.slice(0, split), raw.slice(split + 1));
  }));
}

function renderGanttChart(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts) {
  renderSimpleTrackerGantt(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts);
}

function wireMondayGanttTitleEdits(mount, tasks, redraw, onOpenEditor, opts) {
  const saveSource = opts.saveSource || "Tracker";
  const saveProjectCode = opts.projectCode || projectCodeForTaskList(tasks);
  $all(".monday-gantt-task-name[data-gantt-field], .monday-gantt-group-name[data-gantt-field]", mount).forEach((node) => {
    node.addEventListener("mousedown", (e) => e.stopPropagation());
    node.addEventListener("click", (e) => e.stopPropagation());
    node.addEventListener("dblclick", (e) => {
      e.stopPropagation();
      const task = tasks.find((t) => t.id === node.dataset.taskId);
      if (task && typeof onOpenEditor === "function") onOpenEditor(task.id);
    });
    node.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); node.blur(); }
    });
    node.addEventListener("blur", () => {
      const task = tasks.find((t) => t.id === node.dataset.taskId);
      if (!task) return;
      const value = node.textContent.trim();
      if (!value) { node.textContent = task.title || "Untitled"; return; }
      task.title = value;
      syncTaskStatusFromSubtasks(task);
      Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource });
      if (typeof redraw === "function") redraw();
    });
  });
}

function taskIsFullyDone(task) {
  if (!task) return false;
  const subs = flattenTaskSubitems((task.subtasks) || []);
  if (!subs.length) return task.status === "Done";
  return taskProgressPct(task) === 100;
}

function setSubtaskCompletionState(subtask, completed) {
  if (!subtask) return;
  subtask.done = completed;
  if (completed) subtask.end = new Date().toISOString().slice(0, 10);
}

function setTaskCompletionState(task, completed) {
  if (!task) return;
  if (task.subtasks && task.subtasks.length) {
    walkNestedSubtasks(task.subtasks, (subtask) => setSubtaskCompletionState(subtask, completed));
    syncTaskStatusFromSubtasks(task);
  } else {
    task.status = completed ? "Done" : "Not Started";
  }
  if (completed) { task.health = "On Track"; task.end = new Date().toISOString().slice(0, 10); }
}

function buildGanttMonthTicks(rangeStart, totalDays, dayWidth) {
  let html = "";
  let i = 0;
  while (i < totalDays) {
    const d = ganttAddDays(rangeStart, i);
    const monthEnd = new Date(d.getFullYear(), d.getMonth() + 1, 0);
    const endIndex = Math.min(ganttDaysBetween(ganttIsoDate(rangeStart), ganttIsoDate(monthEnd)) + 1, totalDays) - 1;
    const spanDays = endIndex - i + 1;
    const label = d.toLocaleDateString("en-US", { month: "long", year: i === 0 ? "numeric" : undefined });
    html += `<div class="gantt-month-tick" style="left:${i * dayWidth}px;width:${spanDays * dayWidth}px;">${label}</div>`;
    i = endIndex + 1;
  }
  return html;
}

function renderLegacyGanttChart(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts) {
  opts = opts || {};
  const mondayStyle = opts.mondayStyle !== false;
  const meetingMode = !!opts.meetingMode;
  const saveSource = opts.saveSource || "Tracker";
  const saveProjectCode = opts.projectCode || projectCodeForTaskList(tasks);
  visibleTasks = visibleTasks || tasks;
  if (!tasks.length) {
    mount.innerHTML = `<div class="gantt-empty">No tasks yet. Use "+ Add Task" to start building this project's plan.</div>`;
    return;
  }
  if (!visibleTasks.length) {
    mount.innerHTML = `<div class="gantt-empty">No tasks match this assignee filter.</div>`;
    return;
  }
  const dayWidth = 26;
  const pcode = projectCodeForTaskList(tasks);
  const memberListId = `gantt-members-${Math.random().toString(36).slice(2, 8)}`;
  const members = (window.AEWTTR.db && window.AEWTTR.db.members) || [];
  const starts = visibleTasks.map(t => ganttParseDate(t.start).getTime());
  const ends = visibleTasks.map(t => ganttParseDate(t.end).getTime());
  const rangeStart = ganttAddDays(new Date(Math.min(...starts)), -2);
  const rangeEnd = ganttAddDays(new Date(Math.max(...ends)), 2);
  const totalDays = Math.max(ganttDaysBetween(ganttIsoDate(rangeStart), ganttIsoDate(rangeEnd)), 7) + 1;
  const timelineWidth = totalDays * dayWidth;

  let ticks = "";
  for (let i = 0; i <= totalDays; i += 7) {
    const d = ganttAddDays(rangeStart, i);
    ticks += `<div class="gantt-tick" style="left:${i * dayWidth}px;">${d.toLocaleDateString("en-US", { month: "short", day: "numeric" })}</div>`;
  }
  const monthTicks = buildGanttMonthTicks(rangeStart, totalDays + 1, dayWidth);
  const todayOffset = ganttDaysBetween(ganttIsoDate(rangeStart), ganttIsoDate(new Date()));
  const todayLine = (todayOffset >= 0 && todayOffset <= totalDays) ? `<div class="gantt-today" style="left:${todayOffset * dayWidth}px;"></div>` : "";

  function mondayMetaRow(t) {
    return `
      <span class="gantt-row-dates">${ganttDateRangeLabel(t.start, t.end)}</span>
      <span class="gantt-row-owner-label">${escapeHtml(t.assignee || "Unassigned")}</span>`;
  }

  let rows = "";
  visibleTasks.forEach(t => {
    const left = ganttDaysBetween(ganttIsoDate(rangeStart), t.start) * dayWidth;
    const width = Math.max((ganttDaysBetween(t.start, t.end) + 1) * dayWidth, 18);
    const statusClass = ganttBarClassForTask(t);
    const isOpen = !!expanded[t.id];
    const subtrackCount = (t.subtasks || []).length;
    // Row height must scale with the real subtask count — it used to cap at
    // 3/4 subtasks (Math.min), so any task with more than that had its 5th+
    // subtask bar drawn past the row's bottom edge and overlapping the next
    // task row instead of the row growing to fit (renderInlineSubtaskBars
    // positions every subtask at top:index*16px with no cap of its own).
    const rowHeight = meetingMode ? 44 + (subtrackCount ? subtrackCount * 12 + 4 : 0) : (mondayStyle ? 52 : 52) + (subtrackCount ? subtrackCount * 16 + 8 : 0);
    const barProgress = taskProgressPct(t);
    rows += `
      <div class="gantt-cell-label gantt-cell-label--monday" data-row-id="${t.id}" style="min-height:${rowHeight}px;">
        <div class="gantt-row-title-wrap">
          ${mondayStyle ? "" : `<button class="cl-drag-handle"${tip("Drag to reorder")} type="button"><i class="bx bx-grid-vertical"></i></button>`}
          <button class="gantt-expand-btn" data-toggle="${t.id}"${tip(isOpen ? "Hide subtasks" : "Show subtasks")}><i class="bx bx-chevron-${isOpen ? "down" : "right"}"></i></button>
          ${mondayStyle ? "" : issueTypeIcon(t)}
          ${mondayStyle
            ? `<input type="text" class="gantt-inline-text gantt-inline-title" data-gantt-field="title" data-task-id="${t.id}" value="${escapeHtml(t.title || "")}"${tip("Task title")}>`
            : `<button class="gantt-row-title gantt-open-link" data-open="${t.id}">${escapeHtml(t.title)}</button>`}
        </div>
        <div class="gantt-row-meta ${mondayStyle ? "gantt-row-meta--monday" : ""}">
          ${mondayStyle ? mondayMetaRow(t) : `${issueKey(pcode, t)}${issuePriority(t.priority)}${issuePoints(t)}
          ${userAvatarHtml(t.assignee || "Unassigned", memberEmailForPerson(t.assignee), 18)}
          <span class="gantt-row-owner">${escapeHtml(t.assignee || "Unassigned")}</span>
          <span style="font-size:11.5px;color:var(--aewttr-muted);">${escapeHtml(t.assignee || "Unassigned")}</span>
          <button class="btn-aewttr-outline btn-aewttr-sm gantt-inline-edit" data-open="${t.id}"${tip("Edit this task")}><i class="bx bx-edit"></i> Edit</button>`}
        </div>
      </div>
      <div class="gantt-cell-bar" style="min-height:${rowHeight}px;">
        <div class="gantt-bar-wrap">
          <button class="gantt-bar ${statusClass} ${mondayStyle ? "gantt-bar--monday" : ""}" data-open="${t.id}" data-drag-task="${t.id}" data-day-width="${dayWidth}" style="left:${left}px; width:${width}px;"${typeof ganttTip === "function" ? ganttTip(ganttTaskTipText(t)) : ""}>
            ${mondayStyle ? `<span class="gantt-bar-progress" style="width:${barProgress}%;"></span>` : escapeHtml(t.title)}
            <span class="gantt-resize-handle left" data-drag-task="${t.id}" data-mode="start"></span>
            <span class="gantt-resize-handle right" data-drag-task="${t.id}" data-mode="end"></span>
          </button>
        </div>
        ${renderInlineSubtaskBars(t, rangeStart, dayWidth)}
      </div>
      ${isOpen ? `
      <div class="gantt-drawer">
        ${t.subtasks && t.subtasks.length ? t.subtasks.map((s, si) => renderSubtaskPreview(t, s, si)).join("") : `<div class="gantt-drawer-empty">No subtasks yet — open the task editor to add timeline detail and file references.</div>`}
      </div>` : ""}`;
  });

  mount.innerHTML = `
    ${mondayStyle && members.length ? `<datalist id="${memberListId}">${members.map((m) => `<option value="${escapeHtml(m.name)}">`).join("")}</datalist>` : ""}
    <div class="gantt-scroll ${mondayStyle ? "gantt-scroll--monday" : ""} ${meetingMode ? "gantt-scroll--meeting" : ""}">
      <div class="gantt-grid ${mondayStyle ? "gantt-grid--monday" : ""} ${meetingMode ? "gantt-grid--meeting" : ""}">
        <div class="gantt-head-label">Task</div>
        <div class="gantt-head-ruler gantt-head-ruler--monday" style="width:${timelineWidth}px;">
          <div class="gantt-month-row">${monthTicks}</div>
          <div class="gantt-week-row">${ticks}${todayLine}</div>
        </div>
        ${rows}
      </div>
    </div>`;

  $all("[data-toggle]", mount).forEach(elx => elx.addEventListener("click", (e) => { e.stopPropagation(); onToggleExpand(elx.dataset.toggle); }));
  $all("[data-open]", mount).forEach(elx => elx.addEventListener("click", (e) => {
    if (e.target.closest(".gantt-inline-text, .gantt-inline-date, .gantt-inline-health, .gantt-resize-handle")) return;
    if (mount.dataset.dragJustFinished) return;
    onOpenEditor(elx.dataset.open);
  }));
  $all("[data-open-sub]", mount).forEach(elx => elx.addEventListener("click", (e) => {
    if (mount.dataset.dragJustFinished) return;
    e.stopPropagation();
    const [taskId, subIndex] = elx.dataset.openSub.split(":");
    onOpenEditor(taskId, +subIndex);
  }));
  $all(".gantt-sub-top", mount).forEach(row => row.addEventListener("click", (e) => {
    if (e.target.closest("input[type=checkbox], button, label, .gantt-inline-text")) return;
    if (mount.dataset.dragJustFinished) return;
    const btn = row.querySelector("[data-open-sub]");
    if (!btn) return;
    const [taskId, subIndex] = btn.dataset.openSub.split(":");
    onOpenEditor(taskId, +subIndex);
  }));
  $all(".gantt-drawer input[type=checkbox]", mount).forEach(cb => cb.addEventListener("change", (e) => {
    e.stopPropagation();
    const t = tasks.find(x => x.id === cb.dataset.task);
    setSubtaskCompletionState(t.subtasks[+cb.dataset.sub], cb.checked);
    syncTaskStatusFromSubtasks(t);
    Repo.save("actionItem", t, { projectCode: saveProjectCode, source: saveSource });
    renderGanttChart(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts);
  }));
  if (mondayStyle) wireGanttInlineEdits(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts);
  wireGanttBarDragging(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, { saveSource, projectCode: saveProjectCode });
  wireGanttRowReordering(mount, tasks, redraw);
}

function wireGanttInlineEdits(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts) {
  const saveSource = opts.saveSource || "Tracker";
  const saveProjectCode = opts.projectCode || projectCodeForTaskList(tasks);
  function persist(task) {
    syncTaskStatusFromSubtasks(task);
    ensureAssigneesFromTask(saveProjectCode, task);
    Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource });
    if (typeof redraw === "function") redraw();
  }
  $all(".gantt-inline-text[data-gantt-field]", mount).forEach((input) => {
    input.addEventListener("click", (e) => e.stopPropagation());
    const commit = () => {
      const task = tasks.find((t) => t.id === input.dataset.taskId);
      if (!task) return;
      const field = input.dataset.ganttField;
      const value = input.value.trim();
      if (field === "title" && !value) { input.value = task.title; return; }
      task[field] = value || (field === "assignee" ? "Unassigned" : task[field]);
      persist(task);
    };
    input.addEventListener("change", commit);
    input.addEventListener("blur", commit);
    if (input.dataset.ganttField === "assignee") wireAssigneeAutocomplete(mount, input);
  });
  $all(".gantt-inline-date[data-gantt-field]", mount).forEach((input) => {
    input.addEventListener("click", (e) => e.stopPropagation());
    input.addEventListener("change", () => {
      const task = tasks.find((t) => t.id === input.dataset.taskId);
      if (!task || !input.value) return;
      const prevStart = task.start;
      const prevEnd = task.end;
      task[input.dataset.ganttField] = input.value;
      cascadeSubtaskDates(tasks, task, prevStart, prevEnd);
      persist(task);
      renderGanttChart(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts);
    });
  });
  $all(".monday-gantt-complete-check[data-task-complete]", mount).forEach((input) => {
    input.addEventListener("click", (e) => e.stopPropagation());
    input.addEventListener("change", () => {
      const task = tasks.find((t) => t.id === input.dataset.taskComplete);
      if (!task) return;
      setTaskCompletionState(task, input.checked);
      persist(task);
      renderGanttChart(mount, tasks, expanded, onToggleExpand, onOpenEditor, redraw, visibleTasks, opts);
    });
  });
  $all(".gantt-inline-health[data-gantt-field]", mount).forEach((sel) => {
    sel.addEventListener("click", (e) => e.stopPropagation());
    sel.addEventListener("change", () => {
      const task = tasks.find((t) => t.id === sel.dataset.taskId);
      if (!task) return;
      task.health = sel.value;
      sel.className = `gantt-inline-health health-${sel.value.replace(/\s+/g, "-")}`;
      persist(task);
    });
  });
}

/* Drag-and-drop reordering of Gantt task rows via the row's drag handle
   (mirrors the Checklist view's reordering — see .cl-drag-handle wiring below). */
function wireGanttRowReordering(mount, tasks, redraw) {
  let draggedId = null;
  const rowSelector = ".monday-gantt-task-label[data-row-id], .gantt-cell-label[data-row-id]";
  $all(rowSelector, mount).forEach(row => {
    const id = row.dataset.rowId;
    const handle = $(".cl-drag-handle", row);
    if (!handle) return;
    // A plain click (mousedown with no following dragstart) must not leave
    // the row permanently draggable=true — native drag-and-drop suppresses
    // the normal "mouseup"/"touchend" event for a gesture that actually
    // becomes a drag, so this reset only fires for the non-drag case.
    handle.addEventListener("mousedown", () => {
      row.draggable = true;
      const resetIfNoDrag = () => { if (!row.classList.contains("dragging")) row.draggable = false; };
      window.addEventListener("mouseup", resetIfNoDrag, { once: true });
    });
    handle.addEventListener("touchstart", () => {
      row.draggable = true;
      const resetIfNoDrag = () => { if (!row.classList.contains("dragging")) row.draggable = false; };
      window.addEventListener("touchend", resetIfNoDrag, { once: true });
    }, { passive: true });
    row.addEventListener("dragstart", (e) => {
      draggedId = id;
      // Deferred so the browser snapshots the row at full opacity before
      // the "dragging" class fades it — otherwise the drag ghost can render
      // blank, leaving no visual feedback under the cursor while dragging.
      setTimeout(() => row.classList.add("dragging"), 0);
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", id);
    });
    row.addEventListener("dragend", () => {
      row.draggable = false;
      row.classList.remove("dragging");
      $all(rowSelector, mount).forEach(el => el.classList.remove("drag-over"));
      draggedId = null;
    });
    row.addEventListener("dragover", (e) => {
      if (!draggedId || draggedId === id) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      row.classList.add("drag-over");
    });
    row.addEventListener("dragleave", () => row.classList.remove("drag-over"));
    row.addEventListener("drop", (e) => {
      e.preventDefault();
      row.classList.remove("drag-over");
      if (!draggedId || draggedId === id) return;
      ganttReorderTask(tasks, draggedId, id);
      Repo.save("actionItem", tasks.find(t => t.id === draggedId), { projectCode: projectCodeForTaskList(tasks), source: "Tracker" });
      draggedId = null;
      redraw();
    });
  });
}

function wireTrackerTableDragDrop(mount, tasks, redraw, opts) {
  opts = opts || {};
  const saveProjectCode = opts.projectCode || projectCodeForTaskList(tasks);
  const saveSource = opts.saveSource || "Tracker";

  let dragState = null;

  function clearHighlights() {
    $all(".tracker-drag-over", mount).forEach(el => el.classList.remove("tracker-drag-over"));
  }

  function wireSource(row, stateFactory) {
    const handle = $(".cl-drag-handle", row);
    if (!handle) return;
    handle.addEventListener("mousedown", () => {
      row.draggable = true;
      const reset = () => { if (!row.classList.contains("dragging")) row.draggable = false; };
      window.addEventListener("mouseup", reset, { once: true });
    });
    handle.addEventListener("touchstart", () => {
      row.draggable = true;
      const reset = () => { if (!row.classList.contains("dragging")) row.draggable = false; };
      window.addEventListener("touchend", reset, { once: true });
    }, { passive: true });
    row.addEventListener("dragstart", e => {
      dragState = stateFactory();
      setTimeout(() => row.classList.add("dragging"), 0);
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", JSON.stringify(dragState));
    });
    row.addEventListener("dragend", () => {
      row.draggable = false;
      row.classList.remove("dragging");
      clearHighlights();
      dragState = null;
    });
  }

  function wireTarget(row, accepts, onDrop) {
    row.addEventListener("dragover", e => {
      if (!dragState || !accepts(dragState)) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      row.classList.add("tracker-drag-over");
    });
    row.addEventListener("dragleave", () => row.classList.remove("tracker-drag-over"));
    row.addEventListener("drop", e => {
      e.preventDefault();
      row.classList.remove("tracker-drag-over");
      if (!dragState || !accepts(dragState)) return;
      onDrop(dragState);
      dragState = null;
    });
  }

  $all("tr.monday-row--tracker[data-id]", mount).forEach(row => {
    wireSource(row, () => ({ type: "task", taskId: row.dataset.id }));
  });

  $all("tr.monday-sub-row-item[data-task-id][data-sub-path]", mount).forEach(row => {
    wireSource(row, () => ({ type: "subtask", taskId: row.dataset.taskId, subPath: row.dataset.subPath }));
  });

  $all("tr.monday-row--tracker[data-id]", mount).forEach(row => {
    const targetId = row.dataset.id;
    const targetTask = () => tasks.find(t => t.id === targetId);
    wireTarget(row,
      ds => {
        if (ds.type === "task") return ds.taskId !== targetId && !isTrackerDivider(targetTask());
        if (ds.type === "subtask") return ds.taskId !== targetId && !isTrackerDivider(targetTask());
        return false;
      },
      ds => {
        const tt = targetTask();
        if (!tt || isTrackerDivider(tt)) return;
        if (ds.type === "task") {
          const dragged = tasks.find(t => t.id === ds.taskId);
          if (!dragged) return;
          tasks.splice(tasks.findIndex(t => t.id === ds.taskId), 1);
          if (!tt.subtasks) tt.subtasks = [];
          tt.subtasks.push({
            text: dragged.title || dragged.text || "",
            assignee: dragged.assignee || dragged.owner || "",
            start: dragged.start || "",
            end: dragged.end || "",
            done: false,
            notes: []
          });
          Repo.save("actionItem", tt, { projectCode: saveProjectCode, source: saveSource }).catch(() => {});
          Repo.remove("actionItem", dragged).catch(() => {});
          redraw();
        } else if (ds.type === "subtask") {
          const srcTask = tasks.find(t => t.id === ds.taskId);
          if (!srcTask || !srcTask.subtasks) return;
          const pathParts = ds.subPath.split(".");
          if (pathParts.length > 1) return;
          const idx = parseInt(pathParts[0], 10);
          if (isNaN(idx) || idx >= srcTask.subtasks.length) return;
          const [sub] = srcTask.subtasks.splice(idx, 1);
          if (!tt.subtasks) tt.subtasks = [];
          tt.subtasks.push(sub);
          if (typeof syncTaskStatusFromSubtasks === "function") syncTaskStatusFromSubtasks(srcTask);
          Repo.save("actionItem", srcTask, { projectCode: saveProjectCode, source: saveSource }).catch(() => {});
          Repo.save("actionItem", tt, { projectCode: saveProjectCode, source: saveSource }).catch(() => {});
          redraw();
        }
      }
    );
  });

  $all("tr.monday-divider-row[data-divider-id]", mount).forEach(row => {
    const dividerId = row.dataset.dividerId;
    wireTarget(row,
      ds => ds.type === "task" || ds.type === "subtask",
      ds => {
        const dividerIdx = tasks.findIndex(t => t.id === dividerId);
        if (dividerIdx < 0) return;
        if (ds.type === "task") {
          const taskIdx = tasks.findIndex(t => t.id === ds.taskId);
          if (taskIdx < 0) return;
          const [moved] = tasks.splice(taskIdx, 1);
          const insertIdx = dividerIdx >= taskIdx ? dividerIdx : dividerIdx + 1;
          tasks.splice(insertIdx, 0, moved);
          Repo.save("actionItem", moved, { projectCode: saveProjectCode, source: saveSource }).catch(() => {});
          redraw();
        } else if (ds.type === "subtask") {
          const srcTask = tasks.find(t => t.id === ds.taskId);
          if (!srcTask || !srcTask.subtasks) return;
          const pathParts = ds.subPath.split(".");
          if (pathParts.length > 1) return;
          const idx = parseInt(pathParts[0], 10);
          if (isNaN(idx) || idx >= srcTask.subtasks.length) return;
          const [sub] = srcTask.subtasks.splice(idx, 1);
          const newTask = {
            id: "task-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8),
            title: sub.text || sub.title || "",
            assignee: sub.assignee || "",
            start: sub.start || "",
            end: sub.end || "",
            done: sub.done || false,
            subtasks: [],
            notes: sub.notes || []
          };
          tasks.splice(dividerIdx + 1, 0, newTask);
          if (typeof syncTaskStatusFromSubtasks === "function") syncTaskStatusFromSubtasks(srcTask);
          Repo.save("actionItem", srcTask, { projectCode: saveProjectCode, source: saveSource }).catch(() => {});
          Repo.save("actionItem", newTask, { projectCode: saveProjectCode, source: saveSource }).catch(() => {});
          redraw();
        }
      }
    );
  });
}

function renderTrackerTableView(mount, tasks, expanded, onToggleExpand, onOpenEditor, onChange, visibleTasks, proj, tableOpts) {
  tableOpts = tableOpts || {};
  const saveSource = tableOpts.saveSource || "Tracker";
  const saveProjectCode = tableOpts.projectCode || projectCodeForTaskList(tasks);
  const afterSave = tableOpts.onAfterSave;
  const meetingMode = !!tableOpts.meetingMode;
  const dividerCollapsed = tableOpts.dividerCollapsed || {};
  const subitemExpanded = tableOpts.subitemExpanded || {};
  visibleTasks = visibleTasks || tasks;

  const extraColumns = (tableOpts.extraColumns || []).slice();
  if (meetingMode) {
    extraColumns.push({
      key: "reviewStatus",
      label: "Review Status",
      options: ["Not Reviewed", "Reviewed - No Change", "Updated"],
      colorPrefix: "review"
    });
  }

  const visiblePlain = trackerPlainTasks(visibleTasks);
  const visibleIds = new Set(visiblePlain.map((t) => t.id));
  const sections = [{
    kind: "ungrouped",
    label: proj ? `${proj.name || proj.id} — Tracker` : "Tasks",
    tasks: visiblePlain,
    collapsed: false
  }];

  if (!visiblePlain.length && trackerPlainTasks(tasks).length) {
    mount.innerHTML = `<div class="empty-state" style="padding:30px;">No tasks match this assignee filter.</div>`;
    return;
  }
  const pcode = projectCodeForTaskList(tasks);
  const focusTaskId = window.AEWTTR.state.trackerFocusTaskId || null;
  const focusSubtask = window.AEWTTR.state.trackerFocusSubtask || null;
  const routeIntent = typeof consumeRouteIntent === "function" ? consumeRouteIntent(`tracker:${pcode || "global"}`) : null;
  if (focusTaskId) window.AEWTTR.state.trackerFocusTaskId = null;
  if (focusSubtask) window.AEWTTR.state.trackerFocusSubtask = null;

  function inlineAddTask(parentDividerId, afterTaskId) {
    if (typeof tableOpts.onCreateTask === "function") {
      tableOpts.onCreateTask(parentDividerId || "", afterTaskId || "");
      return;
    }
    const today = new Date().toISOString().slice(0, 10);
    const id = uid("g");
    const newTask = {
      id,
      itemType: "task",
      workItemLevel: "Task",
      title: "New task",
      assignee: tableOpts.defaultAssignee || "Unassigned",
      start: today,
      end: "",
      health: "On Track",
      status: "Not Started",
      parentDividerId: parentDividerId || "",
      subtasks: [],
      notes: []
    };
    if (meetingMode) newTask.reviewStatus = "Not Reviewed";
    if (afterTaskId) {
      const afterIdx = tasks.findIndex((t) => t.id === afterTaskId);
      tasks.splice(afterIdx >= 0 ? afterIdx + 1 : tasks.length, 0, newTask);
    } else {
      tasks.push(newTask);
    }
    tasks.forEach((item, idx) => { if (item) item._sortOrder = idx; });
    expanded[id] = true;
    if (parentDividerId) dividerCollapsed[parentDividerId] = false;
    else dividerCollapsed.__ungrouped__ = false;
    window.AEWTTR.state.trackerFocusTaskId = id;
    ensureAssigneesFromTask(saveProjectCode, newTask);
    if (typeof tableOpts.onBeforeSave === "function") tableOpts.onBeforeSave();
    Repo.save("actionItem", newTask, { projectCode: saveProjectCode, source: saveSource }).catch(() => {
      if (typeof tableOpts.onSaveError === "function") tableOpts.onSaveError();
    });
    if (typeof afterSave === "function") afterSave();
    onChange();
  }

  function persistTask(task) {
    if (isTrackerDivider(task)) {
      syncDividerMetadata(task);
      ensureDividerRolesOnProjectPeople(saveProjectCode, task);
    } else {
      ensureAssigneesFromTask(saveProjectCode, task);
    }
    if (typeof tableOpts.onBeforeSave === "function") tableOpts.onBeforeSave();
    Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource })
      .then(() => { if (typeof afterSave === "function") afterSave(); })
      .catch(() => { if (typeof tableOpts.onSaveError === "function") tableOpts.onSaveError(); });
  }

  function resolveSub(task, pathOrIdx) {
    if (pathOrIdx == null || pathOrIdx === "") return null;
    if (typeof pathOrIdx === "number" || /^\d+$/.test(String(pathOrIdx))) {
      const idx = Number(pathOrIdx);
      if (!String(pathOrIdx).includes(".")) return (task.subtasks || [])[idx] || null;
    }
    return getSubtaskAtPath(task, String(pathOrIdx));
  }

  renderMondayTable(mount, {
    mode: "tracker",
    extraColumns,
    tasks: visiblePlain,
    allTasks: tasks,
    sections,
    subitemExpanded,
    reorderable: true,
    focusTaskId,
    focusSubtask,
    ownerField: "assignee",
    startField: "start",
    endField: "end",
    expanded,
    hideHeader: !!tableOpts.hideHeader,
    hideGroupHeader: !!tableOpts.hideGroupHeader,
    emptyMessage: trackerPlainTasks(tasks).length ? "" : "No tasks yet — right-click to add or use the + button.",
    groupHeader: null,
    healthColumn: { key: "health", label: "Health", options: ["On Track", "At Risk", "Off Track"], colorPrefix: "health" },
    inlineAddLabel: "+ Add task",
    onInlineAdd: () => inlineAddTask(""),
    onToggleExpand,
    onToggleSubitemExpand: (key) => {
      subitemExpanded[key] = !subitemExpanded[key];
      onChange();
    },
    onToggleDivider: (id) => {
      dividerCollapsed[id] = !dividerCollapsed[id];
      onChange();
    },
    onEditDivider: () => {},
    onDeleteDivider: async (id) => {
      const divider = tasks.find((t) => t.id === id && isTrackerDivider(t));
      if (!divider) return;
      const ok = await confirmDialog({
        title: "Delete divider",
        message: `Delete "${divider.title}"? Tasks under it move to Ungrouped.`,
        confirmLabel: "Delete",
        danger: true
      });
      if (!ok) return;
      const reparented = [];
      trackerPlainTasks(tasks).forEach((t) => {
        if (t.parentDividerId === id) {
          reparented.push({ task: t, prev: t.parentDividerId });
          t.parentDividerId = "";
          persistTask(t);
        }
      });
      const i = tasks.findIndex((t) => t.id === id);
      if (i >= 0) tasks.splice(i, 1);
      toast("Divider deleted", "success");
      onChange();
      Repo.remove("actionItem", divider).catch(() => {
        tasks.splice(i, 0, divider);
        reparented.forEach(({ task, prev }) => { task.parentDividerId = prev; persistTask(task); });
        onChange();
      });
    },
    onAddTaskUnderDivider: (dividerId, insertAfter) => inlineAddTask(dividerId || "", insertAfter || ""),
    onToggleComplete: (task) => {
      setTaskCompletionState(task, !taskIsFullyDone(task));
      persistTask(task);
      onChange();
    },
    onReorder: (moved) => {
      persistTask(moved);
      onChange();
    },
    onReorderDivider: (movedId) => {
      const moved = tasks.find((t) => t && t.id === movedId && isTrackerDivider(t));
      if (!moved) return;
      persistTask(moved);
      onChange();
    },
    onDeleteTask: async (task) => {
      const ok = await confirmDialog({
        title: "Delete task",
        message: `Delete "${task.title}" and all its subitems?`,
        confirmLabel: "Delete",
        danger: true
      });
      if (!ok) return;
      const i = tasks.findIndex((t) => t.id === task.id);
      if (i >= 0) tasks.splice(i, 1);
      toast("Task deleted", "success");
      onChange();
      Repo.remove("actionItem", task).catch(() => {
        tasks.splice(i, 0, task);
        onChange();
      });
    },
    onDeleteSubtask: async (task, pathOrIdx) => {
      const sub = resolveSub(task, pathOrIdx);
      const ok = await confirmDialog({
        title: "Delete subitem",
        message: `Delete "${(sub && sub.text) || "this subitem"}"${(sub && (sub.subtasks || []).length) ? " and nested subitems" : ""}?`,
        confirmLabel: "Delete",
        danger: true
      });
      if (!ok) return;
      const snapshot = JSON.parse(JSON.stringify(task.subtasks || []));
      removeSubtaskAtPath(task, String(pathOrIdx));
      syncTaskStatusFromSubtasks(task);
      toast("Subitem deleted", "success");
      onChange();
      Repo.save("actionItem", task, { projectCode: saveProjectCode, source: saveSource })
        .then(() => { if (typeof afterSave === "function") afterSave(); })
        .catch(() => {
          task.subtasks = snapshot;
          syncTaskStatusFromSubtasks(task);
          onChange();
        });
    },
    onAddSubtask: (taskId, parentPath) => {
      const task = tasks.find((t) => t.id === taskId);
      if (!task) return;
      task.subtasks = task.subtasks || [];
      if (parentPath) {
        const parent = getSubtaskAtPath(task, parentPath);
        if (!parent) return;
        const childPath = `${parentPath}.${(parent.subtasks || []).length}`;
        addNestedSubtask(parent, { text: "New subitem", start: parent.start || task.start, end: "" });
        subitemExpanded[`${task.id}:${parentPath}`] = true;
        window.AEWTTR.state.trackerFocusSubtask = { taskId, path: childPath };
      } else {
        const childPath = String(task.subtasks.length);
        task.subtasks.push(normalizeTaskSubtask(task, { text: "New subitem", start: task.start, end: "" }));
        window.AEWTTR.state.trackerFocusSubtask = { taskId, path: childPath };
      }
      syncTaskStatusFromSubtasks(task);
      persistTask(task);
      expanded[taskId] = true;
      onChange();
    },
    onAddSubtaskSibling: (taskId, siblingPath) => {
      const task = tasks.find((t) => t.id === taskId);
      const parts = String(siblingPath || "").split(".").map(Number);
      if (!task || !parts.length || parts.some(Number.isNaN)) return;
      const siblingIndex = parts.pop();
      let list = task.subtasks || [];
      let parent = null;
      for (const index of parts) {
        parent = list[index];
        if (!parent) return;
        parent.subtasks = parent.subtasks || [];
        list = parent.subtasks;
      }
      const current = list[siblingIndex];
      if (!current) return;
      const newSubtask = normalizeTaskSubtask(task, {
        text: "New subtask",
        start: current.start || task.start,
        end: "",
        assignee: current.assignee || ""
      });
      const newIndex = siblingIndex + 1;
      list.splice(newIndex, 0, newSubtask);
      const parentPath = parts.join(".");
      if (parentPath) subitemExpanded[`${task.id}:${parentPath}`] = true;
      window.AEWTTR.state.trackerFocusSubtask = {
        taskId,
        path: parentPath ? `${parentPath}.${newIndex}` : String(newIndex)
      };
      syncTaskStatusFromSubtasks(task);
      persistTask(task);
      expanded[taskId] = true;
      onChange();
    },
    onFieldChange: (task, key, value, subPath) => {
      const sub = subPath != null && subPath !== "" ? resolveSub(task, subPath) : null;
      const prevValue = key.startsWith("subtask.")
        ? (sub ? sub[key.split(".")[1]] : undefined)
        : (sub ? sub[key] : task[key]);
      if (key.startsWith("subtask.")) {
        const field = key.split(".")[1];
        if (sub) sub[field] = value;
      } else if (sub) {
        sub[key] = value;
        if (meetingMode && key === "notes" && typeof recordActiveMeetingNotePosted === "function") {
          const newest = Array.isArray(value) && value[0];
          const freshMs = newest && (typeof meetingNoteTimestampMs === "function"
            ? meetingNoteTimestampMs(newest)
            : (newest.date ? Date.parse(`${newest.date}T${(newest.time || "00:00")}:00`) : null));
          const isFresh = newest && newest.text && !newest.editedAt && freshMs != null && (Date.now() - freshMs) < 20000;
          if (isFresh) {
            recordActiveMeetingNotePosted(saveProjectCode, newest.text, {
              kind: "subitem",
              projectId: saveProjectCode,
              taskId: task.id,
              taskTitle: task.title || "Untitled",
              subitemText: (sub && sub.text) || "Untitled subitem"
            });
          }
        }
      } else {
        task[key] = value;
        if (meetingMode && key === "notes" && typeof recordActiveMeetingNotePosted === "function") {
          const newest = Array.isArray(value) && value[0];
          const freshMs = newest && (typeof meetingNoteTimestampMs === "function"
            ? meetingNoteTimestampMs(newest)
            : (newest.date ? Date.parse(`${newest.date}T${(newest.time || "00:00")}:00`) : null));
          const isFresh = newest && newest.text && !newest.editedAt && freshMs != null && (Date.now() - freshMs) < 20000;
          if (isFresh) {
            recordActiveMeetingNotePosted(saveProjectCode, newest.text, {
              kind: "task",
              projectId: saveProjectCode,
              taskId: task.id,
              taskTitle: task.title || "Untitled"
            });
          }
        } else if (meetingMode && typeof recordActiveMeetingTaskChange === "function"
          && (key === "reviewStatus" || key === "health" || key === "status")) {
          recordActiveMeetingTaskChange(saveProjectCode, task, key, value, { prev: prevValue });
        }
      }
      if (key === "start" || key === "end" || key.startsWith("subtask.")) {
        if (key === "start" || key === "end") cascadeSubtaskDates(tasks, task, task.start, task.end);
      }
      syncTaskStatusFromSubtasks(task);
      persistTask(task);
    },
    onToggleSubtaskDone: (task, pathOrIdx, checked) => {
      const sub = resolveSub(task, pathOrIdx);
      if (!sub) return;
      setSubtaskCompletionState(sub, checked);
      // Cascade done to nested children when marking a parent complete.
      if (checked) {
        walkNestedSubtasks(sub.subtasks || [], (child) => setSubtaskCompletionState(child, true));
      }
      syncTaskStatusFromSubtasks(task);
      persistTask(task);
      onChange();
    }
  });
  wireTrackerContextMenu(mount, tasks, onOpenEditor, onChange, { saveSource, projectCode: saveProjectCode, proj, subitemExpanded, expanded });

  if (routeIntent && routeIntent.task) {
    const task = tasks.find((item) => item.id === routeIntent.task);
    if (task) {
      expanded[task.id] = true;
      if (routeIntent.sub != null && task.subtasks && task.subtasks[Number(routeIntent.sub)]) {
        openSubtaskSidePanel(task, tasks, onChange, Number(routeIntent.sub), false, { presentation: "sidebar" });
      } else {
        openTaskSidePanel(task, tasks, onChange, { presentation: "sidebar" });
      }
    }
  }
}
function projectCodeForTaskList(tasks) {
  const db = window.AEWTTR.db;
  return Object.keys(db.ganttTasks).find(pid => db.ganttTasks[pid] === tasks) || "";
}
function ganttReorderTask(tasks, draggedId, targetId) {
  const from = tasks.findIndex(t => t.id === draggedId);
  const to = tasks.findIndex(t => t.id === targetId);
  if (from < 0 || to < 0 || from === to) return;
  const [moved] = tasks.splice(from, 1);
  tasks.splice(to, 0, moved);
}

function subtaskFileRowHtml(doc) {
  const file = doc || { fileName: "", location: "Windchill" };
  return `<div class="subtask-file-row">
    <input type="text" class="input-aewttr sf-name" placeholder="File name" value="${escapeHtml(file.fileName)}">
    <select class="select-aewttr sf-loc">
      <option ${file.location === "Windchill" ? "selected" : ""}>Windchill</option>
      <option ${file.location === "SPO" ? "selected" : ""}>SPO</option>
    </select>
    <button type="button" class="btn-aewttr-ghost btn-aewttr-sm sf-remove">&times;</button>
  </div>`;
}

function subtaskCardHtml(task, rawSubtask) {
  const s = normalizeTaskSubtask(task, rawSubtask);
  return `<div class="subtask-card">
    <div class="subtask-card-top">
      <input type="checkbox" class="st-done" ${s.done ? "checked" : ""}>
      <input type="text" class="input-aewttr st-text" placeholder="Subtask title" value="${escapeHtml(s.text)}">
      <button type="button" class="btn-aewttr-ghost btn-aewttr-sm st-remove">&times;</button>
    </div>
    <div class="subtask-card-row">
      <input type="text" class="input-aewttr st-assignee" placeholder="Owner" value="${escapeHtml(s.assignee)}">
      <input type="date" class="input-aewttr st-start" value="${s.start}">
      <input type="date" class="input-aewttr st-end" value="${s.end}">
    </div>
    <div class="subtask-card-files">
      <div class="st-files-list">${s.relatedDocs.map(subtaskFileRowHtml).join("")}</div>
      <button type="button" class="btn-aewttr-ghost btn-aewttr-sm st-add-file"><i class="bx bx-paperclip"></i> Add File</button>
    </div>
  </div>`;
}

function readSubtaskCard(row, task) {
  const subtask = normalizeTaskSubtask(task, {
    text: $(".st-text", row).value.trim(),
    assignee: $(".st-assignee", row).value.trim(),
    done: $(".st-done", row).checked,
    start: $(".st-start", row).value,
    end: $(".st-end", row).value,
    relatedDocs: $all(".subtask-file-row", row).map(fileRow => ({
      fileName: $(".sf-name", fileRow).value.trim(),
      location: $(".sf-loc", fileRow).value
    })).filter(doc => doc.fileName)
  });
  return subtask;
}

function closeTaskSidePanel() {
  document.removeEventListener("keydown", tspEscHandler);
  document.querySelectorAll(".task-side-panel, .task-side-backdrop").forEach((node) => node.remove());
  const taskModalBackdrop = document.querySelector(".task-editor-modal")?.closest(".aewttr-modal-backdrop");
  if (taskModalBackdrop) taskModalBackdrop.remove();
}
function tspEscHandler(e) { if (e.key === "Escape") closeTaskSidePanel(); }

function mountTaskEditorShell(presentation) {
  closeTaskSidePanel();
  const isModal = presentation === "modal";
  let root;
  let backdrop;
  if (isModal) {
    backdrop = el(`<div class="aewttr-modal-backdrop task-editor-backdrop"></div>`);
    root = el(`<div class="aewttr-modal wide task-editor-modal"></div>`);
    backdrop.appendChild(root);
    document.body.appendChild(backdrop);
    backdrop.addEventListener("mousedown", (e) => { if (e.target === backdrop) closeTaskSidePanel(); });
  } else {
    backdrop = el(`<div class="task-side-backdrop"></div>`);
    root = el(`<div class="task-side-panel"></div>`);
    document.body.appendChild(backdrop);
    document.body.appendChild(root);
    backdrop.addEventListener("click", closeTaskSidePanel);
  }
  document.addEventListener("keydown", tspEscHandler);
  return { root, isModal };
}

function taskEditorFrameHtml(title, bodyHtml, footHtml, isModal) {
  if (isModal) {
    return `
      <div class="aewttr-modal-head"><h3>${title}</h3><button class="aewttr-modal-close" id="tsp-close" type="button" aria-label="Close">&times;</button></div>
      <div class="aewttr-modal-body tsp-body">${bodyHtml}</div>
      <div class="aewttr-modal-foot tsp-foot">${footHtml}</div>
    `;
  }
  return `
    <div class="tsp-head"><h3>${title}</h3><button class="aewttr-modal-close" id="tsp-close" type="button" aria-label="Close">&times;</button></div>
    <div class="tsp-body">${bodyHtml}</div>
    <div class="tsp-foot">${footHtml}</div>
  `;
}

function openTaskSidePanel(task, tasks, onChange, opts) {
  opts = opts || {};
  if (!task) return;
  const { root: panel, isModal } = mountTaskEditorShell(opts.presentation || "sidebar");
  const bodyHtml = `
        <div class="tsp-row-actions">
          <button class="btn-danger-outline btn-aewttr-sm" id="tsp-delete" type="button"><i class="bx bx-trash"></i> Delete</button>
        </div>
        <div class="form-row"><label>Title</label><input class="input-aewttr" id="tsp-title" value="${escapeHtml(task.title)}"></div>
        <div class="form-row"><label>Assignee</label>
          <div class="traveler-picker traveler-picker--inline" data-owner-wrap="tsp-assignee">
            <div id="tsp-assignee-sel" class="traveler-chip-list"></div>
            <input class="input-aewttr" id="tsp-assignee-input" placeholder="Search people…">
            <div id="tsp-assignee-sugg" class="traveler-suggestions"></div>
          </div>
        </div>
        <div class="form-row"><label>Start date</label><input type="date" class="input-aewttr" id="tsp-start" value="${task.start}"></div>
        <div class="form-row"><label>End date</label><input type="date" class="input-aewttr" id="tsp-end" value="${task.end}"></div>
        <div class="form-row"><label>Health</label>
          <select class="select-aewttr" id="tsp-health">${["On Track", "At Risk", "Off Track"].map(s => `<option ${task.health === s ? "selected" : (!task.health && s === "On Track" ? "selected" : "")}>${s}</option>`).join("")}</select>
        </div>
        <div class="form-row"><label>Est. Effort (hrs)</label>
          <input type="number" class="input-aewttr" id="tsp-effort" value="${task.estimatedEffort || ""}" placeholder="e.g. 10" min="0">
        </div>
        <div class="form-row"><label>Progress</label>
          <div class="tsp-progress-preview">${taskProgressBarHtml(task)}</div>
        </div>
        <div class="form-row"><label>Subtasks</label>
          <p style="font-size:11.5px;color:var(--aewttr-muted);margin:0 0 10px;">Each subtask opens in its own editor. Add new ones here, or use the inline Edit buttons from the gantt.</p>
          <div class="tsp-subtask-list">
            ${(task.subtasks || []).map((s, index) => {
              const subtask = normalizeTaskSubtask(task, s);
              return `<button type="button" class="tsp-subtask-item" data-open-sub="${index}">
                <span class="tsp-subtask-copy">
                  <strong>${escapeHtml(subtask.text || "Untitled subtask")}</strong>
                  <span>${subtask.assignee || "Unassigned"} · ${fmtDate(subtask.start)} – ${fmtDate(subtask.end)}</span>
                </span>
                <span class="kc-badge">${subtask.relatedDocs.length} file${subtask.relatedDocs.length === 1 ? "" : "s"}</span>
              </button>`;
            }).join("") || `<div class="gantt-drawer-empty">No subtasks yet.</div>`}
          </div>
          <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="tsp-add-subtask"><i class="bx bx-plus"></i> Add Subtask</button>
        </div>`;
  const footHtml = `
        <button class="btn-aewttr-ghost" id="tsp-cancel" type="button">Cancel</button>
        <button class="btn-aewttr" id="tsp-save" type="button">Save Changes</button>`;
  panel.innerHTML = taskEditorFrameHtml("Edit Gantt Task", bodyHtml, footHtml, isModal);
  const db = window.AEWTTR.db;
  const currentAssigneeMember = task.assignee && task.assignee !== "Unassigned"
    ? (db.members || []).find((m) => m.name === task.assignee)
    : null;
  const pendingAssignee = task.assignee && task.assignee !== "Unassigned"
    ? [{ name: task.assignee, email: currentAssigneeMember ? currentAssigneeMember.email : "" }]
    : [];
  wirePeoplePicker(panel, pendingAssignee, { mount: "tsp-assignee-sel", input: "tsp-assignee-input", suggestions: "tsp-assignee-sugg" }, { singleSelect: true, allowManualEmail: false });
  $("#tsp-close", panel).addEventListener("click", closeTaskSidePanel);
  $("#tsp-cancel", panel).addEventListener("click", closeTaskSidePanel);
  const subtaskOpts = { presentation: opts.presentation || "sidebar" };
  $("#tsp-add-subtask", panel).addEventListener("click", async () => {
    if (!task.subtasks) task.subtasks = [];
    task.subtasks.push(normalizeTaskSubtask(task, { text: "", start: task.start, end: "", relatedDocs: [] }));
    Repo.save("actionItem", task, { projectCode: projectCodeForTaskList(tasks), source: "Tracker" });
    closeTaskSidePanel();
    onChange();
    openSubtaskSidePanel(task, tasks, onChange, task.subtasks.length - 1, true, subtaskOpts);
  });
  $all("[data-open-sub]", panel).forEach(btn => btn.addEventListener("click", () => openSubtaskSidePanel(task, tasks, onChange, +btn.dataset.openSub, false, subtaskOpts)));
  $("#tsp-delete", panel).addEventListener("click", async () => {
    const ok = await confirmDialog({ title: "Delete task", message: "Delete this task?", confirmLabel: "Delete", danger: true });
    if (!ok) return;
    const i = tasks.findIndex(t => t.id === task.id);
    if (i < 0) return;
    tasks.splice(i, 1);
    closeTaskSidePanel();
    toast("Task deleted", "success");
    onChange();
    Repo.remove("actionItem", task).catch(() => {
      tasks.splice(i, 0, task);
      onChange();
    });
  });
  $("#tsp-save", panel).addEventListener("click", async () => {
    const title = $("#tsp-title", panel).value.trim();
    if (!title) { toast("Title is required", "error"); return; }
    const start = $("#tsp-start", panel).value, end = $("#tsp-end", panel).value;
    if (start && end && start > end) { toast("End date must be on or after start date", "error"); return; }
    const prevStart = task.start, prevEnd = task.end;
    const prevAssignee = task.assignee;
    task.title = title;
    task.assignee = pendingAssignee[0] ? pendingAssignee[0].name : "Unassigned";
    task.start = start;
    task.end = end;
    task.health = $("#tsp-health", panel).value;
    task.estimatedEffort = +$("#tsp-effort", panel).value || 0;
    (task.subtasks || []).forEach(subtask => {
      if (!subtask.start) subtask.start = task.start;
      if (!subtask.health) subtask.health = task.health;
    });
    cascadeSubtaskDates(tasks, task, prevStart, prevEnd);
    syncTaskStatusFromSubtasks(task);
    const savedProjectCode = projectCodeForTaskList(tasks);
    ensureAssigneesFromTask(savedProjectCode, task);
    Repo.save("actionItem", task, { projectCode: savedProjectCode, source: "Tracker" });
    if (task.assignee !== prevAssignee) notifyTaskAssignee(task, savedProjectCode);
    closeTaskSidePanel();
    toast("Task updated", "success");
    onChange();
  });
}

function openSubtaskSidePanel(task, tasks, onChange, subIndex, isNew, opts) {
  opts = opts || {};
  if (!task || !task.subtasks) return;
  const subPath = String(subIndex);
  const pathParts = subPath.split(".").map(Number);
  if (pathParts.some(Number.isNaN) || !pathParts.length) return;
  const itemIndex = pathParts.pop();
  let itemList = task.subtasks;
  for (const index of pathParts) {
    const parent = itemList[index];
    if (!parent) return;
    parent.subtasks = parent.subtasks || [];
    itemList = parent.subtasks;
  }
  if (!itemList[itemIndex]) return;
  const subtask = normalizeTaskSubtask(task, itemList[itemIndex]);
  itemList[itemIndex] = subtask;
  const { root: panel, isModal } = mountTaskEditorShell(opts.presentation || "sidebar");
  const bodyHtml = `
        <div class="form-row"><label>Parent task</label><div class="tsp-context">${escapeHtml(task.title)}</div></div>
        <div class="form-row"><label>Subtask title</label><input class="input-aewttr" id="ssp-title" value="${escapeHtml(subtask.text)}"></div>
        <div class="form-grid-2">
          <div class="form-row"><label>Owner</label><input class="input-aewttr" id="ssp-assignee" value="${escapeHtml(subtask.assignee)}"></div>
          <div class="form-row"><label>Health</label><select class="select-aewttr" id="ssp-health">${["On Track", "At Risk", "Off Track"].map(s => `<option ${subtask.health === s ? "selected" : ""}>${s}</option>`).join("")}</select></div>
        </div>
        <div class="form-grid-2">
          <div class="form-row"><label>Start date</label><input type="date" class="input-aewttr" id="ssp-start" value="${subtask.start}"></div>
          <div class="form-row"><label>End date</label><input type="date" class="input-aewttr" id="ssp-end" value="${subtask.end}"></div>
        </div>
        <div class="form-row"><label>Complete</label>
          <label class="ssp-link-row"><input type="checkbox" id="ssp-done" ${subtask.done ? "checked" : ""}> <span>Mark subitem done</span></label>
        </div>
        <label class="ssp-link-row"><input type="checkbox" id="ssp-linked" ${subtask.linked === false ? "" : "checked"}> <span>Link to task dates <small>— follow the parent task when its dates change</small></span></label>
        <div class="form-row"><label>Files</label>
          <div id="ssp-files">${subtask.relatedDocs.map(subtaskFileRowHtml).join("")}</div>
          <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="ssp-add-file"><i class="bx bx-paperclip"></i> Add File</button>
        </div>`;
  const footHtml = `
        <button class="btn-danger-outline" id="ssp-delete" type="button">Delete</button>
        <button class="btn-aewttr-ghost" id="ssp-cancel" type="button">Cancel</button>
        <button class="btn-aewttr" id="ssp-save" type="button">Save Subtask</button>`;
  panel.innerHTML = taskEditorFrameHtml(isNew ? "Add Subtask" : "Edit Subtask", bodyHtml, footHtml, isModal);
  wireAssigneeAutocomplete(panel, "ssp-assignee");

  function wireFileRemovers() {
    $all(".sf-remove", panel).forEach(btn => btn.addEventListener("click", () => btn.closest(".subtask-file-row").remove()));
  }
  wireFileRemovers();
  $("#tsp-close", panel).addEventListener("click", closeTaskSidePanel);
  $("#ssp-cancel", panel).addEventListener("click", closeTaskSidePanel);
  $("#ssp-add-file", panel).addEventListener("click", () => {
    $("#ssp-files", panel).insertAdjacentHTML("beforeend", subtaskFileRowHtml());
    wireFileRemovers();
  });
  $("#ssp-delete", panel).addEventListener("click", async () => {
    itemList.splice(itemIndex, 1);
    syncTaskStatusFromSubtasks(task);
    Repo.save("actionItem", task, { projectCode: projectCodeForTaskList(tasks), source: "Tracker" });
    closeTaskSidePanel();
    toast("Subtask deleted", "success");
    onChange();
  });
  $("#ssp-save", panel).addEventListener("click", async () => {
    const next = normalizeTaskSubtask(task, {
      text: $("#ssp-title", panel).value.trim(),
      assignee: $("#ssp-assignee", panel).value.trim(),
      done: $("#ssp-done", panel).checked,
      health: $("#ssp-health", panel).value,
      start: $("#ssp-start", panel).value,
      end: $("#ssp-end", panel).value,
      linked: $("#ssp-linked", panel).checked,
      relatedDocs: $all(".subtask-file-row", panel).map(row => ({
        fileName: $(".sf-name", row).value.trim(),
        location: $(".sf-loc", row).value
      })).filter(doc => doc.fileName)
    });
    if (!next.text) { toast("Subtask title is required", "error"); return; }
    if (next.start && next.end && next.start > next.end) { toast("End date must be on or after start date", "error"); return; }
    itemList[itemIndex] = next;
    syncTaskStatusFromSubtasks(task);
    const savedProjectCode = projectCodeForTaskList(tasks);
    ensureAssigneesFromTask(savedProjectCode, task);
    Repo.save("actionItem", task, { projectCode: savedProjectCode, source: "Tracker" });
    closeTaskSidePanel();
    toast("Subtask updated", "success");
    onChange();
  });
}

function openGanttTaskModal(tasks, onDone) {
  const today = new Date().toISOString().slice(0, 10);
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>Add Task</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body">
      <div class="form-row"><label>Title</label><input class="input-aewttr" id="gt-title"></div>
      <div class="form-grid-2">
        <div class="form-row"><label>Assignee</label><input class="input-aewttr" id="gt-assignee"></div>
        <div class="form-row"><label>Health</label>
          <select class="select-aewttr" id="gt-health">${["On Track", "At Risk", "Off Track"].map(s => `<option>${s}</option>`).join("")}</select>
        </div>
      </div>
      <div class="form-grid-2">
        <div class="form-row"><label>Start date</label><input type="date" class="input-aewttr" id="gt-start" value="${today}"></div>
        <div class="form-row"><label>End date</label><input type="date" class="input-aewttr" id="gt-end" value=""></div>
      </div>
      <div class="form-row"><label>Initial subtasks (optional, one per line)</label><textarea class="textarea-aewttr" id="gt-subtasks"></textarea></div>
      <p style="font-size:11.5px;color:var(--aewttr-muted);margin:0;">After you add the task, the full editor lets you assign each subtask its own dates plus file references in SPO or Windchill.</p>
    </div>
    <div class="aewttr-modal-foot">
      <button class="btn-aewttr-ghost" id="gt-cancel">Cancel</button>
      <button class="btn-aewttr" id="gt-save">Add Task</button>
    </div>
  `);
  wireAssigneeAutocomplete(modal, "gt-assignee");
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#gt-cancel", modal).addEventListener("click", closeModal);
  $("#gt-save", modal).addEventListener("click", async () => {
    const title = $("#gt-title", modal).value.trim();
    if (!title) { toast("Title is required", "error"); return; }
    const start = $("#gt-start", modal).value, end = $("#gt-end", modal).value;
    if (start && end && start > end) { toast("End date must be on or after start date", "error"); return; }
    const id = uid("g");
    const newTask = {
      id, title, assignee: $("#gt-assignee", modal).value.trim() || "Unassigned",
      start, end, health: $("#gt-health", modal).value, status: "Not Started",
      subtasks: $("#gt-subtasks", modal).value.split("\n").map(s => s.trim()).filter(Boolean).map(text => ({
        text, assignee: "", done: false, health: $("#gt-health", modal).value, start, end, relatedDocs: []
      }))
    };
    syncTaskStatusFromSubtasks(newTask);
    tasks.push(newTask);
    const newTaskProjectCode = projectCodeForTaskList(tasks);
    ensureAssigneesFromTask(newTaskProjectCode, newTask);
    Repo.save("actionItem", newTask, { projectCode: newTaskProjectCode, source: "Tracker" });
    notifyTaskAssignee(newTask, newTaskProjectCode);
    closeModal();
    toast("Task added", "success");
    onDone(id);
  });
}

/* ---------- Project Boards (per-project custom boards) ---------- */
function drawEngChecklists(body, proj) {
  drawProjectBoardsTab(body, proj, null);
}

function drawProjectSettings(body, proj) {
  const db = window.AEWTTR.db;
  const extra = ensureProjectExtra(proj.id);
  const selectedPortfolios = new Set(projectPortfolios(proj));
  const selectedLocations = new Set(projectLocations(proj));
  const selectedConfigEnd = new Set();
  const initialConfigEnd = normalizeConfigEndItemName(proj.configEndItem || "");
  if (initialConfigEnd) selectedConfigEnd.add(initialConfigEnd);
  const selectedProgram = new Set(); if (proj.program) String(proj.program).split(",").forEach(p => { const v = normalizeProgramName(p.trim()); if (v) selectedProgram.add(v); });
  const selectedTaskOrder = new Set(); if (proj.taskOrder) String(proj.taskOrder).split(",").forEach(p => { const v = normalizeTaskOrderName(p.trim()); if (v) selectedTaskOrder.add(v); });
  const selectedFundingType = new Set(); if (proj.fundingType) String(proj.fundingType).split(",").forEach(p => { const v = normalizeFundingTypeName(p.trim()); if (v) selectedFundingType.add(v); });
  const selectedFiscalYear = new Set(); if (proj.fiscalYear) String(proj.fiscalYear).split(",").forEach(p => { const v = normalizeFiscalYearName(p.trim()); if (v) selectedFiscalYear.add(v); });
  const selectedFundingStatus = new Set(); if (proj.fundingStatus) String(proj.fundingStatus).split(",").forEach(p => { const v = normalizeFundingStatusName(p.trim()); if (v) selectedFundingStatus.add(v); });

  let pendingCover = proj.coverImage || "";
  let pendingCoverFile = null;
  let saveTimer = null;
  let savePendingLabel = null;
  let saving = false;

  body.innerHTML = `
    <div class="project-settings-shell">
      <header class="project-settings-bar">
        <div class="project-settings-bar-copy">
          <h2>Project settings</h2>
          <span class="ps-autosave-status" id="ps-autosave-status" data-state="saved">All changes saved</span>
        </div>
        <button type="button" class="btn-danger-outline btn-aewttr-sm" id="ps-delete"><i class="bx bx-trash"></i> Delete project</button>
      </header>

      <div class="project-settings-grid">
        <section class="ps-panel">
          <div class="ps-panel-title">Identity</div>
          <div class="ps-identity-row">
            <div class="cover-upload-row ps-cover-compact">
              <div class="cover-upload-preview" id="ps-cover-preview" style="${proj.coverImage ? `background-image:url('${proj.coverImage}');` : `background:${projectBannerColor(proj)};`}"></div>
              <div>
                <input type="file" accept="image/*" id="ps-cover-file" style="display:none;">
                <button type="button" class="btn-aewttr-outline btn-aewttr-sm" id="ps-cover-pick">Upload</button>
                ${proj.coverImage || pendingCover ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="ps-cover-clear">Remove</button>` : `<button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="ps-cover-clear" hidden>Remove</button>`}
              </div>
            </div>
            <div class="ps-identity-fields">
              <div class="form-row"><label>Project name</label><input class="input-aewttr" id="ps-name" value="${escapeHtml(proj.name)}"></div>
              <div class="form-row"><label>Description</label><textarea class="textarea-aewttr" id="ps-desc" rows="3" placeholder="One or two sentences describing this project.">${escapeHtml(proj.description || "")}</textarea></div>
            </div>
          </div>
          <div class="form-row" style="margin-bottom:0;">
            <label>Portfolios ${glossaryTip("Portfolio")}</label>
            ${portfolioPickerHtml(Array.from(selectedPortfolios), "ps-portfolios")}
          </div>
        </section>

        <section class="ps-panel">
          <div class="ps-panel-title">Status &amp; schedule</div>
          <div class="form-grid-2">
            <div class="form-row"><label>Technical Status ${glossaryTip("Lifecycle")}</label>
              <select class="select-aewttr" id="ps-tech-status">
                <option value="" ${!proj.technicalStatus ? "selected" : ""}>Not set (auto from risks)</option>
                <option value="Green" ${proj.technicalStatus === "Green" ? "selected" : ""}>Green — On track</option>
                <option value="Amber" ${proj.technicalStatus === "Amber" ? "selected" : ""}>Amber — Watch</option>
                <option value="Red" ${proj.technicalStatus === "Red" ? "selected" : ""}>Red — At risk</option>
              </select>
            </div>
            <div class="form-row"><label>Priority</label>
              <select class="select-aewttr" id="ps-priority">
                <option value="" ${proj.priority === "" ? "selected" : ""}>None</option>
                <option value="High" ${proj.priority === "High" ? "selected" : ""}>High</option>
                <option value="Medium" ${proj.priority === "Medium" ? "selected" : ""}>Medium</option>
                <option value="Low" ${proj.priority === "Low" ? "selected" : ""}>Low</option>
              </select>
            </div>
            <div class="form-row"><label>Lifecycle ${glossaryTip("Lifecycle")}</label>
              <select class="select-aewttr" id="ps-lifecycle">
                <option value="" ${!proj.lifecycleStatus ? "selected" : ""}>None</option>
                <option value="Planned" ${proj.lifecycleStatus === "Planned" || proj.lifecycleStatus === "Planning" ? "selected" : ""}>Planned</option>
                <option value="Awaiting Funding" ${proj.lifecycleStatus === "Awaiting Funding" ? "selected" : ""}>Awaiting Funding</option>
                <option value="Active" ${proj.lifecycleStatus === "Active" ? "selected" : ""}>Active</option>
                <option value="Paused" ${proj.lifecycleStatus === "Paused" || proj.lifecycleStatus === "On Hold" ? "selected" : ""}>Paused</option>
                <option value="Complete" ${proj.lifecycleStatus === "Complete" || proj.lifecycleStatus === "Completed" ? "selected" : ""}>Complete</option>
              </select>
            </div>
            <div class="form-row"><label>Start</label><input type="date" class="input-aewttr" id="ps-startdate" value="${escapeHtml(proj.startDate)}"></div>
            <div class="form-row"><label>Due</label><input type="date" class="input-aewttr" id="ps-duedate" value="${escapeHtml(proj.dueDate)}"></div>
            <div class="form-row"><label>Completed</label><input type="date" class="input-aewttr" id="ps-completiondate" value="${escapeHtml(proj.completionDate)}"></div>
            <div class="form-row">
              <label>Change request</label>
              <div class="travel-choice-row">${travelChoiceGroup("ps-crr", ["Yes", "No"], proj.changeRequestRequired ? "Yes" : "No")}</div>
            </div>
          </div>
          <div class="form-row" style="margin-bottom:0;">
            <label>Locations</label>
            ${locationPickerHtml(Array.from(selectedLocations), "ps-locations")}
          </div>
        </section>

        <section class="ps-panel ps-panel--roles">
          <div class="ps-panel-title">Roles</div>
          <div class="project-role-picker-grid">
            ${ASSIGNABLE_PROJECT_ROLE_FIELDS.filter((f) => f.key !== "contractor").map((field) => projectRolePickerHtml(proj, field)).join("")}
            ${projectContractorCompanyPickerHtml(proj)}
          </div>
        </section>

        <section class="ps-panel">
          <div class="ps-panel-title">Program &amp; funding</div>
          <div class="form-grid-2">
            <div class="form-row"><label>Program ${glossaryTip("Program")}</label>${tagPickerHtml(Array.from(selectedProgram), "ps-program", { emptyText: "No program set.", placeholder: "Search or add program…", hint: "" })}</div>
            <div class="form-row"><label>Contract ${glossaryTip("Contract")}</label><input class="input-aewttr" id="ps-contract" value="${escapeHtml(proj.contract)}"></div>
            <div class="form-row"><label>Task Order</label>${tagPickerHtml(Array.from(selectedTaskOrder), "ps-taskorder", { emptyText: "No task order set.", placeholder: "Search or add task order…", hint: "" })}</div>
            <div class="form-row"><label>Funding Type</label>${tagPickerHtml(Array.from(selectedFundingType), "ps-fundingtype", { emptyText: "No funding type set.", placeholder: "Search or add funding type…", hint: "" })}</div>
            <div class="form-row"><label>Fiscal Year</label>${tagPickerHtml(Array.from(selectedFiscalYear), "ps-fiscalyear", { emptyText: "No fiscal year set.", placeholder: "e.g. FY26", hint: "" })}</div>
            <div class="form-row"><label>Funding Status</label>${tagPickerHtml(Array.from(selectedFundingStatus), "ps-fundingstatus", { emptyText: "No funding status set.", placeholder: "Search or add status…", hint: "" })}</div>
            <div class="form-row"><label>Funded on Contract ($)</label><input class="input-aewttr" id="ps-funded-amount" type="number" min="0" step="1" placeholder="0" value="${proj.fundedOnContractAmount || ""}"></div>
            <div class="form-row"><label>Reimbursable Amount ($)</label><input class="input-aewttr" id="ps-reimb-amount" type="number" min="0" step="1" placeholder="0" value="${proj.reimbursableAmount || ""}"></div>
          </div>
          <div class="form-row"><label>Funding Notes</label><textarea class="textarea-aewttr" id="ps-fundingnotes" rows="2" placeholder="Additional funding details, notes on changes…">${escapeHtml(proj.fundingNotes || "")}</textarea></div>
          <div class="form-row" style="margin-bottom:0;"><label>Configuration End Item ${glossaryTip("End-item config")}</label>${configEndItemPickerHtml(Array.from(selectedConfigEnd), "ps-config")}</div>
        </section>

        <section class="ps-panel">
          <div class="ps-panel-title">Classification</div>
          <div class="form-grid-2">
            <div class="form-row"><label>Project Type</label>
              <select class="select-aewttr" id="ps-project-type">
                <option value="" ${!proj.projectType ? "selected" : ""}>Not set</option>
                <option value="Development" ${proj.projectType === "Development" ? "selected" : ""}>Development</option>
                <option value="Sustainment" ${proj.projectType === "Sustainment" ? "selected" : ""}>Sustainment</option>
                <option value="Development &amp; Sustainment" ${proj.projectType === "Development & Sustainment" ? "selected" : ""}>Development &amp; Sustainment</option>
                <option value="Other" ${proj.projectType === "Other" ? "selected" : ""}>Other</option>
              </select>
            </div>
            <div class="form-row"><label>ATO</label><input class="input-aewttr" id="ps-ato" value="${escapeHtml(proj.ato || "")}" placeholder="e.g. ATO reference or expiration date"></div>
            <div class="form-row">
              <label>Aqu Only</label>
              <div class="travel-choice-row">${travelChoiceGroup("ps-aquonly", ["Yes", "No"], proj.aquOnly ? "Yes" : "No")}</div>
            </div>
            <div class="form-row"><label>Projects</label><input class="input-aewttr" id="ps-projects" value="${escapeHtml(proj.projects || "")}" placeholder="Related projects"></div>
          </div>
        </section>

        <section class="ps-panel">
          <div class="ps-panel-title">Scope &amp; objectives</div>
          <div class="form-row"><label>Scope</label><textarea class="textarea-aewttr" id="ps-scope" rows="4" placeholder="What this project covers.">${escapeHtml(proj.scope || "")}</textarea></div>
          <div class="form-row" style="margin-bottom:0;"><label>Objectives</label><textarea class="textarea-aewttr" id="ps-objectives" rows="4" placeholder="What this project is trying to achieve.">${escapeHtml(proj.objectives || "")}</textarea></div>
        </section>

        <section class="ps-panel">
          <div class="ps-panel-title">Handoff</div>
          <div class="form-row" style="margin-bottom:0;"><label>Handoff notes</label><textarea class="textarea-aewttr" id="ps-handoff" rows="4">${escapeHtml(extra.handoff)}</textarea></div>
        </section>

        <section class="ps-panel">
          <div class="ps-panel-title">SharePoint</div>
          <div class="form-row" style="margin-bottom:0;">
            <label>Project SharePoint folder URL</label>
            <input class="input-aewttr" id="ps-spo-folder" value="${escapeHtml(proj.sharepointFolderUrl || "")}" placeholder="https://tenant.sharepoint.com/sites/…/Shared Documents/…">
            <p class="form-hint" style="margin:5px 0 0;font-size:11.5px;color:var(--aewttr-muted);">Paste the URL to this project's SharePoint document folder. Shown in the Documents tab and used as the export destination.</p>
          </div>
        </section>
      </div>
    </div>
  `;

  const statusEl = $("#ps-autosave-status", body);
  function setAutosaveState(state, message) {
    if (!statusEl) return;
    statusEl.dataset.state = state;
    statusEl.textContent = message;
  }

  async function persistSettings({ fromCover } = {}) {
    if (typeof reanchorProject === "function") reanchorProject(proj);
    const liveDb = window.AEWTTR.db;
    const name = $("#ps-name", body).value.trim();
    if (!name) {
      setAutosaveState("error", "Project name is required");
      return false;
    }
    proj.name = name;
    proj.description = $("#ps-desc", body).value.trim();
    proj.portfolios = Array.from(selectedPortfolios);
    rememberPortfolioNames(proj.portfolios);
    proj.program = Array.from(selectedProgram).join(", ") || "";
    if (proj.program) rememberProgramNames(Array.from(selectedProgram));
    proj.contract = $("#ps-contract", body).value.trim();
    proj.taskOrder = Array.from(selectedTaskOrder).join(", ") || "";
    if (proj.taskOrder) rememberTaskOrderNames(Array.from(selectedTaskOrder));
    proj.fundingType = Array.from(selectedFundingType).join(", ") || "";
    if (proj.fundingType) rememberFundingTypeNames(Array.from(selectedFundingType));
    proj.fiscalYear = Array.from(selectedFiscalYear).join(", ") || "";
    if (proj.fiscalYear) rememberFiscalYearNames(Array.from(selectedFiscalYear));
    proj.fundingStatus = Array.from(selectedFundingStatus).join(", ") || "";
    if (proj.fundingStatus) rememberFundingStatusNames(Array.from(selectedFundingStatus));
    proj.fundedOnContractAmount = parseFloat($("#ps-funded-amount", body).value) || 0;
    proj.reimbursableAmount = parseFloat($("#ps-reimb-amount", body).value) || 0;
    proj.fundingNotes = ($("#ps-fundingnotes", body) ? $("#ps-fundingnotes", body).value.trim() : "") || "";
    proj.configEndItem = Array.from(selectedConfigEnd)[0] || "";
    if (proj.configEndItem) rememberConfigEndItemNames([proj.configEndItem]);
    proj.scope = $("#ps-scope", body).value.trim();
    proj.objectives = $("#ps-objectives", body).value.trim();
    proj.technicalStatus = $("#ps-tech-status", body).value;
    proj.priority = $("#ps-priority", body).value;
    proj.lifecycleStatus = $("#ps-lifecycle", body).value;
    const crrChecked = $(`input[name="ps-crr"]:checked`, body);
    proj.changeRequestRequired = !!crrChecked && crrChecked.value === "Yes";
    proj.locations = Array.from(selectedLocations);
    proj.startDate = $("#ps-startdate", body).value;
    proj.dueDate = $("#ps-duedate", body).value;
    proj.completionDate = $("#ps-completiondate", body).value;
    extra.handoff = $("#ps-handoff", body).value;
    const spoFolderEl = $("#ps-spo-folder", body);
    if (spoFolderEl) proj.sharepointFolderUrl = spoFolderEl.value.trim();
    const projTypeEl = $("#ps-project-type", body);
    if (projTypeEl) proj.projectType = projTypeEl.value;
    const atoEl = $("#ps-ato", body);
    if (atoEl) proj.ato = atoEl.value.trim();
    const aquOnlyChecked = $(`input[name="ps-aquonly"]:checked`, body);
    proj.aquOnly = !!aquOnlyChecked && aquOnlyChecked.value === "Yes";
    const projectsEl = $("#ps-projects", body);
    if (projectsEl) proj.projects = projectsEl.value.trim();
    if (fromCover || pendingCoverFile || pendingCover !== (proj.coverImage || "")) {
      try {
        proj.coverImage = await persistProjectCoverImage(proj, pendingCover, pendingCoverFile);
        pendingCoverFile = null;
        pendingCover = proj.coverImage || "";
      } catch (e) {
        setAutosaveState("error", (e && e.friendly) || e.message || "Cover upload failed");
        return false;
      }
    }
    proj.updated = new Date().toISOString().slice(0, 10);
    liveDb.projectExtra[proj.id] = extra;
    saving = true;
    setAutosaveState("saving", "Saving…");
    try {
      await Repo.save("project", proj);
      if (typeof syncProjectPulseGroup === "function") {
        try { await syncProjectPulseGroup(proj); } catch (e) { console.warn("project group sync", e); }
      }
      setAutosaveState("saved", "All changes saved");
      return true;
    } catch (e) {
      setAutosaveState("error", (e && e.friendly) || "Couldn’t save");
      return false;
    } finally {
      saving = false;
    }
  }

  function scheduleAutosave(opts) {
    clearTimeout(saveTimer);
    clearTimeout(savePendingLabel);
    setAutosaveState("pending", "Saving soon…");
    saveTimer = setTimeout(() => {
      saveTimer = null;
      persistSettings(opts || {});
    }, 900);
  }

  wireProjectRolePickers(body, proj, () => {
    // Role writes already went through Repo.save. Never force "saved" while a
    // form autosave is still pending — that race let background refresh remount
    // the form and wipe text that hadn't been flushed yet.
    if (saving || saveTimer) return;
    if (statusEl && (statusEl.dataset.state === "pending" || statusEl.dataset.state === "saving")) return;
    setAutosaveState("saved", "All changes saved");
  });
  wireProjectContractorCompanyPicker(body, proj, () => {
    if (saving || saveTimer) return;
    if (statusEl && (statusEl.dataset.state === "pending" || statusEl.dataset.state === "saving")) return;
    setAutosaveState("saved", "All changes saved");
  });
  wirePortfolioPicker(body, selectedPortfolios, "ps-portfolios", () => scheduleAutosave());
  wireLocationPicker(body, selectedLocations, "ps-locations", () => scheduleAutosave());
  wireConfigEndItemPicker(body, selectedConfigEnd, "ps-config", () => scheduleAutosave());
  wireTagPicker(body, selectedProgram, "ps-program", { normalize: normalizeProgramName, getKnown: getKnownProgramNames, remember: rememberProgramNames, singleSelect: false, onChange: () => scheduleAutosave() });
  wireTagPicker(body, selectedTaskOrder, "ps-taskorder", { normalize: normalizeTaskOrderName, getKnown: getKnownTaskOrderNames, remember: rememberTaskOrderNames, singleSelect: false, onChange: () => scheduleAutosave() });
  wireTagPicker(body, selectedFundingType, "ps-fundingtype", { normalize: normalizeFundingTypeName, getKnown: getKnownFundingTypeNames, remember: rememberFundingTypeNames, singleSelect: false, onChange: () => scheduleAutosave() });
  wireTagPicker(body, selectedFiscalYear, "ps-fiscalyear", { normalize: normalizeFiscalYearName, getKnown: getKnownFiscalYearNames, remember: rememberFiscalYearNames, singleSelect: false, onChange: () => scheduleAutosave() });
  wireTagPicker(body, selectedFundingStatus, "ps-fundingstatus", { normalize: normalizeFundingStatusName, getKnown: getKnownFundingStatusNames, remember: rememberFundingStatusNames, singleSelect: false, onChange: () => scheduleAutosave() });

  ["ps-name", "ps-desc", "ps-contract", "ps-scope", "ps-objectives", "ps-handoff", "ps-funded-amount", "ps-reimb-amount", "ps-fundingnotes"].forEach((id) => {
    const field = $(`#${id}`, body);
    if (!field) return;
    field.addEventListener("input", () => scheduleAutosave());
    field.addEventListener("change", () => scheduleAutosave());
  });
  ["ps-tech-status", "ps-priority", "ps-lifecycle", "ps-startdate", "ps-duedate", "ps-completiondate"].forEach((id) => {
    const field = $(`#${id}`, body);
    if (field) field.addEventListener("change", () => scheduleAutosave());
  });
  $all(`input[name="ps-crr"]`, body).forEach((input) => input.addEventListener("change", () => scheduleAutosave()));

  $("#ps-cover-pick", body).addEventListener("click", () => $("#ps-cover-file", body).click());
  $("#ps-cover-file", body).addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;
    pendingCoverFile = file;
    const reader = new FileReader();
    reader.onload = () => {
      pendingCover = reader.result;
      $("#ps-cover-preview", body).style.background = `url('${pendingCover}') center/cover`;
      const clearBtn = $("#ps-cover-clear", body);
      if (clearBtn) clearBtn.hidden = false;
      scheduleAutosave({ fromCover: true });
    };
    reader.readAsDataURL(file);
  });
  const clearBtn = $("#ps-cover-clear", body);
  if (clearBtn) clearBtn.addEventListener("click", () => {
    pendingCover = "";
    pendingCoverFile = null;
    $("#ps-cover-preview", body).style.backgroundImage = "none";
    $("#ps-cover-preview", body).style.background = projectBannerColor(proj);
    clearBtn.hidden = true;
    scheduleAutosave({ fromCover: true });
  });

  $("#ps-delete", body).addEventListener("click", async () => {
    const ok = await confirmDialog({
      title: "Delete project",
      message: `Delete ${proj.id} — ${proj.name}? This cannot be undone in this session.`,
      confirmLabel: "Delete",
      danger: true
    });
    if (!ok) return;
    clearTimeout(saveTimer);
    clearTimeout(savePendingLabel);
    db.projects = db.projects.filter(p => p.id !== proj.id);
    Repo.remove("project", proj);
    toast("Project deleted", "success");
    navigate("projects");
  });

  // Background refresh updates window.AEWTTR.db in place but never
  // re-renders on its own — this tab is where PM/Engineer/ISSO/Range POC/
  // Contractor actually get reassigned, so it's exactly where a stale view
  // is most likely to cause real confusion ("I changed the ISSO and it
  // didn't update"). Skips the redraw while an autosave is in flight or
  // still in its 450ms debounce window (see setAutosaveState/saveTimer
  // above) so a background refresh never clobbers an edit that hasn't
  // landed yet, in addition to the usual focused-input/open-modal guards.
  if (window.AEWTTR._projectSettingsLiveRefreshHandler) {
    window.removeEventListener("pulse:data-refreshed", window.AEWTTR._projectSettingsLiveRefreshHandler);
  }
  window.AEWTTR._projectSettingsLiveRefreshHandler = () => {
    if (!document.querySelector(".project-settings-shell")) return;
    if (document.querySelector(".aewttr-modal-backdrop")) return;
    const active = document.activeElement;
    if (active && (active.tagName === "INPUT" || active.tagName === "TEXTAREA" || active.tagName === "SELECT" || active.isContentEditable)) return;
    if (saving || saveTimer || (statusEl && statusEl.dataset.state !== "saved")) return;
    const current = (window.AEWTTR.db.projects || []).find((p) => p.id === proj.id);
    if (!current) return;
    const scroller = document.querySelector(".aewttr-content");
    const scrollTop = scroller ? scroller.scrollTop : 0;
    drawProjectSettings(body, current);
    if (scroller) scroller.scrollTop = scrollTop;
  };
  window.addEventListener("pulse:data-refreshed", window.AEWTTR._projectSettingsLiveRefreshHandler);
}

function drawProjectTickets(body, proj) {
  const db = window.AEWTTR.db;
  if (!window.AEWTTR.state.projectTicketFilters) window.AEWTTR.state.projectTicketFilters = {};
  if (!window.AEWTTR.state.projectTicketFilters[proj.id]) window.AEWTTR.state.projectTicketFilters[proj.id] = { status: "All", type: "All", search: "" };
  const st = window.AEWTTR.state.projectTicketFilters[proj.id];
  const STATUS_TABS = ["All", "Open", "In Progress", "Resolved"];
  const TYPES = ["All", "Blocker", "Bug", "Access", "Platform", "Question"];

  function projectTickets() {
    return db.tickets.filter(t => t.project === proj.id || (t.affected || []).includes(proj.id));
  }

  function draw() {
    const allRows = projectTickets();
    const rows = allRows.filter(t => {
      if (st.status !== "All" && t.status !== st.status) return false;
      if (st.type !== "All" && t.type !== st.type) return false;
      if (st.search && !(t.title.toLowerCase().includes(st.search.toLowerCase()) || t.id.toLowerCase().includes(st.search.toLowerCase()))) return false;
      return true;
    });
    body.innerHTML = `
      <div class="tickets-project-page">
        <div class="tickets-project-head">
          <div><h3>Project tickets</h3><p>Blockers, bugs, access needs, and questions linked to this project.</p></div>
          <button class="btn-aewttr" id="pt-new-ticket"><i class="bx bx-plus"></i> New Project Ticket</button>
        </div>
        <div class="tickets-controls tickets-controls--project">
          <div class="search-box tickets-search"><i class="bx bx-search"></i><input id="pt-search" placeholder="Search this project's tickets" value="${escapeHtml(st.search)}"></div>
          <div class="filter-pills tickets-status-filters" id="pt-status-pills">
            ${STATUS_TABS.map(f => {
              const count = f === "All" ? allRows.length : allRows.filter((ticket) => ticket.status === f).length;
              return `<button class="filter-pill ${st.status === f ? "active" : ""}" data-f="${f}">${f}<span>${count}</span></button>`;
            }).join("")}
          </div>
          <select class="select-aewttr tickets-type-filter" id="pt-type-filter" aria-label="Filter project tickets by type">
            ${TYPES.map(t => `<option value="${t}" ${st.type === t ? "selected" : ""}>${t === "All" ? "All Types" : t}</option>`).join("")}
          </select>
        </div>
        <div class="tickets-table-shell">
          <div class="tickets-table-meta"><strong>${rows.length}</strong> of ${allRows.length} project ticket${allRows.length === 1 ? "" : "s"}</div>
          <div class="tickets-table-scroll">
            <table class="aewttr-table tickets-table">
              <thead><tr><th>Ticket</th><th>Title</th><th>Type</th><th>Status</th><th>Opened</th><th>Reporter</th></tr></thead>
              <tbody>
                ${rows.length ? rows.map(t => `
                  <tr data-id="${t.id}">
                    <td><span class="tickets-id">${escapeHtml(t.id)}</span></td>
                    <td><strong>${escapeHtml(t.title)}</strong></td>
                    <td><span class="type-badge">${escapeHtml(t.type)}</span></td>
                    <td>${statusPill(t.status)}</td>
                    <td>${fmtDate(t.opened)}</td>
                    <td>${escapeHtml(t.reporter)}</td>
                  </tr>`).join("") : `<tr><td colspan="6"><div class="tickets-empty"><strong>No project tickets found</strong><span>Adjust the filters or create a ticket for this project.</span></div></td></tr>`}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    `;
    $all("tr[data-id]", body).forEach(r => r.addEventListener("click", () => openTicketModal(r.dataset.id, draw)));
    $all("[data-f]", $("#pt-status-pills", body)).forEach(b => b.addEventListener("click", () => { st.status = b.dataset.f; draw(); }));
    $("#pt-type-filter", body).addEventListener("change", (e) => { st.type = e.target.value; draw(); });
    $("#pt-search", body).addEventListener("input", (e) => { st.search = e.target.value; draw(); });
    $("#pt-new-ticket", body).addEventListener("click", () => openNewTicketModalWithOptions({ projectId: proj.id, lockProject: true }, draw));
  }
  draw();
}

function drawProjectReporting(body, proj) {
  const extra = ensureProjectExtra(proj.id);
  if (!extra.reportConfig) extra.reportConfig = {};
  const cfg = extra.reportConfig;
  if (cfg.isMilestone === undefined) cfg.isMilestone = false;
  if (!cfg.slideTypes) cfg.slideTypes = { status: true, risks: false, milestones: true, tasks: false };
  if (!cfg.slideContent) cfg.slideContent = { overallRag: "", statusSummary: "", techRag: "", techNarrative: "", objectives: "", deliverablesRag: "", deliverablesText: "", milestonesRag: "",
    riskRows: [{ cat: "Schedule", rag: "", notes: "" }, { cat: "Budget", rag: "", notes: "" }, { cat: "Technical", rag: "", notes: "" }]
  };
  if (!cfg.slideContent.riskRows) cfg.slideContent.riskRows = [{ cat: "Schedule", rag: "", notes: "" }, { cat: "Budget", rag: "", notes: "" }, { cat: "Technical", rag: "", notes: "" }];
  if (!cfg.taskMilestones) cfg.taskMilestones = {};
  if (!cfg.techBullets) cfg.techBullets = [];

  const sc = cfg.slideContent;
  // This is the one project-level milestone store used by every reporting
  // surface (Project, Portfolio, Program, End Item Config, and PPTX export).
  const tm = typeof getProjectReportingMilestones === "function"
    ? getProjectReportingMilestones(proj)
    : cfg.taskMilestones;

  const MC = [
    { key: "contractAwarded", label: "Contract Award",      short: "Contract", color: "#7f7f7f", sym: "★" },
    { key: "fat",             label: "First Article Test",  short: "FAT",      color: "#70d6c8", sym: "△" },
    { key: "sat",             label: "Site Acceptance Test", short: "SAT",     color: "#bdd7ee", sym: "△" },
    { key: "add",             label: "ADD / MFR",           short: "ADD",      color: "#c55a11", sym: "△" },
    { key: "fielding",        label: "Fielding",            short: "Field",    color: "#7030a0", sym: "△" },
    { key: "complete",        label: "Project Completion",  short: "Complete", color: "#ff0000", sym: "★" }
  ];

  const allTasks = (window.AEWTTR.db.ganttTasks && window.AEWTTR.db.ganttTasks[proj.id]) || [];
  const plainTasks = trackerPlainTasks(allTasks);

  // Auto-seed complete date from tracker status
  let autoSeededMilestones = false;
  plainTasks.forEach(function(task) {
    if (!tm[task.id]) tm[task.id] = {};
    if (/complete|done/i.test(String(task.status || "")) && !tm[task.id].complete) {
      tm[task.id].complete = task.dueDate || task.end || "";
      autoSeededMilestones = !!tm[task.id].complete;
    }
  });

  function saveConfig() {
    proj.updated = new Date().toISOString().slice(0, 10);
    if (typeof Repo !== "undefined" && Repo && typeof Repo.save === "function") {
      Repo.save("project", proj);
    } else if (typeof aewttrSaveStore === "function") {
      aewttrSaveStore();
    }
  }

  if (autoSeededMilestones) saveConfig();

  function setRagClass(sel, val) {
    sel.classList.remove("rep-rag--green", "rep-rag--amber", "rep-rag--red");
    if (val === "Green") sel.classList.add("rep-rag--green");
    else if (val === "Amber") sel.classList.add("rep-rag--amber");
    else if (val === "Red") sel.classList.add("rep-rag--red");
  }

  function renderMilestoneTable() {
    const wrap = $("#rep-milestone-table-wrap", body);
    if (!wrap) return;
    if (!plainTasks.length) {
      wrap.innerHTML = `<div class="rep-mt-empty">No tasks in tracker yet. Add tasks in the Tracker tab then return here to set milestone dates.</div>`;
      return;
    }
    wrap.innerHTML = `
      <div class="rep-mt-scroll">
        <table class="rep-mt-tbl">
          <thead>
            <tr>
              <th class="rep-mt-th rep-mt-th--sel" title="Include in report">In Report</th>
              <th class="rep-mt-th rep-mt-th--name">Task</th>
              ${MC.map(function(c) {
                return `<th class="rep-mt-th rep-mt-th--date" style="--mc:#${c.color.replace("#","")}">
                  <span class="rep-mt-sym" style="color:${c.color}">${c.sym}</span>${c.short}
                </th>`;
              }).join("")}
            </tr>
          </thead>
          <tbody>
            ${plainTasks.map(function(task) {
              const t = tm[task.id] || {};
              const isComplete = /complete|done/i.test(String(task.status || ""));
              return `<tr class="rep-mt-row${t.inReport ? " rep-mt-row--on" : ""}" data-task-id="${task.id}">
                <td class="rep-mt-td rep-mt-td--sel">
                  <button type="button" class="rep-mt-sel${t.inReport ? " rep-mt-sel--on" : ""}" data-task-id="${task.id}" title="${t.inReport ? "Remove from milestones slide" : "Add to milestones slide"}">
                    <i class="bx ${t.inReport ? "bxs-check-square" : "bx-checkbox"}"></i>
                  </button>
                </td>
                <td class="rep-mt-td rep-mt-td--name">
                  <span class="rep-mt-name">${escapeHtml(task.title || "Unnamed task")}</span>
                  ${isComplete ? `<span class="rep-mt-badge">Done</span>` : ""}
                </td>
                ${MC.map(function(c) {
                  const val = t[c.key] || "";
                  return `<td class="rep-mt-td rep-mt-td--date">
                    <input type="date" class="rep-mt-date${val ? " rep-mt-date--set" : ""}"
                      data-task-id="${task.id}" data-col="${c.key}"
                      value="${escapeHtml(val)}" style="--mc:${c.color}">
                  </td>`;
                }).join("")}
              </tr>`;
            }).join("")}
          </tbody>
        </table>
      </div>
    `;

    $all(".rep-mt-sel", wrap).forEach(function(btn) {
      btn.addEventListener("click", function() {
        const id = btn.dataset.taskId;
        const current = tm[id] || {};
        const next = typeof updateProjectReportingMilestone === "function"
          ? updateProjectReportingMilestone(proj, id, { inReport: !current.inReport })
          : (tm[id] = Object.assign(current, { inReport: !current.inReport }));
        if (next) tm[id] = next;
        saveConfig();
        renderMilestoneTable();
        renderTimeline();
        renderSlidePreview();
      });
    });

    $all(".rep-mt-date", wrap).forEach(function(inp) {
      inp.addEventListener("change", function() {
        const id = inp.dataset.taskId;
        const col = inp.dataset.col;
        const current = tm[id] || {};
        const next = typeof updateProjectReportingMilestone === "function"
          ? updateProjectReportingMilestone(proj, id, { [col]: inp.value })
          : (tm[id] = Object.assign(current, { [col]: inp.value }));
        if (next) tm[id] = next;
        inp.classList.toggle("rep-mt-date--set", !!inp.value);
        saveConfig();
        renderTimeline();
        renderSlidePreview();
      });
    });
  }

  function renderTimeline() {
    const wrap = $("#rep-timeline-wrap", body);
    if (!wrap) return;

    const reportTasks = plainTasks.filter(function(t) { return tm[t.id] && tm[t.id].inReport; });
    if (!reportTasks.length) {
      wrap.innerHTML = `<div class="rep-tl-empty">Select tasks above (check <strong>In Report</strong>) to see them on the timeline.</div>`;
      return;
    }

    const allMs = [];
    reportTasks.forEach(function(task) {
      const t = tm[task.id] || {};
      MC.forEach(function(c) { if (t[c.key]) allMs.push(new Date(t[c.key]).getTime()); });
    });
    if (!allMs.length) {
      wrap.innerHTML = `<div class="rep-tl-empty">Enter milestone dates in the table above to see the timeline.</div>`;
      return;
    }

    const minT = Math.min.apply(null, allMs);
    const maxT = Math.max.apply(null, allMs);
    const pad = Math.max((maxT - minT) * 0.1, 86400000 * 45);
    const rangeMin = minT - pad;
    const rangeMax = maxT + pad;
    const span = rangeMax - rangeMin;

    function toX(ds) { return ((new Date(ds).getTime() - rangeMin) / span) * 100; }

    const startY = new Date(rangeMin).getFullYear();
    const endY = new Date(rangeMax).getFullYear();
    const years = [];
    for (var y = startY; y <= endY; y++) years.push(y);

    const todayMs = Date.now();
    const todayX = ((todayMs - rangeMin) / span) * 100;
    const todayVisible = todayX >= 0 && todayX <= 100;

    wrap.innerHTML = `
      <div class="rep-tl">
        <div class="rep-tl-legend">
          ${MC.map(function(c) {
            return `<span class="rep-tl-leg-item"><span class="rep-tl-leg-sym" style="color:${c.color}">${c.sym}</span>${c.label}</span>`;
          }).join("")}
          ${todayVisible ? `<span class="rep-tl-leg-item rep-tl-leg-today"><span class="rep-tl-leg-line"></span>Today</span>` : ""}
        </div>
        <div class="rep-tl-stage">
          <div class="rep-tl-axis-row">
            <div class="rep-tl-row-label" style="visibility:hidden;"></div>
            <div class="rep-tl-axis">
              ${years.map(function(yr) {
                const x = ((new Date(yr, 0, 1).getTime() - rangeMin) / span) * 100;
                if (x < 0 || x > 100) return "";
                return `<div class="rep-tl-yr" style="left:${x.toFixed(2)}%"><span>${yr}</span></div>`;
              }).join("")}
              ${todayVisible ? `<div class="rep-tl-today-tick" style="left:${todayX.toFixed(2)}%"></div>` : ""}
            </div>
          </div>
          <div class="rep-tl-rows">
            ${reportTasks.map(function(task) {
              const t = tm[task.id] || {};
              const markers = MC.map(function(c) { return { c: c, v: t[c.key] }; }).filter(function(m) { return m.v; });
              const xs = markers.map(function(m) { return toX(m.v); });
              const minX = xs.length ? Math.min.apply(null, xs) : 0;
              const maxX = xs.length ? Math.max.apply(null, xs) : 0;
              return `<div class="rep-tl-row">
                <div class="rep-tl-row-label" title="${escapeHtml(task.title || "")}">${escapeHtml(task.title || "")}</div>
                <div class="rep-tl-track">
                  ${xs.length >= 2 ? `<div class="rep-tl-connector" style="left:${minX.toFixed(2)}%;width:${(maxX - minX).toFixed(2)}%"></div>` : ""}
                  ${todayVisible ? `<div class="rep-tl-today-line" style="left:${todayX.toFixed(2)}%"></div>` : ""}
                  ${markers.map(function(m) {
                    return `<div class="rep-tl-marker" style="left:${toX(m.v).toFixed(2)}%;color:${m.c.color}" title="${m.c.label}: ${m.v}">
                      <span class="rep-tl-marker-sym">${m.c.sym}</span>
                      <span class="rep-tl-marker-date">${m.v.slice(5)}</span>
                    </div>`;
                  }).join("")}
                </div>
              </div>`;
            }).join("")}
          </div>
        </div>
      </div>
    `;
  }

  function buildGanttHtml() {
    var reportTasks = plainTasks.filter(function(t) { return tm[t.id] && tm[t.id].inReport; });
    if (!reportTasks.length) return '<div class="rep-sld-empty-msg">Select tasks below (In Report)</div>';
    var allMs = [];
    reportTasks.forEach(function(task) {
      var t = tm[task.id] || {};
      MC.forEach(function(c) { if (t[c.key]) allMs.push(new Date(t[c.key]).getTime()); });
    });
    if (!allMs.length) return '<div class="rep-sld-empty-msg">Enter milestone dates below</div>';
    var minT = Math.min.apply(null, allMs);
    var maxT = Math.max.apply(null, allMs);
    var startMoDate = new Date(minT);
    var startMo = new Date(startMoDate.getFullYear(), 0, 1);
    var endMoDate = new Date(maxT);
    var endMo = new Date(endMoDate.getFullYear() + 1, 0, 1);
    var rMin = startMo.getTime();
    var rMax = endMo.getTime();
    var span = rMax - rMin;
    var mos = [];
    var cur = new Date(startMo.getFullYear(), startMo.getMonth(), 1);
    while (cur.getTime() < rMax) { mos.push(new Date(cur)); cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1); }
    var byYear = {}, yearsOrder = [];
    mos.forEach(function(m) {
      var yr = m.getFullYear();
      if (!byYear[yr]) { byYear[yr] = []; yearsOrder.push(yr); }
      byYear[yr].push(m);
    });
    function toFrac(ds) { return ((new Date(ds).getTime() - rMin) / span) * 100; }
    var todayFrac = ((Date.now() - rMin) / span) * 100;
    var showToday = todayFrac >= 0 && todayFrac <= 100;
    var MONTHS = "JFMAMJJASOND";
    var hdrHtml = '<div class="rep-sld-gantt-hdr">' +
      '<div class="rep-sld-gantt-lbl-col rep-sld-gantt-project-head">Project</div>' +
      '<div class="rep-sld-gantt-hdr-cols">' +
      yearsOrder.map(function(yr) {
        return '<div class="rep-sld-gantt-yr-grp" style="flex:' + byYear[yr].length + '">' +
          '<div class="rep-sld-gantt-yr-label">' + yr + '</div>' +
          '<div class="rep-sld-gantt-mo-row">' +
          byYear[yr].map(function(m) { return '<div class="rep-sld-gantt-mo">' + MONTHS[m.getMonth()] + '</div>'; }).join("") +
          '</div></div>';
      }).join("") +
      '</div></div>';
    var rowsHtml = reportTasks.slice(0, 7).map(function(task) {
      var t = tm[task.id] || {};
      var markers = MC.map(function(c) { return { c: c, v: t[c.key] }; }).filter(function(m) { return m.v; });
      var fracs = markers.map(function(m) { return toFrac(m.v); });
      var minF = fracs.length ? Math.min.apply(null, fracs) : null;
      var maxF = fracs.length ? Math.max.apply(null, fracs) : null;
      return '<div class="rep-sld-gantt-row">' +
        '<div class="rep-sld-gantt-lbl-col">' + escapeHtml((task.title || "").slice(0, 16)) + '</div>' +
        '<div class="rep-sld-gantt-track">' +
        (showToday ? '<div class="rep-sld-gantt-today" style="left:' + todayFrac.toFixed(1) + '%"></div>' : "") +
        (minF !== null && maxF !== null && fracs.length >= 2 ? '<div class="rep-sld-gantt-bar" style="left:' + minF.toFixed(1) + '%;width:' + (maxF - minF).toFixed(1) + '%"></div>' : "") +
        markers.map(function(m, i) { return '<span class="rep-sld-gantt-sym" style="left:' + fracs[i].toFixed(1) + '%;color:' + m.c.color + '">' + m.c.sym + '</span>'; }).join("") +
        '</div></div>';
    }).join("");
    var legendHtml = '<div class="rep-sld-gantt-legend">' +
      MC.map(function(c) { return '<span class="rep-sld-gantt-legend-item"><span style="color:' + c.color + '">' + c.sym + '</span>' + c.short + '</span>'; }).join('') +
      '</div>';
    return '<div class="rep-sld-gantt">' + hdrHtml + '<div class="rep-sld-gantt-body">' + rowsHtml + '</div>' + legendHtml + '</div>';
  }

  function scaleSlidePreview() {
    var previewEl = $("#rep-slide-preview", body);
    if (!previewEl) return;
    var frame = previewEl.querySelector(".rep-slide-frame");
    var mock = frame && frame.querySelector(".rep-sld-mock");
    if (!frame || !mock) return;
    var w = frame.getBoundingClientRect().width;
    if (!w) return;
    var scale = w / 960;
    mock.style.transformOrigin = "top left";
    mock.style.transform = "scale(" + scale + ")";
    frame.style.height = Math.ceil(540 * scale) + "px";
  }

  function renderSlidePreview() {
    var wrap = $("#rep-slide-preview", body);
    if (!wrap) return;
    function ragBg(v) { return RAG_BG[v] || "transparent"; }
    var techBullets = cfg.techBullets || [];
    var descText = proj.description || proj.scope || sc.objectives || "";
    var descLines = descText.split(/[\n•]+/).map(function(s) { return s.trim(); }).filter(Boolean).slice(0, 6);
    /* Match the image source used by the PPTX exporter: selected images take
       priority, then the cover image and regular project images. */
    var selectedPhotos = (proj.exportSelections && proj.exportSelections.photos) || [];
    var listedPhotos = Array.isArray(proj.images)
      ? proj.images.map(function(img) { return typeof img === "string" ? img : (img.url || img.fileUrl || ""); }).filter(Boolean)
      : [];
    var photoUrls = (selectedPhotos.length ? selectedPhotos : (proj.coverImage ? [proj.coverImage] : []).concat(listedPhotos))
      .filter(Boolean).filter(function(url, index, all) { return all.indexOf(url) === index; }).slice(0, 4);
    var roster = (window.AEWTTR.db.projectPeople && window.AEWTTR.db.projectPeople[proj.id]) || [];
    var pocEntry = proj.pm ? roster.find(function(person) { return person.id === proj.pm; }) : null;
    if (!pocEntry) pocEntry = roster.find(function(person) { return /pm|product manager|project manager/i.test(person.role || ""); }) || roster[0] || null;
    var pocMember = pocEntry && pocEntry.memberId
      ? (window.AEWTTR.db.members || []).find(function(member) { return member.id === pocEntry.memberId; })
      : null;
    var poc = (pocEntry && pocEntry.label) || (pocMember && pocMember.name) || proj.poc || proj.lead || proj.pm || "";
    var pocPhone = (pocMember && (pocMember.phone || pocMember.workPhone)) || (pocEntry && pocEntry.phone) || proj.pocPhone || proj.phone || "";
    var pocEmail = (pocEntry && pocEntry.email) || (pocMember && pocMember.email) || proj.pocEmail || proj.email || "";
    var asOf = new Date().toLocaleDateString("en-US", { month: "numeric", day: "numeric", year: "numeric" });
    var ganttHtml = buildGanttHtml();
    wrap.innerHTML =
      '<div class="rep-slide-frame"><div class="rep-sld-mock">' +
      '<div class="rep-sld-cui-top">CONTROLLED UNCLASSIFIED INFORMATION</div>' +
      '<div class="rep-sld-title-area">' +
        '<img class="rep-sld-seal" src="assets/images/ttsd-seal-template.png" alt="TTSD seal">' +
        '<div class="rep-sld-title-text">' + escapeHtml(proj.name || proj.id || "PROJECT") + '</div>' +
        '<div class="rep-sld-poc-col">' +
          (poc ? '<div class="rep-sld-poc-line"><strong>POC:</strong> ' + escapeHtml(poc) + '</div>' : "") +
          (pocPhone ? '<div class="rep-sld-poc-line">Ph: ' + escapeHtml(pocPhone) + '</div>' : "") +
          (pocEmail ? '<div class="rep-sld-poc-line">' + escapeHtml(pocEmail) + '</div>' : "") +
        '</div>' +
        '<img class="rep-sld-seal" src="assets/images/aewttr-seal-template.png" alt="AEWTTR seal">' +
      '</div>' +
      '<div class="rep-sld-quads">' +
        '<div class="rep-sld-quad rep-sld-quad--tl">' +
          (photoUrls.length
            ? '<div class="rep-sld-photo-grid rep-sld-photo-grid--' + Math.min(photoUrls.length, 4) + '">' + photoUrls.map(function(u) { return '<img src="' + escapeHtml(u) + '" class="rep-sld-photo" alt="">'; }).join("") + '</div>'
            : '<div class="rep-sld-photo-ph"><i class="bx bx-image-alt"></i><br>Photos</div>'
          ) +
        '</div>' +
        '<div class="rep-sld-vdiv"></div>' +
        '<div class="rep-sld-quad rep-sld-quad--tr">' +
          '<div class="rep-sld-qhead">Technical Status</div>' +
          '<ul class="rep-sld-blist">' +
            (techBullets.length
              ? techBullets.map(function(b) { return '<li>' + escapeHtml(b) + '</li>'; }).join("")
              : '<li class="rep-sld-hint">Add bullets below</li>'
            ) +
          '</ul>' +
          '<table class="rep-sld-rtbl"><thead><tr><th>Risk Category</th><th>Status</th><th>Comments</th></tr></thead><tbody>' +
            (sc.riskRows || []).map(function(r) {
              return '<tr><td>' + escapeHtml(r.cat) + '</td>' +
                '<td class="rep-sld-rtbl-rag" style="background:' + ragBg(r.rag) + '">&nbsp;</td>' +
                '<td>' + escapeHtml((r.notes || "").slice(0, 55)) + '</td></tr>';
            }).join("") +
          '</tbody></table>' +
        '</div>' +
        '<div class="rep-sld-hdiv"></div>' +
        '<div class="rep-sld-quad rep-sld-quad--bl">' +
          '<div class="rep-sld-qhead">Project Description</div>' +
          '<ul class="rep-sld-blist">' +
            (descLines.length
              ? descLines.map(function(l) { return '<li>' + escapeHtml(l) + '</li>'; }).join("")
              : '<li class="rep-sld-hint">Add description in Project Settings</li>'
            ) +
          '</ul>' +
        '</div>' +
        '<div class="rep-sld-quad rep-sld-quad--br">' +
          '<div class="rep-sld-qhead">Milestones</div>' +
          ganttHtml +
        '</div>' +
      '</div>' +
      '<div class="rep-sld-footer-line"></div>' +
      '<div class="rep-sld-footer">' +
        '<img class="rep-sld-footer-navair" src="assets/images/navair-wordmark-template.png" alt="NAVAIR">' +
        '<div class="rep-sld-footer-cui">CONTROLLED UNCLASSIFIED INFORMATION</div>' +
        '<div class="rep-sld-footer-asof">As of ' + asOf + '</div>' +
      '</div>' +
      '</div></div>';
    scaleSlidePreview();
  }

  function renderTechBullets() {
    var wrap = $("#rep-tech-bullets-wrap", body);
    if (!wrap) return;
    if (!cfg.techBullets) cfg.techBullets = [];
    var bullets = cfg.techBullets;
    wrap.innerHTML =
      '<div class="rep-tb-list">' +
      bullets.map(function(b, i) {
        return '<div class="rep-tb-item">' +
          '<span class="rep-tb-dot">•</span>' +
          '<input type="text" class="input-aewttr rep-tb-inp" data-bidx="' + i + '" value="' + escapeHtml(b) + '" placeholder="Bullet text…">' +
          '<button type="button" class="rep-tb-del" data-bidx="' + i + '" title="Remove"><i class="bx bx-x"></i></button>' +
          '</div>';
      }).join("") +
      '<button type="button" class="rep-tb-add" id="rep-tb-add"><i class="bx bx-plus"></i> Add bullet</button>' +
      '</div>';
    $all(".rep-tb-inp", wrap).forEach(function(inp) {
      var dt = null;
      inp.addEventListener("input", function() {
        var idx = parseInt(inp.dataset.bidx, 10);
        cfg.techBullets[idx] = inp.value;
        clearTimeout(dt);
        dt = setTimeout(function() { saveConfig(); renderSlidePreview(); }, 25);
      });
    });
    $all(".rep-tb-del", wrap).forEach(function(btn) {
      btn.addEventListener("click", function() {
        var idx = parseInt(btn.dataset.bidx, 10);
        cfg.techBullets.splice(idx, 1);
        saveConfig();
        renderTechBullets();
        renderSlidePreview();
      });
    });
    var addBtn = $("#rep-tb-add", wrap);
    if (addBtn) {
      addBtn.addEventListener("click", function() {
        if (!cfg.techBullets) cfg.techBullets = [];
        cfg.techBullets.push("");
        saveConfig();
        renderTechBullets();
        renderSlidePreview();
        var inputs = $all(".rep-tb-inp", wrap);
        if (inputs.length) inputs[inputs.length - 1].focus();
      });
    }
  }

  // RAG palette shared between the editor pills and the preview table
  var RAG_BG = { Green: "#22c55e", Yellow: "#f59e0b", Red: "#ef4444" };
  function ragPillStyle(rag) { var bg = RAG_BG[rag]; return bg ? ' style="background:' + bg + ';color:#fff;"' : ""; }

  // Precompute values needed for static BL/TL cards
  var _descText = proj.description || proj.scope || sc.objectives || "";
  var _descLines = _descText.split(/[\n•]+/).map(function(s) { return s.trim(); }).filter(Boolean).slice(0, 6);
  var _selPhotos  = (proj.exportSelections && proj.exportSelections.photos) || [];
  var _listedPh   = Array.isArray(proj.images) ? proj.images.map(function(img) { return typeof img === "string" ? img : (img.url || img.fileUrl || ""); }).filter(Boolean) : [];
  var _photoUrls  = (_selPhotos.length ? _selPhotos : (proj.coverImage ? [proj.coverImage] : []).concat(_listedPh)).filter(Boolean).filter(function(u, i, a) { return a.indexOf(u) === i; }).slice(0, 4);

  // Preview visibility state
  if (cfg.previewVisible === undefined) cfg.previewVisible = true;

  // Risk register for per-risk selection
  var _allRisks = (typeof projectRisks === "function" ? projectRisks(proj.id) : []).filter(function(r) { return r.status !== "Closed"; });
  if (!cfg.selectedRisks) cfg.selectedRisks = [];

  body.innerHTML = `
    <div class="rep-shell">

      <!-- Action bar -->
      <div class="rep-bar">
        <div class="rep-bar-right">
          <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="rep-toggle-preview">
            <i class="bx ${cfg.previewVisible ? "bx-hide" : "bx-show"}"></i> ${cfg.previewVisible ? "Hide" : "Show"} Preview
          </button>
          <button type="button" class="btn-aewttr-outline btn-aewttr-sm" id="rep-export-project-slide"><i class="bx bx-file"></i> Export project slide</button>
          <button type="button" class="btn-aewttr btn-aewttr-sm" id="rep-export-full-deck"><i class="bx bxs-slideshow"></i> Export full deck</button>
        </div>
      </div>

      <!-- Two-column layout -->
      <div class="rep-dual ${cfg.previewVisible ? "" : "rep-dual--no-preview"}">

        <!-- LEFT: sticky slide preview -->
        <div class="rep-dual-left" style="${cfg.previewVisible ? "" : "display:none"}">
          <div class="rep-preview-eyebrow">
            <i class="bx bx-slideshow"></i> Slide Preview
            <span class="rep-preview-eyebrow-hint">— updates as you edit</span>
          </div>
          <div class="rep-preview-wrap">
            <div id="rep-slide-preview"></div>
          </div>
        </div>

        <!-- RIGHT: quadrant editing cards -->
        <div class="rep-dual-right">

          <!-- ↖ TOP LEFT — Photos -->
          <div class="rep-qcard">
            <div class="rep-qcard-hd rep-qcard-hd--tl">
              <span class="rep-qcard-pos">↖ Top Left</span>
              <span class="rep-qcard-name">Photos</span>
            </div>
            <div class="rep-qcard-bd">
              <div class="rep-photo-info">
                <i class="bx bx-image-alt"></i>
                <div>
                  <div class="rep-photo-label">Photos are pulled from the <strong>Photos tab</strong></div>
                  <div class="rep-photo-sub">Up to 4 photos appear in this quadrant. Add or manage them in the Photos tab.</div>
                </div>
              </div>
              ${_photoUrls.length
                ? '<div class="rep-photo-thumbs">' + _photoUrls.map(function(u) { return '<img src="' + escapeHtml(u) + '" class="rep-photo-thumb" alt="">'; }).join("") + '</div>'
                : '<div class="rep-photo-empty">No photos added yet — go to the Photos tab to upload.</div>'
              }
            </div>
          </div>

          <!-- ↗ TOP RIGHT — Technical Status -->
          <div class="rep-qcard">
            <div class="rep-qcard-hd rep-qcard-hd--tr">
              <span class="rep-qcard-pos">↗ Top Right</span>
              <span class="rep-qcard-name">Technical Status</span>
            </div>
            <div class="rep-qcard-bd">
              <p class="rep-qcard-sub">Status Bullets</p>
              <p class="rep-qcard-hint">Bullet points that appear in the slide's technical status section</p>
              <div id="rep-tech-bullets-wrap"></div>
              <div class="rep-qcard-rule"></div>
              <p class="rep-qcard-sub">Risk Summary</p>
              <p class="rep-qcard-hint">Click the badge to cycle Green → Yellow → Red</p>
              <div class="rep-risk-rows">
                ${sc.riskRows.map(function(row, i) {
                  return '<div class="rep-risk-row">' +
                    '<div class="rep-risk-cat">' + escapeHtml(row.cat) + '</div>' +
                    '<button type="button" class="rep-risk-pill rep-risk-pill--' + (row.rag || "none") + ' rep-risk-rag-btn"' + ragPillStyle(row.rag) + ' data-risk-idx="' + i + '">' + (row.rag || "—") + '</button>' +
                    '<input type="text" class="input-aewttr rep-risk-note rep-risk-notes-inp" data-risk-idx="' + i + '" placeholder="Comment…" value="' + escapeHtml(row.notes || "") + '">' +
                  '</div>';
                }).join("")}
              </div>
              ${_allRisks.length ? `
              <div class="rep-qcard-rule"></div>
              <p class="rep-qcard-sub">Risks Slide — Select Items</p>
              <p class="rep-qcard-hint">Choose which risks from the risk register to include in the risks export slide.</p>
              <div class="rep-risk-checklist" id="rep-risk-checklist">
                ${_allRisks.map(function(r) {
                  const checked = cfg.selectedRisks.indexOf(r.id) >= 0;
                  const ragColor = r.rating === "Red" ? "#ef4444" : r.rating === "Yellow" || r.rating === "Amber" ? "#f59e0b" : r.rating === "Green" ? "#22c55e" : "";
                  return '<label class="rep-risk-check-row' + (checked ? " rep-risk-check-row--on" : "") + '">' +
                    '<input type="checkbox" class="rep-risk-check-cb" data-risk-id="' + escapeHtml(r.id) + '"' + (checked ? " checked" : "") + '>' +
                    (ragColor ? '<span class="rep-risk-check-dot" style="background:' + ragColor + '"></span>' : '<span class="rep-risk-check-dot rep-risk-check-dot--none"></span>') +
                    '<span class="rep-risk-check-name">' + escapeHtml(r.name || r.description || "Unnamed risk") + '</span>' +
                    (r.owner ? '<span class="rep-risk-check-owner">' + escapeHtml(r.owner) + '</span>' : '') +
                  '</label>';
                }).join("")}
              </div>
              ` : `<div class="rep-qcard-rule"></div><p class="rep-qcard-hint" style="color:var(--aewttr-muted);">No open risks in the risk register. Add risks in the Risks tab to select them for export.</p>`}
            </div>
          </div>

          <!-- ↙ BOTTOM LEFT — Project Description -->
          <div class="rep-qcard">
            <div class="rep-qcard-hd rep-qcard-hd--bl">
              <span class="rep-qcard-pos">↙ Bottom Left</span>
              <span class="rep-qcard-name">Project Description</span>
            </div>
            <div class="rep-qcard-bd">
              <p class="rep-desc-from">Pulled from <strong>Project Settings → Description</strong>. Edit it there to update this quadrant on the slide.</p>
              <div class="rep-desc-box${_descLines.length ? "" : " rep-desc-empty"}">
                ${_descLines.length
                  ? '<ul style="margin:0;padding-left:16px">' + _descLines.map(function(l) { return '<li>' + escapeHtml(l) + '</li>'; }).join("") + '</ul>'
                  : 'No description set — add one in Project Settings.'
                }
              </div>
            </div>
          </div>

          <!-- ↘ BOTTOM RIGHT — Milestones -->
          <div class="rep-qcard">
            <div class="rep-qcard-hd rep-qcard-hd--br">
              <span class="rep-qcard-pos">↘ Bottom Right</span>
              <span class="rep-qcard-name">Milestones</span>
            </div>
            <div class="rep-qcard-bd rep-qcard-bd--ms">
              <div style="padding:0 14px">
                <p class="rep-qcard-sub">Task Milestone Dates</p>
                <p class="rep-qcard-hint">Check <strong>In Report</strong> to include a task on the Gantt. Enter dates for each event type.</p>
              </div>
              <div id="rep-milestone-table-wrap"></div>
              <div class="rep-qcard-rule" style="margin:14px 14px"></div>
              <div style="padding:0 14px">
                <p class="rep-qcard-sub">Milestone Timeline</p>
              </div>
              <div id="rep-timeline-wrap" style="padding:0 14px 12px"></div>
            </div>
          </div>

        </div><!-- /rep-dual-right -->
      </div><!-- /rep-dual -->
    </div><!-- /rep-shell -->
  `;

  renderMilestoneTable();
  renderTimeline();
  renderSlidePreview();
  renderTechBullets();

  const RAG_CYCLE = ["", "Green", "Yellow", "Red"];
  $all(".rep-risk-rag-btn", body).forEach(function(btn) {
    btn.addEventListener("click", function() {
      const idx = parseInt(btn.dataset.riskIdx, 10);
      const row = sc.riskRows[idx];
      if (!row) return;
      const cur = RAG_CYCLE.indexOf(row.rag || "");
      row.rag = RAG_CYCLE[(cur + 1) % RAG_CYCLE.length];
      btn.className = "rep-risk-pill rep-risk-pill--" + (row.rag || "none") + " rep-risk-rag-btn";
      btn.textContent = row.rag || "—";
      var _pillBg = RAG_BG[row.rag] || "";
      btn.style.background = _pillBg;
      btn.style.color = _pillBg ? "#fff" : "";
      saveConfig();
      renderSlidePreview();
    });
  });

  $all(".rep-risk-notes-inp", body).forEach(function(inp) {
    var t = null;
    inp.addEventListener("input", function() {
      const idx = parseInt(inp.dataset.riskIdx, 10);
      if (sc.riskRows[idx]) sc.riskRows[idx].notes = inp.value;
      clearTimeout(t);
      t = setTimeout(function() { saveConfig(); renderSlidePreview(); }, 25);
    });
  });

  // Preview toggle
  const togglePreviewBtn = $("#rep-toggle-preview", body);
  if (togglePreviewBtn) {
    togglePreviewBtn.addEventListener("click", function() {
      cfg.previewVisible = !cfg.previewVisible;
      saveConfig();
      drawProjectReporting(body, proj);
    });
  }

  // Per-risk selection checkboxes
  $all(".rep-risk-check-cb", body).forEach(function(cb) {
    cb.addEventListener("change", function() {
      var id = cb.dataset.riskId;
      var idx = cfg.selectedRisks.indexOf(id);
      if (cb.checked && idx < 0) cfg.selectedRisks.push(id);
      else if (!cb.checked && idx >= 0) cfg.selectedRisks.splice(idx, 1);
      var row = cb.closest(".rep-risk-check-row");
      if (row) row.classList.toggle("rep-risk-check-row--on", cb.checked);
      saveConfig();
    });
  });

  async function runProjectPptxExport(mode, button) {
    const exp = window.AEWTTR && window.AEWTTR.ProjectPptxExport;
    if (!exp) { toast("Export module not loaded.", "error"); return; }
    const isFullDeck = mode === "fullDeck";
    const otherButton = $(isFullDeck ? "#rep-export-project-slide" : "#rep-export-full-deck", body);
    button.disabled = true;
    if (otherButton) otherButton.disabled = true;
    button.innerHTML = `<i class="bx bx-loader-alt bx-spin"></i> Building…`;
    const openTarget = reserveStatusExportOpen(`${proj.id || "project"}-${isFullDeck ? "status-briefing" : "project-update"}.pptx`);
    try {
      const result = await exp.exportProjectStatusPptxToSharePoint(proj, isFullDeck ? {
        exportMode: "fullDeck",
        groupName: proj.name || proj.id,
        groupType: "project",
        includePhotos: true,
        includeRisks: Array.isArray(cfg.selectedRisks) && cfg.selectedRisks.length > 0,
        selectedRiskIds: cfg.selectedRisks || [],
        includeTasks: !!(cfg.slideTypes && cfg.slideTypes.tasks),
        taskFilter: "open",
        riskFilter: "open",
        popup: openTarget
      } : { exportMode: "projectSlide", popup: openTarget });
      if (result.mode === "sharepoint" && result.fileUrl) {
        toastStatusPptxSaved(result.fileUrl, result.fileName);
      } else if (result.uploadError) {
        const detail = (result.uploadError && (result.uploadError.friendly || result.uploadError.message)) || "upload failed";
        toast(`SharePoint upload failed (${detail}) — downloaded a local copy instead.`, "warn");
      } else {
        toast("PowerPoint downloaded.", "success");
      }
    } catch (e) {
      closeStatusExportOpen(openTarget);
      toast((e && e.message) || "Export failed.", "error");
    }
    finally {
      button.disabled = false;
      if (otherButton) otherButton.disabled = false;
      button.innerHTML = isFullDeck
        ? `<i class="bx bxs-slideshow"></i> Export full deck`
        : `<i class="bx bx-file"></i> Export project slide`;
    }
  }

  const projectSlideButton = $("#rep-export-project-slide", body);
  if (projectSlideButton) projectSlideButton.addEventListener("click", () => runProjectPptxExport("projectSlide", projectSlideButton));
  const fullDeckButton = $("#rep-export-full-deck", body);
  if (fullDeckButton) fullDeckButton.addEventListener("click", () => runProjectPptxExport("fullDeck", fullDeckButton));
}
