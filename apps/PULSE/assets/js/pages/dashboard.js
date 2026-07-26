/* ---------- Hierarchical "My Work" tree ---------- */
function myWorkHierarchy() {
  const db = window.AEWTTR.db || {};
  const user = db.user || {};
  const member = typeof currentUserMember === "function" ? currentUserMember() : null;

  function isMine(assignee) {
    if (!assignee) return false;
    if (assignee === user.name) return true;
    return !!(member && typeof memberMatchesAssignee === "function" && memberMatchesAssignee(member, assignee));
  }

  const result = [];
  (db.projects || []).forEach((project) => {
    const allItems = (db.ganttTasks && db.ganttTasks[project.id]) || [];
    const tasks = allItems.filter((t) => !(typeof isTrackerDivider === "function" ? isTrackerDivider(t) : t.itemType === "divider"));

    function relevantTaskEntry(task) {
      if (task.status === "Done") return null;
      const directlyMine = isMine(task.assignee);
      const mySubtasks = (task.subtasks || []).filter((s) => s.status !== "Done" && isMine(s.assignee));
      if (!directlyMine && !mySubtasks.length) return null;
      return { task, directlyMine, mySubtasks };
    }

    const entries = tasks.map(relevantTaskEntry).filter(Boolean);
    if (entries.length) result.push({ project, entries });
  });
  return result;
}

/* ---------- HTML helpers ---------- */
function dashHealthDot(health) {
  const map = { "On Track": "green", "At Risk": "amber", "Off Track": "red", "Blocked": "red" };
  const c = map[health] || "green";
  return `<span class="mwt-health-dot mwt-health-dot--${c}" title="${escapeHtml(health || "On Track")}"></span>`;
}

function dashStatusChip(status) {
  if (!status || status === "Not Started") return "";
  const map = { "In Progress": "blue", "Done": "green", "Blocked": "red", "On Hold": "amber" };
  const c = map[status] || "muted";
  return `<span class="mwt-status-chip mwt-status-chip--${c}">${escapeHtml(status)}</span>`;
}

/* ---------- Main dashboard renderer ---------- */
PAGE_RENDERERS.dashboard = function () {
  const db = window.AEWTTR.db;
  const alerts = typeof getAppAlertCounts === "function" ? getAppAlertCounts() : { docreview: 0, travel: 0, overview: 0 };
  const hierarchy = myWorkHierarchy();
  const totalAssignedTasks = hierarchy.reduce((total, pr) => total + pr.entries.reduce((n, e) => n + 1 + e.mySubtasks.length, 0), 0);
  setTopbar("Dashboard", `Welcome back, ${db.user.name.split(" ")[0]}.`, "");

  if (!window.AEWTTR.state.dashExpanded) window.AEWTTR.state.dashExpanded = {};

  const canAdmin = typeof canCurrentUserAccessAdmin === "function" ? canCurrentUserAccessAdmin() : false;
  const canTravelApprove = typeof canAccessTravelApprovals === "function" ? canAccessTravelApprovals() : false;

  /* App launcher definitions — main route + sub-tools */
  const apps = [
    {
      route: "overview", title: "Overview", icon: "bx-bar-chart-alt-2",
      desc: "Portfolio-wide command view",
      tools: [
        { label: "Portfolio health", route: "overview" },
        { label: "Workload view", route: "overview" },
        { label: "Approvals", route: "overview" }
      ]
    },
    {
      route: "projects", title: "Projects", icon: "bx-folder-open",
      desc: "All projects and their status",
      tools: [
        { label: "All projects", route: "projects" },
        { label: "Tracker", route: "projects" },
        { label: "Risk registers", route: "projects" }
      ]
    },
    {
      route: "weekly", title: "Weekly Meeting", icon: "bx-conversation",
      desc: "Team roundtables and notes",
      tools: [
        { label: "Run meeting", route: "weekly" },
        { label: "Meeting history", route: "weekly" },
        { label: "Check-ins", route: "weekly" }
      ]
    },
    {
      route: "travel", title: "Travel", icon: "bx-world",
      alertKey: "travel",
      desc: "Submit and track travel requests",
      tools: [
        { label: "My travel", route: "travel/mine" },
        { label: "Submit request", route: "travel/submit" },
        { label: "All travel", route: "travel/all" },
        ...(canTravelApprove ? [{ label: "Approvals", route: "travel/approve" }] : [])
      ]
    },
    {
      route: "docreview", title: "Document Review", icon: "bx-file-blank",
      alertKey: "docreview",
      desc: "Review, concurrence, and signatures",
      tools: [
        { label: "Review queue", route: "docreview" },
        { label: "Pending signatures", route: "docreview" },
        { label: "All documents", route: "docreview" }
      ]
    },
    ...(canAdmin ? [{
      route: "admin", title: "Admin", icon: "bx-shield-quarter",
      desc: "Setup, users, and system logs",
      tools: [
        { label: "SharePoint setup", route: "admin" },
        { label: "Users", route: "admin" },
        { label: "Activity log", route: "logs" }
      ]
    }] : [])
  ];

  /* Hierarchy tree HTML */
  function taskEntryHtml(entry, projectId) {
    const { task, directlyMine, mySubtasks } = entry;
    const due = task.end ? `Due ${fmtDate(task.end)}` : "";
    const taskHtml = `
      <button type="button" class="mwt-task${directlyMine ? " mwt-task--mine" : ""}" data-route="projects/${projectId}/tracker">
        ${dashHealthDot(task.health)}
        <div class="mwt-task-body">
          <span class="mwt-task-title">${escapeHtml(task.title)}</span>
          <span class="mwt-task-meta">${[due, task.assignee && !directlyMine ? escapeHtml(task.assignee) : ""].filter(Boolean).join(" · ")}</span>
        </div>
        ${dashStatusChip(task.status)}
      </button>`;

    const subsHtml = mySubtasks.map((sub) => `
      <button type="button" class="mwt-subtask" data-route="projects/${projectId}/tracker">
        <i class="bx bx-subdirectory-right mwt-sub-icon"></i>
        <div class="mwt-task-body">
          <span class="mwt-task-title">${escapeHtml(sub.title)}</span>
          <span class="mwt-task-meta">${sub.end ? `Due ${fmtDate(sub.end)}` : ""}</span>
        </div>
        ${dashStatusChip(sub.status)}
      </button>`).join("");

    return taskHtml + (mySubtasks.length && !directlyMine ? subsHtml : "");
  }

  function projectBlockHtml({ project, entries }) {
    const pKey = `proj:${project.id}`;
    const taskCount = entries.reduce((n, e) => n + 1 + e.mySubtasks.length, 0);
    const stored = window.AEWTTR.state.dashExpanded[pKey];
    const expanded = typeof stored === "boolean" ? stored : (totalAssignedTasks <= 8 && hierarchy.length <= 2);
    return `
      <div class="mwt-project">
        <button type="button" class="mwt-project-toggle" data-dash-expand="${pKey}" data-dash-expanded="${expanded}">
          <i class="bx bx-chevron-${expanded ? "down" : "right"} mwt-chevron"></i>
          <i class="bx bx-folder-open mwt-proj-icon"></i>
          <span class="mwt-proj-name">${escapeHtml(project.name || project.id)}</span>
          <em class="mwt-count">${taskCount}</em>
        </button>
        <div class="mwt-project-body${expanded ? "" : " mwt-hidden"}">
          ${entries.map((e) => taskEntryHtml(e, project.id)).join("")}
        </div>
      </div>`;
  }

  const myWorkHtml = hierarchy.length
    ? hierarchy.map(projectBlockHtml).join("")
    : `<div class="empty-state">No open tasks assigned to you.</div>`;

  /* App launcher HTML */
  const appsHtml = apps.map((app) => {
    const count = app.alertKey ? (alerts[app.alertKey] || 0) : 0;
    return `
      <div class="dash-app${count ? " dash-app--alert" : ""}">
        <button type="button" class="dash-app-main" data-route="${app.route}">
          <span class="dash-app-icon-wrap">
            <i class="bx ${app.icon}"></i>
          </span>
          <span class="dash-app-label">
            <strong>${app.title}</strong>
            <span>${app.desc}</span>
          </span>
          ${count ? `<span class="dash-app-badge">${count}</span>` : ""}
        </button>
        <div class="dash-app-tools">
          ${app.tools.map((t) => `<button type="button" class="dash-app-tool" data-route="${t.route}">${t.label}</button>`).join("")}
        </div>
      </div>`;
  }).join("");


  /* Quick stats */
  const activeProjects = (db.projects || []).filter((p) => p.lifecycleStatus === "Active" || (!p.lifecycleStatus && p.status !== "Complete" && p.status !== "Cancelled")).length;
  const allMyTasks = hierarchy.reduce((n, pr) => n + pr.entries.length, 0);
  const overdueCount = hierarchy.reduce((n, pr) => n + pr.entries.filter((e) => {
    const due = e.task.end || e.task.dueDate;
    return due && new Date(due) < new Date();
  }).length, 0);

  $("#page-content").innerHTML = `
    <div class="dashboard-v2">
      <div class="dashboard-v2-main">
        <div class="dash-stats-bar">
          <div class="dash-stat">
            <span class="dash-stat-num">${activeProjects}</span>
            <span class="dash-stat-label">Active Projects</span>
          </div>
          <div class="dash-stat">
            <span class="dash-stat-num">${totalAssignedTasks}</span>
            <span class="dash-stat-label">Open Tasks</span>
          </div>
          <div class="dash-stat">
            <span class="dash-stat-num" style="${overdueCount ? "color:var(--aewttr-red)" : ""}">${overdueCount}</span>
            <span class="dash-stat-label">Overdue</span>
          </div>
        </div>
        <section class="dash-apps-section">
          <div class="dash-section-label">Quick Access</div>
          <div class="dash-apps-grid">${appsHtml}</div>
        </section>
      </div>
      <aside class="dashboard-v2-side">
        <section class="aewttr-card aewttr-card-pad dashboard-my-panel">
          <div class="dashboard-my-head">
            <div class="side-panel-title">My Work</div>
            ${totalAssignedTasks ? `<span class="dashboard-my-total">${totalAssignedTasks}</span>` : ""}
          </div>
          <p class="dashboard-my-sub">Your assigned tasks across all projects.</p>
          <div class="mwt-tree${totalAssignedTasks > 8 ? " mwt-tree--dense" : ""}">${myWorkHtml}</div>
        </section>
        <section class="aewttr-card aewttr-card-pad dashboard-my-panel">
          <div class="side-panel-title">Weekly Meeting</div>
          <p class="dashboard-my-sub">Run the team roundtable, review tasks, capture notes.</p>
          <button type="button" class="btn-aewttr-outline btn-aewttr-sm" data-route="weekly" style="width:100%;margin-top:8px;"${tip("Open weekly meeting")}>Open Weekly Meeting</button>
        </section>
      </aside>
    </div>
  `;

  /* Wire click events */
  $all("[data-route]", $("#page-content")).forEach((b) => b.addEventListener("click", () => navigate(b.dataset.route)));

  /* Wire collapse toggles */
  $all("[data-dash-expand]", $("#page-content")).forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const key = btn.dataset.dashExpand;
      window.AEWTTR.state.dashExpanded[key] = btn.dataset.dashExpanded !== "true";
      PAGE_RENDERERS.dashboard();
    });
  });
};
