PAGE_RENDERERS.overview = function () {
  if (!window.AEWTTR.state.overviewView) window.AEWTTR.state.overviewView = "My";
  if (!window.AEWTTR.state.overviewMyTasksFilter) window.AEWTTR.state.overviewMyTasksFilter = "All";
  if (!window.AEWTTR.state.overviewWorkloadSort) window.AEWTTR.state.overviewWorkloadSort = "active";
  if (!window.AEWTTR.state.overviewTeamTab) window.AEWTTR.state.overviewTeamTab = "dashboard";
  if (!window.AEWTTR.state.overviewFilters) {
    window.AEWTTR.state.overviewFilters = { search: "", team: "All", status: "All", fundingType: "All", portfolio: "All", endItem: "All", ato: "All" };
  }

  const db = window.AEWTTR.db;
  const canSeeTeamOverview = !!(canManageMeetings && canManageMeetings()) ||
    ["Admin", "Manager"].includes((db.user && db.user.role) || "");
  const views = [
    { key: "My", label: "My Overview" },
    ...(canSeeTeamOverview ? [{ key: "Team", label: "Team Overview" }] : [])
  ];

  /* ---- task predicates ---- */

  function taskIsDone(task) {
    return task.status === "Done" || task.status === "Completed" ||
      task.status === "Closed" || task.done === true;
  }

  function taskIsActive(task) {
    return !taskIsDone(task);
  }

  function taskIsBlocked(task) {
    if (!taskIsActive(task)) return false;
    const s = String(task.status || "").toLowerCase();
    return s === "blocked" || s === "on hold" || s === "waiting" ||
      !!(task.blockedReason && String(task.blockedReason).trim());
  }

  /* ---- project helpers ---- */

  function allProjectTasks(projId) {
    return trackerPlainTasks(db.ganttTasks[projId] || []);
  }

  function projectEndItem(proj) {
    return (proj.configEndItem && proj.configEndItem.trim()) ||
      (proj.portfolios && proj.portfolios[0]) ||
      "Unclassified";
  }

  function isPlaceholderRisk(r) {
    return String(r.title || "").trim().toLowerCase() === "new risk";
  }

  function projectRiskCount(projId) {
    const extra = (db.projectExtra && db.projectExtra[projId]) || {};
    return (extra.risks || []).filter(r => {
      if (isPlaceholderRisk(r)) return false;
      const s = String(r.status || "").toLowerCase();
      return s !== "closed" && s !== "resolved" && s !== "accepted";
    }).length;
  }

  function projectCriticalRiskCount(projId) {
    const extra = (db.projectExtra && db.projectExtra[projId]) || {};
    return (extra.risks || []).filter(r => {
      if (isPlaceholderRisk(r)) return false;
      const s = String(r.status || "").toLowerCase();
      if (s === "closed" || s === "resolved" || s === "accepted") return false;
      return ["high", "critical"].some(k =>
        String(r.likelihood || "").toLowerCase().includes(k) ||
        String(r.impact || "").toLowerCase().includes(k) ||
        String(r.severity || "").toLowerCase().includes(k)
      );
    }).length;
  }

  function parentProjectName(proj) {
    if (!proj.parentProjectId) return "";
    const parent = (db.projects || []).find(p => p.id === proj.parentProjectId);
    return parent ? (parent.name || parent.id) : proj.parentProjectId;
  }

  function projectPeopleList(projId) {
    return (db.projectPeople && db.projectPeople[projId]) || [];
  }

  function projectPMDisplayName(proj) {
    const roster = projectPeopleList(proj.id);
    const pmEntry = roster.find(p =>
      String(p.role || "").toLowerCase().includes("pm") ||
      String(p.role || "").toLowerCase().includes("project manager") ||
      String(p.type || "").toLowerCase() === "pm"
    );
    if (pmEntry) return pmEntry.label || "";
    const entry = typeof resolvedAssignedPerson === "function"
      ? resolvedAssignedPerson(proj, "pm") : null;
    return (entry && entry.label) || "";
  }

  function myRoleOnProject(proj) {
    const user = db.user || {};
    const member = typeof currentUserMember === "function" ? currentUserMember() : null;
    const roster = projectPeopleList(proj.id);

    const rosterEntry = roster.find(p => {
      if (member && p.memberId && p.memberId === member.id) return true;
      if (p.label && member && String(p.label).toLowerCase() === String(member.name || "").toLowerCase()) return true;
      if (p.label && String(p.label).toLowerCase() === String(user.name || "").toLowerCase()) return true;
      return false;
    });
    if (rosterEntry) return rosterEntry.role || rosterEntry.type || "Member";

    const pmEntry = typeof resolvedAssignedPerson === "function"
      ? resolvedAssignedPerson(proj, "pm") : null;
    if (pmEntry && member && (pmEntry.id === member.id ||
      String(pmEntry.label || "").toLowerCase() === String(user.name || "").toLowerCase())) return "PM";

    const engEntry = typeof resolvedAssignedPerson === "function"
      ? resolvedAssignedPerson(proj, "engineer") : null;
    if (engEntry && member && (engEntry.id === member.id ||
      String(engEntry.label || "").toLowerCase() === String(user.name || "").toLowerCase())) return "Engineer";

    return null;
  }

  function userIsOnProject(proj) {
    return !!(myRoleOnProject(proj));
  }

  function derivedProjectStatus(proj) {
    const tasks = allProjectTasks(proj.id);
    const active = tasks.filter(taskIsActive);
    if (!active.length) {
      if (tasks.length && tasks.every(taskIsDone)) return "Completed";
      return "No Active Work";
    }
    if (active.some(t => taskIsBlocked(t))) return "Needs Attention";
    return "Active";
  }

  function derivedStatusPill(status) {
    const cls = {
      "Active": "ov-status-active",
      "Needs Attention": "ov-status-attention",
      "No Active Work": "ov-status-inactive",
      "Completed": "ov-status-done"
    }[status] || "ov-status-inactive";
    return `<span class="ov-derived-status ${cls}">${escapeHtml(status)}</span>`;
  }

  function projectNextDue(proj) {
    const tasks = allProjectTasks(proj.id).filter(t => taskIsActive(t) && t.end);
    if (!tasks.length) return proj.dueDate || "";
    return tasks.sort((a, b) => String(a.end).localeCompare(String(b.end)))[0].end;
  }

  function asOfLabel() {
    return new Date().toLocaleString("en-US", {
      month: "short", day: "numeric", year: "numeric",
      hour: "numeric", minute: "2-digit"
    });
  }

  /* ---- MY OVERVIEW ---- */

  function getMyData() {
    const user = db.user || {};
    const member = typeof currentUserMember === "function" ? currentUserMember() : null;
    const myEmail = String(user.email || "").toLowerCase();
    const myName  = String(user.name  || "").toLowerCase();

    const allMyTasks = [];
    (db.projects || []).forEach(proj => {
      allProjectTasks(proj.id).forEach(task => {
        const isMe = (member && typeof memberMatchesAssignee === "function" && memberMatchesAssignee(member, task.assignee)) ||
          (user.name && task.assignee === user.name);
        if (isMe) allMyTasks.push({ ...task, projectId: proj.id, projectName: proj.name || "Untitled project", proj });
      });
    });

    const myProjects = (db.projects || []).filter(userIsOnProject);

    const docs = typeof getAllDocReviewRecords === "function" ? getAllDocReviewRecords() : [];

    const isMe = r =>
      (r.email && String(r.email).toLowerCase() === myEmail) ||
      (r.name  && String(r.name ).toLowerCase() === myName);

    const docsNeedingMe = docs.filter(doc => {
      if (doc.isArchived) return false;
      const col = doc._column || "";
      if (["Review Complete", "Signed", "Archived"].includes(col)) return false;
      return Array.isArray(doc.reviewers) && doc.reviewers.some(isMe);
    });

    const thisYear = new Date().getFullYear().toString();
    const yearStart = thisYear + "-01-01";

    const docsReviewedThisYear = docs.filter(doc =>
      Array.isArray(doc.reviewers) && doc.reviewers.some(r =>
        isMe(r) && (String(r.signedAt || r.decidedAt || "").slice(0, 4) === thisYear ||
          String(doc.updatedAt || doc.createdAt || "").slice(0, 4) === thisYear)
      ) && ["Review Complete", "Signed", "Archived", "Concurrence"].some(k =>
        String(doc._column || "").includes(k)
      )
    );

    const docsSignedByMe = docs.filter(doc =>
      Array.isArray(doc.reviewers) && doc.reviewers.some(r =>
        isMe(r) && r.isSigner && r.signedAt && String(r.signedAt).slice(0, 4) === thisYear
      )
    );

    const tasksCompletedThisYear = allMyTasks.filter(t =>
      taskIsDone(t) && String(t.updatedAt || t.completedAt || t.end || "").slice(0, 4) === thisYear
    );

    const sessions = (db.weeklyMeeting && db.weeklyMeeting.sessions) || [];
    const today = new Date().toISOString().slice(0, 10);
    const meetingsAttended = sessions.filter(s => {
      if (!s.attendance || !s.date) return false;
      if (s.date > today) return false;
      if (s.date.slice(0, 4) !== thisYear) return false;
      return Object.entries(s.attendance).some(([id, status]) => {
        if (status !== "In") return false;
        const m = (db.members || []).find(mb => mb.id === id);
        return m && (String(m.name || "").toLowerCase() === myName ||
          String(m.email || "").toLowerCase() === myEmail);
      });
    });

    const myTravel = (db.travelRequests || []).filter(r =>
      (r.status === "Pending" || r.status === "Submitted") && (
        String(r.requester || "").toLowerCase() === myName ||
        String(r.requesterEmail || "").toLowerCase() === myEmail
      )
    );
    const travelThisYear = (db.travelRequests || []).filter(r =>
      String(r.requester || "").toLowerCase() === myName ||
      String(r.requesterEmail || "").toLowerCase() === myEmail
    ).filter(r => String(r.createdAt || r.startDate || "").slice(0, 4) === thisYear);

    const upcomingMeetings = sessions.filter(s => (s.date || "") >= today && s.status !== "complete").length;

    return {
      allMyTasks, myProjects, docsNeedingMe, myTravel, upcomingMeetings,
      thisYear, docsReviewedThisYear, docsSignedByMe, tasksCompletedThisYear,
      meetingsAttended, travelThisYear, myEmail, myName
    };
  }

  function renderMyOverview() {
    const user = db.user || {};
    const {
      allMyTasks, myProjects, docsNeedingMe, myTravel, upcomingMeetings,
      thisYear, docsReviewedThisYear, docsSignedByMe, tasksCompletedThisYear,
      meetingsAttended, travelThisYear
    } = getMyData();

    const activeTasks  = allMyTasks.filter(taskIsActive);
    const blockedTasks = allMyTasks.filter(taskIsBlocked);
    const hour         = new Date().getHours();
    const greeting     = hour < 12 ? "Good morning" : hour < 18 ? "Good afternoon" : "Good evening";
    const firstName    = (user.name || "there").split(" ")[0];

    const taskFilter = window.AEWTTR.state.overviewMyTasksFilter || "All";

    const sortMyTasks = tasks => [...tasks].sort((a, b) => {
      if (taskIsDone(a) !== taskIsDone(b)) return taskIsDone(a) ? 1 : -1;
      if (!!a.end !== !!b.end) return a.end ? -1 : 1;
      return String(a.end || "").localeCompare(String(b.end || ""));
    });

    const filteredMap = {
      "All":      allMyTasks,
      "Open":     activeTasks,
      "No Date":  allMyTasks.filter(t => taskIsActive(t) && !t.end),
      "Completed": allMyTasks.filter(taskIsDone)
    };
    const filteredTasks = sortMyTasks(filteredMap[taskFilter] || allMyTasks);

    const taskFilterMeta = [
      { key: "All",       label: `All (${allMyTasks.length})` },
      { key: "Open",      label: `Open (${activeTasks.length})` },
      { key: "No Date",   label: "No Date" },
      { key: "Completed", label: "Completed" }
    ];

    const attentionItems = [];
    if (docsNeedingMe.length) attentionItems.push({ icon: "bx-file-blank", label: `${docsNeedingMe.length} document${docsNeedingMe.length !== 1 ? "s" : ""} pending review`, route: "docreview" });
    if (myTravel.length)      attentionItems.push({ icon: "bx-plane",      label: `${myTravel.length} travel request${myTravel.length !== 1 ? "s" : ""} awaiting action`,   route: "travel" });

    return `<div class="overview-grid overview-grid--my">

      <section class="overview-hero overview-hero--my">
        <div class="overview-hero-copy">
          <div class="overview-kicker">My Overview · ${thisYear}</div>
          <h2>Your work at a glance</h2>
          <p>${greeting}, ${escapeHtml(firstName)}. You have ${activeTasks.length} open task${activeTasks.length !== 1 ? "s" : ""} across ${myProjects.length} assigned project${myProjects.length !== 1 ? "s" : ""}. Updated ${asOfLabel()}.</p>
          <div class="overview-hero-actions">
            <button class="btn-aewttr-outline btn-aewttr-sm" data-route="projects">My Projects</button>
            <button class="btn-aewttr-outline btn-aewttr-sm" data-route="docreview">Document Review</button>
            <button class="btn-aewttr-outline btn-aewttr-sm" data-route="weekly">Weekly Meeting</button>
          </div>
        </div>
        <div class="overview-hero-stats">
          <div class="overview-metric-grid">
            <article class="overview-metric-card" ${tip(`Tasks you completed in ${thisYear}`)}><span>Tasks Completed</span><strong>${tasksCompletedThisYear.length}</strong><small>${thisYear}</small></article>
            <article class="overview-metric-card" ${tip(`Documents reviewed or concurred on in ${thisYear}`)}><span>Docs Reviewed</span><strong>${docsReviewedThisYear.length}</strong><small>${thisYear}</small></article>
            <article class="overview-metric-card" ${tip(`Documents you signed or approved in ${thisYear}`)}><span>Docs Signed</span><strong>${docsSignedByMe.length}</strong><small>${thisYear}</small></article>
            <article class="overview-metric-card" ${tip(`Weekly meeting sessions attended in ${thisYear}`)}><span>Meetings</span><strong>${meetingsAttended.length}</strong><small>${thisYear}</small></article>
            <article class="overview-metric-card" ${tip(`Travel requests submitted in ${thisYear}`)}><span>Travel Requests</span><strong>${travelThisYear.length}</strong><small>${thisYear}</small></article>
            <article class="overview-metric-card" ${tip("Projects you are currently assigned to")}><span>Projects</span><strong>${myProjects.length}</strong><small>assigned</small></article>
          </div>
        </div>
      </section>

      ${attentionItems.length ? `
      <section class="overview-panel overview-panel--attention">
        <div class="overview-panel-head"><div>
          <div class="side-panel-title">Needs Attention</div>
          <div class="overview-panel-sub">Items waiting on you.</div>
        </div></div>
        <div class="overview-spotlight-list">
          ${attentionItems.map(item => `
            <button class="overview-spotlight-row" ${item.route ? `data-route="${escapeHtml(item.route)}"` : `data-ov-task-filter="${escapeHtml(item.filter || "")}"`}>
              <div><strong><i class="bx ${item.icon}" style="margin-right:6px;vertical-align:-2px;"></i>${item.label}</strong></div>
              <i class="bx bx-chevron-right" style="color:var(--aewttr-muted);font-size:18px;"></i>
            </button>
          `).join("")}
        </div>
      </section>` : ""}

      <section class="overview-panel overview-panel--tasks">
        <div class="overview-panel-head"><div>
          <div class="side-panel-title">My Tasks</div>
          <div class="overview-panel-sub">Project tasks assigned to you.</div>
        </div></div>
        <div class="filter-pills" style="margin-bottom:14px;">
          ${taskFilterMeta.map(f => `<button class="filter-pill${taskFilter === f.key ? " active" : ""}" data-ov-task-filter="${escapeHtml(f.key)}">${escapeHtml(f.label)}</button>`).join("")}
        </div>
        ${filteredTasks.length ? `
        <div class="overview-table-shell">
          <table class="aewttr-table">
            <thead><tr><th>Task</th><th>Project</th><th>Due</th><th>Priority</th><th>Status</th></tr></thead>
            <tbody>
              ${filteredTasks.map(task => `<tr data-route="projects/${escapeHtml(task.projectId)}/workspace">
                  <td><strong>${escapeHtml(task.title || "Untitled")}</strong></td>
                  <td>${escapeHtml(task.projectName)}</td>
                  <td>${task.end ? fmtDate(task.end) : "—"}</td>
                  <td>${priorityTag(task.priority)}</td>
                  <td>${statusPill(task.status || "Not Started")}</td>
                </tr>`).join("")}
            </tbody>
          </table>
        </div>` : `<div class="empty-state" style="padding:30px 0;">No tasks match this filter.</div>`}
      </section>

      <section class="overview-panel overview-panel--projects">
        <div class="overview-panel-head"><div>
          <div class="side-panel-title">My Projects</div>
          <div class="overview-panel-sub">Projects where you hold a role. Click to open the workspace.</div>
        </div></div>
        ${myProjects.length ? `
        <div class="overview-spotlight-list">
          ${myProjects.map(proj => {
            const tasks  = allProjectTasks(proj.id);
            const open   = tasks.filter(taskIsActive).length;
            const risks  = projectRiskCount(proj.id);
            return `<button class="overview-spotlight-row" data-route="projects/${escapeHtml(proj.id)}/workspace">
              <div>
                <strong>${escapeHtml(proj.name || "Untitled project")}</strong>
                <span>${escapeHtml(myRoleOnProject(proj) || "Member")} · ${escapeHtml(proj.team || "—")}</span>
              </div>
              <div style="display:flex;gap:6px;align-items:center;flex-shrink:0;">
                ${open  ? `<span class="kc-badge">${open} open</span>` : ""}
                ${risks ? `<span class="kc-badge kc-badge--warn">${risks} risk${risks !== 1 ? "s" : ""}</span>` : ""}
                <i class="bx bx-chevron-right" style="color:var(--aewttr-muted);font-size:18px;"></i>
              </div>
            </button>`;
          }).join("")}
        </div>` : `<div class="empty-state" style="padding:20px 0;">You are not assigned to any projects.</div>`}
      </section>
    </div>`;
  }

  /* ---- TEAM OVERVIEW ---- */

  function getAllTasksFlat() {
    const result = [];
    (db.projects || []).forEach(proj => {
      allProjectTasks(proj.id).forEach(task => {
        result.push({ ...task, projectId: proj.id, proj });
      });
    });
    return result;
  }

  function buildTeamMetrics(projects) {
    const allTasks = getAllTasksFlat();
    const docs = typeof getAllDocReviewRecords === "function" ? getAllDocReviewRecords() : [];
    const docsInReview = docs.filter(d => !d.isArchived && !["Review Complete", "Signed", "Archived"].includes(d._column || ""));
    /* Matches the Travel page's Awaiting Finance filter, so the tile and the
       queue it links to always report the same number. The old filter tested a
       "Pending" status that does not exist. */
    const pendingTravel = (db.travelRequests || []).filter(r =>
      r.status === "Submitted" && r.chargeObjectStatus === "Pending"
      && (typeof travelCategory !== "function" || travelCategory(r) !== "Leave"));
    const totalRisks = projects.reduce((sum, p) => sum + projectRiskCount(p.id), 0);
    const criticalRisks = projects.reduce((sum, p) => sum + projectCriticalRiskCount(p.id), 0);
    const totalFunded = projects.reduce((sum, p) => sum + (parseFloat(p.fundedOnContractAmount) || 0), 0);
    const totalReimb = projects.reduce((sum, p) => sum + (parseFloat(p.reimbursableAmount) || 0), 0);
    const totalTravelCost = (db.travelRequests || []).reduce((sum, r) => sum + (parseFloat(r.estimatedCost || r.totalCost) || 0), 0);

    const activeProjects = projects.filter(p => {
      const s = derivedProjectStatus(p);
      return s === "Active" || s === "Needs Attention";
    });

    const openTasks = allTasks.filter(taskIsActive).length;
    const doneTasks = allTasks.filter(taskIsDone).length;
    return {
      activeProjects: activeProjects.length,
      totalProjects: projects.length,
      activeTasks: openTasks,
      completedTasks: doneTasks,
      openRisks: totalRisks,
      criticalRisks,
      docsInReview: docsInReview.length,
      pendingTravel: pendingTravel.length,
      totalFunded,
      totalReimb,
      totalTravelCost
    };
  }

  const teamTabs = [
    { key: "dashboard",  label: "Dashboard",  icon: "bx-tachometer" },
    { key: "portfolio",  label: "Portfolio",  icon: "bx-briefcase" },
    { key: "workload",   label: "Workload",   icon: "bx-bar-chart-alt-2" },
    { key: "resources",  label: "Resources",  icon: "bx-sitemap" },
    { key: "risks",      label: "Risks",      icon: "bx-error-alt" },
    { key: "operations", label: "Operations", icon: "bx-list-ul" }
  ];

  function renderTeamOverview(projects) {
    const metrics = buildTeamMetrics(projects);
    const activeTab = window.AEWTTR.state.overviewTeamTab || "portfolio";

    let tabContent = "";
    if (activeTab === "dashboard")   tabContent = renderDashboardTab(projects);
    else if (activeTab === "portfolio")  tabContent = renderPortfolioTab(projects);
    else if (activeTab === "workload")   tabContent = renderTeamWorkload(projects);
    else if (activeTab === "resources")  tabContent = renderResourcesGraph(projects);
    else if (activeTab === "risks")      tabContent = renderRiskRegister(projects);
    else if (activeTab === "operations") tabContent = renderOperationsQueues();

    return `<div class="overview-grid overview-team-overview">

      <section class="ov-team-stats-bar">
        <div class="ov-team-stat${metrics.activeProjects > 0 ? " ov-stat-accent" : ""}" ${tip("Projects with at least one active task")}>
          <strong>${metrics.activeProjects}</strong><span>Active projects</span><small>${metrics.totalProjects} total</small>
        </div>
        <div class="ov-team-stat" ${tip("Tasks across all projects not yet complete")}>
          <strong>${metrics.activeTasks}</strong><span>Open tasks</span>
        </div>
        <div class="ov-team-stat ov-stat-accent" ${tip("Tasks completed across all projects")}>
          <strong>${metrics.completedTasks}</strong><span>Completed tasks</span>
        </div>
        <div class="ov-team-stat${metrics.openRisks > 0 ? " ov-stat-warn" : ""}" ${tip("Open risks across all projects (not Closed/Resolved/New)")}>
          <strong>${metrics.openRisks}</strong><span>Open risks</span>${metrics.criticalRisks ? `<small style="color:var(--aewttr-red);">${metrics.criticalRisks} critical</small>` : ""}
        </div>
        <div class="ov-team-stat" ${tip("Total funded amount on contract across active portfolio")}>
          <strong style="color:var(--aewttr-success,#27ae60);">$${metrics.totalFunded ? (metrics.totalFunded / 1000).toFixed(0) + "K" : "0"}</strong><span>Contract Funding</span><small>$${(metrics.totalReimb / 1000).toFixed(0)}K reimb</small>
        </div>
        <div class="ov-team-stat" ${tip("Document review records not yet complete")}>
          <strong>${metrics.docsInReview}</strong><span>Docs in review</span>
        </div>
        <div class="ov-team-stat" ${tip("Submitted travel requests still waiting on a charge object from finance")}>
          <strong>${metrics.pendingTravel}</strong><span>Awaiting C/O</span><small>$${(metrics.totalTravelCost / 1000).toFixed(1)}K est</small>
        </div>
        <div class="ov-team-stats-actions">
          <button class="btn-aewttr-outline btn-aewttr-sm" data-route="projects"><i class="bx bx-briefcase"></i> All Projects</button>
          <button class="btn-aewttr-outline btn-aewttr-sm" data-route="people"><i class="bx bx-group"></i> Team Members</button>
        </div>
      </section>

      <section class="overview-panel overview-panel--tabs">
        <div class="ov-team-tabs">
          ${teamTabs.map(t => `<button class="ov-team-tab${t.key === activeTab ? " active" : ""}" data-team-tab="${t.key}">
            <i class="bx ${t.icon}"></i> ${t.label}
          </button>`).join("")}
        </div>
        <div class="ov-team-tab-content">
          ${tabContent}
        </div>
      </section>
    </div>`;
  }

  function matchesFilter(proj, f) {
    if (!f) return true;
    if (f.search) {
      const q = f.search.toLowerCase();
      const hay = [proj.id, proj.name, proj.team, proj.program, proj.configEndItem, proj.ato, projectPMDisplayName(proj)].join(" ").toLowerCase();
      if (!hay.includes(q)) return false;
    }
    if (f.team && f.team !== "All" && proj.team !== f.team) return false;
    if (f.status && f.status !== "All" && derivedProjectStatus(proj) !== f.status) return false;
    if (f.fundingType && f.fundingType !== "All" && proj.fundingType !== f.fundingType) return false;
    if (f.portfolio && f.portfolio !== "All") {
      const portfolios = Array.isArray(proj.portfolios) ? proj.portfolios : (proj.portfolios ? [proj.portfolios] : []);
      if (!portfolios.some(p => p === f.portfolio)) return false;
    }
    if (f.endItem && f.endItem !== "All" && !String(proj.configEndItem || "").split(",").map(s => s.trim()).includes(f.endItem)) return false;
    if (f.ato && f.ato !== "All" && (proj.ato || "") !== f.ato) return false;
    return true;
  }

  function donutChartSvg(segments, { size = 130, strokeW = 22, centerLabel = "", centerSub = "" } = {}) {
    const total = segments.reduce((s, x) => s + x.value, 0);
    if (!total) return `<div class="ov-donut-empty">No data</div>`;
    const r = size / 2 - strokeW / 2 - 1;
    const circ = 2 * Math.PI * r;
    const cx = size / 2, cy = size / 2;
    let cum = 0;
    const arcs = segments.filter(s => s.value > 0).map(({ value, color }) => {
      const pct = value / total;
      const dash = (pct * circ).toFixed(2);
      const gap = (circ - pct * circ).toFixed(2);
      const offset = (-(cum / total) * circ).toFixed(2);
      cum += value;
      return `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${color}" stroke-width="${strokeW}" stroke-dasharray="${dash} ${gap}" stroke-dashoffset="${offset}" transform="rotate(-90 ${cx} ${cy})"/>`;
    });
    const labelY = centerSub ? cy - 2 : cy + 6;
    return `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}" aria-hidden="true" style="flex-shrink:0;">
      <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="var(--aewttr-border)" stroke-width="${strokeW}"/>
      ${arcs.join("")}
      <text x="${cx}" y="${labelY}" text-anchor="middle" font-size="18" font-weight="700" fill="var(--aewttr-text)">${escapeHtml(String(centerLabel))}</text>
      ${centerSub ? `<text x="${cx}" y="${cy + 16}" text-anchor="middle" font-size="10" fill="var(--aewttr-muted)">${escapeHtml(centerSub)}</text>` : ""}
    </svg>`;
  }

  /* ---- DASHBOARD TAB (leadership executive view) ---- */

  function renderDashboardTab(projects) {
    const docs = typeof getAllDocReviewRecords === "function" ? getAllDocReviewRecords() : [];
    const docsNeedingAction = docs.filter(d => {
      if (d.isArchived) return false;
      const col = d._column || "";
      return !["Review Complete", "Signed", "Archived"].includes(col);
    }).length;
    /* "Pending" is not a travel status in this app — the real ones are
       Submitted / Approved / Denied / Completed / Withdrawn / Cancelled — so the
       old count silently matched only Submitted and called it "pending", which
       told a lead nothing about what actually needed doing. These three are the
       questions the Team Overview is asked: what is stuck with finance, who is
       about to travel, and who is about to be out. Definitions match the Travel
       page's own filters so the numbers agree with the queue they link to. */
    const allTravel = db.travelRequests || [];
    const isLeaveReq = (r) => (typeof travelCategory === "function" ? travelCategory(r) === "Leave" : false);
    const isUpcoming = (r) => (typeof isUpcomingOrCurrentTravel === "function" ? isUpcomingOrCurrentTravel(r) : false);
    const travelAwaitingCo = allTravel.filter(r =>
      r.status === "Submitted" && r.chargeObjectStatus === "Pending" && !isLeaveReq(r)).length;
    const travelUpcoming = allTravel.filter(r => isUpcoming(r) && !isLeaveReq(r)).length;
    const leaveUpcoming = allTravel.filter(r => isUpcoming(r) && isLeaveReq(r)).length;

    // Per-project health — measured by open/completed tasks + risks
    const projectHealth = projects.map(proj => {
      const tasks = allProjectTasks(proj.id);
      const open = tasks.filter(taskIsActive);
      const done = tasks.filter(taskIsDone);
      const extra = (db.projectExtra && db.projectExtra[proj.id]) || {};
      const openRisks = (extra.risks || []).filter(r => {
        if (isPlaceholderRisk(r)) return false;
        const s = String(r.status || "").toLowerCase();
        return !["closed", "resolved", "accepted"].includes(s);
      });
      const criticalRisks = openRisks.filter(r =>
        ["high", "critical"].some(k =>
          String(r.likelihood || "").toLowerCase().includes(k) ||
          String(r.impact || "").toLowerCase().includes(k) ||
          String(r.severity || "").toLowerCase().includes(k)
        )
      );
      const health = criticalRisks.length ? "attention"
        : open.length ? "active"
        : done.length ? "complete"
        : "inactive";
      const pctDone = (open.length + done.length) > 0
        ? Math.round((done.length / (open.length + done.length)) * 100) : null;
      return { proj, open, done, openRisks, criticalRisks, risks: openRisks.length, health, pctDone };
    });

    // Aggregate KPIs (the shared stats bar above the tabs already covers
    // projects/tasks/risks/funding totals — only track what's unique here).
    const allCritical = projectHealth.flatMap(p => p.criticalRisks);

    // Project health cards sorted: attention → active → complete → inactive
    const healthOrder = { attention: 0, active: 1, complete: 2, inactive: 3 };
    const sorted = [...projectHealth].sort((a, b) =>
      (healthOrder[a.health] ?? 9) - (healthOrder[b.health] ?? 9) ||
      (a.proj.name || "").localeCompare(b.proj.name || "")
    );

    const clsMap   = { attention: "ov-health-card--amber", active: "ov-health-card--green", complete: "ov-health-card--blue", inactive: "ov-health-card--gray" };
    const labelMap  = { attention: "AT RISK", active: "ACTIVE", complete: "COMPLETE", inactive: "INACTIVE" };
    const pillMap   = { attention: "ov-status-attention", active: "ov-status-active", complete: "ov-status-done", inactive: "ov-status-inactive" };

    const cardsHtml = sorted.map(({ proj, open, done, risks, health, pctDone }) => {
      const pm = projectPMDisplayName(proj);
      return `<button class="ov-health-card ${clsMap[health]}" data-route="projects/${escapeHtml(proj.id)}/workspace">
        <div class="ov-health-card-head">
          <div class="ov-health-card-name">${escapeHtml(proj.name || "Untitled")}</div>
          <span class="ov-derived-status ${pillMap[health]}">${labelMap[health]}</span>
        </div>
        ${pm ? `<div class="ov-health-card-meta"><span><i class="bx bx-user"></i>${escapeHtml(pm)}</span>${proj.team ? `<span><i class="bx bx-group"></i>${escapeHtml(proj.team)}</span>` : ""}</div>` : ""}
        <div class="ov-health-card-stats">
          <span><b>${open.length}</b> open</span>
          <span class="ov-hcs-done"><b>${done.length}</b> done${pctDone !== null ? ` (${pctDone}%)` : ""}</span>
          ${risks ? `<span class="ov-hcs-warn"><b>${risks}</b> risk${risks !== 1 ? "s" : ""}</span>` : ""}
        </div>
      </button>`;
    }).join("") || `<div class="empty-state">No projects in portfolio.</div>`;

    // Needs attention panel — critical risks + docs pending
    const groups = [];
    if (allCritical.length) {
      groups.push({
        icon: "bx-error-alt", label: "Critical Risks", color: "var(--aewttr-amber)",
        items: projectHealth.filter(p => p.criticalRisks.length).flatMap(p =>
          p.criticalRisks.slice(0, 3).map(r => ({ title: r.title || r.description || "Risk", sub: p.proj.name || p.proj.id, route: `projects/${escapeHtml(p.proj.id)}/workspace` }))
        ).slice(0, 8)
      });
    }
    if (docsNeedingAction) {
      groups.push({
        icon: "bx-file-blank", label: "Documents Pending", color: "var(--aewttr-muted)",
        items: [{ title: `${docsNeedingAction} document${docsNeedingAction !== 1 ? "s" : ""} awaiting review`, sub: "Open Document Review →", route: "docreview" }]
      });
    }
    const travelItems = [];
    if (travelAwaitingCo) {
      travelItems.push({ title: `${travelAwaitingCo} request${travelAwaitingCo !== 1 ? "s" : ""} awaiting a charge object`,
        sub: "Blocked with finance — open Operations →", tab: "operations" });
    }
    if (travelUpcoming) {
      travelItems.push({ title: `${travelUpcoming} trip${travelUpcoming !== 1 ? "s" : ""} upcoming`,
        sub: "Approved or submitted, still ahead →", tab: "operations" });
    }
    if (leaveUpcoming) {
      travelItems.push({ title: `${leaveUpcoming} leave request${leaveUpcoming !== 1 ? "s" : ""} upcoming`,
        sub: "Team members out soon →", tab: "operations" });
    }
    if (travelItems.length) {
      groups.push({ icon: "bx-navigation", label: "Travel & Leave", color: "var(--aewttr-muted)", items: travelItems });
    }

    const healthCounts = { attention: 0, active: 0, complete: 0, inactive: 0 };
    projectHealth.forEach(p => { healthCounts[p.health] = (healthCounts[p.health] || 0) + 1; });
    const healthDonut = donutChartSvg([
      { value: healthCounts.attention, color: "var(--aewttr-amber)" },
      { value: healthCounts.active, color: "var(--aewttr-green)" },
      { value: healthCounts.complete, color: "var(--aewttr-blue)" },
      { value: healthCounts.inactive, color: "var(--aewttr-muted)" }
    ], { size: 92, strokeW: 14, centerLabel: String(projects.length), centerSub: "projects" });

    const escalationHtml = groups.length
      ? groups.map(g => `<div class="ov-dash-escalation-group">
          <div class="ov-dash-escalation-head"><i class="bx ${g.icon}" style="color:${g.color}"></i><strong>${escapeHtml(g.label)}</strong><span class="ov-queue-count">${g.items.length}</span></div>
          ${g.items.map(item => `<button class="overview-spotlight-row" ${item.tab ? `data-ov-team-tab="${escapeHtml(item.tab)}"` : `data-route="${escapeHtml(item.route)}"`}>
            <div><strong>${escapeHtml(item.title)}</strong><span>${escapeHtml(item.sub)}</span></div>
            <i class="bx bx-chevron-right" style="color:var(--aewttr-muted);font-size:18px;"></i>
          </button>`).join("")}
        </div>`).join("")
      : `<div class="empty-state" style="padding:20px 8px;text-align:center;"><i class="bx bx-check-circle" style="font-size:26px;color:var(--aewttr-green);display:block;margin-bottom:8px;"></i>Everything looks good.</div>`;

    return `<div class="ov-dash-layout">
      <div class="ov-dash-main">
        <div class="ov-dash-section-head">
          <div>
            <div class="side-panel-title" style="margin-bottom:2px;">Portfolio Health</div>
            <div class="overview-panel-sub">${projects.length} project${projects.length !== 1 ? "s" : ""} · sorted by status · click any card to open workspace</div>
          </div>
          <div class="ov-dash-health-donut">
            ${healthDonut}
            <div class="ov-dash-health-donut-legend">
              <span><i style="background:var(--aewttr-amber)"></i>${healthCounts.attention} at risk</span>
              <span><i style="background:var(--aewttr-green)"></i>${healthCounts.active} active</span>
              <span><i style="background:var(--aewttr-blue)"></i>${healthCounts.complete} complete</span>
              <span><i style="background:var(--aewttr-muted)"></i>${healthCounts.inactive} inactive</span>
            </div>
          </div>
        </div>
        <div class="ov-health-grid">${cardsHtml}</div>
      </div>
      <aside class="ov-dash-aside">
        <div class="side-panel-title" style="margin-bottom:12px;">Needs Attention</div>
        ${escalationHtml}
      </aside>
    </div>`;
  }

  /* ---- RISK REGISTER TAB ---- */

  function renderRiskRegister(projects) {
    const allRisks = [];
    projects.forEach(proj => {
      const extra = (db.projectExtra && db.projectExtra[proj.id]) || {};
      (extra.risks || []).forEach(r => {
        if (isPlaceholderRisk(r)) return;
        const s = String(r.status || "").toLowerCase();
        if (["closed", "resolved", "accepted"].includes(s)) return;
        allRisks.push({ ...r, projectId: proj.id, projectName: proj.name || proj.id });
      });
    });

    const sevOrder = { critical: 0, high: 1, medium: 2, low: 3 };
    function sev(r) {
      const combo = [r.likelihood, r.impact, r.severity, r.priority].map(v => String(v || "").toLowerCase()).join(" ");
      if (combo.includes("critical")) return "critical";
      if (combo.includes("high")) return "high";
      if (combo.includes("medium")) return "medium";
      if (combo.includes("low")) return "low";
      return "";
    }
    function sevStyle(s) {
      if (s === "critical" || s === "high") return "ov-status-attention";
      if (s === "medium") return "ov-status-inactive";
      if (s === "low") return "ov-status-done";
      return "ov-status-inactive";
    }

    allRisks.sort((a, b) => {
      const sa = sevOrder[sev(a)] ?? 99, sb = sevOrder[sev(b)] ?? 99;
      return sa !== sb ? sa - sb : String(a.projectName).localeCompare(String(b.projectName));
    });

    const highCount = allRisks.filter(r => ["critical", "high"].includes(sev(r))).length;

    return `<div class="ov-section-head" style="margin-bottom:16px;">
      <div>
        <div class="side-panel-title">Open Risk Register</div>
        <div class="overview-panel-sub">${allRisks.length} open risk${allRisks.length !== 1 ? "s" : ""} · ${highCount} high/critical · closed, resolved, and accepted excluded</div>
      </div>
    </div>
    ${allRisks.length ? `<div class="overview-table-shell">
      <table class="aewttr-table">
        <thead><tr><th>Risk</th><th>Project</th><th>Likelihood</th><th>Impact</th><th>Owner</th><th>Status</th></tr></thead>
        <tbody>
          ${allRisks.map(r => {
            const s = sev(r);
            const lSev = String(r.likelihood || "").toLowerCase();
            const iSev = String(r.impact || "").toLowerCase();
            return `<tr style="cursor:pointer;" data-route="projects/${escapeHtml(r.projectId)}/workspace">
              <td style="max-width:280px;">
                <strong>${escapeHtml(r.title || r.description || "Risk")}</strong>
                ${r.mitigation ? `<div style="font-size:11px;color:var(--aewttr-muted);margin-top:2px;">${escapeHtml(String(r.mitigation).slice(0, 90))}${String(r.mitigation).length > 90 ? "…" : ""}</div>` : ""}
              </td>
              <td><span class="kc-badge">${escapeHtml(r.projectName)}</span></td>
              <td>${r.likelihood ? `<span class="ov-derived-status ${sevStyle(lSev)}">${escapeHtml(r.likelihood)}</span>` : "—"}</td>
              <td>${r.impact ? `<span class="ov-derived-status ${sevStyle(iSev)}">${escapeHtml(r.impact)}</span>` : "—"}</td>
              <td>${escapeHtml(r.owner || r.riskOwner || "—")}</td>
              <td>${r.status ? `<span class="ov-derived-status ov-status-inactive">${escapeHtml(r.status)}</span>` : "—"}</td>
            </tr>`;
          }).join("")}
        </tbody>
      </table>
    </div>` : `<div class="empty-state" style="padding:32px 0;text-align:center;">
      <i class="bx bx-check-shield" style="font-size:28px;color:var(--aewttr-green);display:block;margin-bottom:8px;"></i>
      No open risks across the portfolio.
    </div>`}`;
  }

  function renderPortfolioTab(projects) {
    const f = window.AEWTTR.state.overviewFilters || {};
    const filtered = projects.filter(p => matchesFilter(p, f));

    const portfolioOptions = Array.from(new Set(projects.flatMap(p => Array.isArray(p.portfolios) ? p.portfolios : []).filter(Boolean))).sort();
    const endItemOptions = Array.from(new Set(projects.flatMap(p => String(p.configEndItem || "").split(",").map(s => s.trim()).filter(Boolean)))).sort();
    const teamOptions = Array.from(new Set(projects.map(p => p.team).filter(Boolean))).sort();
    const hasFilter = !!(f.search || (f.team && f.team !== "All") || (f.status && f.status !== "All") || (f.portfolio && f.portfolio !== "All") || (f.endItem && f.endItem !== "All"));

    return `
      <div class="overview-filters" style="margin-bottom:16px;">
        <div class="search-box overview-filter-search"><i class="bx bx-search" aria-hidden="true"></i><input class="input-aewttr" id="ov-sys-search" placeholder="Search name, PM, team…" value="${escapeHtml(f.search || "")}"></div>
        ${portfolioOptions.length ? `<select class="select-aewttr" id="ov-sys-portfolio">
          <option value="All">All Portfolios</option>
          ${portfolioOptions.map(o => `<option value="${escapeHtml(o)}"${f.portfolio === o ? " selected" : ""}>${escapeHtml(o)}</option>`).join("")}
        </select>` : ""}
        ${endItemOptions.length ? `<select class="select-aewttr" id="ov-sys-enditem">
          <option value="All">All End Items</option>
          ${endItemOptions.map(o => `<option value="${escapeHtml(o)}"${f.endItem === o ? " selected" : ""}>${escapeHtml(o)}</option>`).join("")}
        </select>` : ""}
        <select class="select-aewttr" id="ov-sys-status">
          <option value="All"${(!f.status || f.status === "All") ? " selected" : ""}>All Statuses</option>
          ${["Active","Needs Attention","No Active Work","Completed"].map(s => `<option value="${escapeHtml(s)}"${f.status === s ? " selected" : ""}>${escapeHtml(s)}</option>`).join("")}
        </select>
        ${teamOptions.length ? `<select class="select-aewttr" id="ov-sys-team">
          <option value="All"${(!f.team || f.team === "All") ? " selected" : ""}>All Teams</option>
          ${teamOptions.map(t => `<option value="${escapeHtml(t)}"${f.team === t ? " selected" : ""}>${escapeHtml(t)}</option>`).join("")}
        </select>` : ""}
        ${hasFilter ? `<button class="btn-aewttr-ghost btn-aewttr-sm" id="ov-sys-clear"><i class="bx bx-x"></i> Clear</button>` : ""}
      </div>
      <div class="overview-table-shell">
        <table class="aewttr-table overview-table">
          <thead><tr>
            <th>Project</th><th>Status</th><th>Open</th><th>Done</th><th>Risks</th><th>PM</th><th>End Item</th><th>Portfolio</th><th>Funding</th><th>Due</th>
          </tr></thead>
          <tbody>
            ${filtered.length ? filtered.map(proj => {
              const tasks = allProjectTasks(proj.id);
              const open = tasks.filter(taskIsActive).length;
              const done = tasks.filter(taskIsDone).length;
              const pct = (open + done) > 0 ? Math.round((done / (open + done)) * 100) : null;
              const risks = projectRiskCount(proj.id);
              const portfolio = (Array.isArray(proj.portfolios) && proj.portfolios[0]) || "—";
              const pmName = projectPMDisplayName(proj);
              const funded = proj.fundedOnContractAmount ? `$${Number(proj.fundedOnContractAmount).toLocaleString()}` : "—";
              return `<tr data-route="projects/${escapeHtml(proj.id)}/workspace" style="cursor:pointer;">
                <td><strong>${escapeHtml(proj.name || "Untitled project")}</strong></td>
                <td>${derivedStatusPill(derivedProjectStatus(proj))}</td>
                <td>${open || "—"}</td>
                <td>${done ? `${done}${pct !== null ? ` <span style="color:var(--aewttr-muted);font-size:11px;">(${pct}%)</span>` : ""}` : "—"}</td>
                <td><span class="${risks > 0 ? "ov-risks-badge" : ""}">${risks || "—"}</span></td>
                <td>${escapeHtml(pmName || "—")}</td>
                <td>${escapeHtml(proj.configEndItem || "—")}</td>
                <td><span class="kc-badge">${escapeHtml(portfolio)}</span></td>
                <td><strong style="color:var(--aewttr-green);">${funded}</strong></td>
                <td>${proj.dueDate ? fmtDate(proj.dueDate) : "—"}</td>
              </tr>`;
            }).join("") : `<tr><td colspan="10"><div class="empty-state" style="padding:20px 0;">No projects match these filters.</div></td></tr>`}
          </tbody>
        </table>
      </div>`;
  }

  function renderEndItemAnalysis(projects) {
    const f = window.AEWTTR.state.overviewFilters || {};

    const groups = {};
    projects.forEach(proj => {
      const ei = projectEndItem(proj);
      if (!groups[ei]) groups[ei] = { name: ei, projects: [] };
      groups[ei].projects.push(proj);
    });
    const sortedGroups = Object.values(groups).sort((a, b) => a.name.localeCompare(b.name));

    const teamOptions = Array.from(new Set(projects.map(p => p.team).filter(Boolean))).sort();
    const fundingOptions = Array.from(new Set(projects.map(p => p.fundingType).filter(Boolean))).sort();
    const hasActiveFilter = !!(f.search || (f.team && f.team !== "All") || (f.status && f.status !== "All") || (f.fundingType && f.fundingType !== "All"));

    return `<div class="ov-section-head">
        <div>
          <div class="side-panel-title">Portfolio by End Item</div>
          <div class="overview-panel-sub">Work grouped by configuration end item. Expand a row to see projects, tasks, and risk exposure.</div>
        </div>
      </div>
      <div class="overview-filters" style="margin-bottom:16px;">
        <div class="search-box overview-filter-search"><i class="bx bx-search" aria-hidden="true"></i><input class="input-aewttr" id="ov-sys-search" placeholder="Search projects, PM, team…" value="${escapeHtml(f.search || "")}"></div>
        <select class="select-aewttr" id="ov-sys-team" aria-label="Filter by team">
          <option value="All" ${(!f.team || f.team === "All") ? "selected" : ""}>All Teams</option>
          ${teamOptions.map(t => `<option value="${escapeHtml(t)}" ${f.team === t ? "selected" : ""}>${escapeHtml(t)}</option>`).join("")}
        </select>
        <select class="select-aewttr" id="ov-sys-status" aria-label="Filter by project status">
          <option value="All" ${(!f.status || f.status === "All") ? "selected" : ""}>All Statuses</option>
          ${["Active","Needs Attention","No Active Work","Completed"].map(s => `<option value="${escapeHtml(s)}" ${f.status === s ? "selected" : ""}>${escapeHtml(s)}</option>`).join("")}
        </select>
        ${fundingOptions.length ? `<select class="select-aewttr" id="ov-sys-funding" aria-label="Filter by funding type">
          <option value="All" ${(!f.fundingType || f.fundingType === "All") ? "selected" : ""}>All Funding</option>
          ${fundingOptions.map(ft => `<option value="${escapeHtml(ft)}" ${f.fundingType === ft ? "selected" : ""}>${escapeHtml(ft)}</option>`).join("")}
        </select>` : ""}
        ${hasActiveFilter ? `<button class="btn-aewttr-ghost btn-aewttr-sm" id="ov-sys-clear">Clear filters</button>` : ""}
      </div>
      <div id="ov-system-groups">
        ${sortedGroups.map(sg => renderEndItemGroup(sg, f)).filter(Boolean).join("") ||
          `<div class="empty-state" style="padding:30px 0;">No projects match these filters.</div>`}
      </div>`;
  }

  function renderEndItemGroup(sg, f) {
    const filteredProjects = sg.projects.filter(p => matchesFilter(p, f));
    if (!filteredProjects.length) return "";

    const allTasks = filteredProjects.flatMap(p => allProjectTasks(p.id));
    const active = allTasks.filter(taskIsActive).length;
    const risks = filteredProjects.reduce((sum, p) => sum + projectRiskCount(p.id), 0);
    const teams = [...new Set(filteredProjects.map(p => p.team).filter(Boolean))];
    const sysId = "ov-sys-" + encodeURIComponent(sg.name).replace(/%/g, "X");

    return `<div class="ov-system-group">
      <button class="ov-system-group-header" data-expand-sys="${sysId}" aria-expanded="false">
        <div class="ov-system-group-title">
          <i class="bx bx-chevron-right ov-expand-icon"></i>
          <strong>${escapeHtml(sg.name)}</strong>
          <span class="ov-system-group-meta">
            ${filteredProjects.length} project${filteredProjects.length !== 1 ? "s" : ""}
            · ${active} active task${active !== 1 ? "s" : ""}
            ${risks ? `· <span class="ov-risks-count">${risks} risk${risks !== 1 ? "s" : ""}</span>` : ""}
          </span>
        </div>
        <div class="ov-system-group-right">
          ${teams.map(t => `<span class="ov-team-badge">${escapeHtml(t)}</span>`).join("")}
        </div>
      </button>
      <div class="ov-system-group-body" id="${sysId}" hidden>
        <div class="overview-table-shell" style="margin:6px 0 12px;">
          <table class="aewttr-table">
            <thead>
              <tr><th>Project</th><th>Status</th><th>Active</th><th>Risks</th><th>PM</th><th>Team</th><th>Portfolio</th><th>Funding</th><th>FY</th><th>Task Order</th></tr>
            </thead>
            <tbody>
              ${filteredProjects.map(proj => {
                const ptasks = allProjectTasks(proj.id);
                const pa = ptasks.filter(taskIsActive).length;
                const pr = projectRiskCount(proj.id);
                const portfolio = (proj.portfolios && proj.portfolios[0]) || "—";
                return `<tr data-route="projects/${escapeHtml(proj.id)}/workspace">
                  <td><strong>${escapeHtml(proj.name || "Untitled project")}</strong></td>
                  <td>${derivedStatusPill(derivedProjectStatus(proj))}</td>
                  <td>${pa}</td>
                  <td><span class="${pr > 0 ? "ov-risks-badge" : ""}">${pr || "—"}</span></td>
                  <td>${escapeHtml(projectPMDisplayName(proj) || "—")}</td>
                  <td>${escapeHtml(proj.team || "—")}</td>
                  <td>${escapeHtml(portfolio)}</td>
                  <td>${escapeHtml(proj.fundingType || "—")}</td>
                  <td>${escapeHtml(proj.fiscalYear ? `FY${proj.fiscalYear}` : "—")}</td>
                  <td>${escapeHtml(proj.taskOrder || "—")}</td>
                </tr>`;
              }).join("")}
            </tbody>
          </table>
        </div>
      </div>
    </div>`;
  }

  function buildTeamWorkloadRows(projects) {
    const allTasks = getAllTasksFlat();
    const members = db.members || [];
    const sortKey = window.AEWTTR.state.overviewWorkloadSort || "active";
    const personMap = {};
    const ensurePerson = (name) => {
      if (!personMap[name]) {
        const member = members.find(m => m.name === name) || { name, id: "" };
        personMap[name] = { member, activeTasks: 0, completedTasks: 0, docsPending: 0, projects: new Set(), endItems: new Set(), roles: new Set(), teams: new Set(), nextDue: null };
      }
      return personMap[name];
    };
    const docs = typeof getAllDocReviewRecords === "function" ? getAllDocReviewRecords() : [];
    docs.forEach(doc => {
      if (doc.isArchived || ["Review Complete", "Signed", "Archived"].includes(doc._column || "")) return;
      (doc.reviewers || []).forEach(reviewer => {
        if ((reviewer.decision || "Pending") !== "Pending") return;
        const member = members.find(m => typeof memberMatchesAssignee === "function" && memberMatchesAssignee(m, reviewer.name))
          || (reviewer.email ? members.find(m => (m.email || "").toLowerCase() === reviewer.email.toLowerCase()) : null);
        const name = (member && member.name) || reviewer.name;
        if (!name) return;
        ensurePerson(name).docsPending += 1;
      });
    });
    allTasks.forEach(task => {
      const assignee = task.assignee;
      if (!assignee || assignee === "Unassigned") return;
      const member = members.find(m => typeof memberMatchesAssignee === "function" && memberMatchesAssignee(m, assignee));
      const name = (member && member.name) || assignee;
      const entry = ensurePerson(name);
      if (taskIsActive(task)) {
        entry.activeTasks++;
        entry.projects.add(task.projectId);
        if (task.proj) {
          entry.endItems.add(projectEndItem(task.proj));
          if (task.proj.team) entry.teams.add(task.proj.team);
        }
        if (task.end && (!entry.nextDue || task.end < entry.nextDue)) entry.nextDue = task.end;
      }
      if (taskIsDone(task)) entry.completedTasks++;
    });
    projects.forEach(proj => {
      projectPeopleList(proj.id).forEach(person => {
        if (!person.label) return;
        const entry = ensurePerson(person.label);
        if (person.role) entry.roles.add(person.role);
        entry.projects.add(proj.id);
        entry.endItems.add(projectEndItem(proj));
        if (proj.team) entry.teams.add(proj.team);
      });
    });
    return Object.values(personMap)
      .filter(e => e.activeTasks > 0 || e.projects.size > 0 || e.docsPending > 0)
      .sort((a, b) => {
        if (sortKey === "active") return b.activeTasks - a.activeTasks;
        if (sortKey === "projects") return b.projects.size - a.projects.size;
        return a.member.name.localeCompare(b.member.name);
      });
  }

  function renderTeamWorkload(projects) {
    const members = db.members || [];
    const sortKey = window.AEWTTR.state.overviewWorkloadSort || "active";
    const rows = buildTeamWorkloadRows(projects);
    const sortButtons = [
      { key: "active", label: "Most Open" },
      { key: "projects", label: "Most Projects" },
      { key: "name", label: "Name" }
    ];

    return `<div class="ov-section-head" style="display:flex;align-items:flex-start;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:16px;">
        <div>
          <div class="side-panel-title">Team Workload</div>
          <div class="overview-panel-sub">Open and completed tasks per person across all projects.</div>
        </div>
        <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;">
          <span style="font-size:11px;color:var(--aewttr-muted);">Sort:</span>
          ${sortButtons.map(sb => `<button class="btn-aewttr-ghost btn-aewttr-sm${sortKey === sb.key ? " active" : ""}" data-ov-wl-sort="${sb.key}">${sb.label}</button>`).join("")}
        </div>
      </div>
      ${rows.length ? `
      <div class="overview-table-shell">
        <table class="aewttr-table">
          <thead><tr><th>Person</th><th>Open Tasks</th><th>Completed</th><th>Doc Reviews</th><th>Projects</th><th>Roles</th><th>Team</th><th>Next Due</th></tr></thead>
          <tbody>
            ${rows.map(entry => {
              const m = members.find(mbr => mbr.name === entry.member.name);
              const total = entry.activeTasks + entry.completedTasks;
              const pct = total > 0 ? Math.round((entry.completedTasks / total) * 100) : null;
              return `<tr ${m ? `data-person-id="${escapeHtml(m.id)}" style="cursor:pointer;"` : ""}>
                <td>
                  <div class="people-row-person">
                    ${typeof userAvatarHtml === "function" ? userAvatarHtml({ name: entry.member.name, email: m && m.email, size: 28 }) : ""}
                    <strong>${escapeHtml(entry.member.name)}</strong>
                  </div>
                </td>
                <td>${entry.activeTasks || "—"}</td>
                <td>${entry.completedTasks ? `${entry.completedTasks}${pct !== null ? ` <span style="color:var(--aewttr-muted);font-size:11px;">(${pct}%)</span>` : ""}` : "—"}</td>
                <td>${entry.docsPending ? `<span style="color:var(--aewttr-amber);font-weight:600;">${entry.docsPending}</span> pending` : "—"}</td>
                <td>${entry.projects.size || "—"}</td>
                <td>${escapeHtml([...entry.roles].join(", ") || "—")}</td>
                <td>${escapeHtml([...entry.teams].join(", ") || "—")}</td>
                <td>${entry.nextDue ? fmtDate(entry.nextDue) : "—"}</td>
              </tr>`;
            }).join("")}
          </tbody>
        </table>
      </div>` : `<div class="empty-state" style="padding:30px 0;">No project workload data found.</div>`}`;
  }

  /* ---- RESOURCES GRAPH ---- */

  function renderResourcesGraph(projects) {
    return renderTeamResourcesHub(projects);
    const allTasks = getAllTasksFlat();
    const members = db.members || [];

    // Build person → projects map
    const personProjects = {};
    const ensurePerson = (name) => {
      if (!personProjects[name]) personProjects[name] = { name, projects: new Set(), taskCount: 0 };
      return personProjects[name];
    };

    allTasks.forEach(task => {
      if (!task.assignee || task.assignee === "Unassigned") return;
      if (!taskIsActive(task)) return;
      const member = members.find(m => typeof memberMatchesAssignee === "function" && memberMatchesAssignee(m, task.assignee));
      const name = (member && member.name) || task.assignee;
      const entry = ensurePerson(name);
      entry.projects.add(task.projectId);
      entry.taskCount++;
    });

    projects.forEach(proj => {
      projectPeopleList(proj.id).forEach(person => {
        if (!person.label) return;
        ensurePerson(person.label).projects.add(proj.id);
      });
    });

    const people = Object.values(personProjects)
      .filter(p => p.projects.size > 0)
      .sort((a, b) => b.taskCount - a.taskCount || a.name.localeCompare(b.name));

    const activeProjects = projects.filter(p => {
      const tasks = allProjectTasks(p.id);
      return tasks.some(taskIsActive) || projectPeopleList(p.id).length > 0;
    }).sort((a, b) => (a.name || a.id).localeCompare(b.name || b.id));

    if (!people.length) {
      return `<div class="ov-section-head"><div class="side-panel-title">Resources</div></div>
        <div class="empty-state" style="padding:40px 0;">No resource assignments found.</div>`;
    }

    // Generate a stable color per project
    const PROJ_COLORS = [
      "#3b6bcc","#e05f2b","#2b9e6a","#9b59b6","#c0392b",
      "#16a085","#d35400","#2980b9","#8e44ad","#27ae60",
      "#e74c3c","#1abc9c","#f39c12","#6c5ce7","#00b894"
    ];
    const projColorMap = {};
    activeProjects.forEach((p, i) => { projColorMap[p.id] = PROJ_COLORS[i % PROJ_COLORS.length]; });

    const PERSON_H = 52;
    const PROJ_H = 44;
    const LEFT_W = 220;
    const RIGHT_W = 220;
    const MID_GAP = 180;
    const PAD_Y = 24;
    const svgW = LEFT_W + MID_GAP + RIGHT_W;
    const svgH = Math.max(people.length * PERSON_H, activeProjects.length * PROJ_H) + PAD_Y * 2;

    const personY = (i) => PAD_Y + i * PERSON_H + PERSON_H / 2;
    const projY = (i) => PAD_Y + i * PROJ_H + PROJ_H / 2;

    // Draw curves from person (right edge of left col) to project (left edge of right col)
    const curves = [];
    people.forEach((person, pi) => {
      const py = personY(pi);
      const x1 = LEFT_W;
      person.projects.forEach(projId => {
        const projIdx = activeProjects.findIndex(p => p.id === projId);
        if (projIdx < 0) return;
        const pjy = projY(projIdx);
        const x2 = LEFT_W + MID_GAP;
        const mx = (x1 + x2) / 2;
        const color = projColorMap[projId] || "#3b6bcc";
        curves.push(`<path d="M${x1},${py} C${mx},${py} ${mx},${pjy} ${x2},${pjy}" fill="none" stroke="${color}" stroke-width="1.5" opacity="0.45"/>`);
      });
    });

    // Person nodes
    const personNodes = people.map((person, i) => {
      const cy = personY(i);
      const m = members.find(mb => mb.name === person.name);
      const initials = person.name.split(" ").map(w => w[0] || "").slice(0, 2).join("").toUpperCase();
      const taskLabel = person.taskCount > 0 ? ` · ${person.taskCount} task${person.taskCount !== 1 ? "s" : ""}` : "";
      return `<g class="ov-res-person-node" data-person-name="${escapeHtml(person.name)}">
        <circle cx="26" cy="${cy}" r="18" fill="var(--aewttr-blue)" opacity="0.15"/>
        <text x="26" y="${cy + 5}" text-anchor="middle" font-size="11" font-weight="600" fill="var(--aewttr-blue)">${escapeHtml(initials)}</text>
        <text x="52" y="${cy - 4}" font-size="12" font-weight="600" fill="var(--aewttr-text)">${escapeHtml(person.name)}</text>
        <text x="52" y="${cy + 10}" font-size="10" fill="var(--aewttr-muted)">${person.projects.size} project${person.projects.size !== 1 ? "s" : ""}${taskLabel}</text>
      </g>`;
    });

    // Project nodes
    const projNodes = activeProjects.map((proj, i) => {
      const cy = projY(i);
      const color = projColorMap[proj.id];
      const tasks = allProjectTasks(proj.id).filter(taskIsActive).length;
      const risks = projectRiskCount(proj.id);
      const x = LEFT_W + MID_GAP;
      return `<g class="ov-res-proj-node" data-proj-id="${escapeHtml(proj.id)}" style="cursor:pointer;">
        <rect x="${x}" y="${cy - 16}" width="10" height="32" rx="3" fill="${color}" opacity="0.8"/>
        <text x="${x + 18}" y="${cy - 3}" font-size="12" font-weight="600" fill="var(--aewttr-text)">${escapeHtml((proj.name || "Untitled project").substring(0, 28))}</text>
        <text x="${x + 18}" y="${cy + 11}" font-size="10" fill="var(--aewttr-muted)">${tasks} active${risks ? ` · ${risks} risk${risks !== 1 ? "s" : ""}` : ""}</text>
      </g>`;
    });

    return `<div class="ov-section-head" style="margin-bottom:16px;">
        <div>
          <div class="side-panel-title">Team Resources</div>
          <div class="overview-panel-sub">How the team is spread across projects. Lines connect people to their assigned projects.</div>
        </div>
      </div>
      <div class="ov-resources-graph-wrap">
        <div class="ov-resources-legend">
          ${activeProjects.map(p => `<span class="ov-res-legend-item"><span class="ov-res-legend-dot" style="background:${projColorMap[p.id]}"></span>${escapeHtml(p.name || "Untitled project")}</span>`).join("")}
        </div>
        <div class="ov-resources-labels">
          <span>Team Members</span>
          <span>Projects</span>
        </div>
        <div class="ov-resources-svg-wrap" style="overflow-x:auto;">
          <svg viewBox="0 0 ${svgW} ${svgH}" width="${svgW}" height="${svgH}" xmlns="http://www.w3.org/2000/svg" class="ov-resources-svg">
            ${curves.join("")}
            ${personNodes.join("")}
            ${projNodes.join("")}
          </svg>
        </div>
      </div>`;
  }

  /* Per-person activity outside the project graph: trips taken, leave taken,
     requests raised, and documents actually reviewed. The graph above answers
     "who is on what"; this answers "what has each person been doing".

     Matching is by email where both sides have one and by name otherwise —
     travel records carry a traveler list with optional emails, and reviewer
     records carry name and email, so neither key is reliably present. */
  function renderTeamActivityTable(people) {
    /* Rows here come from project assignments and task assignees, which carry a
       display name only. Travel and reviewer records may key on either, so look
       up an email from the member list where one exists. */
    const members = db.members || [];
    const withEmail = (people || []).map((row) => {
      const match = members.find((m) => String(m.name || "").trim().toLowerCase() === String(row.name || "").trim().toLowerCase());
      return { name: row.name, email: (match && match.email) || row.email || "" };
    });
    people = withEmail;
    const travel = db.travelRequests || [];
    const docs = typeof getAllDocReviewRecords === "function" ? getAllDocReviewRecords() : [];
    const key = (value) => String(value || "").trim().toLowerCase();

    function matchesPerson(person, name, email) {
      const personEmail = key(person.email);
      if (personEmail && key(email) === personEmail) return true;
      return key(name) === key(person.name);
    }

    /* A person "went on" a trip if they are listed as a traveller. The
       requester is counted separately, since raising a request for someone else
       is not the same as travelling. */
    function isTravellerOn(request, person) {
      const travellers = request.travelers && request.travelers.length
        ? request.travelers
        : [{ name: request.requester, email: request.requesterEmail }];
      return travellers.some((t) => t && matchesPerson(person, t.name, t.email));
    }

    const COUNTED_AS_TAKEN = ["Approved", "Completed"];

    const rows = people.map((person) => {
      const mine = travel.filter((r) => isTravellerOn(r, person));
      const isLeave = (r) => (typeof travelCategory === "function" ? travelCategory(r) === "Leave" : false);
      const taken = (r) => COUNTED_AS_TAKEN.includes(r.status);
      return {
        name: person.name,
        trips: mine.filter((r) => !isLeave(r) && taken(r)).length,
        leave: mine.filter((r) => isLeave(r) && taken(r)).length,
        submitted: travel.filter((r) => matchesPerson(person, r.requester, r.requesterEmail)).length,
        reviewed: docs.reduce((sum, doc) => sum + ((doc.reviewers || []).some((rv) =>
          matchesPerson(person, rv.name, rv.email) && rv.decision && rv.decision !== "Pending") ? 1 : 0), 0),
      };
    }).sort((a, b) => (b.trips + b.leave + b.submitted + b.reviewed) - (a.trips + a.leave + a.submitted + a.reviewed)
      || a.name.localeCompare(b.name));

    const body = rows.length
      ? rows.map((r) => `<tr data-person-name="${escapeHtml(r.name)}">
          <td><strong>${escapeHtml(r.name)}</strong></td>
          <td class="ov-activity-num">${r.trips || "—"}</td>
          <td class="ov-activity-num">${r.leave || "—"}</td>
          <td class="ov-activity-num">${r.submitted || "—"}</td>
          <td class="ov-activity-num">${r.reviewed || "—"}</td>
        </tr>`).join("")
      : `<tr><td colspan="5"><div class="empty-state">No team members yet.</div></td></tr>`;

    return `<div class="ov-section-head" style="margin:26px 0 12px;">
        <div>
          <div class="side-panel-title">Activity by person</div>
          <div class="overview-panel-sub">Trips and leave counted once approved or completed. Requests counts what each person raised, including for others. Reviews counts documents they have actually decided on, not ones still awaiting them.</div>
        </div>
      </div>
      <div class="ov-activity-scroll">
        <table class="aewttr-table ov-activity-tbl">
          <thead>
            <tr>
              <th>Person</th>
              <th class="ov-activity-num">Trips</th>
              <th class="ov-activity-num">Leave</th>
              <th class="ov-activity-num">Requests raised</th>
              <th class="ov-activity-num">Docs reviewed</th>
            </tr>
          </thead>
          <tbody>${body}</tbody>
        </table>
      </div>`;
  }

  function renderTeamResourcesHub(projects) {
    const activeTasks = getAllTasksFlat().filter(taskIsActive);
    const state = window.AEWTTR.state.teamResources || { view: "people", query: "", expanded: {} };
    if (!state.expanded) state.expanded = {};
    window.AEWTTR.state.teamResources = state;
    const people = new Map();
    const projectIndex = new Map(projects.map((project) => [project.id, project]));
    const projectStats = new Map(projects.map((project) => [project.id, {
      id: project.id, name: project.name || project.id, people: new Set(), tasks: 0
    }]));
    const person = (name) => {
      const key = String(name || "Unassigned").trim() || "Unassigned";
      if (!people.has(key)) people.set(key, { name: key, projects: new Set(), tasks: 0, roles: new Set() });
      return people.get(key);
    };
    projects.forEach((project) => projectPeopleList(project.id).forEach((entry) => {
      if (!entry.label) return;
      const row = person(entry.label);
      row.projects.add(project.id);
      if (entry.role || entry.type) row.roles.add(entry.role || entry.type);
      projectStats.get(project.id).people.add(row.name);
    }));
    activeTasks.forEach((task) => {
      const project = projectStats.get(task.projectId);
      if (project) { project.tasks++; }
      if (!task.assignee || task.assignee === "Unassigned") return;
      const row = person(task.assignee);
      row.projects.add(task.projectId);
      row.tasks++;
      if (project) project.people.add(row.name);
    });
    const peopleRows = [...people.values()].filter((row) => row.projects.size).sort((a, b) => b.tasks - a.tasks || a.name.localeCompare(b.name));
    const projectRows = [...projectStats.values()].filter((row) => row.people.size || row.tasks).sort((a, b) => b.tasks - a.tasks || a.name.localeCompare(b.name));
    window.AEWTTR.state.resourceDirectory = { people: peopleRows, projects: projectRows, projectIndex };
    const query = String(state.query || "").toLowerCase().trim();
    const matchesPerson = (row) => !query || [row.name, ...[...row.projects].map((id) => (projectIndex.get(id) || {}).name || id)].join(" ").toLowerCase().includes(query);
    const matchesProject = (row) => !query || [row.name, ...row.people].join(" ").toLowerCase().includes(query);
    const personList = peopleRows.filter(matchesPerson);
    const projectList = projectRows.filter(matchesProject);
    function personDetailHtml(row) {
      const links = [...row.projects].map((id) => {
        const project = projectIndex.get(id);
        return `<button type="button" class="resource-detail-link" data-route="projects/${escapeHtml(id)}/workspace"><strong>${escapeHtml((project && project.name) || id)}</strong><span>Open project workspace</span><i class="bx bx-chevron-right"></i></button>`;
      }).join("");
      return `<div class="resource-detail-inline">
        <div class="resource-detail-metrics"><span><b>${row.tasks}</b> active tasks</span><span><b>${row.projects.size}</b> projects</span></div>
        <h4>Project assignments</h4>${links || `<div class="empty-state">No assignments yet.</div>`}
      </div>`;
    }
    function projectDetailHtml(row) {
      const names = [...row.people].map((name) => `<div class="resource-detail-person"><span>${userAvatarHtml(name, "", 30)}</span>${escapeHtml(name)}</div>`).join("");
      return `<div class="resource-detail-inline">
        <div class="resource-detail-metrics"><span><b>${row.tasks}</b> active tasks</span><span><b>${row.people.size}</b> assigned</span></div>
        <h4>Assigned people</h4><div class="resource-detail-people">${names || `<div class="empty-state">No assignments yet.</div>`}</div>
      </div>`;
    }
    const peopleHtml = personList.map((row) => {
      const key = `person:${row.name}`;
      const open = !!state.expanded[key];
      return `<div class="resource-row-wrap">
        <button type="button" class="resource-person-row${open ? " is-expanded" : ""}" data-resource-toggle="${escapeHtml(key)}">
          ${userAvatarHtml(row.name, "", 40)}<span class="resource-row-copy"><strong>${escapeHtml(row.name)}</strong><small>${row.tasks} active · ${row.projects.size} project${row.projects.size === 1 ? "" : "s"}</small></span>
          <span class="resource-row-projects">${[...row.projects].slice(0, 3).map((id) => `<em>${escapeHtml((projectIndex.get(id) || {}).name || id)}</em>`).join("")}</span>
          <i class="bx bx-chevron-${open ? "up" : "right"}"></i>
        </button>
        ${open ? personDetailHtml(row) : ""}
      </div>`;
    }).join("") || `<div class="empty-state">No people match this search.</div>`;
    const projectsHtml = projectList.map((row) => {
      const key = `project:${row.id}`;
      const open = !!state.expanded[key];
      return `<div class="resource-row-wrap">
        <button type="button" class="resource-project-row${open ? " is-expanded" : ""}" data-resource-toggle="${escapeHtml(key)}">
          <span class="resource-project-mark"></span><span class="resource-row-copy"><strong>${escapeHtml(row.name)}</strong><small>${row.people.size} assigned · ${row.tasks} active</small></span>
          <span class="resource-avatar-stack">${[...row.people].slice(0, 4).map((name) => userAvatarHtml(name, "", 28)).join("")}</span>
          <i class="bx bx-chevron-${open ? "up" : "right"}"></i>
        </button>
        ${open ? projectDetailHtml(row) : ""}
      </div>`;
    }).join("") || `<div class="empty-state">No projects match this search.</div>`;
    const PROJ_COLORS = ["#3b6bcc","#e05f2b","#2b9e6a","#9b59b6","#c0392b","#16a085","#d35400","#2980b9","#8e44ad","#27ae60","#e74c3c","#1abc9c","#f39c12","#6c5ce7","#00b894"];
    const projColorMap = {};
    projectList.forEach((row, i) => { projColorMap[row.id] = PROJ_COLORS[i % PROJ_COLORS.length]; });
    // A wide, generous design canvas — the <svg> below has no fixed pixel
    // width/height of its own, so this coordinate space just sets the
    // aspect ratio; CSS stretches it to fill whatever room the card
    // actually has (previously a hardcoded 620px canvas sat inside an
    // ~1100px+ container, wasting most of the width and reading tiny).
    const PERSON_H = 64; const PROJ_H = 56; const LEFT_W = 260; const RIGHT_W = 260; const MID_GAP = 320; const PAD_Y = 32;
    const svgW = LEFT_W + MID_GAP + RIGHT_W;
    const svgH = Math.max(personList.length * PERSON_H, projectList.length * PROJ_H, 140) + PAD_Y * 2;
    const pY = (i) => PAD_Y + i * PERSON_H + PERSON_H / 2;
    const rY = (i) => PAD_Y + i * PROJ_H + PROJ_H / 2;
    const curves = [];
    personList.forEach((row, pi) => {
      [...row.projects].forEach(projId => {
        const ri = projectList.findIndex(r => r.id === projId);
        if (ri < 0) return;
        const x1 = LEFT_W, x2 = LEFT_W + MID_GAP, mx = (x1 + x2) / 2;
        const color = projColorMap[projId] || "#3b6bcc";
        curves.push(`<path d="M${x1},${pY(pi)} C${mx},${pY(pi)} ${mx},${rY(ri)} ${x2},${rY(ri)}" fill="none" stroke="${color}" stroke-width="2" opacity="0.4"/>`);
      });
    });
    const personNodes = personList.map((row, i) => {
      const cy = pY(i);
      const initials = row.name.split(" ").map(w => w[0]||"").slice(0,2).join("").toUpperCase();
      return `<g class="ov-res-person-node" data-resource-person="${escapeHtml(row.name)}" style="cursor:pointer;">
        <circle cx="26" cy="${cy}" r="21" fill="var(--aewttr-blue)" opacity="0.12"/>
        <text x="26" y="${cy+5}" text-anchor="middle" font-size="12.5" font-weight="700" fill="var(--aewttr-blue)">${escapeHtml(initials)}</text>
        <text x="58" y="${cy-5}" font-size="14.5" font-weight="600" fill="var(--aewttr-text)">${escapeHtml(row.name.length > 26 ? row.name.slice(0,25)+"…" : row.name)}</text>
        <text x="58" y="${cy+13}" font-size="12" fill="var(--aewttr-muted)">${row.projects.size} project${row.projects.size!==1?"s":""} · ${row.tasks} task${row.tasks!==1?"s":""}</text>
      </g>`;
    });
    const projNodes = projectList.map((row, i) => {
      const cy = rY(i); const color = projColorMap[row.id]; const x = LEFT_W + MID_GAP;
      return `<g class="ov-res-proj-node" data-resource-project="${escapeHtml(row.id)}" style="cursor:pointer;">
        <rect x="${x}" y="${cy-17}" width="10" height="34" rx="3" fill="${color}" opacity="0.85"/>
        <text x="${x+20}" y="${cy-4}" font-size="14.5" font-weight="600" fill="var(--aewttr-text)">${escapeHtml((row.name||row.id).slice(0,34))}</text>
        <text x="${x+20}" y="${cy+14}" font-size="12" fill="var(--aewttr-muted)">${row.people.size} assigned · ${row.tasks} active</text>
      </g>`;
    });
    const legendHtml = projectList.map(row => `<span class="ov-res-legend-item"><span class="ov-res-legend-dot" style="background:${projColorMap[row.id]}"></span>${escapeHtml(row.name||row.id)}</span>`).join("");
    const mapHtml = `<div class="ov-resources-graph-wrap">
      ${legendHtml ? `<div class="ov-resources-legend">${legendHtml}</div>` : ""}
      <div class="ov-resources-labels"><span>Team Members</span><span>Projects</span></div>
      <div class="ov-resources-svg-wrap">
        <svg viewBox="0 0 ${svgW} ${svgH}" xmlns="http://www.w3.org/2000/svg" class="ov-resources-svg">
          ${curves.join("")}${personNodes.join("")}${projNodes.join("")}
        </svg>
      </div>
    </div>`;
    const content = state.view === "projects" ? projectsHtml : state.view === "map" ? mapHtml : peopleHtml;
    return `<section class="resource-hub"><header class="resource-hub-head"><div><h3>Team Resources</h3><p>Browse staffing, active workload, and project assignments.</p></div><div class="resource-summary"><span><b>${peopleRows.length}</b> people</span><span><b>${projectRows.length}</b> projects</span><span><b>${activeTasks.length}</b> active tasks</span></div></header>
      <div class="resource-controls"><div class="resource-view-tabs">${[["people","People","bx-user"],["projects","Projects","bx-briefcase"],["map","Connections","bx-share-alt"]].map(([key,label,icon]) => `<button type="button" class="${state.view === key ? "is-active" : ""}" data-resource-view="${key}"><i class="bx ${icon}"></i>${label}</button>`).join("")}</div><label class="resource-search"><i class="bx bx-search"></i><input id="resource-search" value="${escapeHtml(state.query || "")}" placeholder="Search people or projects"></label></div>
      <div class="resource-scroll-list">${content}</div>
      ${renderTeamActivityTable(peopleRows)}</section>`;
  }

  function renderAllProjectsTable(projects) {
    const f = window.AEWTTR.state.overviewFilters || {};
    const filtered = projects.filter(p => matchesFilter(p, f));

    return `<div class="ov-section-head" style="margin-top:24px;">
        <div class="side-panel-title">All Projects</div>
        <div class="overview-panel-sub">Complete project inventory with derived status, task counts, funding, and key dates.</div>
      </div>
      <div class="overview-table-shell">
        <table class="aewttr-table overview-table">
          <thead>
            <tr><th>Project</th><th>End Item</th><th>Portfolio</th><th>PM</th><th>Team</th><th>Status</th><th>Active</th><th>Risks</th><th>Funding</th><th>FY</th><th>Task Order</th><th>Start</th><th>Due</th></tr>
          </thead>
          <tbody>
            ${filtered.length ? filtered.map(proj => {
              const tasks = allProjectTasks(proj.id);
              const active = tasks.filter(taskIsActive).length;
              const risks = projectRiskCount(proj.id);
              const portfolio = (proj.portfolios && proj.portfolios[0]) || "—";
              return `<tr data-route="projects/${escapeHtml(proj.id)}/workspace">
                <td><strong>${escapeHtml(proj.name || "Untitled project")}</strong></td>
                <td>${escapeHtml(projectEndItem(proj))}</td>
                <td>${escapeHtml(portfolio)}</td>
                <td>${escapeHtml(projectPMDisplayName(proj) || "—")}</td>
                <td>${escapeHtml(proj.team || "—")}</td>
                <td>${derivedStatusPill(derivedProjectStatus(proj))}</td>
                <td>${active}</td>
                <td><span class="${risks > 0 ? "ov-risks-badge" : ""}">${risks || "—"}</span></td>
                <td>${escapeHtml(proj.fundingType || "—")}</td>
                <td>${escapeHtml(proj.fiscalYear ? `FY${proj.fiscalYear}` : "—")}</td>
                <td>${escapeHtml(proj.taskOrder || "—")}</td>
                <td>${proj.startDate ? fmtDate(proj.startDate) : "—"}</td>
                <td>${proj.dueDate ? fmtDate(proj.dueDate) : "—"}</td>
              </tr>`;
            }).join("") : `<tr><td colspan="13"><div class="empty-state" style="padding:20px 0;">No projects match these filters.</div></td></tr>`}
          </tbody>
        </table>
      </div>`;
  }

  function renderOperationsQueues() {
    const docs = typeof getAllDocReviewRecords === "function" ? getAllDocReviewRecords() : [];
    const docsInReview = docs
      .filter(d => !d.isArchived && !["Review Complete", "Signed", "Archived"].includes(d._column || ""))
      .slice(0, 8);
    const pendingTravel = (db.travelRequests || [])
      .filter(r => r.status === "Pending" || r.status === "Submitted")
      .slice(0, 8);

    return `<div class="ov-section-head" style="margin-bottom:16px;">
        <div class="side-panel-title">Documents &amp; Travel</div>
        <div class="overview-panel-sub">Active review queue and pending travel requests. Click any row to open the full workflow.</div>
      </div>
      <div class="overview-mini-grid">
        <div>
          <div class="overview-mini-title">Document Review Queue <span class="ov-queue-count">${docsInReview.length}</span></div>
          <div class="overview-list-table">
            ${docsInReview.length
              ? docsInReview.map(doc => `<button class="overview-list-row" data-route="docreview">
                  <div><strong>${escapeHtml(doc.title || "Untitled")}</strong><span>${escapeHtml(doc._column || "Not Started")} · ${escapeHtml(doc.submitter || "Unknown")}</span></div>
                </button>`).join("")
              : `<div class="empty-state">No documents in review.</div>`}
          </div>
          <div style="margin-top:10px;"><button class="btn-aewttr-outline btn-aewttr-sm" data-route="docreview">Open Document Review</button></div>
        </div>
        <div>
          <div class="overview-mini-title">Pending Travel Requests <span class="ov-queue-count">${pendingTravel.length}</span></div>
          <div class="overview-list-table">
            ${pendingTravel.length
              ? pendingTravel.map(r => `<div class="overview-list-row">
                  <div><strong>${escapeHtml(r.id || "Request")}</strong><span>${escapeHtml(r.requester || "Unknown")} · ${escapeHtml(r.destination || "No destination")}</span></div>
                  <span class="status-pill">${escapeHtml(r.status)}</span>
                </div>`).join("")
              : `<div class="empty-state">No pending travel requests.</div>`}
          </div>
          <div style="margin-top:10px;"><button class="btn-aewttr-outline btn-aewttr-sm" data-route="travel">Open Travel</button></div>
        </div>
      </div>`;
  }

  function buildResourceReportData(projects) {
    const activeTasks = getAllTasksFlat().filter(taskIsActive);
    const projectIndex = new Map(projects.map(project => [project.id, project]));
    const projectRows = new Map(projects.map(project => [project.id, {
      id: project.id,
      name: project.name || project.id,
      people: new Set(),
      tasks: 0
    }]));
    const peopleRows = new Map();
    const ensurePerson = (name) => {
      const key = String(name || "Unassigned").trim() || "Unassigned";
      if (!peopleRows.has(key)) peopleRows.set(key, { name: key, projects: new Set(), tasks: 0, roles: new Set() });
      return peopleRows.get(key);
    };

    projects.forEach(project => projectPeopleList(project.id).forEach(entry => {
      if (!entry.label) return;
      const person = ensurePerson(entry.label);
      person.projects.add(project.id);
      if (entry.role || entry.type) person.roles.add(entry.role || entry.type);
      const projectRow = projectRows.get(project.id);
      if (projectRow) projectRow.people.add(person.name);
    }));
    activeTasks.forEach(task => {
      const project = projectRows.get(task.projectId);
      if (project) project.tasks++;
      if (!task.assignee || task.assignee === "Unassigned") return;
      const person = ensurePerson(task.assignee);
      person.projects.add(task.projectId);
      person.tasks++;
      if (project) project.people.add(person.name);
    });

    const query = String((window.AEWTTR.state.teamResources || {}).query || "").toLowerCase().trim();
    const people = [...peopleRows.values()]
      .filter(row => row.projects.size)
      .filter(row => !query || [row.name, ...[...row.projects].map(id => (projectIndex.get(id) || {}).name || id)].join(" ").toLowerCase().includes(query))
      .sort((a, b) => b.tasks - a.tasks || a.name.localeCompare(b.name));
    const resourceProjects = [...projectRows.values()]
      .filter(row => row.people.size || row.tasks)
      .filter(row => !query || [row.name, ...row.people].join(" ").toLowerCase().includes(query))
      .sort((a, b) => b.tasks - a.tasks || a.name.localeCompare(b.name));
    return { people, projects: resourceProjects, projectIndex, query };
  }

  function buildMyOverviewReportModel() {
    const user = db.user || {};
    const data = getMyData();
    const activeTasks = data.allMyTasks.filter(taskIsActive);
    const taskFilter = window.AEWTTR.state.overviewMyTasksFilter || "All";
    const filteredMap = {
      "All": data.allMyTasks,
      "Open": activeTasks,
      "No Date": data.allMyTasks.filter(task => taskIsActive(task) && !task.end),
      "Completed": data.allMyTasks.filter(taskIsDone)
    };
    const filteredTasks = [...(filteredMap[taskFilter] || data.allMyTasks)].sort((a, b) => {
      if (taskIsDone(a) !== taskIsDone(b)) return taskIsDone(a) ? 1 : -1;
      if (!!a.end !== !!b.end) return a.end ? -1 : 1;
      return String(a.end || "").localeCompare(String(b.end || ""));
    });
    const attentionRows = [
      ...data.docsNeedingMe.map(doc => [
        "Document Review",
        doc.title || "Untitled",
        doc._column || "Pending",
        doc.submitter || doc.owner || "-"
      ]),
      ...data.myTravel.map(request => [
        "Travel",
        request.id || request.title || "Request",
        request.status || "Pending",
        request.destination || "-"
      ])
    ];
    const sections = [];
    if (attentionRows.length) {
      sections.push({
        title: "Needs Attention",
        note: "Items currently waiting on the report owner.",
        columns: [
          { label: "Workflow", weight: 1 },
          { label: "Item", weight: 2.2 },
          { label: "Status", weight: 1 },
          { label: "Owner / Destination", weight: 1.5 }
        ],
        rows: attentionRows
      });
    }
    sections.push({
      title: `My Tasks - ${taskFilter}`,
      note: `The task table reflects the active "${taskFilter}" filter from My Overview.`,
      columns: [
        { label: "Task", weight: 2.4 },
        { label: "Project", weight: 1.8 },
        { label: "Due", weight: 0.9 },
        { label: "Priority", weight: 0.8 },
        { label: "Status", weight: 1.1 }
      ],
      rows: filteredTasks.map(task => [
        task.title || "Untitled",
        task.projectName || task.projectId || "-",
        task.end ? fmtDate(task.end) : "-",
        task.priority || "-",
        task.status || "Not Started"
      ]),
      emptyMessage: "No tasks match the active filter."
    });
    sections.push({
      title: "My Projects",
      note: "Projects where the report owner holds an assigned role.",
      columns: [
        { label: "Project", weight: 2.1 },
        { label: "Role", weight: 1.1 },
        { label: "Team", weight: 1.1 },
        { label: "Status", weight: 1.2 },
        { label: "Open Tasks", weight: 0.7 },
        { label: "Open Risks", weight: 0.7 },
        { label: "Next Due", weight: 0.9 }
      ],
      rows: data.myProjects.map(project => {
        const tasks = allProjectTasks(project.id);
        return [
          project.name || project.id,
          myRoleOnProject(project) || "Member",
          project.team || "-",
          derivedProjectStatus(project),
          tasks.filter(taskIsActive).length,
          projectRiskCount(project.id),
          projectNextDue(project) ? fmtDate(projectNextDue(project)) : "-"
        ];
      }),
      emptyMessage: "No assigned projects."
    });
    return {
      reportLabel: "Personal Overview Report",
      scopeLabel: "My Overview",
      storageFolder: "Personal",
      fileBaseName: "PULSE-My-Overview",
      title: "Your work at a glance",
      subtitle: `A complete snapshot of the active personal view for ${user.name || "the current user"}. Task filter: ${taskFilter}.`,
      generatedAt: asOfLabel(),
      generatedBy: user.name || "",
      summary: [
        { label: "Tasks Completed", value: data.tasksCompletedThisYear.length, detail: data.thisYear, tone: "accent" },
        { label: "Docs Reviewed", value: data.docsReviewedThisYear.length, detail: data.thisYear },
        { label: "Docs Signed", value: data.docsSignedByMe.length, detail: data.thisYear },
        { label: "Meetings", value: data.meetingsAttended.length, detail: `${data.upcomingMeetings} upcoming` },
        { label: "Travel Requests", value: data.travelThisYear.length, detail: data.thisYear },
        { label: "Projects", value: data.myProjects.length, detail: "assigned" }
      ],
      sections
    };
  }

  function buildTeamOverviewReportModel(projects) {
    const user = db.user || {};
    const metrics = buildTeamMetrics(projects);
    const activeTab = window.AEWTTR.state.overviewTeamTab || "portfolio";
    const tabLabel = (teamTabs.find(tab => tab.key === activeTab) || {}).label || "Portfolio";
    const sections = [];

    if (activeTab === "dashboard") {
      sections.push({
        title: "Portfolio Health Summary",
        note: "Dashboard view: projects by status with open and completed task counts.",
        columns: [
          { label: "Project", weight: 2 },
          { label: "Status", weight: 1 },
          { label: "Open Tasks", weight: 0.9 },
          { label: "Completed", weight: 0.9 },
          { label: "% Done", weight: 0.7 },
          { label: "Risks", weight: 0.7 },
          { label: "PM", weight: 1.5 }
        ],
        rows: projects.map(proj => {
          const tasks = allProjectTasks(proj.id);
          const open = tasks.filter(taskIsActive).length;
          const done = tasks.filter(taskIsDone).length;
          const pct = (open + done) > 0 ? Math.round((done / (open + done)) * 100) + "%" : "—";
          const risks = projectRiskCount(proj.id);
          const status = derivedProjectStatus(proj);
          return [proj.name || proj.id, status, open, done, pct, risks, projectPMDisplayName(proj) || "—"];
        }).sort((a, b) => {
          const w = { "Needs Attention": 0, Active: 1, "No Active Work": 2, Completed: 3 };
          return (w[a[1]] ?? 9) - (w[b[1]] ?? 9);
        }),
        emptyMessage: "No projects in portfolio."
      });
    } else if (activeTab === "risks") {
      const allRisks = [];
      projects.forEach(proj => {
        const extra = (db.projectExtra && db.projectExtra[proj.id]) || {};
        (extra.risks || []).forEach(r => {
          if (isPlaceholderRisk(r)) return;
          const s = String(r.status || "").toLowerCase();
          if (["closed", "resolved", "accepted"].includes(s)) return;
          allRisks.push({ ...r, projectName: proj.name || proj.id });
        });
      });
      sections.push({
        title: "Open Risk Register",
        note: `${allRisks.length} open risks across the portfolio. Closed, resolved, and accepted risks excluded.`,
        columns: [
          { label: "Risk", weight: 2 },
          { label: "Project", weight: 1.5 },
          { label: "Likelihood", weight: 0.9 },
          { label: "Impact", weight: 0.9 },
          { label: "Owner", weight: 1.2 },
          { label: "Status", weight: 0.9 }
        ],
        rows: allRisks.map(r => [
          r.title || r.description || "Risk",
          r.projectName,
          r.likelihood || "-",
          r.impact || "-",
          r.owner || r.riskOwner || "-",
          r.status || "-"
        ]),
        emptyMessage: "No open risks across the portfolio."
      });
    } else if (activeTab === "portfolio") {
      const filters = window.AEWTTR.state.overviewFilters || {};
      const filteredProjects = projects.filter(project => matchesFilter(project, filters));
      const groupMap = {};
      filteredProjects.forEach(project => {
        const name = projectEndItem(project);
        if (!groupMap[name]) groupMap[name] = [];
        groupMap[name].push(project);
      });
      const activeFilters = [
        filters.search ? `Search: ${filters.search}` : "",
        filters.team && filters.team !== "All" ? `Team: ${filters.team}` : "",
        filters.status && filters.status !== "All" ? `Status: ${filters.status}` : "",
        filters.fundingType && filters.fundingType !== "All" ? `Funding: ${filters.fundingType}` : "",
        filters.portfolio && filters.portfolio !== "All" ? `Portfolio: ${filters.portfolio}` : "",
        filters.endItem && filters.endItem !== "All" ? `End Item: ${filters.endItem}` : "",
        filters.ato && filters.ato !== "All" ? `ATO: ${filters.ato}` : ""
      ].filter(Boolean);
      sections.push({
        title: "Portfolio by End Item",
        note: activeFilters.length ? `Active filters - ${activeFilters.join("; ")}` : "No portfolio filters are active.",
        columns: [
          { label: "End Item", weight: 2.2 },
          { label: "Projects", weight: 0.8 },
          { label: "Active Tasks", weight: 0.9 },
          { label: "Open Risks", weight: 0.8 },
          { label: "Teams", weight: 2 }
        ],
        rows: Object.entries(groupMap).sort(([a], [b]) => a.localeCompare(b)).map(([name, rows]) => [
          name,
          rows.length,
          rows.flatMap(project => allProjectTasks(project.id)).filter(taskIsActive).length,
          rows.reduce((sum, project) => sum + projectRiskCount(project.id), 0),
          [...new Set(rows.map(project => project.team).filter(Boolean))].join(", ") || "-"
        ]),
        emptyMessage: "No end items match the active filters."
      });
      sections.push({
        title: "Project Inventory",
        note: "All projects visible under the current portfolio filters.",
        columns: [
          { label: "Project", weight: 1.8 },
          { label: "End Item", weight: 1.3 },
          { label: "PM", weight: 1.2 },
          { label: "Team", weight: 0.9 },
          { label: "Status", weight: 1 },
          { label: "Active", weight: 0.55 },
          { label: "Risks", weight: 0.55 },
          { label: "Funding", weight: 0.9 },
          { label: "FY", weight: 0.45 },
          { label: "Due", weight: 0.8 }
        ],
        rows: filteredProjects.map(project => [
          project.name || project.id,
          projectEndItem(project),
          projectPMDisplayName(project) || "-",
          project.team || "-",
          derivedProjectStatus(project),
          allProjectTasks(project.id).filter(taskIsActive).length,
          projectRiskCount(project.id),
          project.fundingType || "-",
          project.fiscalYear ? `FY${project.fiscalYear}` : "-",
          project.dueDate ? fmtDate(project.dueDate) : "-"
        ]),
        emptyMessage: "No projects match the active filters."
      });
    } else if (activeTab === "workload") {
      const workloadRows = buildTeamWorkloadRows(projects);
      const sortLabel = ({ active: "Most Active", projects: "Most Projects", name: "Name" })[window.AEWTTR.state.overviewWorkloadSort || "active"];
      sections.push({
        title: "Team Project Workload",
        note: `Sorted by ${sortLabel}. Includes project-related assignments shown in the active view.`,
        columns: [
          { label: "Person", weight: 1.6 },
          { label: "Open", weight: 0.55 },
          { label: "Completed", weight: 0.65 },
          { label: "Projects", weight: 0.6 },
          { label: "Roles", weight: 1.3 },
          { label: "Teams", weight: 1.2 },
          { label: "Next Due", weight: 0.8 }
        ],
        rows: workloadRows.map(entry => [
          entry.member.name,
          entry.activeTasks,
          entry.completedTasks,
          entry.projects.size,
          [...entry.roles].join(", ") || "-",
          [...entry.teams].join(", ") || "-",
          entry.nextDue ? fmtDate(entry.nextDue) : "-"
        ]),
        emptyMessage: "No active project workload data."
      });
    } else if (activeTab === "resources") {
      const resourceData = buildResourceReportData(projects);
      const resourceView = (window.AEWTTR.state.teamResources || {}).view || "people";
      if (resourceView === "projects") {
        sections.push({
          title: "Resources by Project",
          note: resourceData.query ? `Search: ${resourceData.query}` : "All project resource assignments in the active view.",
          columns: [
            { label: "Project", weight: 2 },
            { label: "Assigned People", weight: 3 },
            { label: "People", weight: 0.7 },
            { label: "Active Tasks", weight: 0.8 }
          ],
          rows: resourceData.projects.map(row => [row.name, [...row.people].join(", ") || "-", row.people.size, row.tasks])
        });
      } else {
        sections.push({
          title: resourceView === "map" ? "Resource Connections" : "Resources by Person",
          note: resourceData.query ? `Search: ${resourceData.query}` : "All people-to-project assignments in the active view.",
          columns: [
            { label: "Person", weight: 1.5 },
            { label: "Projects", weight: 3 },
            { label: "Roles", weight: 1.5 },
            { label: "Project Count", weight: 0.8 },
            { label: "Active Tasks", weight: 0.8 }
          ],
          rows: resourceData.people.map(row => [
            row.name,
            [...row.projects].map(id => (resourceData.projectIndex.get(id) || {}).name || id).join(", ") || "-",
            [...row.roles].join(", ") || "-",
            row.projects.size,
            row.tasks
          ])
        });
      }
    } else {
      const docs = typeof getAllDocReviewRecords === "function" ? getAllDocReviewRecords() : [];
      const docsInReview = docs.filter(doc => !doc.isArchived && !["Review Complete", "Signed", "Archived"].includes(doc._column || "")).slice(0, 8);
      const pendingTravel = (db.travelRequests || []).filter(request => request.status === "Pending" || request.status === "Submitted").slice(0, 8);
      sections.push({
        title: "Document Review Queue",
        note: "The active documents shown in Operations.",
        columns: [
          { label: "Document", weight: 2.5 },
          { label: "Stage", weight: 1.2 },
          { label: "Submitter", weight: 1.5 },
          { label: "Updated", weight: 1 }
        ],
        rows: docsInReview.map(doc => [doc.title || "Untitled", doc._column || "Not Started", doc.submitter || "Unknown", doc.updatedAt ? fmtDate(doc.updatedAt) : "-"]),
        emptyMessage: "No documents in review."
      });
      sections.push({
        title: "Pending Travel Requests",
        note: "The active travel requests shown in Operations.",
        columns: [
          { label: "Request", weight: 1.2 },
          { label: "Requester", weight: 1.5 },
          { label: "Destination", weight: 1.8 },
          { label: "Dates", weight: 1.4 },
          { label: "Status", weight: 1 }
        ],
        rows: pendingTravel.map(request => [
          request.id || request.title || "Request",
          request.requester || "Unknown",
          request.destination || "-",
          [request.startDate ? fmtDate(request.startDate) : "", request.endDate ? fmtDate(request.endDate) : ""].filter(Boolean).join(" - ") || "-",
          request.status || "Pending"
        ]),
        emptyMessage: "No pending travel requests."
      });
    }

    return {
      reportLabel: "Team Overview Report",
      scopeLabel: `Team Overview - ${tabLabel}`,
      storageFolder: `Team-${activeTab}`,
      fileBaseName: `PULSE-Team-Overview-${activeTab}`,
      title: "Mission work at a glance",
      subtitle: `A complete snapshot of the active Team Overview ${tabLabel} view.`,
      generatedAt: asOfLabel(),
      generatedBy: user.name || "",
      summary: [
        { label: "Active Projects", value: metrics.activeProjects, detail: `${metrics.totalProjects} total`, tone: "accent" },
        { label: "Open Tasks", value: metrics.activeTasks },
        { label: "Completed Tasks", value: metrics.completedTasks, tone: "accent" },
        { label: "Open Risks", value: metrics.openRisks, tone: metrics.openRisks ? "warning" : "" },
        { label: "Docs in Review", value: metrics.docsInReview },
        { label: "Travel awaiting C/O", value: metrics.pendingTravel }
      ],
      sections
    };
  }

  /* ---- main render ---- */

  function render() {
    let view = window.AEWTTR.state.overviewView;
    if (!views.some(v => v.key === view)) { view = "My"; window.AEWTTR.state.overviewView = view; }

    const projects = db.projects || [];
    const bodyHtml = view === "Team" && canSeeTeamOverview
      ? renderTeamOverview(projects)
      : renderMyOverview();

    setTopbar(
      "Overview",
      view === "Team"
        ? `Team operations overview — as of ${asOfLabel()}.`
        : `Your work overview — as of ${asOfLabel()}.`,
      ""
    );

    $("#page-content").innerHTML = `
      <div class="overview-toolbar">
        <div class="filter-pills">
          ${views.map(v => `<button class="filter-pill${v.key === view ? " active" : ""}" data-overview-view="${v.key}">${v.label}</button>`).join("")}
        </div>
        <div class="overview-toolbar-actions">
          <button class="btn-aewttr overview-report-button" id="overview-generate-pptx" ${tip("Export this view as a comprehensive PowerPoint briefing deck")}><i class="bx bxs-slideshow"></i> Export PPTX Briefing</button>
          <button class="btn-aewttr-outline" id="overview-export-xlsx" ${tip("Export project metrics as a themed Excel workbook")}><i class="bx bx-spreadsheet"></i> Export Excel</button>
        </div>
      </div>
      <div id="overview-view-root">${bodyHtml}</div>
    `;

    /* navigation */
    $all("[data-overview-view]", $("#page-content")).forEach(btn => {
      btn.addEventListener("click", () => { window.AEWTTR.state.overviewView = btn.dataset.overviewView; render(); });
    });
    $all("[data-route]", $("#page-content")).forEach(btn => {
      btn.addEventListener("click", () => navigate(btn.dataset.route));
    });
    /* Needs Attention rows that point at a Team Overview tab switch in place
       rather than navigating away — the detail is already on this page. */
    $all("[data-ov-team-tab]", $("#page-content")).forEach(btn => {
      btn.addEventListener("click", () => {
        window.AEWTTR.state.overviewTeamTab = btn.dataset.ovTeamTab;
        render();
      });
    });
    $all("[data-person-id]", $("#page-content")).forEach(row => {
      row.addEventListener("click", () => navigate(`people/${row.dataset.personId}`));
    });

    /* My Tasks filter pills */
    $all("[data-ov-task-filter]", $("#page-content")).forEach(btn => {
      btn.addEventListener("click", () => { window.AEWTTR.state.overviewMyTasksFilter = btn.dataset.ovTaskFilter; render(); });
    });

    /* Workload sort */
    $all("[data-ov-wl-sort]", $("#page-content")).forEach(btn => {
      btn.addEventListener("click", () => { window.AEWTTR.state.overviewWorkloadSort = btn.dataset.ovWlSort; render(); });
    });

    /* Team inner tabs */
    $all("[data-team-tab]", $("#page-content")).forEach(btn => {
      btn.addEventListener("click", () => { window.AEWTTR.state.overviewTeamTab = btn.dataset.teamTab; render(); });
    });

    /* Resources graph: click project node to navigate */
    $all(".ov-res-proj-node", $("#page-content")).forEach(node => {
      node.addEventListener("click", () => {
        const id = node.dataset.projId;
        if (id) navigate(`projects/${projectRouteKeyById(id)}/workspace`);
      });
    });

    $all("[data-resource-view]", $("#page-content")).forEach((button) => {
      button.addEventListener("click", () => {
        window.AEWTTR.state.teamResources.view = button.dataset.resourceView;
        render();
      });
    });
    const resourceSearch = $("#resource-search", $("#page-content"));
    if (resourceSearch) resourceSearch.addEventListener("input", () => {
      window.AEWTTR.state.teamResources.query = resourceSearch.value;
      render();
    });
    /* Expand/collapse the row inline instead of opening a modal — assignment
       detail (project links / assigned people) reads directly in the list. */
    $all("[data-resource-toggle]", $("#page-content")).forEach((button) => button.addEventListener("click", () => {
      const key = button.dataset.resourceToggle;
      window.AEWTTR.state.teamResources.expanded[key] = !window.AEWTTR.state.teamResources.expanded[key];
      render();
    }));
    $all(".resource-detail-inline [data-route]", $("#page-content")).forEach((button) => button.addEventListener("click", (e) => {
      e.stopPropagation();
      navigate(button.dataset.route);
    }));

    /* Portfolio filters */
    const sysSearch = $("#ov-sys-search");
    if (sysSearch) {
      sysSearch.addEventListener("input", e => { window.AEWTTR.state.overviewFilters.search = e.target.value; });
      sysSearch.addEventListener("keydown", e => { if (e.key === "Enter") render(); });
      sysSearch.addEventListener("blur", () => render());
    }
    const sysTeam = $("#ov-sys-team");
    if (sysTeam) sysTeam.addEventListener("change", e => { window.AEWTTR.state.overviewFilters.team = e.target.value; render(); });
    const sysStatus = $("#ov-sys-status");
    if (sysStatus) sysStatus.addEventListener("change", e => { window.AEWTTR.state.overviewFilters.status = e.target.value; render(); });
    const sysFunding = $("#ov-sys-funding");
    if (sysFunding) sysFunding.addEventListener("change", e => { window.AEWTTR.state.overviewFilters.fundingType = e.target.value; render(); });
    const sysPortfolio = $("#ov-sys-portfolio");
    if (sysPortfolio) sysPortfolio.addEventListener("change", e => { window.AEWTTR.state.overviewFilters.portfolio = e.target.value; render(); });
    const sysEndItem = $("#ov-sys-enditem");
    if (sysEndItem) sysEndItem.addEventListener("change", e => { window.AEWTTR.state.overviewFilters.endItem = e.target.value; render(); });
    const sysAto = $("#ov-sys-ato");
    if (sysAto) sysAto.addEventListener("change", e => { window.AEWTTR.state.overviewFilters.ato = e.target.value; render(); });
    const sysClear = $("#ov-sys-clear");
    if (sysClear) sysClear.addEventListener("click", () => {
      window.AEWTTR.state.overviewFilters = { search: "", team: "All", status: "All", fundingType: "All", portfolio: "All", endItem: "All", ato: "All" };
      render();
    });

    /* PPTX briefing deck export */
    const pptxBtn = $("#overview-generate-pptx");
    if (pptxBtn) {
      pptxBtn.onclick = async () => {
        const service = window.AEWTTR && window.AEWTTR.OverviewPptxService;
        if (!service || typeof service.generate !== "function") {
          toast("The PPTX export service is unavailable. Refresh PULSE and try again.", "error");
          return;
        }
        const originalHtml = pptxBtn.innerHTML;
        const officeApi = window.AEWTTR && window.AEWTTR.OfficeDesktop;
        const openTarget = typeof isSharePointMode === "function" && isSharePointMode()
          && officeApi && typeof officeApi.reserveSharePointFileWindow === "function"
          ? officeApi.reserveSharePointFileWindow("PULSE-Overview.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
          : null;
        pptxBtn.disabled = true;
        pptxBtn.innerHTML = `<i class="bx bx-loader-alt bx-spin"></i> Building deck…`;
        try {
          const isTeam = view === "Team" && canSeeTeamOverview;
          const result = await service.generate(isTeam, window.AEWTTR.db, projects, { popup: openTarget });
          if (result && result.mode === "sharepoint" && result.fileUrl) {
            if (typeof toastStatusPptxSaved === "function") toastStatusPptxSaved(result.fileUrl, result.fileName);
            else toast("PPTX saved and opened in SharePoint.", "success");
          } else if (result && result.uploadError) {
            const detail = (result.uploadError && (result.uploadError.friendly || result.uploadError.message)) || "upload failed";
            toast("SharePoint upload failed (" + detail + ") — downloaded a local copy instead.", "warn");
          } else {
            toast("PPTX briefing deck downloaded.", "success");
          }
        } catch (error) {
          if (officeApi && typeof officeApi.closeReservedSharePointFileWindow === "function") {
            officeApi.closeReservedSharePointFileWindow(openTarget);
          }
          const message = (error && (error.friendly || error.message)) || "Could not generate the PPTX briefing.";
          toast(message, "error");
        } finally {
          pptxBtn.disabled = false;
          pptxBtn.innerHTML = originalHtml;
        }
      };
    }

    /* Themed, multi-tab Excel workbook export: Brand (logo/wordmark), Summary
       (KPIs + rendered chart images), Projects (full inventory). The KPI/chart
       numbers reuse the same model the PPTX briefing builds from, so the two
       export formats always agree. */
    const xlsxBtn = $("#overview-export-xlsx");
    if (xlsxBtn) {
      xlsxBtn.onclick = async () => {
        const exportService = window.AEWTTR && window.AEWTTR.ExportService;
        const pptxService = window.AEWTTR && window.AEWTTR.OverviewPptxService;
        if (!exportService || typeof exportService.exportMultiSheetXlsx !== "function") {
          toast("Excel export is unavailable in this package.", "error");
          return;
        }
        const isTeamExport = view === "Team" && canSeeTeamOverview;
        const fileName = isTeamExport ? "PULSE-team-overview.xlsx" : "PULSE-my-overview.xlsx";
        const workbookTitle = isTeamExport ? "PULSE Team Overview" : "PULSE My Overview";
        const subtitle = `Project metrics and status | Reporting as of ${asOfLabel()}`;
        const popup = typeof exportService.reserveOpen === "function"
          ? exportService.reserveOpen(fileName, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
          : null;
        const originalHtml = xlsxBtn.innerHTML;
        xlsxBtn.disabled = true;
        xlsxBtn.innerHTML = `<i class="bx bx-loader-alt bx-spin"></i> Building workbook…`;
        try {
          const model = pptxService && typeof pptxService.buildModel === "function"
            ? pptxService.buildModel(isTeamExport, window.AEWTTR.db, projects)
            : null;
          const metrics = (model && model.metrics) || {};

          const totalTasksForChart = metrics.totalTasks || 0;
          const openTasksForChart = metrics.activeTasks || 0;
          const closedTasksForChart = isTeamExport ? (metrics.closedTasks || 0) : (metrics.completedTasks || 0);
          const statusGroups = (isTeamExport && model && model.statusGroups) || {};

          const summaryRows = isTeamExport
            ? [
                ["Total Projects", metrics.totalProjects || 0],
                ["Active Projects", metrics.activeProjects || 0],
                ["Total Tasks", totalTasksForChart],
                ["Open Tasks", openTasksForChart],
                ["Closed Tasks", closedTasksForChart],
                ["On Track Projects", statusGroups["On Track"] || 0],
                ["At Risk Projects", statusGroups["At Risk"] || 0],
                ["Blocked Projects", statusGroups["Blocked"] || 0],
                ["Open Risks", metrics.openRisks || 0],
                ["Team Size", metrics.teamSize || 0],
                ["Documents In Review", metrics.documentsInReview || 0],
                ["Pending Travel", metrics.pendingTravel || 0]
              ]
            : [
                ["Total Projects", metrics.totalProjects || 0],
                ["Total Tasks", totalTasksForChart],
                ["Open Tasks", openTasksForChart],
                ["Closed Tasks", closedTasksForChart],
                ["Documents Reviewed", metrics.documentsReviewed || 0],
                ["Documents Signed", metrics.documentsSigned || 0],
                ["Travel Requests", metrics.travelRequests || 0]
              ];

          const summaryCharts = [{
            type: "bar",
            title: "Task Breakdown",
            categories: ["Total", "Open", "Closed"],
            series: [{ name: "Tasks", values: [totalTasksForChart, openTasksForChart, closedTasksForChart] }],
            pointColors: ["2F66FF", "C77800", "2B7A5E"],
            pxW: 480, pxH: 300
          }];
          if (isTeamExport) {
            summaryCharts.push({
              type: "doughnut",
              title: "Project Health",
              categories: ["On Track", "At Risk", "Blocked"],
              series: [{ name: "Projects", values: [statusGroups["On Track"] || 0, statusGroups["At Risk"] || 0, statusGroups["Blocked"] || 0] }],
              colors: ["2B7A5E", "C77800", "B42318"],
              pxW: 480, pxH: 300
            });
          }

          const projectColumns = ["Project", "End Item", "Portfolio", "PM", "Team", "Derived Status", "Total Tasks", "Open Tasks", "Closed Tasks", "Open Risks", "Funding Type", "Fiscal Year", "Task Order", "Start Date", "Due Date"];
          const projectRows = projects.map(proj => {
            const tasks = allProjectTasks(proj.id);
            const closedCount = tasks.filter(taskIsDone).length;
            return [
              proj.name || "Untitled project",
              projectEndItem(proj),
              (proj.portfolios && proj.portfolios[0]) || "",
              projectPMDisplayName(proj),
              proj.team || "",
              derivedProjectStatus(proj),
              tasks.length,
              tasks.length - closedCount,
              closedCount,
              projectRiskCount(proj.id),
              proj.fundingType || "",
              proj.fiscalYear || "",
              proj.taskOrder || "",
              proj.startDate || "",
              proj.dueDate || ""
            ];
          });

          const sheetDefs = [
            {
              name: "Summary",
              title: workbookTitle,
              subtitle,
              sectionLabel: isTeamExport ? "TEAM SUMMARY" : "MY SUMMARY",
              footerLabel: workbookTitle,
              columns: ["Metric", "Value"],
              rows: summaryRows,
              charts: summaryCharts,
              tabColor: "2F66FF"
            }
          ];

          if (isTeamExport) {
            const portfolioMap = new Map();
            projects.forEach(proj => {
              const names = (Array.isArray(proj.portfolios) && proj.portfolios.length) ? proj.portfolios : ["Unassigned"];
              const tasks = allProjectTasks(proj.id);
              const open = tasks.filter(taskIsActive).length;
              const closed = tasks.filter(taskIsDone).length;
              const risks = projectRiskCount(proj.id);
              names.forEach(name => {
                if (!portfolioMap.has(name)) portfolioMap.set(name, { name, projects: 0, open: 0, closed: 0, risks: 0 });
                const row = portfolioMap.get(name);
                row.projects += 1;
                row.open += open;
                row.closed += closed;
                row.risks += risks;
              });
            });
            const portfolioRows = [...portfolioMap.values()].sort((a, b) => b.open - a.open);
            sheetDefs.push({
              name: "Portfolios",
              title: workbookTitle,
              subtitle,
              sectionLabel: "PORTFOLIO ROLLUP",
              footerLabel: workbookTitle,
              columns: ["Portfolio", "Projects", "Open Tasks", "Closed Tasks", "Open Risks"],
              rows: portfolioRows.map(row => [row.name, row.projects, row.open, row.closed, row.risks]),
              charts: portfolioRows.length ? [{
                type: "bar",
                title: "Open Tasks by Portfolio",
                categories: portfolioRows.slice(0, 8).map(row => row.name),
                series: [{ name: "Open Tasks", values: portfolioRows.slice(0, 8).map(row => row.open) }],
                pxW: 560, pxH: 320
              }] : [],
              tabColor: "7030A0"
            });

            const workloadRows = buildTeamWorkloadRows(projects);
            sheetDefs.push({
              name: "Workload",
              title: workbookTitle,
              subtitle,
              sectionLabel: "TEAM WORKLOAD",
              footerLabel: workbookTitle,
              columns: ["Person", "Open Tasks", "Completed Tasks", "Doc Reviews Pending", "Projects", "Roles", "Team", "Next Due"],
              rows: workloadRows.map(entry => [
                entry.member.name,
                entry.activeTasks,
                entry.completedTasks,
                entry.docsPending,
                entry.projects.size,
                [...entry.roles].join(", "),
                [...entry.teams].join(", "),
                entry.nextDue || ""
              ]),
              charts: workloadRows.length ? [{
                type: "bar",
                horizontal: true,
                title: "Top Open Workload",
                categories: workloadRows.slice(0, 8).map(entry => entry.member.name),
                series: [{ name: "Open Tasks", values: workloadRows.slice(0, 8).map(entry => entry.activeTasks) }],
                pxW: 560, pxH: 320
              }] : [],
              tabColor: "C77800"
            });

            const allRisks = [];
            projects.forEach(proj => {
              (typeof projectRisks === "function" ? projectRisks(proj.id) : []).forEach(risk => {
                if (typeof isPlaceholderRisk === "function" && isPlaceholderRisk(risk)) return;
                const status = String(risk.status || "").toLowerCase();
                if (["closed", "resolved", "accepted"].includes(status)) return;
                allRisks.push({ project: proj.name || "Untitled project", ...risk });
              });
            });
            const ratingCounts = { Red: 0, Amber: 0, Green: 0 };
            allRisks.forEach(risk => { ratingCounts[risk.rating] = (ratingCounts[risk.rating] || 0) + 1; });
            sheetDefs.push({
              name: "Risks",
              title: workbookTitle,
              subtitle,
              sectionLabel: "OPEN RISKS",
              footerLabel: workbookTitle,
              columns: ["Project", "Risk", "Category", "Rating", "Likelihood", "Impact", "Status", "Due", "Owner"],
              rows: allRisks.map(risk => [risk.project, risk.name, risk.category, risk.rating, risk.likelihood, risk.impact, risk.status, risk.due || "", risk.owner || ""]),
              emptyMessage: "No open risks.",
              charts: allRisks.length ? [{
                type: "doughnut",
                title: "Risks by Rating",
                categories: ["Red", "Amber", "Green"],
                series: [{ name: "Risks", values: [ratingCounts.Red, ratingCounts.Amber, ratingCounts.Green] }],
                colors: ["B42318", "C77800", "2B7A5E"],
                pxW: 480, pxH: 300
              }] : [],
              tabColor: "B42318"
            });
          } else {
            const myData = getMyData();
            const taskColumns = ["Task", "Project", "Status", "Priority", "Due"];
            const taskRows = myData.allMyTasks.map(task => [
              task.name || task.title || "Untitled task",
              task.projectName || "",
              task.status || (taskIsDone(task) ? "Done" : "Open"),
              task.priority || "",
              task.end || ""
            ]);
            const priorityCounts = {};
            myData.allMyTasks.filter(taskIsActive).forEach(task => {
              const key = task.priority || "None";
              priorityCounts[key] = (priorityCounts[key] || 0) + 1;
            });
            const priorityEntries = Object.entries(priorityCounts).sort((a, b) => b[1] - a[1]);
            sheetDefs.push({
              name: "My Tasks",
              title: workbookTitle,
              subtitle,
              sectionLabel: "MY TASKS",
              footerLabel: workbookTitle,
              columns: taskColumns,
              rows: taskRows,
              emptyMessage: "No tasks assigned.",
              charts: priorityEntries.length ? [{
                type: "bar",
                title: "Open Tasks by Priority",
                categories: priorityEntries.map(([name]) => name),
                series: [{ name: "Open Tasks", values: priorityEntries.map(([, count]) => count) }],
                pxW: 480, pxH: 300
              }] : [],
              tabColor: "C77800"
            });

            const docsAndTravelRows = [
              ...myData.docsNeedingMe.map(doc => ["Document", doc.title || "Untitled", doc._column || "Pending", doc.submitter || "", doc.updatedAt || ""]),
              ...myData.myTravel.map(req => ["Travel", req.destination || req.id || "Request", req.status || "Pending", req.requester || "", req.startDate || ""])
            ];
            sheetDefs.push({
              name: "Documents & Travel",
              title: workbookTitle,
              subtitle,
              sectionLabel: "DOCUMENTS & TRAVEL",
              footerLabel: workbookTitle,
              columns: ["Type", "Item", "Status", "Owner", "Date"],
              rows: docsAndTravelRows,
              emptyMessage: "No documents or travel requests need attention.",
              tabColor: "7030A0"
            });
          }

          sheetDefs.push({
            name: "Projects",
            title: workbookTitle,
            subtitle,
            sectionLabel: isTeamExport ? "TEAM PROJECT INVENTORY" : "MY PROJECT INVENTORY",
            footerLabel: workbookTitle,
            columns: projectColumns,
            rows: projectRows,
            tabColor: "2B7A5E"
          });

          const result = await exportService.exportMultiSheetXlsx(fileName, sheetDefs, {
            popup,
            folderName: "Overview",
            title: workbookTitle
          });
          if (result.mode === "sharepoint") toast("Excel workbook saved and opened in SharePoint.", "success");
          else if (result.uploadError) toast("SharePoint upload failed — downloaded a local copy instead.", "warn");
          else toast("Excel workbook downloaded.", "success");
        } catch (error) {
          const message = (error && (error.friendly || error.message)) || "Could not build the Excel workbook.";
          toast(message, "error");
        } finally {
          xlsxBtn.disabled = false;
          xlsxBtn.innerHTML = originalHtml;
        }
      };
    }
  }

  render();
};

function overviewProgressBarHtml(pct) {
  if (pct == null) return `<span class="overview-row-sub">No tasks</span>`;
  return `<span class="overview-progress-mini-track"><span class="overview-progress-mini-fill" style="width:${pct}%;"></span></span>${pct}%`;
}
