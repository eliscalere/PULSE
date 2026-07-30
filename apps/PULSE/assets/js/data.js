/* AEWTTR workspace — SharePoint-first empty data shape */

function buildEmptyStore(currentUser) {
  return {
    projects: [],
    projectExtra: {},
    projectPeople: {},
    projectContractors: {},
    projectImports: {},
    projectFinance: {},
    ganttTasks: {},
    engChecklists: {},
    travelRequests: [],
    teamEvents: [],
    debriefs: [],
    // Current 8-status workflow + legacy SharePoint ReviewColumn values so
    // loaders never crash on db.docs[column].push when seeding from SP.
    docs: {
      "Not Started": [],
      "In Review": [],
      "Changes Requested": [],
      "Review Complete": [],
      "Awaiting Final Pack": [],
      "Signing in Progress": [],
      "Signed": [],
      "Archived": [],
      Concurrence: [],
      "Final / Signed": [],
      "Awaiting Signature": [],
      Archive: []
    },
    docReviewerGroups: [],
    groups: [],
    groupReportConfigs: {},
    checklistBoards: [],
    checklistTasks: {},
    tickets: [],
    weeklyMeeting: { roster: [], sessions: [], currentSession: null, meetingStatus: "idle", projectMeetings: {} },
    members: [],
    issues: [],
    appSettings: null,
    locationConfig: null,
    user: {
      id: (currentUser && currentUser.id) || "u0",
      name: (currentUser && (currentUser.displayName || currentUser.name)) || "User",
      email: (currentUser && currentUser.email) || "",
      role: (currentUser && currentUser.role) || "Member",
      isAdmin: !!(currentUser && currentUser.isAdmin),
      spo: "—",
      powerbi: "—"
    }
  };
}

function ensureProjectExtra(pid) {
  const db = window.AEWTTR.db;
  if (!db.projectExtra) db.projectExtra = {};
  if (!db.projectExtra[pid]) db.projectExtra[pid] = { status: "Not Started", history: [], risks: [], handoff: "", meetingNotes: "", notes: [] };
  if (db.projectExtra[pid].meetingNotes == null) db.projectExtra[pid].meetingNotes = "";
  if (!Array.isArray(db.projectExtra[pid].notes)) db.projectExtra[pid].notes = [];
  return db.projectExtra[pid];
}

/* Portfolios/Programs/End Item Configs are just name tags shared across
   projects — not their own SharePoint record — so their Reporting-tab slide
   content (Technical Status bullets, Description, which projects' milestones
   to show) lives in its own keyed store, persisted via
   Repo.save("groupReportConfig", cfg). */
function groupReportConfigKey(groupType, groupName) {
  return `${groupType}:${groupName}`;
}

function ensureGroupReportConfig(groupType, groupName) {
  const db = window.AEWTTR.db;
  if (!db.groupReportConfigs) db.groupReportConfigs = {};
  const key = groupReportConfigKey(groupType, groupName);
  if (!db.groupReportConfigs[key]) {
    db.groupReportConfigs[key] = {
      groupKey: key,
      groupType,
      groupName,
      techBullets: [],
      description: "",
      otherMilestones: [],
      photoProjectId: "",
      photoUrl: "",
      photoName: "",
      includedProjectIds: [],
      riskRows: [{ cat: "Schedule", rag: "", notes: "" }, { cat: "Budget", rag: "", notes: "" }, { cat: "Technical", rag: "", notes: "" }]
    };
  }
  const cfg = db.groupReportConfigs[key];
  if (!Array.isArray(cfg.techBullets)) cfg.techBullets = [];
  if (cfg.description == null) cfg.description = "";
  if (!Array.isArray(cfg.otherMilestones)) cfg.otherMilestones = [];
  if (cfg.photoProjectId == null) cfg.photoProjectId = "";
  if (cfg.photoUrl == null) cfg.photoUrl = "";
  if (cfg.photoName == null) cfg.photoName = "";
  if (!Array.isArray(cfg.includedProjectIds)) cfg.includedProjectIds = [];
  if (!Array.isArray(cfg.riskRows) || !cfg.riskRows.length) {
    cfg.riskRows = [{ cat: "Schedule", rag: "", notes: "" }, { cat: "Budget", rag: "", notes: "" }, { cat: "Technical", rag: "", notes: "" }];
  }
  return cfg;
}

/* ---------- shared reporting milestones ----------
   Project, Portfolio, Program, End Item Config, and PowerPoint reporting all
   describe the same project's schedule.  Keep the typed milestone data in
   one normalized, project-level store instead of letting each reporting
   surface create its own copy.  `taskMilestones` remains the persisted field
   so existing SharePoint ReportConfigJson data stays fully compatible. */
function getProjectReportingMilestones(projectOrId) {
  const projectId = typeof projectOrId === "string" ? projectOrId : (projectOrId && projectOrId.id);
  if (!projectId) return {};
  const extra = ensureProjectExtra(projectId);
  if (!extra.reportConfig || typeof extra.reportConfig !== "object") extra.reportConfig = {};
  const milestones = extra.reportConfig.taskMilestones;
  if (!milestones || typeof milestones !== "object" || Array.isArray(milestones)) {
    extra.reportConfig.taskMilestones = {};
  }
  return extra.reportConfig.taskMilestones;
}

function updateProjectReportingMilestone(projectOrId, taskId, patch) {
  const projectId = typeof projectOrId === "string" ? projectOrId : (projectOrId && projectOrId.id);
  if (!projectId || !taskId) return null;
  const milestones = getProjectReportingMilestones(projectId);
  const current = milestones[taskId] && typeof milestones[taskId] === "object" ? milestones[taskId] : {};
  const next = Object.assign(current, patch || {});
  milestones[taskId] = next;
  if (typeof window !== "undefined" && typeof window.dispatchEvent === "function" && typeof CustomEvent === "function") {
    window.dispatchEvent(new CustomEvent("pulse:reporting-milestones-changed", {
      detail: { projectId: projectId, taskId: taskId, milestones: milestones }
    }));
  }
  return next;
}

/* ---------- project-level milestone dates (single set per project) ----------
   Replaces the per-task taskMilestones picker above for the single-project
   Reporting tab: one Contract/FAT/SAT/ADD/Fielding/Complete date directly on
   the project, not attached to any tracker task. Portfolio/Program/EIS group
   views read this same store to build a cross-project milestone table. */
const PROJECT_MILESTONE_KEYS = ["contractAwarded", "fat", "sat", "add", "fielding", "complete"];

function getProjectMilestones(projectOrId) {
  const projectId = typeof projectOrId === "string" ? projectOrId : (projectOrId && projectOrId.id);
  if (!projectId) return {};
  const extra = ensureProjectExtra(projectId);
  if (!extra.reportConfig || typeof extra.reportConfig !== "object") extra.reportConfig = {};
  let milestones = extra.reportConfig.projectMilestones;
  if (!milestones || typeof milestones !== "object" || Array.isArray(milestones)) {
    milestones = extra.reportConfig.projectMilestones = {};
  }
  PROJECT_MILESTONE_KEYS.forEach((key) => { if (milestones[key] == null) milestones[key] = ""; });
  return milestones;
}

function updateProjectMilestone(projectOrId, key, value) {
  const projectId = typeof projectOrId === "string" ? projectOrId : (projectOrId && projectOrId.id);
  if (!projectId || !PROJECT_MILESTONE_KEYS.includes(key)) return null;
  const milestones = getProjectMilestones(projectId);
  milestones[key] = value || "";
  if (typeof window !== "undefined" && typeof window.dispatchEvent === "function" && typeof CustomEvent === "function") {
    window.dispatchEvent(new CustomEvent("pulse:project-milestones-changed", {
      detail: { projectId: projectId, key: key, milestones: milestones }
    }));
  }
  return milestones;
}

function normalizeRelatedDocs(list) {
  return (list || []).map(doc => ({
    fileName: doc && doc.fileName ? doc.fileName : "",
    location: doc && doc.location ? doc.location : "Windchill"
  })).filter(doc => doc.fileName);
}

function normalizeSubtask(subtask, parentTask) {
  const start = subtask && subtask.start ? subtask.start : (parentTask && parentTask.start ? parentTask.start : "");
  // Subtasks may be open-ended even when their parent has a due date.
  const end = subtask && subtask.end ? subtask.end : "";
  const nested = (subtask && (subtask.subtasks || subtask.children)) || [];
  return {
    id: (subtask && subtask.id) || "",
    text: subtask && subtask.text ? subtask.text : "",
    assignee: subtask && subtask.assignee ? subtask.assignee : "",
    done: !!(subtask && subtask.done),
    health: (subtask && subtask.health) || (parentTask && parentTask.health) || "On Track",
    start,
    end,
    linked: !(subtask && subtask.linked === false),
    notes: Array.isArray(subtask && subtask.notes) ? subtask.notes : [],
    relatedDocs: normalizeRelatedDocs(subtask && subtask.relatedDocs),
    // Nested children — no depth/count cap; UI collapses for navigability.
    subtasks: nested.map((child) => normalizeSubtask(child, {
      start,
      end,
      health: (subtask && subtask.health) || (parentTask && parentTask.health) || "On Track"
    }))
  };
}

function blankDividerMetadata(src) {
  const s = src || {};
  const meta = {
    pm: s.pm || "",
    engineer: s.engineer || "",
    isso: s.isso || "",
    rangePoc: s.rangePoc || "",
    contract: s.contract || "",
    contractor: s.contractor || "",
    taskOrder: s.taskOrder || "",
    fundingType: s.fundingType || "",
    fiscalYear: s.fiscalYear || "",
    fundingStatus: s.fundingStatus || "",
    changeRequestRequired: !!s.changeRequestRequired,
    configEndItem: s.configEndItem || "",
    program: s.program || "",
    product: s.product || "",
    locations: Array.isArray(s.locations) ? s.locations.slice() : [],
    portfolios: Array.isArray(s.portfolios) ? s.portfolios.slice() : [],
    priority: s.priority || "",
    rag: s.rag || "Green",
    lifecycleStatus: s.lifecycleStatus || "",
    scope: s.scope || "",
    objectives: s.objectives || "",
    description: s.description || "",
    startDate: s.startDate || s.start || "",
    dueDate: s.dueDate || s.end || s.due || "",
    completionDate: s.completionDate || "",
    isMilestone: !!(s.isMilestone || (s.metadata && s.metadata.isMilestone))
  };
  // Preserve client identity keys so SharePoint reloads keep the same local ids
  // (parentDividerLocalId / task↔divider links survive refresh).
  if (s.clientLocalId) meta.clientLocalId = s.clientLocalId;
  else if (s.id) meta.clientLocalId = s.id;
  if (s.projectCode) meta.projectCode = s.projectCode;
  else if (s._projectCode) meta.projectCode = s._projectCode;
  if (s.parentDividerLocalId) meta.parentDividerLocalId = s.parentDividerLocalId;
  return meta;
}

function isTrackerDivider(item) {
  return !!(item && (item.itemType === "divider" || item.workItemLevel === "Divider"));
}

function normalizeGanttTask(task) {
  if (isTrackerDivider(task)) {
    const meta = blankDividerMetadata(Object.assign({}, task.metadata || {}, task));
    return {
      ...task,
      itemType: "divider",
      workItemLevel: "Divider",
      title: task.title || "Untitled divider",
      assignee: task.assignee || "",
      parentDividerId: "",
      start: meta.startDate || "",
      end: meta.dueDate || "",
      startDate: meta.startDate || "",
      dueDate: meta.dueDate || "",
      completionDate: meta.completionDate || "",
      health: task.health || "On Track",
      status: task.status || "Not Started",
      notes: Array.isArray(task.notes) ? task.notes : [],
      subtasks: [],
      metadata: meta,
      pm: meta.pm,
      engineer: meta.engineer,
      isso: meta.isso,
      rangePoc: meta.rangePoc,
      contract: meta.contract,
      contractor: meta.contractor,
      taskOrder: meta.taskOrder,
      fundingType: meta.fundingType,
      fiscalYear: meta.fiscalYear,
      fundingStatus: meta.fundingStatus,
      changeRequestRequired: meta.changeRequestRequired,
      configEndItem: meta.configEndItem,
      program: meta.program,
      product: meta.product,
      locations: meta.locations,
      portfolios: meta.portfolios,
      priority: meta.priority,
      rag: meta.rag,
      lifecycleStatus: meta.lifecycleStatus,
      scope: meta.scope,
      objectives: meta.objectives,
      description: meta.description,
      isMilestone: !!(task.isMilestone || meta.isMilestone)
    };
  }
  return {
    ...task,
    itemType: task.itemType || "task",
    workItemLevel: task.workItemLevel || "Task",
    parentDividerId: task.parentDividerId || "",
    assignee: task.assignee || "Unassigned",
    start: task.start || new Date().toISOString().slice(0, 10),
    // A blank end date means the work is open-ended. Do not turn it into a
    // same-day due date during normalization.
    end: task.end || "",
    health: task.health || "On Track",
    isMilestone: false,
    notes: Array.isArray(task.notes) ? task.notes : (task.notes ? [{ id: uid("nt"), author: "Notes", date: "", time: "", text: String(task.notes) }] : []),
    subtasks: (task.subtasks || []).map(sub => normalizeSubtask(sub, task))
  };
}

function normalizeStoreShape(db) {
  if (!db || typeof db !== "object") db = {};
  const storeDefaults = buildEmptyStore(db.user);
  Object.keys(storeDefaults).forEach((key) => {
    if (key === "user") return;
    if (db[key] == null) db[key] = JSON.parse(JSON.stringify(storeDefaults[key]));
  });
  if (!db.user || typeof db.user !== "object") db.user = { ...storeDefaults.user };
  if (!db.user.id) {
    const me = (db.members || []).find(member => member.name === db.user.name);
    if (me) db.user.id = me.id;
  }
  db.travelRequests = (db.travelRequests || []).map(r => ({
    ...r,
    requestType: r.requestType || "Project Travel",
    projectIds: r.projectIds || []
  }));
  db.issues = Array.isArray(db.issues) ? db.issues : [];
  if (db.appSettings != null && typeof db.appSettings !== "object") db.appSettings = null;
  // Guarantee every known doc-review column bag exists (cache / old clients
  // may only have the pre-signing 5-bucket shape).
  if (!db.docs || typeof db.docs !== "object") db.docs = {};
  [
    "Not Started", "In Review", "Changes Requested", "Review Complete",
    "Awaiting Final Pack", "Signing in Progress", "Signed", "Archived",
    "Concurrence", "Final / Signed", "Awaiting Signature", "Archive"
  ].forEach((column) => {
    if (!Array.isArray(db.docs[column])) db.docs[column] = [];
  });
  (db.projects || []).forEach(p => {
    if (p.coverImage == null) p.coverImage = "";
    if (!Array.isArray(p.images)) p.images = [];
    if (p.description == null) p.description = "";
    if (p.scope == null) p.scope = "";
    if (!Array.isArray(p.deliverables)) {
      p.deliverables = String(p.deliverables || p.objectives || "")
        .split(/\r?\n/)
        .map(item => item.replace(/^\s*[-*•]\s*/, "").trim())
        .filter(Boolean);
    }
    delete p.objectives;
    if (p.projectType == null) p.projectType = p.parentProjectId ? "Subproject" : "Project";
    if (p.parentProjectId == null) p.parentProjectId = "";
    if (p.pm == null) p.pm = "";
    if (p.engineer == null) p.engineer = "";
    if (p.isso == null) p.isso = "";
    if (p.rangePoc == null) p.rangePoc = "";
    if (p.contract == null) p.contract = "";
    if (p.contractor == null) p.contractor = "";
    if (p.taskOrder == null) p.taskOrder = "";
    if (p.fundingType == null) p.fundingType = "";
    if (p.fiscalYear == null) p.fiscalYear = "";
    if (p.fundingStatus == null) p.fundingStatus = "";
    if (p.changeRequestRequired == null) p.changeRequestRequired = false;
    if (p.configEndItem == null) p.configEndItem = "";
    if (p.program == null) p.program = "";
    if (p.product == null) p.product = "";
    if (p.locations == null) p.locations = [];
    if (!Array.isArray(p.portfolios)) p.portfolios = [];
    if (p.mainFolderUrl == null) p.mainFolderUrl = "";
    if (p.exportSelections == null || typeof p.exportSelections !== "object") p.exportSelections = {};
    if (p.startDate == null) p.startDate = "";
    if (p.dueDate == null) p.dueDate = "";
    if (p.completionDate == null) p.completionDate = "";
  });
  db.projectExtra = db.projectExtra || {};
  Object.keys(db.projectExtra).forEach(pid => {
    if (!db.projectExtra[pid] || typeof db.projectExtra[pid] !== "object") {
      db.projectExtra[pid] = { status: "Not Started", history: [], risks: [], handoff: "", meetingNotes: "" };
    }
    db.projectExtra[pid].meetingNotes = db.projectExtra[pid].meetingNotes || "";
  });
  db.projectPeople = db.projectPeople || {};
  Object.keys(db.projectPeople).forEach(projectId => {
    db.projectPeople[projectId] = (db.projectPeople[projectId] || []).map(person => ({
      id: person.id || uid("ppj"),
      type: person.type || "member",
      memberId: person.memberId || "",
      label: person.label || "",
      role: person.role || "",
      company: person.company || "",
      email: person.email || ""
    }));
  });
  db.projectImports = db.projectImports || {};
  Object.keys(db.projectImports).forEach(projectId => {
    const entry = db.projectImports[projectId] || {};
    entry.docs = (entry.docs || []).map(doc => ({
      id: doc.id || uid("imp"),
      type: doc.type || "communication",
      fileName: doc.fileName || "",
      uploadedOn: doc.uploadedOn || new Date().toISOString().slice(0, 10),
      uploadedBy: doc.uploadedBy || "",
      accessOnly: doc.accessOnly !== false,
      parsed: doc.parsed || {},
      applied: !!doc.applied
    }));
    db.projectImports[projectId] = entry;
  });
  db.projectFinance = db.projectFinance || {};
  Object.keys(db.projectFinance).forEach(projectId => {
    const finance = db.projectFinance[projectId] || {};
    finance.summary = finance.summary || {};
    finance.snapshot = {
      totalCeiling: Number(finance.snapshot && finance.snapshot.totalCeiling || 0),
      fundedToDate: Number(finance.snapshot && finance.snapshot.fundedToDate || 0),
      obligated: Number(finance.snapshot && finance.snapshot.obligated || 0),
      accrued: Number(finance.snapshot && finance.snapshot.accrued || 0),
      actuals: Number(finance.snapshot && finance.snapshot.actuals || 0),
      forecastAtComplete: Number(finance.snapshot && finance.snapshot.forecastAtComplete || 0)
    };
    finance.clins = (finance.clins || []).map(clin => ({
      id: clin.id || uid("clin"),
      code: clin.code || "",
      title: clin.title || "",
      type: clin.type || "",
      funded: Number(clin.funded || 0),
      obligated: Number(clin.obligated || 0),
      spent: Number(clin.spent || 0),
      status: clin.status || "Active"
    }));
    finance.invoices = (finance.invoices || []).map(invoice => ({
      id: invoice.id || uid("inv"),
      invoiceNo: invoice.invoiceNo || "",
      vendor: invoice.vendor || "",
      amount: Number(invoice.amount || 0),
      received: invoice.received || "",
      status: invoice.status || "In Review",
      due: invoice.due || ""
    }));
    finance.watchlist = (finance.watchlist || []).map(item => ({
      id: item.id || uid("fw"),
      title: item.title || "",
      owner: item.owner || "",
      due: item.due || "",
      status: item.status || "Open",
      note: item.note || ""
    }));
    finance.communications = (finance.communications || []).map(item => ({
      id: item.id || uid("fc"),
      subject: item.subject || "",
      counterparty: item.counterparty || "",
      date: item.date || "",
      action: item.action || ""
    }));
    db.projectFinance[projectId] = finance;
  });
  Object.keys(db.ganttTasks || {}).forEach(projectId => {
    db.ganttTasks[projectId] = (db.ganttTasks[projectId] || []).map(normalizeGanttTask);
  });
  ["engChecklists", "checklistTasks"].forEach(key => {
    Object.keys(db[key] || {}).forEach(boardId => {
      const board = db[key][boardId];
      Object.keys(board || {}).forEach(column => {
        board[column] = (board[column] || []).map(task => ({
          ...task,
          subtasks: (task.subtasks || []).map(sub => normalizeSubtask(sub)),
          relatedDocs: normalizeRelatedDocs(task.relatedDocs)
        }));
      });
    });
  });
  function normalizeSession(session, idPrefix) {
    let notesFeed = (session && Array.isArray(session.notesFeed)) ? session.notesFeed.slice() : [];
    if (!notesFeed.length && session && session.notes && String(session.notes).trim()) {
      notesFeed = [{
        id: uid("note"),
        author: "Meeting notes",
        date: (session && session.date) || new Date().toISOString().slice(0, 10),
        time: "",
        text: String(session.notes).trim()
      }];
    }
    notesFeed = notesFeed
      .filter((note) => note && String(note.text || "").trim())
      .map((note) => ({
        id: note.id || uid("note"),
        author: note.author === "Legacy note" ? "Meeting notes" : (note.author || "Unknown"),
        date: note.date || new Date().toISOString().slice(0, 10),
        time: note.time || "",
        text: String(note.text || "").trim()
      }));
    const sessionStatus = session && session.sessionStatus ? session.sessionStatus : "active";
    const agenda = (session && Array.isArray(session.agenda)) ? session.agenda
      .filter((item) => item && typeof item === "object")
      .map((item, index) => ({
        id: String(item.id || `agenda-${index + 1}`),
        title: String(item.title || ""),
        notes: String(item.notes || ""),
        done: !!item.done,
        author: String(item.author || ""),
        createdAt: item.createdAt || ""
      })) : [];
    // Keep the property absent for legacy sessions. The meeting editor uses
    // that distinction to migrate an existing AgendaJson into document blocks
    // rather than interpreting an absent document as a deliberately blank one.
    const rawDocBlocks = (session && Array.isArray(session.docBlocks)) ? session.docBlocks : null;
    const normalized = {
      id: (session && session.id) || uid(idPrefix),
      date: (session && session.date) || new Date().toISOString().slice(0, 10),
      agenda,
      notesFeed,
      guests: (session && session.guests) || [],
      attendance: (session && session.attendance) || {},
      activity: (session && session.activity) || [],
      sessionStatus,
      startedAt: (session && session.startedAt) || null,
      endedAt: (session && session.endedAt) || null,
      tasksUpdated: (session && session.tasksUpdated) || 0,
      tasksReviewed: (session && session.tasksReviewed) || 0,
      capturedNotes: (session && Array.isArray(session.capturedNotes)) ? session.capturedNotes : [],
      ganttChanges: (session && Array.isArray(session.ganttChanges)) ? session.ganttChanges : [],
      minutesLog: (session && Array.isArray(session.minutesLog)) ? session.minutesLog : [],
      guestAttendees: (session && Array.isArray(session.guestAttendees)) ? session.guestAttendees : [],
      docHtml: (session && typeof session.docHtml === "string") ? session.docHtml : "",
      _spId: (session && session._spId) || undefined,
      _projectCode: (session && session._projectCode) || "",
      scheduledTitle: (session && session.scheduledTitle) || null,
      scheduledTime: (session && session.scheduledTime) || null,
      agendaText: (session && session.agendaText) || null,
      createdBy: (session && session.createdBy) || null,
      createdAt: (session && session.createdAt) || null
    };
    if (rawDocBlocks !== null) {
      const blockTypes = ["action", "bullet", "h2", "para"];
      normalized.docBlocks = rawDocBlocks.filter((block) => block && typeof block === "object").map((block, index) => ({
        id: String(block.id || `doc-block-${index + 1}`),
        type: blockTypes.includes(block.type) ? block.type : "para",
        content: String(block.content || ""),
        done: !!block.done
      }));
    }
    return normalized;
  }
  function meetingStatusForData(data) {
    if (data.meetingStatus) return data.meetingStatus;
    if (data.currentSession) {
      return data.currentSession.sessionStatus === "ended" ? "ended" : "active";
    }
    return (data.sessions && data.sessions.length) ? "ended" : "idle";
  }
  if (!db.weeklyMeeting) db.weeklyMeeting = { roster: [], sessions: [], currentSession: null, meetingStatus: "idle" };
  db.weeklyMeeting.roster = db.weeklyMeeting.roster || [];
  db.weeklyMeeting.sessions = (db.weeklyMeeting.sessions || []).map(session => ({
    ...normalizeSession(session, "wm"),
    attendees: session.attendees || []
  }));
  if (db.weeklyMeeting.currentSession) {
    db.weeklyMeeting.currentSession = normalizeSession(db.weeklyMeeting.currentSession, "wm");
    if (db.weeklyMeeting.currentSession.sessionStatus === "ended") db.weeklyMeeting.currentSession = null;
  }
  db.weeklyMeeting.meetingStatus = meetingStatusForData(db.weeklyMeeting);
  db.weeklyMeeting.projectMeetings = db.weeklyMeeting.projectMeetings || {};
  Object.keys(db.weeklyMeeting.projectMeetings).forEach(projectId => {
    if (!db.weeklyMeeting.projectMeetings[projectId] || typeof db.weeklyMeeting.projectMeetings[projectId] !== "object") {
      db.weeklyMeeting.projectMeetings[projectId] = { sessions: [], currentSession: null, meetingStatus: "idle" };
    }
    const meeting = db.weeklyMeeting.projectMeetings[projectId];
    meeting.sessions = (meeting.sessions || []).map(session => ({
      ...normalizeSession(session, "pms"),
      attendees: session.attendees || []
    }));
    if (meeting.currentSession) {
      meeting.currentSession = normalizeSession(meeting.currentSession, "pmc");
      if (meeting.currentSession.sessionStatus === "ended") meeting.currentSession = null;
    }
    meeting.meetingStatus = meetingStatusForData(meeting);
    delete meeting.participants;
    delete meeting.scorecard;
    delete meeting.roadblocks;
  });
  return db;
}

const LOCAL_STORE_KEY = "pulse-local-db";
const LOCAL_STORE_MAX_AGE_MS = 90 * 24 * 60 * 60 * 1000; // 90 days

function aewttrLoadStore() {
  try {
    const raw = localStorage.getItem(LOCAL_STORE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && parsed.data && parsed.savedAt && (Date.now() - Number(parsed.savedAt)) < LOCAL_STORE_MAX_AGE_MS) {
        return normalizeStoreShape(parsed.data);
      }
    }
  } catch (e) { /* storage unavailable or corrupt */ }
  return normalizeStoreShape(buildEmptyStore({ name: "Local User", role: "Admin", isAdmin: true }));
}

function aewttrSaveStore() {
  if (!window.AEWTTR || !window.AEWTTR.db) return;
  try {
    const payload = JSON.stringify({ savedAt: Date.now(), data: window.AEWTTR.db });
    localStorage.setItem(LOCAL_STORE_KEY, payload);
  } catch (e) { /* storage full or blocked */ }
}

function aewttrResetStore() {
  const fresh = normalizeStoreShape(buildEmptyStore({ name: "Local User", role: "Admin", isAdmin: true }));
  window.AEWTTR.db = fresh;
  try { localStorage.removeItem(LOCAL_STORE_KEY); } catch (e) { /* ignore */ }
}

/* ---------- technical-status bullets ----------

   Bullets were plain strings. Sub-bullets need an indent level, so they are now
   { text, level }. Existing saved data is all strings, and a string is simply a
   level-0 bullet — normalising on read means no migration and no risk of losing
   someone's saved slide content.

   The depth cap is 4 because that is what a slide can carry legibly; PowerPoint
   itself allows more, but a fifth indent on a quadrant card is unreadable. */
const TECH_BULLET_MAX_LEVEL = 4;

function normalizeTechBullet(bullet) {
  if (bullet && typeof bullet === "object") {
    const level = Math.min(TECH_BULLET_MAX_LEVEL, Math.max(0, parseInt(bullet.level, 10) || 0));
    return { text: String(bullet.text == null ? "" : bullet.text), level };
  }
  return { text: String(bullet == null ? "" : bullet), level: 0 };
}

/* Levels are per-item, so nothing stops a list from jumping 0 → 2 — which is an
   orphan with no parent to hang off, and renders as a stray deep indent on the
   slide. Outdenting a parent is the easy way to create one. Clamping each bullet
   to at most one deeper than the bullet above heals that on read, so the outline
   is always well-formed no matter how it was edited. */
function normalizeTechBullets(list) {
  let ceiling = 0;
  return (Array.isArray(list) ? list : []).map((item) => {
    const bullet = normalizeTechBullet(item);
    if (bullet.level > ceiling) bullet.level = ceiling;
    ceiling = Math.min(TECH_BULLET_MAX_LEVEL, bullet.level + 1);
    return bullet;
  });
}

/* Flat text, for anywhere that wants the bullet without its indent. */
function techBulletText(bullet) {
  return normalizeTechBullet(bullet).text;
}

function techBulletLevel(bullet) {
  return normalizeTechBullet(bullet).level;
}
