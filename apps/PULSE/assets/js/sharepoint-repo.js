/* AEWTTR-PULSE — SharePoint repository layer.

   Bridges the SharePoint REST adapter to the in-memory `db` shape every
   page module already reads/writes (db.projects, db.ganttTasks,
   db.travelRequests, db.weeklyMeeting, etc.). Page code never calls
   sharePointAdapter directly — it keeps mutating local objects exactly as
   before, and calls Repo.save(kind, obj) / Repo.remove(kind, obj) where it
   used to call aewttrSaveStore(). In local-fallback mode those calls just
   delegate to today's localStorage write; in SharePoint mode they write the
   one changed record to the right list. */

const SP_LISTS = {
  appRole: "PULSE App Roles",
  project: "PULSE Projects",
  risk: "PULSE Risks",
  actionItem: "PULSE Action Items",
  meetingSession: "PULSE Meetings",
  travelRequest: "PULSE Travel Requests",
  travelDebrief: "PULSE Travel Debriefs",
  teamEvent: "PULSE Team Events",
  docReviewItem: "PULSE Document Review",
  docReviewerGroup: "PULSE Doc Reviewer Groups",
  auditLog: "PULSE Audit Log",
  issue: "PULSE Issues",
  ticket: "PULSE Issues",
  notification: "PULSE Notifications",
  notificationConfig: "PULSE Notification Config",
  appSettings: "PULSE App Settings",
  locationConfig: "PULSE Location Config"
};

/* Persisted to localStorage (not just an in-memory variable) so the Admin >
   SharePoint Setup > Diagnostics panel can still show the error that caused
   a failed save even after the user reloads the page to "try again" —
   which otherwise wipes the only record of what went wrong. */
const LAST_SP_ERROR_KEY = "aewttrLastSpError";
let _lastSpError = null;
function recordSpError(err) {
  _lastSpError = err && err.friendly ? err : { friendly: String(err && err.message || err), raw: "" };
  try {
    localStorage.setItem(LAST_SP_ERROR_KEY, JSON.stringify({ ..._lastSpError, recordedAt: new Date().toISOString() }));
  } catch (e) { /* localStorage unavailable — in-memory value still works this session */ }
  return _lastSpError;
}
function getLastSpError() {
  if (_lastSpError) return _lastSpError;
  try {
    const stored = localStorage.getItem(LAST_SP_ERROR_KEY);
    if (stored) return JSON.parse(stored);
  } catch (e) { /* ignore malformed/unavailable storage */ }
  return null;
}
function clearLastSpError() {
  _lastSpError = null;
  try { localStorage.removeItem(LAST_SP_ERROR_KEY); } catch (e) { /* ignore */ }
}

function isSharePointMode() { return window.AEWTTR.mode === "sharepoint"; }
function currentSiteUrl() { return window.AEWTTR.siteUrl; }

/* ---------- ProjectCode <-> SharePoint item id map ---------- */
let _projectIdMap = { byCode: {}, byId: {} };
function registerProjectId(code, spId) {
  _projectIdMap.byCode[code] = spId;
  _projectIdMap.byId[spId] = code;
}
function spIdForProjectCode(code) { return _projectIdMap.byCode[code]; }
function projectCodeForSpId(spId) { return _projectIdMap.byId[spId] || ""; }

const CHECKLIST_COLUMN_TO_STATUS = { "To Do": "Not Started", "In Progress": "In Progress", "Complete": "Complete" };
const STATUS_TO_CHECKLIST_COLUMN = { "Not Started": "To Do", "In Progress": "In Progress", "Blocked": "To Do", "On Hold": "To Do", "Complete": "Complete", "Done": "Complete" };

function safeJsonParse(text, fallback) {
  if (!text) return fallback;
  try { return JSON.parse(text); } catch (e) { return fallback; }
}

function normalizeFieldAlias(value) {
  return String(value || "").replace(/[^a-z0-9]/gi, "").toLowerCase();
}

function getSchemaFieldsForList(listTitle) {
  return (typeof SHAREPOINT_SCHEMA !== "undefined" && SHAREPOINT_SCHEMA[listTitle] && SHAREPOINT_SCHEMA[listTitle].fields) || [];
}

function normalizeListItemToSchema(item, fieldDefinitions, schemaFields, listTitle) {
  const normalized = { ...(item || {}) };
  const fields = fieldDefinitions || [];
  (schemaFields || []).forEach((schemaField) => {
    if (normalized[schemaField.name] != null && normalized[schemaField.name] !== "") return;
    const match = typeof findSchemaFieldMatch === "function"
      ? findSchemaFieldMatch(fields, schemaField)
      : null;
    const matchedField = match && match.found && match.field ? match.field : fields.find((field) => {
      if (!field) return false;
      const targetAlias = normalizeFieldAlias(schemaField.name);
      const aliases = new Set([
        normalizeFieldAlias(field.InternalName),
        normalizeFieldAlias(field.Title),
        normalizeFieldAlias(field.StaticName),
        normalizeFieldAlias(field.EntityPropertyName),
        ...((field.aliases || []).map(normalizeFieldAlias))
      ]);
      return aliases.has(targetAlias);
    });
    if (matchedField && Object.prototype.hasOwnProperty.call(normalized, matchedField.InternalName)) {
      normalized[schemaField.name] = normalized[matchedField.InternalName];
      if (typeof pushDebugLog === "function" && matchedField.InternalName !== schemaField.name) {
        pushDebugLog("info", "sharepoint-load", `Remapped ${listTitle || "list"} field ${schemaField.name}`, {
          fromInternalName: matchedField.InternalName,
          toSchemaName: schemaField.name
        });
      }
      return;
    }
    /* Fallback when field-definition probes were skipped or incomplete: match
       any key on the REST payload whose normalized alias equals the schema
       field (e.g. Subtasks_x0020_Json → SubtasksJson). */
    const targetAlias = normalizeFieldAlias(schemaField.name);
    const fallbackKey = Object.keys(normalized).find((key) => {
      if (!key || key === "Id" || key === "ID" || key.startsWith("@") || key.startsWith("odata")) return false;
      if (key === schemaField.name) return false;
      return normalizeFieldAlias(key) === targetAlias;
    });
    if (fallbackKey != null && normalized[fallbackKey] != null && normalized[fallbackKey] !== "") {
      normalized[schemaField.name] = normalized[fallbackKey];
    }
  });
  return normalized;
}

function normalizeItemsToSchema(items, fieldDefinitions, listTitle) {
  const schemaFields = getSchemaFieldsForList(listTitle);
  return (items || []).map((item) => normalizeListItemToSchema(item, fieldDefinitions, schemaFields, listTitle));
}

/* ---------- mappers: project ---------- */

function serializeProjectBoards(projectId) {
  const db = window.AEWTTR.db || {};
  return (db.checklistBoards || [])
    .filter((board) => board && board.projectId === projectId)
    .map((board) => ({
      id: board.id,
      name: board.name || "Board",
      projectId: board.projectId,
      team: board.team || "All",
      type: board.type === "kanban" ? "kanban" : "table",
      statuses: Array.isArray(board.statuses) && board.statuses.length ? board.statuses.slice() : ["To Do", "In Progress", "Complete"],
      columnMeta: board.columnMeta || {},
      customFields: Array.isArray(board.customFields) ? board.customFields.slice() : [],
      tasks: (db.checklistTasks && db.checklistTasks[board.id]) || {}
    }));
}

function hydrateProjectBoards(db, projectId, boardsPayload) {
  const boards = Array.isArray(boardsPayload) ? boardsPayload : [];
  if (!db.checklistBoards) db.checklistBoards = [];
  if (!db.checklistTasks) db.checklistTasks = {};
  const keepIds = new Set();
  boards.forEach((raw) => {
    if (!raw || !raw.id) return;
    keepIds.add(raw.id);
    const board = {
      id: raw.id,
      name: raw.name || "Board",
      projectId: projectId || raw.projectId || "",
      team: raw.team || "All",
      type: raw.type === "kanban" ? "kanban" : "table",
      statuses: Array.isArray(raw.statuses) && raw.statuses.length ? raw.statuses.slice() : ["To Do", "In Progress", "Complete"],
      columnMeta: raw.columnMeta || {},
      customFields: Array.isArray(raw.customFields) ? raw.customFields.slice() : []
    };
    const existingIdx = db.checklistBoards.findIndex((b) => b.id === board.id);
    if (existingIdx >= 0) db.checklistBoards[existingIdx] = board;
    else db.checklistBoards.push(board);
    db.checklistTasks[board.id] = raw.tasks && typeof raw.tasks === "object" ? raw.tasks : {};
    board.statuses.forEach((status) => {
      if (!db.checklistTasks[board.id][status]) db.checklistTasks[board.id][status] = [];
    });
  });
  // Drop boards for this project that are no longer in BoardsJson — otherwise
  // a deleted board comes back on the next SharePoint poll / page load.
  db.checklistBoards = db.checklistBoards.filter((b) => {
    if (!b || b.projectId !== projectId) return true;
    if (keepIds.has(b.id)) return true;
    if (db.checklistTasks) delete db.checklistTasks[b.id];
    return false;
  });
}

function projectToSpItem(proj) {
  const db = window.AEWTTR.db;
  const extra = (db.projectExtra && db.projectExtra[proj.id]) || { history: [], risks: [], handoff: "", meetingNotes: "" };
  const people = (db.projectPeople && db.projectPeople[proj.id]) || [];
  return {
    Title: proj.name,
    ProjectCode: proj.id,
    ProjectStatus: typeof computeProjectStatus === "function" ? computeProjectStatus(proj) : "Not Started",
    Team: proj.team,
    Priority: proj.priority,
    Tools: proj.tools || "",
    Stakeholders: proj.stakeholders || "",
    Description: proj.description || "",
    Scope: proj.scope || "",
    DeliverablesJson: JSON.stringify(typeof projectDeliverables === "function" ? projectDeliverables(proj) : (proj.deliverables || [])),
    CoverImageUrl: proj.coverImage || "",
    ProjectImagesJson: JSON.stringify(proj.images || []),
    HistoryJson: JSON.stringify(extra.history || []),
    RisksJson: JSON.stringify(extra.risks || []),
    Handoff: extra.handoff || "",
    MeetingNotes: extra.meetingNotes || "",
    PeopleJson: JSON.stringify(people),
    BoardsJson: JSON.stringify(serializeProjectBoards(proj.id)),
    LinkSubtaskDates: proj.linkSubtaskDates !== false,
    LastUpdated: proj.updated ? `${proj.updated}T00:00:00Z` : null,
    IsArchived: false,
    ProjectType: proj.projectType || (proj.parentProjectId ? "Subproject" : "Project"),
    ParentProjectId: proj.parentProjectId ? (spIdForProjectCode(proj.parentProjectId) || 0) : 0,
    Pm: proj.pm || "",
    Engineer: proj.engineer || "",
    Isso: proj.isso || "",
    RangePoc: proj.rangePoc || "",
    Contract: proj.contract || "",
    Contractor: proj.contractor || "",
    TaskOrder: proj.taskOrder || "",
    FundingType: proj.fundingType || "",
    FiscalYear: proj.fiscalYear || "",
    FundingStatus: proj.fundingStatus || "",
    ChangeRequestRequired: !!proj.changeRequestRequired,
    ConfigEndItem: proj.configEndItem || "",
    Program: proj.program || "",
    Product: proj.product || "",
    LocationsJson: JSON.stringify(proj.locations || []),
    PortfoliosJson: JSON.stringify(typeof projectPortfolios === "function" ? projectPortfolios(proj) : (proj.portfolios || [])),
    MainFolderUrl: proj.mainFolderUrl || "",
    ExportSelectionsJson: JSON.stringify(proj.exportSelections || {}),
    ReportConfigJson: JSON.stringify(extra.reportConfig || {}),
    ProjectStartDate: proj.startDate ? `${proj.startDate}T00:00:00Z` : null,
    ProjectDueDate: proj.dueDate ? `${proj.dueDate}T00:00:00Z` : null,
    ProjectCompletionDate: proj.completionDate ? `${proj.completionDate}T00:00:00Z` : null,
    NotesJson: JSON.stringify(extra.notes || []),
    FundedOnContractAmount: proj.fundedOnContractAmount || 0,
    ReimbursableAmount: proj.reimbursableAmount || 0,
    FundingNotes: proj.fundingNotes || "",
    TechnicalStatus: proj.technicalStatus || "",
    LifecycleStatus: proj.lifecycleStatus || "",
    ClassificationType: proj.projectType || "",
    ATO: proj.ato || "",
    AquOnly: !!proj.aquOnly,
    RelatedProjects: proj.projects || ""
  };
}

function spItemToProject(item) {
  const parsedPortfolios = safeJsonParse(item.PortfoliosJson, null);
  const portfolios = Array.isArray(parsedPortfolios)
    ? parsedPortfolios.map((n) => String(n == null ? "" : n).trim()).filter(Boolean)
    : [];
  const proj = {
    _spId: item.Id,
    id: item.ProjectCode || `P${item.Id}`,
    name: item.Title || "",
    team: item.Team || "",
    priority: item.Priority || "Medium",
    effort: "TBD",
    updated: item.LastUpdated ? item.LastUpdated.slice(0, 10) : new Date().toISOString().slice(0, 10),
    tools: item.Tools || "",
    stakeholders: item.Stakeholders || "",
    coverImage: item.CoverImageUrl || "",
    images: safeJsonParse(item.ProjectImagesJson, []),
    description: item.Description || "",
    scope: item.Scope || "",
    deliverables: safeJsonParse(item.DeliverablesJson, null)
      || String(item.Objectives || "").split(/\r?\n/).map((value) => value.replace(/^\s*[-*•]\s*/, "").trim()).filter(Boolean),
    linkSubtaskDates: item.LinkSubtaskDates !== false,
    projectType: item.ProjectType || (item.ParentProjectId ? "Subproject" : "Project"),
    parentProjectId: item.ParentProjectId ? projectCodeForSpId(item.ParentProjectId) : "",
    pm: item.Pm || "",
    engineer: item.Engineer || "",
    isso: item.Isso || "",
    rangePoc: item.RangePoc || "",
    contract: item.Contract || "",
    contractor: item.Contractor || "",
    taskOrder: item.TaskOrder || "",
    fundingType: item.FundingType || "",
    fiscalYear: item.FiscalYear || "",
    fundingStatus: item.FundingStatus || "",
    changeRequestRequired: !!item.ChangeRequestRequired,
    configEndItem: item.ConfigEndItem || "",
    program: item.Program || "",
    product: item.Product || "",
    locations: safeJsonParse(item.LocationsJson, []),
    portfolios,
    mainFolderUrl: item.MainFolderUrl || "",
    exportSelections: safeJsonParse(item.ExportSelectionsJson, {}),
    startDate: item.ProjectStartDate ? item.ProjectStartDate.slice(0, 10) : "",
    dueDate: item.ProjectDueDate ? item.ProjectDueDate.slice(0, 10) : "",
    completionDate: item.ProjectCompletionDate ? item.ProjectCompletionDate.slice(0, 10) : "",
    fundedOnContractAmount: Number(item.FundedOnContractAmount || 0),
    reimbursableAmount: Number(item.ReimbursableAmount || 0),
    fundingNotes: item.FundingNotes || "",
    technicalStatus: item.TechnicalStatus || "",
    lifecycleStatus: item.LifecycleStatus || "",
    projectType: item.ClassificationType || "",
    ato: item.ATO || "",
    aquOnly: !!item.AquOnly,
    projects: item.RelatedProjects || ""
  };
  const extra = {
    status: item.ProjectStatus || "Not Started",
    history: safeJsonParse(item.HistoryJson, []),
    risks: safeJsonParse(item.RisksJson, []),
    handoff: item.Handoff || "",
    meetingNotes: item.MeetingNotes || "",
    reportConfig: safeJsonParse(item.ReportConfigJson, {}),
    notes: safeJsonParse(item.NotesJson, [])
  };
  const people = safeJsonParse(item.PeopleJson, []);
  const boards = safeJsonParse(item.BoardsJson, []);
  return { proj, extra, people, boards };
}

function riskRatingFromScore(likelihood, impact) {
  const score = Number(likelihood || 0) * Number(impact || 0);
  if (score >= 15) return "Red";
  if (score >= 6) return "Amber";
  return "Green";
}

function riskToSpItem(risk) {
  return {
    Title: risk.name || risk.title || "Untitled risk",
    ProjectId: risk.projectId ? (spIdForProjectCode(risk.projectId) || 0) : 0,
    Portfolio: risk.portfolio || "",
    RiskName: risk.name || risk.title || "",
    RiskDescription: risk.description || risk.text || "",
    OwnerName: risk.owner || risk.ownerName || "",
    OwnerEmail: risk.ownerEmail || "",
    Likelihood: Number(risk.likelihood || 0),
    Impact: Number(risk.impact || 0),
    RiskRating: risk.rating || risk.rag || risk.level || riskRatingFromScore(risk.likelihood, risk.impact),
    Category: risk.category || "Operational",
    MitigationPlan: risk.mitigationPlan || "",
    ResponseStrategy: risk.responseStrategy || "Mitigate",
    DueDate: risk.due ? `${risk.due}T00:00:00Z` : null,
    RiskStatus: risk.status || "Open",
    LastReviewedDate: risk.lastReviewedDate ? `${risk.lastReviewedDate}T00:00:00Z` : null,
    ReviewNotesJson: JSON.stringify(risk.reviewNotes || []),
    MitigationsJson: JSON.stringify(risk.mitigations || []),
    ResidualLikelihood: Number(risk.residualLikelihood || 0),
    ResidualImpact: Number(risk.residualImpact || 0)
  };
}

function spItemToRisk(item) {
  const projectId = item.ProjectId ? projectCodeForSpId(item.ProjectId) : "";
  return {
    _spId: item.Id,
    id: "risk" + item.Id,
    projectId,
    portfolio: item.Portfolio || "",
    name: item.RiskName || item.Title || "",
    title: item.RiskName || item.Title || "",
    description: item.RiskDescription || "",
    text: item.RiskDescription || item.RiskName || item.Title || "",
    owner: item.OwnerName || "",
    ownerName: item.OwnerName || "",
    ownerEmail: item.OwnerEmail || "",
    likelihood: Number(item.Likelihood || 0),
    impact: Number(item.Impact || 0),
    rating: item.RiskRating || riskRatingFromScore(item.Likelihood, item.Impact),
    level: item.RiskRating || riskRatingFromScore(item.Likelihood, item.Impact),
    category: item.Category || "",
    mitigationPlan: item.MitigationPlan || "",
    responseStrategy: item.ResponseStrategy || "Mitigate",
    due: item.DueDate ? item.DueDate.slice(0, 10) : "",
    status: item.RiskStatus || "Open",
    lastReviewedDate: item.LastReviewedDate ? item.LastReviewedDate.slice(0, 10) : "",
    reviewNotes: safeJsonParse(item.ReviewNotesJson, []),
    mitigations: safeJsonParse(item.MitigationsJson, []),
    residualLikelihood: Number(item.ResidualLikelihood || 0),
    residualImpact: Number(item.ResidualImpact || 0)
  };
}

/* ---------- mappers: action item (tracker task / checklist item) ---------- */

/* task.notes is an array of {id,author,date,time,text} chat messages
   (openTaskNotesModal, app.js — shared by the Tracker and Weekly Meeting
   popups). Notes went straight into a text column here, so a pre-chat
   record's Notes value is a plain legacy string, not JSON — preserve it as
   a single message instead of silently discarding it (safeJsonParse alone
   would just fall back to [] on a parse failure). */
function parseActionItemNotes(text) {
  if (!text) return [];
  const parsed = safeJsonParse(text, null);
  if (Array.isArray(parsed)) return parsed;
  const legacy = String(text).trim();
  return legacy ? [{ id: uid("nt"), author: "Notes", date: "", time: "", text: legacy }] : [];
}

function actionItemParentDividerSpId(task) {
  if (!task) return 0;
  if (task._parentDividerSpId) return task._parentDividerSpId;
  const dividerId = task.parentDividerId || (task.metadata && task.metadata.parentDividerLocalId) || "";
  if (!dividerId) return 0;
  const siblings = typeof actionItemSiblings === "function" ? actionItemSiblings(task) : null;
  const list = siblings || ((window.AEWTTR.db.ganttTasks || {})[task._projectCode] || []);
  const divider = (list || []).find((t) => t && t.id === dividerId);
  return (divider && divider._spId) || 0;
}

function resolveActionItemProjectSpId(projectCode) {
  if (!projectCode) return 0;
  let spId = spIdForProjectCode(projectCode);
  if (spId) return spId;
  const db = window.AEWTTR && window.AEWTTR.db;
  const proj = ((db && db.projects) || []).find((p) => p && p.id === projectCode);
  if (proj && proj._spId) {
    registerProjectId(proj.id, proj._spId);
    return proj._spId;
  }
  return 0;
}

function actionItemToSpItem(task, extra) {
  extra = extra || {};
  const projectCode = extra.projectCode || task._projectCode || (task.metadata && task.metadata.projectCode) || "";
  const column = extra.column || task._column;
  const isChecklist = (extra.source || task._source) === "Checklist";
  const isDivider = !!(task && (task.itemType === "divider" || task.workItemLevel === "Divider"));
  let meta = null;
  if (isDivider) {
    const from = Object.assign({}, task.metadata || {}, {
      pm: task.pm,
      engineer: task.engineer,
      isso: task.isso,
      rangePoc: task.rangePoc,
      contract: task.contract,
      contractor: task.contractor,
      taskOrder: task.taskOrder,
      fundingType: task.fundingType,
      fiscalYear: task.fiscalYear,
      fundingStatus: task.fundingStatus,
      changeRequestRequired: task.changeRequestRequired,
      configEndItem: task.configEndItem,
      program: task.program,
      product: task.product,
      locations: task.locations,
      portfolios: task.portfolios,
      priority: task.priority,
      rag: task.rag,
      lifecycleStatus: task.lifecycleStatus,
      scope: task.scope,
      objectives: task.objectives,
      description: task.description,
      startDate: task.startDate || task.start,
      dueDate: task.dueDate || task.end || task.due,
      completionDate: task.completionDate,
      clientLocalId: (task.metadata && task.metadata.clientLocalId) || task.id,
      projectCode
    });
    meta = typeof blankDividerMetadata === "function" ? blankDividerMetadata(from) : from;
    task.metadata = meta;
    task.workItemLevel = "Divider";
    task.itemType = "divider";
  } else if (task.metadata && typeof task.metadata === "object") {
    meta = task.metadata;
  }
  return {
    Title: task.title,
    ProjectId: resolveActionItemProjectSpId(projectCode) || 0,
    ParentActionItemId: task.parentActionItemId || 0,
    ParentDividerId: isDivider ? 0 : actionItemParentDividerSpId(task),
    Position: typeof task._sortOrder === "number" ? task._sortOrder : 0,
    SortOrder: typeof task._sortOrder === "number" ? task._sortOrder : 0,
    WorkItemLevel: isDivider ? "Divider" : (task.parentActionItemId ? "Subtask" : (task.workItemLevel || "Task")),
    Source: isChecklist ? "Checklist" : "Tracker",
    AssignedToName: isDivider ? (task.assignee || "") : (task.assignee || (task.owner || "Unassigned")),
    StartDate: (isDivider ? (meta && (meta.startDate || task.start)) : task.start)
      ? `${isDivider ? (meta.startDate || task.start) : task.start}T00:00:00Z`
      : null,
    DueDate: (isDivider ? (meta && (meta.dueDate || task.end || task.due)) : (task.end || task.due))
      ? `${isDivider ? (meta.dueDate || task.end || task.due) : (task.end || task.due)}T00:00:00Z`
      : null,
    ActionStatus: isChecklist ? (CHECKLIST_COLUMN_TO_STATUS[column] || "Not Started") : (task.status || "Not Started"),
    Health: task.health || "On Track",
    BlockedReason: task.blockedReason || "",
    OnHoldReason: task.onHoldReason || "",
    Notes: JSON.stringify(Array.isArray(task.notes) ? task.notes : parseActionItemNotes(task.notes)),
    SubtasksJson: JSON.stringify(task.subtasks || []),
    MetadataJson: (() => {
      const base = meta && typeof meta === "object" ? Object.assign({}, meta) : {};
      // Stable client id + project code survive reload even when SharePoint
      // remaps Id / ProjectId lookups fail temporarily.
      base.clientLocalId = task.id || base.clientLocalId || "";
      if (projectCode) base.projectCode = projectCode;
      if (task.reviewStatus) base.reviewStatus = task.reviewStatus;
      // Milestones live on project dividers only (not tasks / subitems).
      if (isDivider) base.isMilestone = !!task.isMilestone;
      else {
        delete base.isMilestone;
        // Keep a local parent id so tasks still regroup under a divider
        // even if ParentDividerId couldn't resolve yet (divider not on SPO).
        if (task.parentDividerId) base.parentDividerLocalId = task.parentDividerId;
        else delete base.parentDividerLocalId;
      }
      return Object.keys(base).length ? JSON.stringify(base) : "";
    })(),
    IsMilestone: (isDivider && task.isMilestone) ? "Yes" : "No"
  };
}

function spItemToActionItem(item) {
  const metadata = safeJsonParse(item.MetadataJson, null);
  const metaObj = metadata && typeof metadata === "object" ? metadata : null;
  const projectCode = projectCodeForSpId(item.ProjectId)
    || (metaObj && metaObj.projectCode)
    || "";
  const isChecklist = item.Source === "Checklist";
  const isDivider = item.WorkItemLevel === "Divider";
  const task = {
    _spId: item.Id,
    _source: isChecklist ? "Checklist" : "Tracker",
    _projectCode: projectCode,
    _column: isChecklist ? (STATUS_TO_CHECKLIST_COLUMN[item.ActionStatus] || "To Do") : undefined,
    // Keep the original client id across refresh so parentDividerLocalId links work.
    id: (metaObj && metaObj.clientLocalId) || ("ai" + item.Id),
    parentActionItemId: item.ParentActionItemId || 0,
    _parentDividerSpId: item.ParentDividerId || 0,
    _sortOrder: typeof item.Position === "number"
      ? item.Position
      : (typeof item.SortOrder === "number" ? item.SortOrder : 0),
    parentDividerId: "",
    workItemLevel: item.WorkItemLevel || (item.ParentActionItemId ? "Subtask" : "Task"),
    itemType: isDivider ? "divider" : "task",
    title: item.Title || "",
    assignee: item.AssignedToName || (isDivider ? "" : "Unassigned"),
    owner: item.AssignedToName || (isDivider ? "" : "Unassigned"),
    start: item.StartDate ? item.StartDate.slice(0, 10) : "",
    end: item.DueDate ? item.DueDate.slice(0, 10) : "",
    due: item.DueDate ? item.DueDate.slice(0, 10) : "",
    status: item.ActionStatus || "Not Started",
    health: item.Health || "On Track",
    blockedReason: item.BlockedReason || "",
    onHoldReason: item.OnHoldReason || "",
    notes: parseActionItemNotes(item.Notes),
    subtasks: safeJsonParse(item.SubtasksJson, []),
    metadata: metaObj || (isDivider ? {} : undefined),
    isMilestone: false
  };
  if (metaObj && metaObj.reviewStatus) task.reviewStatus = metaObj.reviewStatus;
  if (isDivider) {
    task.isMilestone = item.IsMilestone === "Yes" || item.IsMilestone === true || !!(metaObj && metaObj.isMilestone);
    task.startDate = task.start;
    task.dueDate = task.end;
    // Flatten common mini-project fields onto the divider for editing convenience.
    const meta = task.metadata || {};
    [
      "pm", "engineer", "isso", "rangePoc", "contract", "contractor",
      "taskOrder", "fundingType", "fiscalYear", "fundingStatus",
      "changeRequestRequired", "configEndItem", "program", "product",
      "locations", "portfolios", "priority", "rag", "scope", "objectives",
      "description", "completionDate", "lifecycleStatus", "isMilestone"
    ].forEach((key) => {
      if (task[key] == null && meta[key] != null) task[key] = meta[key];
    });
    if (meta.startDate && !task.startDate) task.startDate = meta.startDate;
    if (meta.dueDate && !task.dueDate) task.dueDate = meta.dueDate;
    if (meta.isMilestone != null) task.isMilestone = !!meta.isMilestone || task.isMilestone;
  }
  return task;
}

/* ---------- mappers: travel ---------- */

function travelRequestToSpItem(r) {
  return {
    Title: r.tripTitle || r.id,
    RequestCode: r.id,
    RequesterName: r.requester,
    RequesterEmail: r.requesterEmail || "",
    FormMode: r.formMode || "Standard",
    RequestType: r.requestType || r.formMode || "Standard",
    TripTitle: r.tripTitle || "",
    TravelersJson: JSON.stringify(r.travelers || []),
    ProjectIdsJson: JSON.stringify((r.projectIds || []).map(spIdForProjectCode).filter(Boolean)),
    Destination: r.destination,
    StartDate: r.start ? `${r.start}T00:00:00Z` : null,
    EndDate: r.end ? `${r.end}T00:00:00Z` : null,
    AllDay: r.allDay !== false,
    StartTime: r.allDay === false ? (r.startTime || "") : "",
    EndTime: r.allDay === false ? (r.endTime || "") : "",
    Purpose: r.purpose || "",
    ImpactIfNotApproved: r.impactIfNotApproved || "",
    AlternativesJson: JSON.stringify(r.alternatives || []),
    EngineeringFormJson: JSON.stringify(r.engineeringForm || null),
    EstCost: r.cost || 0,
    TravelType: r.type || "TDY",
    TravelNotes: r.notes || "",
    ReqStatus: r.status || "Submitted",
    ChargeObject: r.chargeObject || "",
    ChargeObjectStatus: r.chargeObjectStatus || "Pending",
    ChargeObjectAssignedBy: r.chargeObjectAssignedBy || "",
    ChargeObjectAssignedAt: r.chargeObjectAssignedAt || "",
    CustomerConcurrenceStatus: r.customerConcurrenceStatus || "Pending",
    CustomerConcurredBy: r.customerConcurredBy || "",
    CustomerConcurredAt: r.customerConcurredAt || "",
    RequiresConcurrence: !!r.requiresConcurrence,
    DocReviewId: r.docReviewId || "",
    DenyReason: r.denyReason || "",
    ExportFileUrl: r.exportFileUrl || "",
    ExportFileName: r.exportFileName || "",
    ExportMimeType: r.exportMimeType || ""
  };
}

function spItemToTravelRequest(item) {
  return {
    _spId: item.Id,
    id: item.RequestCode || `TR-${item.Id}`,
    requester: item.RequesterName || "",
    requesterEmail: item.RequesterEmail || "",
    formMode: item.FormMode || (item.RequestType === "Engineering" ? "Engineering" : "Standard"),
    requestType: item.RequestType || item.FormMode || "Standard",
    tripTitle: item.TripTitle || item.Title || "",
    travelers: safeJsonParse(item.TravelersJson, []),
    projectIds: safeJsonParse(item.ProjectIdsJson, []).map(projectCodeForSpId).filter(Boolean),
    destination: item.Destination || "",
    start: item.StartDate ? item.StartDate.slice(0, 10) : "",
    end: item.EndDate ? item.EndDate.slice(0, 10) : "",
    allDay: item.AllDay === false || item.AllDay === 0
      ? false
      : (item.AllDay === true || item.AllDay === 1
        ? true
        : !(item.StartTime || item.EndTime)),
    startTime: item.StartTime || "",
    endTime: item.EndTime || "",
    purpose: item.Purpose || "",
    impactIfNotApproved: item.ImpactIfNotApproved || "",
    alternatives: safeJsonParse(item.AlternativesJson, []),
    engineeringForm: safeJsonParse(item.EngineeringFormJson, null),
    cost: item.EstCost || 0,
    type: item.TravelType || "TDY",
    notes: item.TravelNotes || "",
    status: item.ReqStatus || "Submitted",
    chargeObject: item.ChargeObject || "",
    chargeObjectStatus: item.ChargeObjectStatus || (item.ChargeObject ? "Assigned" : "Pending"),
    chargeObjectAssignedBy: item.ChargeObjectAssignedBy || "",
    chargeObjectAssignedAt: item.ChargeObjectAssignedAt || "",
    customerConcurrenceStatus: item.CustomerConcurrenceStatus || (item.ChargeObject ? "Concurred" : "Pending"),
    customerConcurredBy: item.CustomerConcurredBy || "",
    customerConcurredAt: item.CustomerConcurredAt || "",
    requiresConcurrence: !!item.RequiresConcurrence,
    docReviewId: item.DocReviewId || "",
    exportFileUrl: item.ExportFileUrl || "",
    exportFileName: item.ExportFileName || "",
    exportMimeType: item.ExportMimeType || ""
  };
}

function travelDebriefToSpItem(d) {
  return {
    Title: d.id,
    RequestCode: d.trId,
    TravelerName: d.travelerName || "",
    TravelerEmail: d.travelerEmail || "",
    TripDates: d.dates || "",
    Systems: d.systems || "",
    Classification: d.classification || "Non-Contractor-Forward",
    Summary: d.summary || "",
    Followup: d.followup || ""
  };
}

function spItemToTravelDebrief(item) {
  return {
    _spId: item.Id,
    id: item.Title || `DB-${item.Id}`,
    trId: item.RequestCode || "",
    travelerName: item.TravelerName || "",
    travelerEmail: item.TravelerEmail || "",
    dates: item.TripDates || "",
    systems: item.Systems || "",
    classification: item.Classification || "Non-Contractor-Forward",
    summary: item.Summary || "",
    followup: item.Followup || "",
    files: []
  };
}

/* ---------- mappers: team events ---------- */

function teamEventToSpItem(event) {
  return {
    Title: event.title || event.id,
    EventCode: event.id,
    EventDate: event.start ? `${event.start}T00:00:00Z` : null,
    AllDay: event.allDay !== false,
    StartTime: event.allDay === false ? (event.startTime || "") : "",
    EndTime: event.allDay === false ? (event.endTime || "") : "",
    Location: event.location || "",
    Description: event.description || "",
    UpdatedAt: event.updatedAt || new Date().toISOString()
  };
}

function spItemToTeamEvent(item) {
  const start = item.EventDate ? item.EventDate.slice(0, 10) : "";
  return {
    _spId: item.Id,
    id: item.EventCode || `TE-${item.Id}`,
    title: item.Title || "",
    start,
    // Legacy local events carried an end date. Team Events are now always
    // one day, but retaining this normalized value keeps old calendar data
    // harmless during the migration.
    end: start,
    allDay: item.AllDay === false || item.AllDay === 0 ? false : !(item.StartTime || item.EndTime),
    startTime: item.StartTime || "",
    endTime: item.EndTime || "",
    location: item.Location || "",
    description: item.Description || "",
    updatedAt: item.UpdatedAt || ""
  };
}

/* ---------- mappers: document review ---------- */

const DOC_COLUMNS = [
  "Not Started",
  "In Review",
  "Changes Requested",
  "Review Complete",
  "Awaiting Final Pack",
  "Signing in Progress",
  "Signed",
  "Archived",
  "Concurrence",
  "Final / Signed",
  "Awaiting Signature"
];

function docReviewItemToSpItem(doc, column) {
  const revisions = Array.isArray(doc.revisions) ? doc.revisions : [];
  const activeRevision = revisions.find((revision) => revision.id === doc.activeRevisionId) || revisions[0] || null;
  const pendingReviewers = (doc.reviewers || []).filter((reviewer) => (reviewer.decision || "Pending") === "Pending");
  return {
    Title: doc.title,
    ProjectId: doc.projectCode ? (spIdForProjectCode(doc.projectCode) || 0) : 0,
    DocType: doc.type || "",
    DocKind: doc.docKind || "",
    ContractorName: doc.contractorName || "",
    PortfolioOrConfigItem: [...(doc.portfolios || []), ...(doc.configEndItems || [])].join("; ") || doc.portfolioOrConfigItem || "",
    PortfoliosJson: JSON.stringify(doc.portfolios || []),
    ConfigEndItemsJson: JSON.stringify(doc.configEndItems || []),
    ReviewerGroupName: doc.reviewerGroupName || "",
    OwnerName: doc.owner || doc.ownerName || doc.submitter || "",
    SubmitterName: doc.submitter || "",
    SubmitterEmail: doc.submitterEmail || "",
    SubmittedDate: doc.date ? `${doc.date}T00:00:00Z` : null,
    DeadlineDate: doc.deadline ? `${doc.deadline}T00:00:00Z` : null,
    ReviewColumn: column,
    RevisionLabel: doc.revisionLabel || (activeRevision ? `Rev ${activeRevision.number || ""}`.trim() : ""),
    SignatureStatus: doc.signatureStatus || (column === "Final / Signed" ? "Signed" : (pendingReviewers.length ? "Pending Review" : "Pending Signature")),
    PendingReviewersJson: JSON.stringify(pendingReviewers),
    ReviewersJson: JSON.stringify(doc.reviewers || []),
    SignersJson: JSON.stringify(doc.signers || (doc.reviewers || []).filter((reviewer) => reviewer.isSigner).map((reviewer) => ({
      name: reviewer.name || "",
      email: reviewer.email || "",
      order: Number(reviewer.signOrder) || 0,
      signedAt: reviewer.signedAt || "",
      status: reviewer.signedAt ? "Signed" : "Pending",
      notifiedAt: reviewer.signNotifiedAt || ""
    }))),
    CommentsJson: JSON.stringify(doc.comments || []),
    SignedDate: doc.signedDate ? `${doc.signedDate}T00:00:00Z` : null,
    FileUrl: (activeRevision && activeRevision.fileUrl) || doc.fileUrl || "",
    FileName: (activeRevision && activeRevision.fileName) || doc.fileName || "",
    FileType: (activeRevision && activeRevision.fileType) || doc.fileType || "",
    ExtractedText: (activeRevision && activeRevision.extractedText) || doc.extractedText || "",
    RevisionHistoryJson: JSON.stringify(revisions),
    AiSuggestionsJson: JSON.stringify(doc.aiReview || null),
    ReviewActivityJson: JSON.stringify(doc.reviewActivity || []),
    CurrentRevisionNumber: doc.currentRevisionNumber || (activeRevision && activeRevision.number) || 0,
    ActiveRevisionId: doc.activeRevisionId || (activeRevision && activeRevision.id) || "",
    IsArchived: !!doc.isArchived,
    ArchivedDate: doc.archivedDate ? `${doc.archivedDate}T00:00:00Z` : null,
    FullyReviewedDate: doc.fullyReviewedDate ? `${doc.fullyReviewedDate}T00:00:00Z` : null,
    FinalPackUrl: doc.finalPackUrl || "",
    FinalPackFileName: doc.finalPackFileName || "",
    FinalPackUploadedAt: doc.finalPackUploadedAt || null,
    FinalPackUploadedBy: doc.finalPackUploadedBy || "",
    SigningSequenceLocked: !!doc.signingSequenceLocked,
    AwaitingFinalPackNotifiedAt: doc.awaitingFinalPackNotifiedAt || "",
    CurrentSignerNotifiedAt: doc.currentSignerNotifiedAt || ""
  };
}

function spItemToDocReviewItem(item) {
  const revisions = safeJsonParse(item.RevisionHistoryJson, []);
  if (!revisions.length && (item.FileUrl || item.FileName)) {
    revisions.push({
      id: item.ActiveRevisionId || `rev-${item.Id}-1`,
      number: item.CurrentRevisionNumber || 1,
      fileUrl: item.FileUrl || "",
      fileName: item.FileName || "",
      fileType: item.FileType || "",
      extractedText: item.ExtractedText || "",
      uploadedBy: item.SubmitterName || "",
      uploadedDate: item.SubmittedDate ? item.SubmittedDate.slice(0, 10) : "",
      source: "Initial Upload"
    });
  }
  revisions.sort((a, b) => (Number(b && b.number) || 0) - (Number(a && a.number) || 0));
  const activeRevision = revisions.find((revision) => revision.id === item.ActiveRevisionId) || revisions[0] || null;
  return {
    _spId: item.Id,
    _column: DOC_COLUMNS.includes(item.ReviewColumn) ? item.ReviewColumn : "Not Started",
    id: "d" + item.Id,
    title: item.Title || "",
    projectCode: item.ProjectId ? projectCodeForSpId(item.ProjectId) : "",
    type: item.DocType || "",
    docKind: item.DocKind || "",
    contractorName: item.ContractorName || "",
    portfolioOrConfigItem: item.PortfolioOrConfigItem || "",
    portfolios: safeJsonParse(item.PortfoliosJson, []),
    configEndItems: safeJsonParse(item.ConfigEndItemsJson, []),
    reviewerGroupName: item.ReviewerGroupName || "",
    owner: item.OwnerName || "",
    ownerName: item.OwnerName || "",
    submitter: item.SubmitterName || "",
    submitterEmail: item.SubmitterEmail || "",
    date: item.SubmittedDate ? item.SubmittedDate.slice(0, 10) : "",
    deadline: item.DeadlineDate ? item.DeadlineDate.slice(0, 10) : "",
    revisionLabel: item.RevisionLabel || "",
    signatureStatus: item.SignatureStatus || "",
    pendingReviewers: safeJsonParse(item.PendingReviewersJson, []),
    reviewers: safeJsonParse(item.ReviewersJson, []),
    signers: safeJsonParse(item.SignersJson, []),
    comments: safeJsonParse(item.CommentsJson, []),
    signedDate: item.SignedDate ? item.SignedDate.slice(0, 10) : undefined,
    fileUrl: (activeRevision && activeRevision.fileUrl) || item.FileUrl || "",
    fileName: (activeRevision && activeRevision.fileName) || item.FileName || "",
    fileType: (activeRevision && activeRevision.fileType) || item.FileType || "",
    extractedText: (activeRevision && activeRevision.extractedText) || item.ExtractedText || "",
    revisions: revisions,
    activeRevisionId: (activeRevision && activeRevision.id) || item.ActiveRevisionId || "",
    currentRevisionNumber: Number(item.CurrentRevisionNumber) || (activeRevision && activeRevision.number) || (revisions.length ? revisions[0].number : 0),
    aiReview: safeJsonParse(item.AiSuggestionsJson, null),
    reviewActivity: safeJsonParse(item.ReviewActivityJson, []),
    isArchived: !!item.IsArchived,
    archivedDate: item.ArchivedDate ? item.ArchivedDate.slice(0, 10) : "",
    fullyReviewedDate: item.FullyReviewedDate ? item.FullyReviewedDate.slice(0, 10) : "",
    finalPackUrl: item.FinalPackUrl || "",
    finalPackFileName: item.FinalPackFileName || "",
    finalPackUploadedAt: item.FinalPackUploadedAt || "",
    finalPackUploadedBy: item.FinalPackUploadedBy || "",
    signingSequenceLocked: !!item.SigningSequenceLocked,
    awaitingFinalPackNotifiedAt: item.AwaitingFinalPackNotifiedAt || "",
    currentSignerNotifiedAt: item.CurrentSignerNotifiedAt || ""
  };
}

/* ---------- mappers: doc reviewer groups ---------- */

function docReviewerGroupToSpItem(group) {
  return {
    Title: group.name || "",
    GroupName: group.name || "",
    MembersJson: JSON.stringify(group.members || []),
    GroupSource: group.source === "project" ? "project" : "manual",
    ProjectId: group.projectId || "",
    CreatedByEmail: group.createdByEmail || "",
    CreatedDate: group.createdDate ? `${group.createdDate}T00:00:00Z` : null
  };
}

function spItemToDocReviewerGroup(item) {
  return {
    _spId: item.Id,
    id: "rg" + item.Id,
    name: item.GroupName || item.Title || "",
    members: safeJsonParse(item.MembersJson, []),
    source: item.GroupSource === "project" ? "project" : "manual",
    projectId: item.ProjectId || "",
    createdByEmail: item.CreatedByEmail || "",
    createdDate: item.CreatedDate ? item.CreatedDate.slice(0, 10) : ""
  };
}

/* ---------- mappers: meeting session + rocks ---------- */

/* A meeting document is a typed list of editable blocks. Normalize it at the
   persistence boundary so an incomplete SharePoint value cannot break the
   notes editor on load. `null` means no document was ever saved; an empty
   array is a valid, intentionally empty document. */
function normalizeMeetingDocBlocks(raw) {
  if (!Array.isArray(raw)) return null;
  const types = ["action", "bullet", "h2", "para"];
  return raw.filter((block) => block && typeof block === "object").map((block, index) => ({
    id: String(block.id || `doc-block-${index + 1}`),
    type: types.includes(block.type) ? block.type : "para",
    content: String(block.content || ""),
    done: !!block.done
  }));
}

function parseMeetingSessionDocBlocks(raw) {
  if (raw === undefined || raw === null || raw === "") return null;
  return normalizeMeetingDocBlocks(typeof raw === "string" ? safeJsonParse(raw, null) : raw);
}

function parseMeetingSessionNotes(rawNotes) {
  if (!rawNotes) return {
    notesFeed: [],
    docBlocks: null,
    sessionStatus: null,
    guests: [],
    capturedNotes: [],
    ganttChanges: [],
    minutesLog: [],
    tasksUpdated: 0,
    tasksReviewed: 0,
    startedAt: null,
    endedAt: null,
    guestAttendees: []
  };
  const parsed = safeJsonParse(rawNotes, null);
  if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
    return {
      notesFeed: Array.isArray(parsed.notesFeed) ? parsed.notesFeed : [],
      docBlocks: normalizeMeetingDocBlocks(parsed.docBlocks),
      docHtml: typeof parsed.docHtml === "string" ? parsed.docHtml : "",
      sessionStatus: parsed.sessionStatus || null,
      guests: Array.isArray(parsed.guests) ? parsed.guests : [],
      capturedNotes: Array.isArray(parsed.capturedNotes) ? parsed.capturedNotes : [],
      ganttChanges: Array.isArray(parsed.ganttChanges) ? parsed.ganttChanges : [],
      minutesLog: Array.isArray(parsed.minutesLog) ? parsed.minutesLog : [],
      tasksUpdated: Number(parsed.tasksUpdated) || 0,
      tasksReviewed: Number(parsed.tasksReviewed) || 0,
      startedAt: parsed.startedAt || null,
      endedAt: parsed.endedAt || null,
      guestAttendees: Array.isArray(parsed.guestAttendees) ? parsed.guestAttendees : [],
      scheduledTitle: parsed.scheduledTitle || null,
      scheduledTime: parsed.scheduledTime || null,
      agendaText: parsed.agendaText || null,
      createdBy: parsed.createdBy || null,
      createdAt: parsed.createdAt || null
    };
  }
  if (Array.isArray(parsed)) {
    return {
      notesFeed: parsed,
      docBlocks: null,
      sessionStatus: null,
      guests: [],
      capturedNotes: [],
      ganttChanges: [],
      minutesLog: [],
      tasksUpdated: 0,
      tasksReviewed: 0,
      startedAt: null,
      endedAt: null,
      guestAttendees: []
    };
  }
  const text = String(rawNotes).trim();
  if (text) {
    return {
      notesFeed: [{
        id: (typeof uid === "function" ? uid("note") : "note-import"),
        author: "Meeting notes",
        date: new Date().toISOString().slice(0, 10),
        time: "",
        text
      }],
      docBlocks: null,
      sessionStatus: null,
      guests: [],
      capturedNotes: [],
      ganttChanges: [],
      minutesLog: [],
      tasksUpdated: 0,
      tasksReviewed: 0,
      startedAt: null,
      endedAt: null,
      guestAttendees: []
    };
  }
  return {
    notesFeed: [],
    docBlocks: null,
    sessionStatus: null,
    guests: [],
    capturedNotes: [],
    ganttChanges: [],
    minutesLog: [],
    tasksUpdated: 0,
    tasksReviewed: 0,
    startedAt: null,
    endedAt: null,
    guestAttendees: []
  };
}

function meetingSessionToSpItem(session, scopeProjectCode) {
  const sessionStatus = session.sessionStatus || "active";
  const docBlocks = normalizeMeetingDocBlocks(session.docBlocks);
  // Existing sites may still have the original Active/Ended choice column.
  // Keep the richer planned state in Notes while writing an accepted legacy
  // value, so agenda planning works before an admin reruns SharePoint Setup.
  const storedSessionStatus = sessionStatus === "planned" ? "active" : sessionStatus;
  return {
    Title: session.id,
    ProjectId: scopeProjectCode ? (spIdForProjectCode(scopeProjectCode) || 0) : 0,
    MeetingSeries: session.series || session.meetingSeries || "",
    SeriesKey: session.seriesKey || "",
    MeetingDate: session.date ? `${session.date}T00:00:00Z` : null,
    AgendaJson: JSON.stringify(session.agenda || []),
    // New sites get a purpose-built column. Notes also contains the document
    // so existing sites without that column retain it after refresh.
    DocBlocksJson: JSON.stringify(docBlocks || []),
    DocHtml: session.docHtml || "",
    Attendees: JSON.stringify(session.attendees || []),
    AttendanceJson: JSON.stringify(session.attendance || {}),
    Notes: JSON.stringify({
      notesFeed: session.notesFeed || [],
      docBlocks: docBlocks || [],
      docHtml: session.docHtml || "",
      sessionStatus,
      guests: session.guests || [],
      capturedNotes: session.capturedNotes || [],
      ganttChanges: session.ganttChanges || [],
      minutesLog: session.minutesLog || [],
      tasksUpdated: session.tasksUpdated || 0,
      tasksReviewed: session.tasksReviewed || 0,
      startedAt: session.startedAt || null,
      endedAt: session.endedAt || null,
      guestAttendees: session.guestAttendees || [],
      scheduledTitle: session.scheduledTitle || null,
      scheduledTime: session.scheduledTime || null,
      agendaText: session.agendaText || null,
      createdBy: session.createdBy || null,
      createdAt: session.createdAt || null
    }),
    DecisionsJson: JSON.stringify(session.decisions || []),
    ActionItemsJson: JSON.stringify(session.actionItems || []),
    CarryForwardJson: JSON.stringify(session.carryForward || []),
    RiskChangesJson: JSON.stringify(session.riskChanges || []),
    ProjectStatusChangesJson: JSON.stringify(session.projectStatusChanges || []),
    ActivityJson: JSON.stringify(session.activity || []),
    SessionStatus: storedSessionStatus
  };
}

function spItemToMeetingSession(item) {
  const parsedNotes = parseMeetingSessionNotes(item.Notes);
  const docBlocks = parseMeetingSessionDocBlocks(item.DocBlocksJson);
  const persistedDocBlocks = docBlocks !== null ? docBlocks : parsedNotes.docBlocks;
  let sessionStatus = ["planned", "active", "ended"].includes(item.SessionStatus) ? item.SessionStatus : null;
  if (["planned", "active", "ended"].includes(parsedNotes.sessionStatus)) sessionStatus = parsedNotes.sessionStatus;
  const session = {
    _spId: item.Id,
    _projectCode: item.ProjectId ? projectCodeForSpId(item.ProjectId) : "",
    id: item.Title || ("wm" + item.Id),
    date: item.MeetingDate ? item.MeetingDate.slice(0, 10) : "",
    series: item.MeetingSeries || "",
    meetingSeries: item.MeetingSeries || "",
    seriesKey: item.SeriesKey || "",
    agenda: safeJsonParse(item.AgendaJson, []),
    notesFeed: parsedNotes.notesFeed,
    guests: parsedNotes.guests,
    capturedNotes: parsedNotes.capturedNotes,
    ganttChanges: parsedNotes.ganttChanges,
    minutesLog: parsedNotes.minutesLog,
    tasksUpdated: parsedNotes.tasksUpdated,
    tasksReviewed: parsedNotes.tasksReviewed,
    startedAt: parsedNotes.startedAt,
    endedAt: parsedNotes.endedAt,
    guestAttendees: parsedNotes.guestAttendees,
    decisions: safeJsonParse(item.DecisionsJson, []),
    actionItems: safeJsonParse(item.ActionItemsJson, []),
    carryForward: safeJsonParse(item.CarryForwardJson, []),
    riskChanges: safeJsonParse(item.RiskChangesJson, []),
    projectStatusChanges: safeJsonParse(item.ProjectStatusChangesJson, []),
    attendees: safeJsonParse(item.Attendees, []),
    attendance: safeJsonParse(item.AttendanceJson, {}),
    activity: safeJsonParse(item.ActivityJson, []),
    sessionStatus: sessionStatus || "active",
    docHtml: item.DocHtml || parsedNotes.docHtml || "",
    scheduledTitle: parsedNotes.scheduledTitle || null,
    scheduledTime: parsedNotes.scheduledTime || null,
    agendaText: parsedNotes.agendaText || null,
    createdBy: parsedNotes.createdBy || null,
    createdAt: parsedNotes.createdAt || null
  };
  // Leave legacy sessions without a saved document alone so the editor can
  // migrate their already-persisted AgendaJson instead of showing it blank.
  if (persistedDocBlocks !== null) session.docBlocks = persistedDocBlocks;
  return session;
}

/* ---------- mappers: audit log ---------- */

function auditLogToSpItem(entry) {
  return {
    Title: (entry.summary || "User action").slice(0, 255),
    ActionTime: entry.ts || new Date().toISOString(),
    ActorEmail: entry.actorEmail || "",
    ActorName: entry.actorName || "",
    ActorRole: entry.actorRole || "",
    Action: entry.action || "Other",
    Area: entry.area || "System",
    Summary: entry.summary || "",
    DetailJson: entry.detail == null ? "" : JSON.stringify(entry.detail),
    Route: entry.route || "",
    RecordId: entry.recordId || ""
  };
}

function spItemToAuditLog(item) {
  return {
    _spId: item.Id,
    id: `LOG-${item.Id}`,
    ts: item.ActionTime || item.Created || "",
    actorEmail: item.ActorEmail || "",
    actorName: item.ActorName || "",
    actorRole: item.ActorRole || "",
    action: item.Action || "Other",
    area: item.Area || "System",
    summary: item.Summary || item.Title || "",
    detail: safeJsonParse(item.DetailJson, null),
    route: item.Route || "",
    recordId: item.RecordId || ""
  };
}

/* ---------- mappers: issue reports ---------- */

function compactIssueJson(value, maxLength) {
  const limit = maxLength || 26000;
  let json = "";
  try { json = JSON.stringify(value == null ? {} : value); } catch (e) { json = JSON.stringify({ unavailable: String(e) }); }
  if (json.length <= limit) return json;
  return JSON.stringify({ truncated: true, preview: json.slice(0, Math.max(0, limit - 80)) });
}

function issueScreenshotForSp(value) {
  const screenshot = String(value || "");
  return /^data:image\/(?:jpeg|png|webp);base64,/i.test(screenshot) && screenshot.length <= 52000 ? screenshot : "";
}

function issueScreenshotPathForSp(value, maxLength) {
  const path = String(value || "").trim();
  return /^(?:https?:\/\/|\/)/i.test(path) ? path.slice(0, maxLength || 4000) : "";
}

function issueToSpItem(issue) {
  const pageTitle = String(issue.pageTitle || "PULSE page").slice(0, 180);
  const issueType = issue.type || "Other";
  return {
    Title: String(issue.title || `${issueType}: ${pageTitle}`).slice(0, 255),
    IssueCode: String(issue.id || "").slice(0, 120),
    IssueType: issueType,
    IssueStatus: issue.status || "New",
    ReportedByName: String(issue.reportedBy || "").slice(0, 255),
    ReporterEmail: String(issue.reporterEmail || "").slice(0, 255),
    PageTitle: pageTitle,
    Route: String(issue.route || "").slice(0, 255),
    PageUrl: String(issue.pageUrl || "").slice(0, 4000),
    OccurredAt: issue.createdAt || new Date().toISOString(),
    Description: String(issue.description || "").slice(0, 18000),
    ExpectedBehavior: String(issue.expectedBehavior || "").slice(0, 12000),
    AdditionalContext: String(issue.additionalContext || "").slice(0, 12000),
    ErrorCodesJson: compactIssueJson(issue.errorCodes || [], 10000),
    LogsJson: compactIssueJson(issue.logs || [], 26000),
    DiagnosticsJson: compactIssueJson(issue.diagnostics || {}, 26000),
    ScreenshotDataUrl: issue.screenshotFileUrl ? "" : issueScreenshotForSp(issue.screenshotDataUrl),
    ScreenshotFileUrl: issueScreenshotPathForSp(issue.screenshotFileUrl, 4000),
    ScreenshotServerRelativeUrl: issueScreenshotPathForSp(issue.screenshotServerRelativeUrl, 4000),
    ScreenshotFileName: String(issue.screenshotFileName || "").slice(0, 255),
    ScreenshotMetaJson: compactIssueJson(issue.screenshotMeta || {}, 5000),
    ResolutionNote: String(issue.resolutionNote || "").slice(0, 12000)
  };
}

function spItemToIssue(item) {
  return {
    _spId: item.Id,
    id: item.IssueCode || `ISS-${item.Id}`,
    title: item.Title || "Issue report",
    type: item.IssueType || "Other",
    status: item.IssueStatus || "New",
    reportedBy: item.ReportedByName || "",
    reporterEmail: item.ReporterEmail || "",
    pageTitle: item.PageTitle || "",
    route: item.Route || "",
    pageUrl: item.PageUrl || "",
    createdAt: item.OccurredAt || item.Created || "",
    description: item.Description || "",
    expectedBehavior: item.ExpectedBehavior || "",
    additionalContext: item.AdditionalContext || "",
    errorCodes: safeJsonParse(item.ErrorCodesJson, []),
    logs: safeJsonParse(item.LogsJson, []),
    diagnostics: safeJsonParse(item.DiagnosticsJson, {}),
    screenshotDataUrl: issueScreenshotForSp(item.ScreenshotDataUrl),
    screenshotFileUrl: issueScreenshotPathForSp(item.ScreenshotFileUrl, 4000),
    screenshotServerRelativeUrl: issueScreenshotPathForSp(item.ScreenshotServerRelativeUrl, 4000),
    screenshotFileName: item.ScreenshotFileName || "",
    screenshotMeta: safeJsonParse(item.ScreenshotMetaJson, {}),
    resolutionNote: item.ResolutionNote || ""
  };
}

/* ---------- mappers: tickets (backed by PULSE Issues list) ---------- */

function _issueTypeToTicketType(t) {
  if (!t) return "Question";
  if (t === "Bug / error" || t === "Data / saving") return "Bug";
  if (t === "Access / permissions") return "Access";
  if (t === "Display / usability") return "Platform";
  if (t === "Feature Request") return "Feature Request";
  if (t === "Blocker") return "Blocker";
  return "Question";
}
function _ticketTypeToIssueType(t) {
  if (!t) return "Other";
  if (t === "Bug") return "Bug / error";
  if (t === "Blocker") return "Bug / error";
  if (t === "Access") return "Access / permissions";
  if (t === "Platform") return "Display / usability";
  if (t === "Feature Request") return "Feature Request";
  return "Other";
}
function _issueStatusToTicketStatus(s) {
  if (!s) return "Open";
  if (s === "In Progress") return "In Progress";
  if (s === "Resolved" || s === "Closed") return "Resolved";
  return "Open";
}
function _ticketStatusToIssueStatus(s) {
  if (!s) return "New";
  if (s === "In Progress") return "In Progress";
  if (s === "Resolved") return "Resolved";
  return "New";
}

function ticketToSpItem(t) {
  return {
    Title: String(t.title || "Support ticket").slice(0, 255),
    IssueCode: String(t.id || "").slice(0, 120),
    IssueType: _ticketTypeToIssueType(t.type),
    IssueStatus: _ticketStatusToIssueStatus(t.status),
    ReportedByName: String(t.reporter || "").slice(0, 255),
    ReporterEmail: String(t.reporterEmail || "").slice(0, 255),
    OccurredAt: t.opened || new Date().toISOString(),
    Route: String(t.project || "").slice(0, 255),
    Description: String(t.detail || "").slice(0, 18000),
    ExpectedBehavior: String(t.workaround || "").slice(0, 12000),
    AdditionalContext: String(t.esdp || "").slice(0, 12000),
    LogsJson: JSON.stringify(t.updates || []).slice(0, 26000)
  };
}

function spItemToTicket(item) {
  return {
    _spId: item.Id,
    id: item.IssueCode || `ISS-${String(item.Id).padStart(4, "0")}`,
    title: item.Title || "Support ticket",
    type: _issueTypeToTicketType(item.IssueType),
    status: _issueStatusToTicketStatus(item.IssueStatus),
    reporter: item.ReportedByName || "",
    reporterEmail: item.ReporterEmail || "",
    opened: item.OccurredAt || item.Created || "",
    project: item.Route || "",
    detail: item.Description || "",
    workaround: item.ExpectedBehavior || "",
    esdp: item.AdditionalContext || "",
    updates: safeJsonParse(item.LogsJson, []),
    affected: item.Route ? [item.Route] : []
  };
}

/* ---------- mappers: notification config ---------- */

function notificationConfigToSpItem(cfg) {
  return {
    Title: "Default",
    TeamsEnabled: cfg.teamsEnabled !== false,
    AlsoSendEmail: cfg.alsoSendEmail !== false
  };
}

function spItemToNotificationConfig(item) {
  return {
    _spId: item.Id,
    teamsEnabled: item.TeamsEnabled !== false,
    alsoSendEmail: item.AlsoSendEmail !== false
  };
}

function notificationConfigDefaults() {
  return { teamsEnabled: true, alsoSendEmail: true };
}

/* ---------- mappers: application settings ---------- */

function appSettingsDefaults() {
  return { cuiMarkingEnabled: false };
}

function appSettingsToSpItem(cfg) {
  return {
    Title: "Default",
    CuiMarkingEnabled: !!(cfg && cfg.cuiMarkingEnabled)
  };
}

function spItemToAppSettings(item) {
  return {
    _spId: item.Id,
    cuiMarkingEnabled: item.CuiMarkingEnabled === true
  };
}

/* ---------- mappers: location config ---------- */

function locationConfigToSpItem(cfg) {
  return {
    Title: "Default",
    LocationsJson: JSON.stringify(cfg.locations || []),
    PortfolioCatalogJson: JSON.stringify(cfg.portfolios || []),
    ContractorCatalogJson: JSON.stringify(cfg.contractors || []),
    ContractCatalogJson: JSON.stringify(cfg.contracts || []),
    ConfigEndItemCatalogJson: JSON.stringify(cfg.configEndItems || []),
    HideUnaffiliatedPeople: cfg.hideUnaffiliatedPeople !== false
  };
}

function spItemToLocationConfig(item) {
  const locations = safeJsonParse(item.LocationsJson, []);
  const portfolios = safeJsonParse(item.PortfolioCatalogJson, []);
  const contractors = safeJsonParse(item.ContractorCatalogJson, []);
  const contracts = safeJsonParse(item.ContractCatalogJson, []);
  const configEndItems = safeJsonParse(item.ConfigEndItemCatalogJson, []);
  return {
    _spId: item.Id,
    locations: Array.isArray(locations) ? locations.map((n) => String(n == null ? "" : n).trim()).filter(Boolean) : [],
    portfolios: Array.isArray(portfolios) ? portfolios.map((n) => String(n == null ? "" : n).trim()).filter(Boolean) : [],
    contractors: Array.isArray(contractors) ? contractors.map((n) => String(n == null ? "" : n).trim()).filter(Boolean) : [],
    contracts: Array.isArray(contracts) ? contracts.map((n) => String(n == null ? "" : n).trim()).filter(Boolean) : [],
    configEndItems: Array.isArray(configEndItems) ? configEndItems.map((n) => String(n == null ? "" : n).trim()).filter(Boolean) : [],
    hideUnaffiliatedPeople: item.HideUnaffiliatedPeople !== false
  };
}

/* ---------- mappers: app roles -> db.members view ---------- */

function appRoleToMember(record) {
  return {
    id: "u" + record.Id,
    _spId: record.Id,
    name: record.UserDisplayName || record.UserEmail || "Unknown",
    email: record.UserEmail || "",
    loginName: record.LoginName || "",
    role: record.JobTitle || record.Role || "Member",
    isAdmin: record.Role === "Admin",
    isFinanceAdmin: record.Role === "Admin" || record.Role === "Finance Admin",
    isDocAdmin: record.Role === "Admin" || record.Role === "Document Admin",
    hiddenFromMeetings: sharePointAdapter.isHiddenFromMeetings(record),
    themeMode: record.ThemeMode || "",
    notificationPrefs: safeJsonParse(record.NotificationPrefsJson, null),
    spo: record.SpoAccess || "None",
    powerbi: record.PowerBiAccess || "None",
    lastLogin: ""
  };
}

let _roleRecords = [];
function rebuildMembersView() {
  const members = _roleRecords
    .filter(sharePointAdapter.isRoleRecordActive)
    .map(appRoleToMember)
    .filter((m) => typeof isHiddenServicePrincipal !== "function" || !isHiddenServicePrincipal(m));
  if (window.AEWTTR && window.AEWTTR.db) {
    window.AEWTTR.db.members = members;
  }
  return members;
}

async function refreshMembersFromSharePoint(siteUrl) {
  if (!siteUrl || typeof sharePointAdapter.getRoleRecords !== "function") return rebuildMembersView();
  try {
    _roleRecords = await sharePointAdapter.getRoleRecords(siteUrl);
  } catch (e) { /* keep prior _roleRecords */ }
  return rebuildMembersView();
}

/* ---------- bulk load ---------- */

function buildSharePointBackedEmptyDb() {
  return JSON.parse(JSON.stringify({
    projects: [],
    projectExtra: {},
    projectPeople: {},
    projectImports: {},
    projectFinance: {},
    ganttTasks: {},
    engChecklists: {},
    travelRequests: [],
    teamEvents: [],
    debriefs: [],
    docs: DOC_COLUMNS.reduce((bags, column) => {
      bags[column] = [];
      return bags;
    }, {}),
    docReviewerGroups: [],
    groups: [],
    checklistBoards: [],
    checklistTasks: {},
    tickets: [],
    weeklyMeeting: { roster: [], sessions: [], currentSession: null, meetingStatus: "idle", projectMeetings: {} },
    auditLog: [],
    issues: [],
    notificationConfig: null,
    appSettings: null,
    locationConfig: null,
    members: []
  }));
}

function getSerializableDb(db) {
  const copy = JSON.parse(JSON.stringify(db || {}));
  delete copy.user;
  delete copy.members;
  return copy;
}

let _lastResolvedListUrl = {}; // listTitle -> serverRelativeUrl, for the diagnostics report
async function listOrEmpty(siteUrl, listTitle) {
  try {
    const [items, info] = await Promise.all([
      sharePointAdapter.getItems(siteUrl, listTitle),
      sharePointAdapter.getListInfo(siteUrl, listTitle).catch(() => null)
    ]);
    _lastResolvedListUrl[listTitle] = (info && info.serverRelativeUrl) || "";
    return items;
  } catch (e) {
    if (e && e.status === 404) { _lastResolvedListUrl[listTitle] = ""; return []; } // list not created yet — setup will make it
    throw e;
  }
}

function isArchivedProjectItem(item) {
  const value = item && item.IsArchived;
  if (value === true || value === 1) return true;
  const normalized = String(value == null ? "" : value).trim().toLowerCase();
  return normalized === "true" || normalized === "yes" || normalized === "1";
}

/* Every list read this boot/refresh recorded here — surfaced in the Admin >
   SharePoint Setup diagnostics panel and (for real lists, not field-schema
   probes) in the visible boot console, so "why is X empty" never requires
   digging through Admin > Logs. */
function recordLoadReport(label, ok, count, error) {
  if (!window.AEWTTR) return;
  window.AEWTTR.lastLoadReport = window.AEWTTR.lastLoadReport || [];
  window.AEWTTR.lastLoadReport.push({
    label, ok, count,
    serverRelativeUrl: _lastResolvedListUrl[label] || "",
    error: error ? String((error && error.friendly) || (error && error.message) || error) : ""
  });
}

async function loadListWithFallback(label, loader, fallbackValue) {
  try {
    const value = await loader();
    const count = Array.isArray(value) ? value.length : undefined;
    recordLoadReport(label, true, count, null);
    if (typeof pushDebugLog === "function") {
      pushDebugLog("success", "sharepoint-load", `${label} loaded`, { count });
    }
    // Per-list successes are NOT written to bootLog — with ~17 lists/field
    // probes loaded every boot and every "Reload List Data" click, that would
    // make the startup log scroll forever. Only failures + one final summary
    // line (see loadAllFromSharePoint) go to bootLog; full per-list detail
    // always lives in the Admin > SharePoint Setup load report.
    return value;
  } catch (error) {
    recordLoadReport(label, false, null, error);
    if (typeof pushDebugLog === "function") {
      pushDebugLog("error", "sharepoint-load", `${label} failed to load; using fallback`, error);
    }
    if (typeof bootLog === "function") {
      bootLog(`${label}: FAILED to load — ${String((error && error.friendly) || (error && error.message) || error)}`, "error");
    }
    return fallbackValue;
  }
}

function bySortOrderThenId(a, b) {
  const sa = typeof a.Position === "number"
    ? a.Position
    : (typeof a.SortOrder === "number" ? a.SortOrder : 1e9);
  const sb = typeof b.Position === "number"
    ? b.Position
    : (typeof b.SortOrder === "number" ? b.SortOrder : 1e9);
  return sa - sb || (a.Id - b.Id);
}

async function loadAllFromSharePoint(siteUrl, options) {
  options = options || {};
  const fast = !!options.fast;
  const a = sharePointAdapter;
  if (window.AEWTTR && !fast) window.AEWTTR.lastLoadReport = [];
  if (fast) _lastResolvedListUrl = {};
  /* Background refreshes skip the verbose load report but still need field
     definitions so SharePoint internal column names remap to schema names.
     getExistingFieldDefinitions is cached in sharepoint-adapter after the
     first full boot — fast refresh reuses that cache with no extra REST. */
  const fieldProbe = (label, listTitle) => {
    const loader = () => a.getExistingFieldDefinitions(siteUrl, listTitle);
    if (fast) return loader().catch(() => []);
    return loadListWithFallback(label, loader, []);
  };
  const issueDataPromise = Promise.all([
    loadListWithFallback(SP_LISTS.issue, () => sharePointAdapter.getItems(siteUrl, SP_LISTS.issue, { top: 500, orderby: "OccurredAt desc" }), []),
    loadListWithFallback(SP_LISTS.appSettings, () => listOrEmpty(siteUrl, SP_LISTS.appSettings), []),
    fieldProbe("PULSE Issues fields", SP_LISTS.issue),
    fieldProbe("PULSE App Settings fields", SP_LISTS.appSettings)
  ]);
  const [roleRecords, projectItems, riskItems, actionItems, travelItems, debriefItems, teamEventItems, docItems, docReviewGroupItems, meetingItems, auditItems, notificationConfigItems, locationConfigItems, ticketItems, projectFieldDefinitions, riskFieldDefinitions, actionFieldDefinitions, travelFieldDefinitions, debriefFieldDefinitions, teamEventFieldDefinitions, docFieldDefinitions, docReviewGroupFieldDefinitions, meetingFieldDefinitions, auditFieldDefinitions, notificationConfigFieldDefinitions, locationConfigFieldDefinitions, ticketFieldDefinitions, roleFieldDefinitions] = await Promise.all([
    loadListWithFallback("PULSE App Roles", () => a.getRoleRecords(siteUrl), []),
    loadListWithFallback(SP_LISTS.project, () => listOrEmpty(siteUrl, SP_LISTS.project), []),
    loadListWithFallback(SP_LISTS.risk, () => listOrEmpty(siteUrl, SP_LISTS.risk), []),
    loadListWithFallback(SP_LISTS.actionItem, () => listOrEmpty(siteUrl, SP_LISTS.actionItem), []),
    loadListWithFallback(SP_LISTS.travelRequest, () => listOrEmpty(siteUrl, SP_LISTS.travelRequest), []),
    loadListWithFallback(SP_LISTS.travelDebrief, () => listOrEmpty(siteUrl, SP_LISTS.travelDebrief), []),
    loadListWithFallback(SP_LISTS.teamEvent, () => listOrEmpty(siteUrl, SP_LISTS.teamEvent), []),
    loadListWithFallback(SP_LISTS.docReviewItem, () => listOrEmpty(siteUrl, SP_LISTS.docReviewItem), []),
    loadListWithFallback(SP_LISTS.docReviewerGroup, () => listOrEmpty(siteUrl, SP_LISTS.docReviewerGroup), []),
    loadListWithFallback(SP_LISTS.meetingSession, () => listOrEmpty(siteUrl, SP_LISTS.meetingSession), []),
    loadListWithFallback(SP_LISTS.auditLog, () => sharePointAdapter.getItems(siteUrl, SP_LISTS.auditLog, { top: 500, orderby: "ActionTime desc" }), []),
    loadListWithFallback(SP_LISTS.notificationConfig, () => listOrEmpty(siteUrl, SP_LISTS.notificationConfig), []),
    loadListWithFallback(SP_LISTS.locationConfig, () => listOrEmpty(siteUrl, SP_LISTS.locationConfig), []),
    loadListWithFallback(SP_LISTS.ticket, () => sharePointAdapter.getItems(siteUrl, SP_LISTS.ticket, { top: 500, orderby: "OccurredAt desc" }), []),
    fieldProbe("PULSE Projects fields", SP_LISTS.project),
    fieldProbe("PULSE Risks fields", SP_LISTS.risk),
    fieldProbe("PULSE Action Items fields", SP_LISTS.actionItem),
    fieldProbe("PULSE Travel Requests fields", SP_LISTS.travelRequest),
    fieldProbe("PULSE Travel Debriefs fields", SP_LISTS.travelDebrief),
    fieldProbe("PULSE Team Events fields", SP_LISTS.teamEvent),
    fieldProbe("PULSE Document Review fields", SP_LISTS.docReviewItem),
    fieldProbe("PULSE Doc Reviewer Groups fields", SP_LISTS.docReviewerGroup),
    fieldProbe("PULSE Meetings fields", SP_LISTS.meetingSession),
    fieldProbe("PULSE Audit Log fields", SP_LISTS.auditLog),
    fieldProbe("PULSE Notification Config fields", SP_LISTS.notificationConfig),
    fieldProbe("PULSE Location Config fields", SP_LISTS.locationConfig),
    fieldProbe("PULSE Tickets fields", SP_LISTS.ticket),
    fieldProbe("PULSE App Roles fields", SP_LISTS.appRole)
  ]);
  const [issueItems, appSettingsItems, issueFieldDefinitions, appSettingsFieldDefinitions] = await issueDataPromise;
  _roleRecords = normalizeItemsToSchema(roleRecords, roleFieldDefinitions, SP_LISTS.appRole)
    .map((item, index) => sharePointAdapter.normalizeRoleRecord(item, index));
  const db = buildSharePointBackedEmptyDb();

  _projectIdMap = { byCode: {}, byId: {} };
  const normalizedProjectItems = normalizeItemsToSchema(projectItems, projectFieldDefinitions, SP_LISTS.project);
  const normalizedRiskItems = normalizeItemsToSchema(riskItems, riskFieldDefinitions, SP_LISTS.risk);
  const normalizedActionItems = normalizeItemsToSchema(actionItems, actionFieldDefinitions, SP_LISTS.actionItem);
  const normalizedTravelItems = normalizeItemsToSchema(travelItems, travelFieldDefinitions, SP_LISTS.travelRequest);
  const normalizedDebriefItems = normalizeItemsToSchema(debriefItems, debriefFieldDefinitions, SP_LISTS.travelDebrief);
  const normalizedTeamEventItems = normalizeItemsToSchema(teamEventItems, teamEventFieldDefinitions, SP_LISTS.teamEvent);
  const normalizedDocItems = normalizeItemsToSchema(docItems, docFieldDefinitions, SP_LISTS.docReviewItem);
  const normalizedDocReviewGroupItems = normalizeItemsToSchema(docReviewGroupItems, docReviewGroupFieldDefinitions, SP_LISTS.docReviewerGroup);
  const normalizedMeetingItems = normalizeItemsToSchema(meetingItems, meetingFieldDefinitions, SP_LISTS.meetingSession);
  const normalizedAuditItems = normalizeItemsToSchema(auditItems, auditFieldDefinitions, SP_LISTS.auditLog);
  const normalizedIssueItems = normalizeItemsToSchema(issueItems, issueFieldDefinitions, SP_LISTS.issue);
  const normalizedNotificationConfigItems = normalizeItemsToSchema(notificationConfigItems, notificationConfigFieldDefinitions, SP_LISTS.notificationConfig);
  const normalizedAppSettingsItems = normalizeItemsToSchema(appSettingsItems, appSettingsFieldDefinitions, SP_LISTS.appSettings);
  const normalizedLocationConfigItems = normalizeItemsToSchema(locationConfigItems, locationConfigFieldDefinitions, SP_LISTS.locationConfig);
  const normalizedTicketItems = normalizeItemsToSchema(ticketItems, ticketFieldDefinitions, SP_LISTS.ticket);

  normalizedProjectItems
    .filter((item) => !isArchivedProjectItem(item))
    .forEach((item) => {
    const { proj, extra, people, boards } = spItemToProject(item);
    registerProjectId(proj.id, item.Id);
    db.projects.push(proj);
    db.projectExtra[proj.id] = extra;
    db.projectPeople[proj.id] = people;
    hydrateProjectBoards(db, proj.id, boards);
  });
  if (typeof pushDebugLog === "function") {
    pushDebugLog("info", "sharepoint-load", "Project records mapped into app state", {
      rawProjectItemCount: normalizedProjectItems.length,
      visibleProjectCount: db.projects.length,
      projectIds: db.projects.map((proj) => proj.id)
    });
  }
  // Fold the archive-filter outcome into the same report entry the diagnostics
  // panel reads, so "list loaded fine but every item got archived-filtered"
  // is visible in one place instead of requiring Admin > Logs.
  if (window.AEWTTR && window.AEWTTR.lastLoadReport) {
    const entry = window.AEWTTR.lastLoadReport.find((r) => r.label === SP_LISTS.project);
    if (entry) {
      entry.archivedFilteredCount = normalizedProjectItems.length - db.projects.length;
      entry.visibleCount = db.projects.length;
    }
  }

  normalizedRiskItems.forEach((item) => {
    const risk = spItemToRisk(item);
    if (!risk.projectId) return;
    if (!db.projectExtra[risk.projectId]) {
      db.projectExtra[risk.projectId] = { status: "Not Started", history: [], risks: [], handoff: "", meetingNotes: "" };
    }
    db.projectExtra[risk.projectId].risks = db.projectExtra[risk.projectId].risks || [];
    db.projectExtra[risk.projectId].risks.push(risk);
  });

  normalizedActionItems.slice().sort(bySortOrderThenId).forEach((item) => {
    const task = spItemToActionItem(item);
    task._sortOrder = typeof item.SortOrder === "number" ? item.SortOrder : null;
    if (!task._projectCode) return; // orphaned record (its project was deleted)
    if (task._source === "Checklist") {
      if (!db.engChecklists[task._projectCode]) db.engChecklists[task._projectCode] = { "To Do": [], "In Progress": [], "Complete": [] };
      const column = task._column || "To Do";
      if (!db.engChecklists[task._projectCode][column]) db.engChecklists[task._projectCode][column] = [];
      db.engChecklists[task._projectCode][column].push(task);
    } else {
      if (!db.ganttTasks[task._projectCode]) db.ganttTasks[task._projectCode] = [];
      db.ganttTasks[task._projectCode].push(task);
    }
  });

  // Resolve ParentDividerId (SharePoint number) → local divider id after all
  // action items for the project are present. Also honor parentDividerLocalId
  // which survives when ParentDividerId was 0 at create time.
  Object.keys(db.ganttTasks || {}).forEach((pid) => {
    const list = db.ganttTasks[pid] || [];
    const bySpId = new Map(list.filter((t) => t._spId).map((t) => [Number(t._spId), t]));
    const byLocalId = new Map();
    list.forEach((t) => {
      if (!t || t.itemType !== "divider") return;
      byLocalId.set(t.id, t);
      const local = t.metadata && t.metadata.clientLocalId;
      if (local) byLocalId.set(local, t);
    });
    list.forEach((task) => {
      if (!task || task.itemType === "divider") return;
      const dividerSp = Number(task._parentDividerSpId || 0);
      if (dividerSp) {
        const divider = bySpId.get(dividerSp);
        if (divider && divider.itemType === "divider") {
          task.parentDividerId = divider.id;
          return;
        }
      }
      const localId = task.metadata && task.metadata.parentDividerLocalId;
      if (localId && byLocalId.has(localId)) task.parentDividerId = byLocalId.get(localId).id;
    });
  });

  db.travelRequests = normalizedTravelItems.map(spItemToTravelRequest);
  db.debriefs = normalizedDebriefItems.map(spItemToTravelDebrief);
  db.teamEvents = normalizedTeamEventItems.map(spItemToTeamEvent);

  // Ensure every current + legacy ReviewColumn bag exists before push.
  // Missing bags used to throw here and abort the ENTIRE SharePoint load
  // (travelRequests already mapped above never reach the caller) — which is
  // why Travel and Document Review both looked empty after the 8-status
  // workflow landed while empty-db still only had the old 5 columns.
  DOC_COLUMNS.forEach((column) => {
    if (!db.docs[column]) db.docs[column] = [];
  });
  normalizedDocItems.forEach((item) => {
    const doc = spItemToDocReviewItem(item);
    const column = doc._column || "Not Started";
    if (!db.docs[column]) db.docs[column] = [];
    db.docs[column].push(doc);
  });
  db.groups = normalizedDocReviewGroupItems.map(spItemToDocReviewerGroup);
  db.docReviewerGroups = db.groups;

  /* Meetings: every session is one list item. The newest item per scope
     (highest Id) is the live "current session"; the rest are the archive. */
  const sessionsByScope = {};
  normalizedMeetingItems.forEach((item) => {
    const session = spItemToMeetingSession(item);
    const scope = session._projectCode || "";
    (sessionsByScope[scope] = sessionsByScope[scope] || []).push({ session, itemId: item.Id });
  });
  Object.keys(sessionsByScope).forEach((scope) => {
    const allEntries = sessionsByScope[scope].sort((x, y) => y.itemId - x.itemId);
    // Planned sessions are scheduled meetings — keep them separate from the live/archive flow.
    const plannedEntries = allEntries.filter((e) => e.session.sessionStatus === "planned");
    const entries = allEntries.filter((e) => e.session.sessionStatus !== "planned");
    const newest = entries[0];
    const newestEnded = newest && newest.session.sessionStatus === "ended";
    const current = newestEnded ? null : newest && newest.session;
    const archive = (newestEnded ? entries : entries.slice(1)).map((entry) => entry.session)
      .sort((x, y) => String(y.date).localeCompare(String(x.date)));
    const meetingStatus = current ? "active" : (entries.length ? "ended" : "idle");
    const scheduledMeetings = plannedEntries.map((e) => {
      const s = e.session;
      return {
        id: s.id, _spId: s._spId,
        title: s.scheduledTitle || s.id,
        date: s.date || "",
        time: s.scheduledTime || "",
        agendaText: s.agendaText || "",
        docHtml: s.docHtml || "",
        createdBy: s.createdBy || "",
        createdAt: s.createdAt || "",
        sessionStatus: "scheduled"
      };
    });
    if (!scope) {
      db.weeklyMeeting.currentSession = current;
      db.weeklyMeeting.sessions = archive;
      db.weeklyMeeting.meetingStatus = meetingStatus;
      db.weeklyMeeting.scheduledMeetings = scheduledMeetings;
    } else {
      db.weeklyMeeting.projectMeetings[scope] = Object.assign(
        { sessions: [], currentSession: null, meetingStatus: "idle", scheduledMeetings: [] },
        db.weeklyMeeting.projectMeetings[scope] || {},
        { currentSession: current, sessions: archive, meetingStatus, scheduledMeetings }
      );
    }
  });

  db.auditLog = normalizedAuditItems.map(spItemToAuditLog).sort((a, b) => String(b.ts).localeCompare(String(a.ts)));
  db.issues = normalizedIssueItems.map(spItemToIssue).sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
  db.tickets = normalizedTicketItems.map(spItemToTicket).sort((a, b) => String(b.opened).localeCompare(String(a.opened)));
  if (normalizedNotificationConfigItems.length) {
    db.notificationConfig = spItemToNotificationConfig(normalizedNotificationConfigItems[0]);
  } else {
    db.notificationConfig = notificationConfigDefaults();
  }
  db.appSettings = normalizedAppSettingsItems.length
    ? spItemToAppSettings(normalizedAppSettingsItems[0])
    : appSettingsDefaults();
  db.locationConfig = normalizedLocationConfigItems.length
    ? spItemToLocationConfig(normalizedLocationConfigItems[0])
    : { locations: [], portfolios: [], contractors: [], hideUnaffiliatedPeople: true };

  db.weeklyMeeting.roster = _roleRecords
    .filter(sharePointAdapter.isRoleRecordActive)
    .filter((record) => !sharePointAdapter.isHiddenFromMeetings(record))
    .map((record) => "u" + record.Id);
  db.members = rebuildMembersView();

  if (typeof bootLog === "function" && window.AEWTTR && window.AEWTTR.lastLoadReport) {
    const realLists = window.AEWTTR.lastLoadReport.filter((r) => !r.label.endsWith(" fields"));
    const failed = realLists.filter((r) => !r.ok);
    bootLog(
      failed.length
        ? `SharePoint data: ${realLists.length - failed.length}/${realLists.length} lists loaded, ${failed.length} failed (see Logs).`
        : `SharePoint data: all ${realLists.length} lists loaded (${db.projects.length} projects).`,
      failed.length ? "error" : "success"
    );
  }
  // NOTE: do NOT call normalizeStoreShape here — db.user is not set yet at
  // this point (both callers in app.js assign it right after this returns).
  // normalizeStoreShape reads db.user.id unconditionally, so calling it here
  // threw on every single load, was silently swallowed by the caller's
  // try/catch, and fell back to an empty database — the exact "my project
  // vanishes on refresh" bug. Callers normalize themselves once .user is set.
  return db;
}

/* ---------- generic save / remove ----------

   Per-record persistence: each Repo.save targets ONE list item. Saves are
   debounced per object (500 ms) so rapid-fire callers — a gantt bar drag
   fires a save per pointer-move — collapse into a single REST call carrying
   the final state. Writes for the same object are chained so a create always
   completes (and stamps _spId) before a follow-up update runs. */

// Tracker updates should feel immediate. The write chain still serializes
// rapid edits safely, but removes the noticeable save delay.
const SAVE_DEBOUNCE_MS = 25;
const _saveTimers = new Map();     // obj -> { timer, kind }        (debounce not yet elapsed)
const _writesInFlight = new Set(); // obj currently mid-REST-call   (debounce elapsed, network pending)
const _saveChains = new WeakMap(); // obj -> promise chain (serializes writes per object)
const _pendingSettlers = new Map(); // obj -> {promise, resolve} shared by every repoSave call in one debounce burst

function mapKindToPayload(kind, obj) {
  switch (kind) {
    case "project": return projectToSpItem(obj);
    case "risk": return riskToSpItem(obj);
    case "actionItem": return actionItemToSpItem(obj);
    case "travelRequest": return travelRequestToSpItem(obj);
    case "travelDebrief": return travelDebriefToSpItem(obj);
    case "teamEvent": return teamEventToSpItem(obj);
    case "docReviewItem": return docReviewItemToSpItem(obj, obj._column || "Not Started");
    case "docReviewerGroup": return docReviewerGroupToSpItem(obj);
    case "meetingSession": return meetingSessionToSpItem(obj, obj._projectCode || "");
    case "issue": return issueToSpItem(obj);
    case "ticket": return ticketToSpItem(obj);
    case "notificationConfig": return notificationConfigToSpItem(obj);
    case "appSettings": return appSettingsToSpItem(obj);
    case "locationConfig": return locationConfigToSpItem(obj);
    default: return null;
  }
}

/* Stamp routing context (projectCode/source/column) onto the object itself so
   later saves without an `extra` argument still know where the record lives. */
function stampExtra(kind, obj, extra) {
  if (!extra) return;
  if (extra.projectCode !== undefined && extra.projectCode !== "") obj._projectCode = extra.projectCode;
  if (extra.source) obj._source = extra.source;
  if (extra.column) obj._column = extra.column;
  if (kind === "actionItem") {
    if (!obj.metadata || typeof obj.metadata !== "object") obj.metadata = {};
    if (!obj.metadata.clientLocalId && obj.id) obj.metadata.clientLocalId = obj.id;
    if (obj._projectCode) obj.metadata.projectCode = obj._projectCode;
    if (obj.parentDividerId) obj.metadata.parentDividerLocalId = obj.parentDividerId;
  }
}

/* The array an action item lives in, used for SortOrder reindexing. */
function actionItemSiblings(obj) {
  const db = window.AEWTTR.db;
  if (obj._source === "Checklist") {
    const board = db.engChecklists[obj._projectCode];
    return (board && board[obj._column]) || null;
  }
  return db.ganttTasks[obj._projectCode] || null;
}

/* Boards live in BoardsJson on PULSE Projects. Older sites may not have the
   column yet — without it, filterItemPayloadForExistingFields silently drops
   the field and board creates/edits appear to save but never persist. */
let _boardsJsonFieldReady = false;
async function ensureBoardsJsonField(siteUrl) {
  if (_boardsJsonFieldReady) return;
  if (!siteUrl || typeof sharePointAdapter === "undefined" || !sharePointAdapter.ensureField) return;
  await sharePointAdapter.ensureField(siteUrl, SP_LISTS.project, {
    name: "BoardsJson",
    type: "Note",
    numLines: 30
  });
  _boardsJsonFieldReady = true;
}

let _deliverablesJsonFieldReady = false;
async function ensureDeliverablesJsonField(siteUrl) {
  if (_deliverablesJsonFieldReady || !siteUrl || typeof sharePointAdapter === "undefined" || !sharePointAdapter.ensureField) return;
  await sharePointAdapter.ensureField(siteUrl, SP_LISTS.project, {
    name: "DeliverablesJson",
    type: "Note",
    numLines: 20
  });
  _deliverablesJsonFieldReady = true;
}

// Agenda uses PULSE Meetings.AgendaJson. Older manually-created sites may
// have the list but not this field, so create it before the first agenda save.
let _meetingAgendaFieldReady = false;
async function ensureMeetingAgendaField(siteUrl) {
  if (_meetingAgendaFieldReady || !siteUrl || typeof sharePointAdapter === "undefined" || !sharePointAdapter.ensureField) return;
  await sharePointAdapter.ensureField(siteUrl, SP_LISTS.meetingSession, {
    name: "AgendaJson", type: "Note", numLines: 20
  });
  _meetingAgendaFieldReady = true;
}

/* PortfoliosJson on projects + Locations/PortfolioCatalog on Location Config
   have the same silent-drop failure mode until the columns exist. */
let _portfoliosJsonFieldReady = false;
async function ensurePortfoliosJsonField(siteUrl) {
  if (_portfoliosJsonFieldReady) return;
  if (!siteUrl || typeof sharePointAdapter === "undefined" || !sharePointAdapter.ensureField) return;
  await sharePointAdapter.ensureField(siteUrl, SP_LISTS.project, {
    name: "PortfoliosJson",
    type: "Note"
  });
  _portfoliosJsonFieldReady = true;
}

/* Reporting settings are stored alongside the project because they are a
   project-level view of existing tracker data.  Ensure older sites receive
   the column before their first Reporting-tab edit is written. */
let _reportConfigJsonFieldReady = false;
async function ensureReportConfigJsonField(siteUrl) {
  if (_reportConfigJsonFieldReady || !siteUrl || typeof sharePointAdapter === "undefined" || !sharePointAdapter.ensureField) return;
  await sharePointAdapter.ensureField(siteUrl, SP_LISTS.project, {
    name: "ReportConfigJson", type: "Note", numLines: 30
  });
  _reportConfigJsonFieldReady = true;
}

let _locationConfigFieldsReady = false;
async function ensureLocationConfigFields(siteUrl) {
  if (_locationConfigFieldsReady) return;
  if (!siteUrl || typeof sharePointAdapter === "undefined" || !sharePointAdapter.ensureField) return;
  await sharePointAdapter.ensureField(siteUrl, SP_LISTS.locationConfig, {
    name: "LocationsJson",
    type: "Note",
    numLines: 10
  });
  await sharePointAdapter.ensureField(siteUrl, SP_LISTS.locationConfig, {
    name: "PortfolioCatalogJson",
    type: "Note",
    numLines: 10
  });
  await sharePointAdapter.ensureField(siteUrl, SP_LISTS.locationConfig, {
    name: "HideUnaffiliatedPeople",
    type: "Boolean"
  });
  await sharePointAdapter.ensureField(siteUrl, SP_LISTS.locationConfig, {
    name: "ContractorCatalogJson",
    type: "Note",
    numLines: 10
  });
  await sharePointAdapter.ensureField(siteUrl, SP_LISTS.locationConfig, {
    name: "ContractCatalogJson",
    type: "Note",
    numLines: 10
  });
  await sharePointAdapter.ensureField(siteUrl, SP_LISTS.locationConfig, {
    name: "ConfigEndItemCatalogJson",
    type: "Note",
    numLines: 10
  });
  _locationConfigFieldsReady = true;
}

/* Issue reporting is intentionally its own list so a user can capture a
   reproducible report without modifying operational records. Ensure fields
   before the first write too: this repairs older sites where the list exists
   but a setup run pre-dates the issue-report feature. */
let _issueFieldsReady = false;
async function ensureIssueFields(siteUrl) {
  if (_issueFieldsReady || !siteUrl || typeof sharePointAdapter === "undefined" || !sharePointAdapter.ensureField) return;
  const schema = typeof SHAREPOINT_SCHEMA !== "undefined" && SHAREPOINT_SCHEMA[SP_LISTS.issue];
  // A report can be the first interaction with a newly deployed package.
  // Ensure the list itself exists before probing its fields; previously the
  // field probe hit a 404 on a fresh site and the Submit button could only
  // surface a generic save failure. Existing lists are simply verified here.
  if (schema && typeof sharePointAdapter.ensureListSchema === "function") {
    const results = await sharePointAdapter.ensureListSchema(siteUrl, { [SP_LISTS.issue]: schema });
    const result = results && results[0];
    const failed = result && (result.errors || []).find((entry) => entry && entry.error);
    if (failed) throw failed.error;
    _issueFieldsReady = true;
    return;
  }
  for (const field of (schema && schema.fields) || []) {
    await sharePointAdapter.ensureField(siteUrl, SP_LISTS.issue, field);
  }
  _issueFieldsReady = true;
}

async function ensureTicketFields(siteUrl) {
  return ensureIssueFields(siteUrl);
}

let _appSettingsFieldsReady = false;
async function ensureAppSettingsFields(siteUrl) {
  if (_appSettingsFieldsReady || !siteUrl || typeof sharePointAdapter === "undefined" || !sharePointAdapter.ensureField) return;
  const schema = typeof SHAREPOINT_SCHEMA !== "undefined" && SHAREPOINT_SCHEMA[SP_LISTS.appSettings];
  for (const field of (schema && schema.fields) || []) {
    await sharePointAdapter.ensureField(siteUrl, SP_LISTS.appSettings, field);
  }
  _appSettingsFieldsReady = true;
}

/* Action-item columns that older sites may be missing. Without them,
   filterItemPayloadForExistingFields silently drops ProjectId / MetadataJson /
   SubtasksJson and creates look successful until refresh (items vanish). */
let _actionItemFieldsReady = false;
async function ensureActionItemFields(siteUrl) {
  if (_actionItemFieldsReady) return;
  if (!siteUrl || typeof sharePointAdapter === "undefined" || !sharePointAdapter.ensureField) return;
  const fields = [
    { name: "ProjectId", type: "Number", required: true },
    { name: "ParentActionItemId", type: "Number" },
    { name: "WorkItemLevel", type: "Choice", choices: ["Task", "Subtask"] },
    { name: "Source", type: "Choice", choices: ["Tracker", "Checklist"] },
    { name: "AssignedToName", type: "Text" },
    { name: "StartDate", type: "DateTime" },
    { name: "DueDate", type: "DateTime" },
    { name: "ActionStatus", type: "Choice", choices: ["Not Started", "In Progress", "Blocked", "On Hold", "Complete", "Done", "Cancelled"] },
    { name: "Health", type: "Choice", choices: ["On Track", "At Risk", "Off Track", "Blocked", "On Hold", "Complete"] },
    { name: "BlockedReason", type: "Note", numLines: 6 },
    { name: "OnHoldReason", type: "Note", numLines: 6 },
    { name: "Notes", type: "Note" },
    { name: "SubtasksJson", type: "Note", numLines: 10 },
    { name: "MetadataJson", type: "Note", numLines: 20 },
    { name: "IsMilestone", type: "Choice", choices: ["Yes", "No"] },
    { name: "Position", type: "Number" },
    { name: "SortOrder", type: "Number" }
  ];
  for (const field of fields) {
    await sharePointAdapter.ensureField(siteUrl, SP_LISTS.actionItem, field);
  }
  _actionItemFieldsReady = true;
}

async function writeRecord(kind, obj) {
  const siteUrl = currentSiteUrl();
  const listTitle = SP_LISTS[kind];
  if (!listTitle) return;

  if (kind === "project") {
    await ensureBoardsJsonField(siteUrl);
    await ensureDeliverablesJsonField(siteUrl);
    await ensurePortfoliosJsonField(siteUrl);
    await ensureReportConfigJsonField(siteUrl);
  }
  if (kind === "locationConfig") {
    await ensureLocationConfigFields(siteUrl);
  }
  if (kind === "issue") {
    await ensureIssueFields(siteUrl);
  }
  if (kind === "ticket") {
    await ensureTicketFields(siteUrl);
  }
  if (kind === "appSettings") {
    await ensureAppSettingsFields(siteUrl);
  }
  if (kind === "meetingSession") {
    await ensureMeetingAgendaField(siteUrl);
  }
  if (kind === "actionItem") {
    await ensureActionItemFields(siteUrl);
    // Recover project map from live db if the boot map was incomplete.
    if (obj._projectCode && !spIdForProjectCode(obj._projectCode)) {
      resolveActionItemProjectSpId(obj._projectCode);
    }
  }

  const payload = mapKindToPayload(kind, obj);
  if (!payload) return;

  if (kind === "actionItem") {
    const siblings = actionItemSiblings(obj);
    if (siblings) {
      payload.Position = siblings.indexOf(obj);
      payload.SortOrder = payload.Position;
    }
    if (!payload.ProjectId) {
      const err = new Error(`Cannot save "${obj.title || "item"}": project "${obj._projectCode || ""}" is not linked to SharePoint. Run SharePoint Setup, then try again.`);
      err.friendly = err.message;
      throw err;
    }
  }
  if ((kind === "notificationConfig" || kind === "appSettings" || kind === "locationConfig") && !obj._spId) {
    // single-row list: adopt the existing row if setup already seeded one
    const rows = await sharePointAdapter.getItems(siteUrl, listTitle, { top: 1 });
    if (rows.length) obj._spId = rows[0].Id;
  }

  if (obj._spId) {
    await sharePointAdapter.updateItem(siteUrl, listTitle, obj._spId, payload);
  } else {
    const created = await sharePointAdapter.createItem(siteUrl, listTitle, payload);
    obj._spId = created && created.Id;
    if (kind === "project") registerProjectId(obj.id, obj._spId);
  }

  // Reordering one task displaces its siblings — persist the ones whose
  // stored SortOrder no longer matches their array position.
  if (kind === "actionItem") {
    const siblings = actionItemSiblings(obj);
    if (siblings) {
      for (let i = 0; i < siblings.length; i++) {
        const sibling = siblings[i];
        if (sibling === obj) {
          sibling._sortOrder = i;
          if (sibling._spId) {
            // Target obj also needs its SortOrder persisted immediately
            // because its initial payload during serialization missed it.
            await sharePointAdapter.updateItem(siteUrl, listTitle, sibling._spId, { Position: i, SortOrder: i });
          }
          continue;
        }
        if (sibling._spId && sibling._sortOrder !== i) {
          await sharePointAdapter.updateItem(siteUrl, listTitle, sibling._spId, { Position: i, SortOrder: i });
          sibling._sortOrder = i;
        }
      }
    }
  }

  if (window.AEWTTR) {
    window.AEWTTR.lastLocalSpWriteAt = Date.now();
    // A hard page refresh right after this write would otherwise rehydrate
    // from the sessionStorage snapshot taken before this write happened —
    // "I created a project, hit refresh, and it's not there" — so keep that
    // snapshot current on every successful write, not just on a full poll.
    if (typeof persistSpDbCache === "function") persistSpDbCache(currentSiteUrl(), window.AEWTTR.db);
  }
}

/* Every repoSave call within one debounce burst (e.g. a gantt-bar drag firing
   a save per pointer-move) shares this single settled-when-actually-written
   promise, created on the burst's first call and resolved once the real
   network write finishes — not when the debounce timer merely fires. */
function getOrCreateSettler(obj) {
  let entry = _pendingSettlers.get(obj);
  if (!entry) {
    let resolveFn;
    let rejectFn;
    const promise = new Promise((resolve, reject) => {
      resolveFn = resolve;
      rejectFn = reject;
    });
    // Swallow the rejection at the source so fire-and-forget callers never
    // trip the global unhandledrejection toast; awaiters still see the error.
    promise.catch(() => {});
    entry = { promise, resolve: resolveFn, reject: rejectFn };
    _pendingSettlers.set(obj, entry);
  }
  return entry;
}

/* Background (debounced) saves retry silently on transient failures before
   bothering the user. Error toasts are deduped so a flaky connection doesn't
   stack a wall of identical messages. */
const BG_SAVE_RETRY_DELAYS_MS = [3000, 10000];
let _lastSaveErrorToastAt = 0;
function toastSaveErrorDeduped(err) {
  if (Date.now() - _lastSaveErrorToastAt < 20000) return;
  _lastSaveErrorToastAt = Date.now();
  toast((err && err.friendly) || "Couldn’t save to SharePoint. Retrying…", "error");
}

function attemptBackgroundWrite(kind, obj, settler, attempt) {
  writeRecordSafe(kind, obj, false).then((ok) => {
    if (ok || (obj && obj._deleted)) {
      _writesInFlight.delete(obj);
      _pendingSettlers.delete(obj);
      settler.resolve();
      return;
    }
    if (attempt < BG_SAVE_RETRY_DELAYS_MS.length) {
      setTimeout(() => attemptBackgroundWrite(kind, obj, settler, attempt + 1), BG_SAVE_RETRY_DELAYS_MS[attempt]);
      return;
    }
    _writesInFlight.delete(obj);
    _pendingSettlers.delete(obj);
    const err = obj._lastSaveError || new Error("Save failed");
    toastSaveErrorDeduped(err);
    settler.reject(err);
  });
}

function scheduleRecordSave(kind, obj) {
  if (obj && obj._deleted) {
    const settler = getOrCreateSettler(obj);
    settler.resolve();
    _pendingSettlers.delete(obj);
    return settler.promise;
  }
  const pending = _saveTimers.get(obj);
  if (pending) clearTimeout(pending.timer);
  const settler = getOrCreateSettler(obj);
  const timer = setTimeout(() => {
    _saveTimers.delete(obj);
    if (obj && obj._deleted) {
      _pendingSettlers.delete(obj);
      settler.resolve();
      return;
    }
    _writesInFlight.add(obj);
    attemptBackgroundWrite(kind, obj, settler, 0);
  }, SAVE_DEBOUNCE_MS);
  _saveTimers.set(obj, { timer, kind });
  return settler.promise;
}

/* Resolves once the write this call triggered has actually reached
   SharePoint — callers that need obj._spId (e.g. creating a project's seed
   tasks right after creating the project) can safely `await` this.
   Rejects when the write fails so awaited autosave / optimistic callers can
   show an error or roll back. Fire-and-forget callers should attach .catch. */
async function repoSave(kind, obj, extra) {
  if (!isSharePointMode()) { aewttrSaveStore(); return; }
  if (!obj) return;
  if (obj._deleted) return;
  stampExtra(kind, obj, extra);
  delete obj._lastSaveError;
  // Creates / explicit immediate writes skip the drag-debounce so ParentDividerId
  // and ProjectId land before the user navigates away.
  if (extra && extra.immediate) {
    const pending = _saveTimers.get(obj);
    if (pending) { clearTimeout(pending.timer); _saveTimers.delete(obj); }
    const settler = getOrCreateSettler(obj);
    _writesInFlight.add(obj);
    let ok = await writeRecordSafe(kind, obj, false);
    if (!ok && !obj._deleted) {
      // One quick retry — a transient blip shouldn't roll back an optimistic
      // edit the user made a second ago.
      await new Promise((r) => setTimeout(r, 1500));
      ok = await writeRecordSafe(kind, obj, true);
    }
    _writesInFlight.delete(obj);
    _pendingSettlers.delete(obj);
    if (ok) settler.resolve();
    else {
      const err = obj._lastSaveError || new Error("Save failed");
      settler.reject(err);
      throw err;
    }
    return settler.promise;
  }
  const pending = scheduleRecordSave(kind, obj);
  // Prevent unhandledrejection when callers ignore the promise; awaiters
  // still see the rejection on the same promise instance.
  pending.catch(() => {});
  return pending;
}

function isMissingSharePointItemError(err) {
  const status = err && (err.status || err.statusCode);
  if (status === 404) return true;
  const text = String((err && err.friendly) || (err && err.message) || err || "").toLowerCase();
  return text.includes("doesn't exist") || text.includes("does not exist") || text.includes("404");
}

async function repoRemove(kind, obj) {
  if (!isSharePointMode()) { aewttrSaveStore(); return; }
  if (!obj) return;
  // Stop any queued writes and mark deleted so in-flight / late saves skip.
  obj._deleted = true;
  const pendingEntry = _saveTimers.get(obj);
  if (pendingEntry) { clearTimeout(pendingEntry.timer); _saveTimers.delete(obj); }
  const pendingSettler = _pendingSettlers.get(obj);
  if (pendingSettler) {
    _pendingSettlers.delete(obj);
    pendingSettler.resolve();
  }
  const siteUrl = currentSiteUrl();
  const listTitle = SP_LISTS[kind];
  const prev = _saveChains.get(obj) || Promise.resolve();
  _writesInFlight.add(obj);
  delete obj._lastRemoveError;
  const next = prev.then(async () => {
    if (!listTitle || !obj._spId) return;
    try {
      await sharePointAdapter.deleteItem(siteUrl, listTitle, obj._spId);
    } catch (e) {
      // Already gone (race with a prior delete / never persisted) — treat as success.
      if (isMissingSharePointItemError(e)) return;
      throw e;
    }
    // Deleting a project also deletes its tracker/checklist records so they
    // don't come back as orphans on the next load.
    if (kind === "project") {
      const db = window.AEWTTR.db;
      const children = (db.ganttTasks[obj.id] || []).slice();
      const board = db.engChecklists[obj.id] || {};
      Object.keys(board).forEach((column) => children.push(...(board[column] || [])));
      for (const child of children) {
        child._deleted = true;
        if (child._spId) await sharePointAdapter.deleteItem(siteUrl, SP_LISTS.actionItem, child._spId).catch(() => {});
      }
    }
    if (window.AEWTTR) {
      window.AEWTTR.lastLocalSpWriteAt = Date.now();
      if (typeof persistSpDbCache === "function") persistSpDbCache(currentSiteUrl(), window.AEWTTR.db);
    }
  }).catch((e) => {
    const formatted = recordSpError(e);
    obj._lastRemoveError = formatted;
    // Don't surface 404s — the local delete already succeeded.
    if (!isMissingSharePointItemError(formatted) && !isMissingSharePointItemError(e)) {
      toastSaveErrorDeduped(formatted);
    }
    // Soft-resolve missing items so callers don't trip unhandledrejection.
    if (isMissingSharePointItemError(formatted) || isMissingSharePointItemError(e)) return;
    throw formatted;
  }).finally(() => {
    _writesInFlight.delete(obj);
  });
  _saveChains.set(obj, next);
  next.catch(() => {});
  return next;
}

/* Fire all pending debounced saves immediately (tab hidden / closing / navigation). */
function repoFlushNow() {
  const pending = Array.from(_saveTimers.entries());
  _saveTimers.clear();
  return Promise.all(pending.map(([obj, entry]) => {
    clearTimeout(entry.timer);
    _writesInFlight.add(obj);
    const settler = _pendingSettlers.get(obj);
    return writeRecordSafe(entry.kind, obj, false).then((ok) => {
      _writesInFlight.delete(obj);
      _pendingSettlers.delete(obj);
      if (!ok && !(obj && obj._deleted)) toastSaveErrorDeduped(obj._lastSaveError);
      if (!settler) return;
      if (ok) settler.resolve();
      else settler.reject(obj._lastSaveError || new Error("Save failed"));
    });
  }));
}

function writeRecordSafe(kind, obj, showToastOnError) {
  const prev = _saveChains.get(obj) || Promise.resolve();
  const next = prev.then(async () => {
    if (obj && obj._deleted) return true;
    delete obj._lastSaveError;
    try {
      await writeRecord(kind, obj);
      return true;
    } catch (e) {
      // A late update after delete (or purge) shouldn't spam the UI.
      if (obj && obj._deleted) return true;
      if (isMissingSharePointItemError(e)) return true;
      throw e;
    }
  }).catch((e) => {
    const formatted = recordSpError(e);
    obj._lastSaveError = formatted;
    if (showToastOnError && !isMissingSharePointItemError(formatted) && !isMissingSharePointItemError(e)) {
      toastSaveErrorDeduped(formatted);
    }
    return false;
  });
  _saveChains.set(obj, next);
  return next;
}

function repoHasPendingChanges() {
  return _saveTimers.size > 0 || _writesInFlight.size > 0;
}

document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "hidden" && isSharePointMode()) repoFlushNow();
});
window.addEventListener("pagehide", () => { if (isSharePointMode()) repoFlushNow(); });

const Repo = { save: repoSave, remove: repoRemove, flush: repoFlushNow, hasPendingChanges: repoHasPendingChanges };
