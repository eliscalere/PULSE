/* AEWTTR-PULSE — site-wide user action audit log (admin-only view). */

const AUDIT_LOG_STORAGE_KEY = "aewttr_audit_log";
const AUDIT_LOG_MAX_LOCAL = 1000;
const AUDIT_LOG_MAX_MEMORY = 2000;
let _lastNavPath = "";

/* Audit log list rotation — create "PULSE Audit Log 2", "3", etc. when the
   active list reaches AUDIT_LIST_CAP items, so each list stays fast to load. */
const AUDIT_LIST_BASE    = "PULSE Audit Log";
const AUDIT_LIST_CAP     = 1000;
const AUDIT_ACTIVE_KEY   = "aewttr_audit_list";
let _auditListName       = null;  // lazily resolved from localStorage
let _auditListCount      = null;  // null = not yet queried from SP

function _auditListSuffix(name) {
  if (name === AUDIT_LIST_BASE) return 1;
  const m = /^PULSE Audit Log (\d+)$/.exec(name);
  return m ? Number(m[1]) : 1;
}
function _auditListForSuffix(n) {
  return n <= 1 ? AUDIT_LIST_BASE : `${AUDIT_LIST_BASE} ${n}`;
}
async function _resolveAuditList(siteUrl) {
  if (!_auditListName) {
    try { _auditListName = localStorage.getItem(AUDIT_ACTIVE_KEY) || AUDIT_LIST_BASE; }
    catch (e) { _auditListName = AUDIT_LIST_BASE; }
  }
  // If we already know the count is under cap, short-circuit.
  if (_auditListCount !== null && _auditListCount < AUDIT_LIST_CAP) {
    _auditListCount++;
    return _auditListName;
  }
  // Query item count from SP.
  try {
    const r = await fetch(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(_auditListName)}')/ItemCount`,
      { credentials: "same-origin", headers: { Accept: "application/json;odata=nometadata" } }
    );
    if (r.ok) {
      _auditListCount = (await r.json()).value || 0;
    } else if (r.status === 404 && _auditListName !== AUDIT_LIST_BASE) {
      // The cached rotated list name was never created — reset to the base list so
      // subsequent writes don't keep hammering a nonexistent list.
      _auditListName = AUDIT_LIST_BASE;
      try { localStorage.removeItem(AUDIT_ACTIVE_KEY); } catch (e2) {}
      _auditListCount = null;
      return _auditListName;
    } else {
      _auditListCount = 0;
    }
  } catch (e) { _auditListCount = 0; }

  if (_auditListCount >= AUDIT_LIST_CAP) {
    const nextName = _auditListForSuffix(_auditListSuffix(_auditListName) + 1);
    if (typeof SHAREPOINT_SCHEMA !== "undefined" && SHAREPOINT_SCHEMA[AUDIT_LIST_BASE] &&
        typeof sharePointAdapter !== "undefined" && sharePointAdapter.ensureListSchema) {
      try {
        await sharePointAdapter.ensureListSchema(siteUrl, {
          [nextName]: Object.assign({}, SHAREPOINT_SCHEMA[AUDIT_LIST_BASE], { title: nextName })
        });
      } catch (e) { console.warn("PULSE audit: failed to create log list", nextName, e); }
    }
    _auditListName  = nextName;
    _auditListCount = 0;
    try { localStorage.setItem(AUDIT_ACTIVE_KEY, _auditListName); } catch (e) {}
  }
  _auditListCount++;
  return _auditListName;
}

function kindToAuditArea(kind) {
  const map = {
    travelRequest: "Travel",
    travelDebrief: "Travel",
    docReviewItem: "Documents",
    project: "Projects",
    actionItem: "Projects",
    meetingSession: "Weekly",
    appRole: "Users"
  };
  return map[kind] || "System";
}

function summarizeRecord(kind, obj, verb) {
  if (!obj) return `${verb} ${kind}`;
  switch (kind) {
    case "travelRequest":
      return `${verb} travel ${obj.id || ""}${obj.destination ? ` — ${obj.destination}` : ""}`.trim();
    case "travelDebrief":
      return `${verb} debrief for ${obj.requestId || obj.id || "trip"}`;
    case "docReviewItem":
      return `${verb} document ${obj.title || obj.id || ""}`.trim();
    case "project":
      return `${verb} project ${obj.id || ""}${obj.name ? ` — ${obj.name}` : ""}`.trim();
    case "actionItem":
      return `${verb} task ${obj.title || obj.id || ""}`.trim();
    case "meetingSession":
      return `${verb} meeting session`;
    case "appRole":
      return `${verb} user role ${obj.email || obj.UserEmail || ""}`.trim();
    default:
      return `${verb} ${kind}${obj.id ? ` ${obj.id}` : ""}`.trim();
  }
}

function currentAuditActor() {
  const user = window.AEWTTR && window.AEWTTR.db && window.AEWTTR.db.user;
  return {
    actorEmail: (user && user.email) || "",
    actorName: (user && user.name) || "Unknown",
    actorRole: typeof currentAppRole === "function" ? currentAppRole() : ((user && user.role) || "Member")
  };
}

function loadLocalAuditLog() {
  try { return JSON.parse(lsGet(AUDIT_LOG_STORAGE_KEY) || "[]"); } catch (e) { return []; }
}

function initAuditLogStore() {
  const db = window.AEWTTR && window.AEWTTR.db;
  if (!db) return;
  if (!Array.isArray(db.auditLog)) {
    db.auditLog = typeof isSharePointMode === "function" && isSharePointMode() ? [] : loadLocalAuditLog();
  }
}

async function persistAuditEntry(entry) {
  try {
    if (typeof isSharePointMode === "function" && isSharePointMode() && typeof SP_LISTS !== "undefined") {
      const siteUrl = typeof currentSiteUrl === "function" ? currentSiteUrl() : "";
      if (!siteUrl || typeof auditLogToSpItem !== "function") return;
      const listTitle = await _resolveAuditList(siteUrl);
      const created = await sharePointAdapter.createItem(siteUrl, listTitle, auditLogToSpItem(entry));
      entry._spId = created && created.Id;
      entry._auditList = listTitle;
      return;
    }
    const db = window.AEWTTR && window.AEWTTR.db;
    if (db && Array.isArray(db.auditLog)) {
      lsSet(AUDIT_LOG_STORAGE_KEY, JSON.stringify(db.auditLog.slice(0, AUDIT_LOG_MAX_LOCAL)));
    }
  } catch (e) {
    console.warn("PULSE: audit log persist failed.", e);
  }
}

function logUserAction(options) {
  options = options || {};
  if (!window.AEWTTR || !window.AEWTTR.db) return null;
  initAuditLogStore();
  const actor = currentAuditActor();
  const summary = String(options.summary || "").trim();
  if (!summary) return null;
  const entry = {
    id: uid("LOG"),
    ts: new Date().toISOString(),
    action: options.action || "Other",
    area: options.area || "System",
    summary: summary.slice(0, 500),
    detail: options.detail || null,
    route: options.route || "",
    recordId: options.recordId ? String(options.recordId) : "",
    actorEmail: actor.actorEmail,
    actorName: actor.actorName,
    actorRole: actor.actorRole
  };
  window.AEWTTR.db.auditLog.unshift(entry);
  if (window.AEWTTR.db.auditLog.length > AUDIT_LOG_MAX_MEMORY) {
    window.AEWTTR.db.auditLog.length = AUDIT_LOG_MAX_MEMORY;
  }
  persistAuditEntry(entry);
  return entry;
}

function hookAuditInstrumentation() {
  if (typeof Repo !== "undefined" && !Repo._auditHooked) {
    const origSave = Repo.save.bind(Repo);
    const origRemove = Repo.remove.bind(Repo);
    Repo.save = async function (kind, obj, extra) {
      const isNew = obj && !obj._spId;
      await origSave(kind, obj, extra);
      if (obj) {
        const action = obj._auditAction || (isNew ? "Create" : "Update");
        const summary = obj._auditSummary || summarizeRecord(kind, obj, action === "Create" ? "Created" : action === "Update" ? "Updated" : action);
        delete obj._auditAction;
        delete obj._auditSummary;
        logUserAction({
          action,
          area: kindToAuditArea(kind),
          summary,
          recordId: obj.id || obj._spId || "",
          detail: { kind, extra: extra || null }
        });
      }
    };
    Repo.remove = async function (kind, obj) {
      if (obj) {
        logUserAction({
          action: "Delete",
          area: kindToAuditArea(kind),
          summary: summarizeRecord(kind, obj, "Deleted"),
          recordId: obj.id || obj._spId || "",
          detail: { kind }
        });
      }
      return origRemove(kind, obj);
    };
    Repo._auditHooked = true;
  }

  if (typeof navigate === "function" && !window._auditNavigateHooked) {
    const origNavigate = navigate;
    // Was dropping the second (query) argument entirely — every caller
    // passing navigate(path, { tr: id }) style deep-link params (travel
    // debrief switching, doc review notification links, admin's "go to
    // this request") silently lost them the moment this hook installed,
    // landing on the right page but with no ?tr=... to pick a specific
    // record.
    window.navigate = function (path, query) {
      const normalized = String(path || "").replace(/^#\/?/, "");
      if (normalized && normalized !== _lastNavPath) {
        _lastNavPath = normalized;
        logUserAction({
          action: "Navigate",
          area: "Navigation",
          summary: `Opened ${normalized}`,
          route: `#/${normalized}`
        });
      }
      return origNavigate(path, query);
    };
    window._auditNavigateHooked = true;
  }
}

window.logUserAction = logUserAction;
window.initAuditLogStore = initAuditLogStore;
hookAuditInstrumentation();
