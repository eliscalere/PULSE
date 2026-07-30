/**
 * PULSE Dashboard — v1.0.0
 *
 * Developer:  Eli Scalere (Contractor)
 * Primary:    elijah.t.scalere.ctr@us.navy.mil
 * Secondary:  eli.scalere@scaleredesign.com  |  (516) 265-2636
 */

/* PULSE — core shell, router, shared utilities */

/* Merge into any AEWTTR bag already seeded by earlier scripts (export.js,
   project-pptx-export.js, etc.). A full reassignment here used to wipe
   ProjectPptxExport / ExportService and surface "unavailable in this package". */
window.AEWTTR = Object.assign(window.AEWTTR || {}, {
  db: null,
  state: {}, // transient UI state per page (filters, active sub-tab, etc.)
  debugLog: []
});

/* ---------- minimal boot screen ----------
   Logo + wordmark + progress bar while the app boots. A small "Logs" toggle
   reveals a short, capped, copyable diagnostic list — enough to hand to
   tech support without becoming an ever-scrolling console. bootLog() itself
   never renders anything once boot has completed (see the bootComplete
   guard below) — this is what previously let post-boot calls (e.g. clicking
   "Reload List Data" in Admin) blow away the already-rendered app UI. */
const BOOT_LOG_CAP = 40;
// Let the opening wordmark complete, then leave a short beat before the
// workspace replaces it. Slow SharePoint boots simply hold the completed mark.
const PULSE_BOOT_MIN_DISPLAY_MS = 1000;
const BOOT_PROGRESS_STEPS = [
  "Loading workspace",
  "Syncing projects",
  "Checking logs",
  "Building dashboard",
  "Finalizing"
];

/* Compact PULSE signal mark. The brand guide calls for a restrained black or
   white identity, so the app icon has no gradient, glow, or status color. */
function pulseMarkSvg(options) {
  const opts = options || {};
  const size = Number(opts.size || 48);
  const animate = !!opts.animate;
  const tone = opts.tone || "dark";
  const title = escapeHtml(opts.title || "PULSE");
  const badgeColor = tone === "light" ? "#FFFFFF" : "#070708";
  const lineColor = tone === "light" ? "#070708" : "#FFFFFF";
  return `
    <svg class="pulse-mark-svg${animate ? " is-animated" : ""}" width="${size}" height="${size}" viewBox="0 0 120 120" fill="none" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="${title}">
      <rect class="pulse-mark-badge" x="4" y="4" width="112" height="112" rx="26" fill="${badgeColor}" />
      <path class="pulse-mark-wave" d="M14 62H37L46 34L62 92L72 62H82L90 48L98 62H106"
        stroke="${lineColor}" stroke-width="7.5" stroke-linecap="round" stroke-linejoin="round"
        fill="none" />
    </svg>
  `;
}

function pulseBrandSvg(opts) {
  const options = opts || {};
  const size = Number(options.size || 48);
  const animate = !!options.animate;
  const tone = options.boot ? "dark" : (options.tone || "light");
  return pulseMarkSvg({ size, animate, tone, title: options.title || "PULSE" });
}

function pulseBrandLockup(opts) {
  const options = opts || {};
  const boot = !!options.boot;
  const compact = !!options.compact;
  const subline = escapeHtml(options.subline || "Project Updates, Logs, Status, and Execution");
  const wordmark = `<div class="pulse-wordmark${boot ? " pulse-wordmark--boot" : ""}" aria-label="PULSE"><span>P</span><span>U</span><span>L</span><span>S</span><span>E</span><i class="pulse-wordmark-dot pulse-wordmark-dot--one" aria-hidden="true"></i><i class="pulse-wordmark-dot pulse-wordmark-dot--two" aria-hidden="true"></i></div>`;

  if (boot) {
    return `
      <div class="pulse-brand-lockup boot">
        ${wordmark}
        <p class="pulse-brand-subline">${subline}</p>
      </div>
    `;
  }

  return `
    <div class="pulse-brand-lockup${compact ? " compact" : ""}">
      <img class="pulse-nav-dots-img" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAACb0lEQVR42q2YTWtTQRSG33MTYtUgCCrdSBA3imCV6s9wI1rostD2V9Tf4j+w1C7FhRs3CgbBj40NrkqxdBH8aJPePG7OlPF6k9yvA2HuwNw5z8yd8545MbkBJklmBtCS9EDSbUkLkn5I6pvZNx/bMrNUTVlw7s/rQB9I+deGwAtgOUA05hxIgAvAdsZpCpxmYMbAZmMQQMshdtzBKGf1ABN3PvH+SpMQaz7pCfMtdYgj4JrDWx3nHeCzT5xSzMbePvM52lX9J5LuS7olybxfxEwSkh55P60DcCeasMx7JqkHdD10rSrAgjunwvsd/6nODhz6asqsIMAOJf2qC/BO0p+Sn2HiYz+a2YkrI5UAzGwg6ZUDlDlMJul59FxLCZemqF6ejbx94wqa1FZCbzcycR4EZxLBhfgfANeDjM9YWDGRiiBWgYM5O/Aa6Pn4ZIq0t3Jg2jNhIohFYAt4C+wDh8CeJ6qn0fgk836So7CXgW6en5kQUb/rk3TmOEui5ycOu+fw+76YLWCxCETuduVtbewc6PnnmWUHwGrhLDrvIHkUmB/IQXSAT3MO8TgC2WgklYcw9JCMQ3RWKg9wS3XTeNvbx5kUXTSVv2xKP3aj21IRC5/lN3AjqaqeZpYC5yTdrXiXOC/pYT0ZlS5KulQhH4T0f6UuwMh/ZS2k/+NKAOEGZGY/JX331UxKAiDpU50dCDG8W/Eu8VXSh7pp3PxqfhQJzjwLV/+1Rooab1emFC9ZEQpCtePwrSYhNjNaMO2Cs+1lYFJLCadALHsBO8xZfR9Y/68ab7DCPivZgZuS7km6KulY0hdJ7128zv4GkKS/w3S3ykLv268AAAAASUVORK5CYII=" width="32" height="32" alt="" aria-hidden="true">
    </div>
  `;
}

/* Role-based landing pages — Admins land on Team Overview, Finance Admins
   on the Awaiting Finance travel queue, Document Admins on Document Review,
   since those are the pages each opens first in practice; everyone else
   keeps the existing Dashboard default. Priority, highest first: a port
   config's explicit defaultRoute (e.g. a Travel-only deployment) always
   wins; then the user's own saved "Default page" preference (Settings →
   Notifications), since that's an explicit personal choice; then role
   defaults — Admin BEFORE Finance Admin/Document Admin, because the "Admin"
   role carries isFinanceAdmin/isDocAdmin too (see sharepoint-adapter.js's
   getCurrentUserRole: `isFinanceAdmin: role === "Admin" || role === "Finance
   Admin"`, same pattern for isDocAdmin) — checking those narrower roles
   first would wrongly route a full Admin to Travel or Document Review
   instead of Team Overview. */
function pulseComputeDefaultRoute() {
  if (window.PULSE_PORT_CONFIG && window.PULSE_PORT_CONFIG.defaultRoute) {
    return window.PULSE_PORT_CONFIG.defaultRoute;
  }
  const user = window.AEWTTR && window.AEWTTR.db && window.AEWTTR.db.user;
  const savedDefault = user && typeof normalizeNotificationPrefs === "function"
    ? normalizeNotificationPrefs(user.notificationPrefs).defaultPage
    : "";
  if (savedDefault) return savedDefault;
  if (user && user.isAdmin) {
    if (window.AEWTTR) {
      window.AEWTTR.state = window.AEWTTR.state || {};
      window.AEWTTR.state.overviewView = "Team";
    }
    return "overview";
  }
  if (user && user.isFinanceAdmin) return "travel/finance";
  if (user && user.isDocAdmin) return "docreview";
  return "dashboard";
}

function wirePulseBrandHome(node) {
  if (!node) return;
  const targetRoute = pulseComputeDefaultRoute();
  node.addEventListener("click", () => {
    if (typeof navigate === "function") navigate(targetRoute);
  });
  node.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      if (typeof navigate === "function") navigate(targetRoute);
    }
  });
}

function ensureBootLogPanel() {
  const root = document.getElementById("aewttr-root");
  if (!root) return null;
  if (document.getElementById("aewttr-boot-screen")) return document.getElementById("aewttr-bootlog-lines");
  root.innerHTML = `
    <div id="aewttr-boot-screen" class="pulse-boot-screen">
      <main class="pulse-loader-panel" role="status" aria-label="Loading PULSE">
        <div class="pulse-loader-brand" aria-hidden="true">
          <span class="pulse-loader-mark">
            <i class="pulse-loader-mark-dot pulse-loader-mark-dot--one"></i>
            <i class="pulse-loader-mark-dot pulse-loader-mark-dot--two"></i>
          </span>
          <div class="pulse-loader-wordmark"><span>P</span><span>U</span><span>L</span><span>S</span><span>E</span></div>
        </div>
        <div class="pulse-loader-progress" aria-hidden="true">
          <span id="aewttr-boot-progress" class="pulse-loader-progress-fill" style="width:6%"></span>
        </div>
        <div class="pulse-loader-meta" aria-live="polite">
          <span id="pulse-boot-status">Loading workspace</span>
          <span id="pulse-boot-percent">6%</span>
        </div>
        <span id="pulse-boot-step" class="pulse-loader-sr">${BOOT_PROGRESS_STEPS[0]}</span>
      </main>
      <div id="aewttr-bootlog-lines" hidden></div>
    </div>
  `;
  return document.getElementById("aewttr-bootlog-lines");
}

function waitForBootMinimum(startedAt) {
  const remaining = PULSE_BOOT_MIN_DISPLAY_MS - (performance.now() - startedAt);
  if (remaining <= 0) return Promise.resolve();
  return new Promise((resolve) => setTimeout(resolve, remaining));
}

/* Fills the boot progress bar to an absolute percentage — called at known
   milestones in bootSharePointMode so it reads as real progress rather than
   an indeterminate spinner. No-op once boot has completed or if the boot
   screen isn't showing (local-mode boot is instant enough to skip it). */
function setBootProgress(pct) {
  if (window.AEWTTR.bootComplete) return;
  const safePct = Math.max(0, Math.min(100, pct));
  const bar = document.getElementById("aewttr-boot-progress");
  const percent = document.getElementById("pulse-boot-percent");
  const step = document.getElementById("pulse-boot-step");
  if (bar) bar.style.width = safePct + "%";
  if (percent) percent.textContent = Math.round(safePct) + "%";
  const index = Math.min(BOOT_PROGRESS_STEPS.length - 1, Math.floor((safePct / 100) * BOOT_PROGRESS_STEPS.length));
  if (step) step.textContent = BOOT_PROGRESS_STEPS[index];
  const status = document.getElementById("pulse-boot-status");
  if (status) status.textContent = BOOT_PROGRESS_STEPS[index];
}

function bootLog(message, type) {
  window.AEWTTR.bootMessages = window.AEWTTR.bootMessages || [];
  const prefix = type === "error" ? "[error]" : (type === "success" ? "[ ok ]" : "[....]");
  window.AEWTTR.bootMessages.push(`${prefix} ${message}`);
  if (window.AEWTTR.bootMessages.length > BOOT_LOG_CAP) window.AEWTTR.bootMessages.shift();
  // Never touch the DOM once the app has finished booting and rendered the
  // real shell — otherwise any later bootLog call (background refresh,
  // "Reload List Data", etc.) would wipe out the live app.
  if (window.AEWTTR.bootComplete) return;
  const lines = ensureBootLogPanel();
  if (!lines) return;
  const row = document.createElement("div");
  row.textContent = `${prefix} ${message}`;
  if (type === "error") row.style.color = "#ff9b9b";
  if (type === "success") row.style.color = "#9af7b8";
  lines.appendChild(row);
  while (lines.children.length > BOOT_LOG_CAP) lines.removeChild(lines.firstChild);
}

/* ---------- tiny helpers ---------- */
function $(sel, root) { return (root || document).querySelector(sel); }
function $all(sel, root) { return Array.from((root || document).querySelectorAll(sel)); }
function el(html) { const d = document.createElement("div"); d.innerHTML = html.trim(); return d.firstElementChild; }
function escapeHtml(s) {
  return String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
function tip(text) {
  if (text == null || text === "") return "";
  return ` data-tip="${escapeHtml(String(text))}"`;
}
function ganttTip(text) {
  if (text == null || text === "") return "";
  return ` data-gantt-tip="${escapeHtml(String(text))}"`;
}

const GLOSSARY = {
  "Portfolio": "A named grouping of related projects sharing a common mission area, sponsor, or resource pool. Portfolios let you filter reports and status decks across multiple projects at once.",
  "Program": "Program Office — a higher-level initiative or office sponsoring multiple portfolios (e.g. EW, SHRAM). Used to aggregate projects in master status decks.",
  "Contract": "The contract number or name under which this project is funded and executed.",
  "Lifecycle": "The current phase of a project: Planned → Awaiting Funding → Active → Paused → Complete.",
  "End-item config": "The hardware or software end-item this project delivers. Used to group related projects across programs for reporting.",
  "ATO": "Authority to Operate — formal DoD approval to deploy an information system after a security review.",
  "Deliverable": "A specific, tangible output — hardware, software, a document, or a service — with a defined due date.",
  "Milestone": "A key checkpoint in the project schedule: a date by which something significant happens (e.g. FAT, SAT, Contract Award)."
};
function glossaryTip(term) {
  if (!GLOSSARY[term]) return "";
  return `<span class="glossary-tip-icon"${tip(GLOSSARY[term])}><i class="bx bx-info-circle" aria-hidden="true"></i></span>`;
}

/* Portaled tooltips — fixed to viewport so overflow:hidden ancestors never clip them. */
const FLOAT_TIP_DELAY_MS = 2000;
let _floatTipEl = null;
let _floatTipTimer = null;
let _floatTipTarget = null;

function ensureFloatTipEl() {
  if (_floatTipEl) return _floatTipEl;
  _floatTipEl = document.createElement("div");
  _floatTipEl.className = "pulse-float-tip";
  _floatTipEl.setAttribute("role", "tooltip");
  _floatTipEl.hidden = true;
  document.body.appendChild(_floatTipEl);
  return _floatTipEl;
}

function hideFloatTip() {
  if (_floatTipTimer) {
    clearTimeout(_floatTipTimer);
    _floatTipTimer = null;
  }
  _floatTipTarget = null;
  if (_floatTipEl) _floatTipEl.hidden = true;
}

function positionFloatTip(target) {
  const tip = ensureFloatTipEl();
  const rect = target.getBoundingClientRect();
  const centerX = rect.left + rect.width / 2;
  tip.style.left = `${Math.min(Math.max(centerX, 12), window.innerWidth - 12)}px`;
  tip.style.top = `${rect.top - 10}px`;
  tip.style.transform = "translate(-50%, -100%)";
  tip.hidden = false;
  const tipRect = tip.getBoundingClientRect();
  if (tipRect.top < 8) {
    tip.style.top = `${rect.bottom + 10}px`;
    tip.style.transform = "translate(-50%, 0)";
  }
  if (tipRect.left < 8) {
    tip.style.left = "8px";
    tip.style.transform = tip.style.transform.replace("translate(-50%", "translate(0");
  } else if (tipRect.right > window.innerWidth - 8) {
    tip.style.left = `${window.innerWidth - 8}px`;
    tip.style.transform = tip.style.transform.replace("translate(-50%", "translate(-100%");
  }
}

function wirePortaledTooltips(root) {
  root = root || document;
  if (root === document && document.documentElement.dataset.portaledTipsWired === "1") return;
  root.addEventListener("mouseover", (e) => {
    const el = e.target.closest("[data-gantt-tip], [data-tip]");
    if (!el || !root.contains(el)) return;
    if (_floatTipTarget === el) return;
    hideFloatTip();
    _floatTipTarget = el;
    const text = el.getAttribute("data-gantt-tip") || el.getAttribute("data-tip");
    if (!text) return;
    _floatTipTimer = setTimeout(() => {
      if (_floatTipTarget !== el) return;
      const tip = ensureFloatTipEl();
      tip.textContent = text;
      positionFloatTip(el);
    }, FLOAT_TIP_DELAY_MS);
  });
  root.addEventListener("mouseout", (e) => {
    const from = e.target.closest("[data-gantt-tip], [data-tip]");
    if (!from) return;
    const to = e.relatedTarget;
    if (to && from.contains(to)) return;
    hideFloatTip();
  });
  root.addEventListener("scroll", hideFloatTip, true);
  if (root === document) document.documentElement.dataset.portaledTipsWired = "1";
}
window.hideFloatTip = hideFloatTip;
function initials(name) {
  if (!name) return "?";
  return name.split(" ").filter(Boolean).slice(0, 2).map(p => p[0].toUpperCase()).join("");
}
function currentUserProfilePhotoUrl() {
  const spUser = window.AEWTTR && window.AEWTTR.currentSpUser;
  const siteUrl = (window.AEWTTR && window.AEWTTR.siteUrl) || "";
  const email = (spUser && spUser.email) || (window.AEWTTR && window.AEWTTR.db && window.AEWTTR.db.user && window.AEWTTR.db.user.email) || "";
  return userPhotoUrlForPerson("", email);
}
function memberEmailForPerson(name, email) {
  const direct = String(email || "").trim();
  if (direct) return direct;
  const key = String(name || "").trim().toLowerCase();
  if (!key) return "";
  const member = (window.AEWTTR.db && window.AEWTTR.db.members || []).find((entry) => String(entry.name || "").trim().toLowerCase() === key);
  return member && member.email ? String(member.email).trim() : "";
}
function userPhotoUrlForPerson(name, email) {
  const siteUrl = (window.AEWTTR && window.AEWTTR.siteUrl) || "";
  const resolvedEmail = memberEmailForPerson(name, email);
  if (!siteUrl || !resolvedEmail) return "";
  return `${siteUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(resolvedEmail)}`;
}
function userAvatarHtml(nameOrOpts, emailMaybe, sizeMaybe) {
  const opts = (nameOrOpts && typeof nameOrOpts === "object") ? nameOrOpts : { name: nameOrOpts, email: emailMaybe, size: sizeMaybe };
  const name = opts.name || "User";
  const email = opts.email || "";
  const size = Number(opts.size || 21);
  const className = opts.className || "kc-avatar";
  const style = opts.style || `width:${size}px;height:${size}px;font-size:${Math.max(9, Math.round(size * 0.42))}px;`;
  const photoUrl = userPhotoUrlForPerson(name, email);
  if (photoUrl) {
    return `<span class="${className} has-photo" style="${style}"><img class="kc-avatar-photo" src="${escapeHtml(photoUrl)}" alt="${escapeHtml(name)}" onerror="this.remove(); this.parentElement.classList.remove('has-photo');"><span class="kc-avatar-fallback">${escapeHtml(initials(name))}</span></span>`;
  }
  return `<span class="${className}" style="${style}">${escapeHtml(initials(name))}</span>`;
}
function uid(prefix) { return prefix + "-" + Math.random().toString(36).slice(2, 9); }
function fmtDate(d) {
  if (!d) return "—";
  const dt = new Date(d + "T00:00:00");
  if (isNaN(dt)) return d;
  return dt.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}
function isOverdue(dateStr) {
  if (!dateStr) return false;
  return new Date(dateStr + "T00:00:00") < new Date(new Date().toDateString());
}
function ragPill(rag) {
  return `<span class="rag-pill rag-${rag}"><span class="dot"></span>${rag}</span>`;
}
function statusPill(status) {
  const cls = "status-" + String(status).replace(/\s+/g, "-");
  return `<span class="status-pill ${cls}">${escapeHtml(status)}</span>`;
}
function priorityTag(p) {
  if (!p) return `<span class="priority-tag priority-None">None</span>`;
  const key = ["High", "Immediate"].includes(p) ? "High" : (p === "Low" || p === "Lower" || p === "None" ? "Low" : "Medium");
  return `<span class="priority-tag priority-${key}">${escapeHtml(p)}</span>`;
}
function lifecyclePill(l) {
  if (!l) return `<span class="lifecycle-pill lifecycle-None">None</span>`;
  const cls = "lifecycle-" + String(l).replace(/\s+/g, "-");
  return `<span class="lifecycle-pill ${cls}">${escapeHtml(l)}</span>`;
}

function currentUserMember() {
  const db = window.AEWTTR.db || {};
  const user = db.user || {};
  const email = String(user.email || "").trim().toLowerCase();
  return (db.members || []).find((member) => {
    if (email && member.email && String(member.email).trim().toLowerCase() === email) return true;
    if (user.id && member.id === user.id) return true;
    return !!(user.name && member.name === user.name);
  }) || null;
}
function memberMatchesAssignee(member, assignee) {
  if (!assignee) return false;
  const name = String(assignee).trim();
  if (!name) return false;
  if (!member) return false;
  const full = member.name.toLowerCase();
  const short = name.toLowerCase();
  return full === short || full.startsWith(short + " ") || full.split(" ")[0] === short;
}
function tasksForCurrentUser() {
  const db = window.AEWTTR.db || {};
  const user = db.user || {};
  const member = currentUserMember();
  const rows = [];
  (db.projects || []).forEach((project) => {
    (db.ganttTasks[project.id] || []).forEach((task) => {
      if (task.status === "Done") return;
      const mine = memberMatchesAssignee(member, task.assignee) || task.assignee === user.name;
      if (mine) rows.push({ project, task });
    });
  });
  return rows.sort((a, b) => String(a.task.end || "").localeCompare(String(b.task.end || "")));
}
/* ---------- Jira-style issue chrome (type icon, key, points, priority) ---------- */
/* Issue "type" is inferred from real task signals so the tracker reads like a Jira
   board without adding a new data field: blocked/off-track → Bug, has subtasks →
   Story, everything else → Task. An explicit task.type always wins if set. */
function issueType(task) {
  if (task && task.type) return task.type;
  if (task && (task.health === "Off Track" || task.status === "Blocked")) return "bug";
  if (task && task.subtasks && task.subtasks.length) return "story";
  return "task";
}
const ISSUE_TYPE_META = {
  epic:    { icon: "bx-bolt-circle", label: "Epic" },
  story:   { icon: "bx-bookmark",    label: "Story" },
  task:    { icon: "bx-check",       label: "Task" },
  bug:     { icon: "bx-circle",      label: "Bug" },
  subtask: { icon: "bx-subdirectory-right", label: "Subtask" }
};
function issueTypeIcon(task) {
  const type = issueType(task);
  const meta = ISSUE_TYPE_META[type] || ISSUE_TYPE_META.task;
  return `<span class="issue-type it-${type}" title="${meta.label}"><i class="bx ${meta.icon}"></i></span>`;
}
/* Stable, human-readable key derived from the project code + a deterministic
   number hashed from the task id, so it never shifts when tasks are reordered. */
function issueSeq(id) {
  const s = String(id == null ? "" : id);
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0;
  return (h % 899) + 100;
}
function issueKey(projectCode, task) {
  const code = String(projectCode || "PULSE").replace(/[^A-Za-z0-9]/g, "").toUpperCase() || "PULSE";
  return `<span class="issue-key">${escapeHtml(code)}-${issueSeq(task.id)}</span>`;
}
function issuePoints(task) {
  const n = task && task.subtasks ? task.subtasks.length : 0;
  if (!n) return "";
  return `<span class="issue-points" title="${n} subtask${n === 1 ? "" : "s"}">${n}</span>`;
}
function issuePriority(p) {
  if (!p || p === "None") return "";
  const key = ["Critical", "Immediate", "Highest"].includes(p) ? "highest"
    : (p === "High" ? "high"
    : (p === "Low" || p === "Lower" ? "low" : "medium"));
  const icon = { highest: "bx-chevrons-up", high: "bx-chevron-up", medium: "bx-minus", low: "bx-chevron-down" }[key];
  return `<span class="jira-prio jp-${key}" title="${escapeHtml(p)} priority"><i class="bx ${icon}"></i></span>`;
}

/* ---------- configurable RAG logic ---------- */
function ragConfigDefaults() {
  return { overdueAmber: 1, overdueRed: 3, behindAmber: 1, behindRed: 2, completionAmberBelow: 60, completionRedBelow: 25 };
}
function getRagConfig() {
  return ragConfigDefaults();
}

/* ---------- AI (CapraGPT) settings — always read from APP_CONFIG.aiReview ---------- */
function getAiConfig() {
  const fallback = (typeof APP_CONFIG !== "undefined" && APP_CONFIG.aiReview) || {};
  return {
    apiKey: fallback.apiKey || "",
    endpoint: fallback.endpoint || "https://api.capragpt.mil/v1/chat/completions",
    model: fallback.model || "gpt-4.1",
    enabled: fallback.enabled !== false
  };
}

/* ---------- configurable location list (Admin > Locations) ---------- */
function locationConfigDefaults() {
  return { locations: [], portfolios: [], contractors: [], configEndItems: [], hideUnaffiliatedPeople: true, meetingAdminEmails: [] };
}
function getLocationConfig() {
  const db = window.AEWTTR.db;
  if (!db.locationConfig) db.locationConfig = locationConfigDefaults();
  if (!Array.isArray(db.locationConfig.locations)) db.locationConfig.locations = [];
  if (!Array.isArray(db.locationConfig.portfolios)) db.locationConfig.portfolios = [];
  if (!Array.isArray(db.locationConfig.contractors)) db.locationConfig.contractors = [];
  if (!Array.isArray(db.locationConfig.configEndItems)) db.locationConfig.configEndItems = [];
  if (db.locationConfig.hideUnaffiliatedPeople == null) db.locationConfig.hideUnaffiliatedPeople = true;
  if (!Array.isArray(db.locationConfig.meetingAdminEmails)) db.locationConfig.meetingAdminEmails = [];
  return db.locationConfig;
}

/* Legacy Location Config flag — superseded by All users / AEWTTR tabs +
   Assign to PULSE. Kept so older config rows still load; pickers and People
   always use AEWTTR-associated members now. */
function hideUnaffiliatedPeopleEnabled() {
  return true;
}

/* True when the signed-in user may use the app: Admin role, or an active
   PULSE App Roles record (associated). */
function isCurrentUserPulseAssociated() {
  if (window.AEWTTR && window.AEWTTR.pulseAssociated) return true;
  if (canCurrentUserAccessAdmin()) return true;
  const diagnosticsRole = window.AEWTTR && window.AEWTTR.spDiagnostics && window.AEWTTR.spDiagnostics.currentRole;
  if (diagnosticsRole && typeof sharePointAdapter !== "undefined" && sharePointAdapter.isUserPulseAssociated) {
    return !!sharePointAdapter.isUserPulseAssociated(diagnosticsRole);
  }
  if (diagnosticsRole && (diagnosticsRole.isAdmin || diagnosticsRole.roleRecordId || diagnosticsRole.associated)) return true;
  const user = window.AEWTTR.db && window.AEWTTR.db.user;
  if (user && user.isAdmin) return true;
  if (user && user.email && window.AEWTTR.db && Array.isArray(window.AEWTTR.db.members)) {
    const email = String(user.email).trim().toLowerCase();
    if (window.AEWTTR.db.members.some((m) => String(m.email || "").trim().toLowerCase() === email)) return true;
  }
  return false;
}

function renderPulseAccessBlockedScreen() {
  const root = $("#aewttr-root");
  if (!root) return;
  const name = (window.AEWTTR.db && window.AEWTTR.db.user && window.AEWTTR.db.user.name)
    || (window.AEWTTR.currentSpUser && window.AEWTTR.currentSpUser.displayName)
    || "there";
  root.innerHTML = `
    <div class="pulse-access-blocked" role="alert">
      <div class="pulse-access-blocked-inner">
        ${pulseBrandLockup({ boot: true, animate: false })}
        <h1 class="pulse-access-blocked-title">You’re signed in, but not set up for PULSE yet</h1>
        <p class="pulse-access-blocked-copy">
          Hi ${escapeHtml(name)} — an admin needs to add you from <strong>Users → All users → Assign to PULSE</strong> before you can open the workspace.
        </p>
        <p class="pulse-access-blocked-hint">If you believe this is a mistake, contact your PULSE admin.</p>
      </div>
    </div>
  `;
  window.AEWTTR.accessBlocked = true;
}

/* ---------- project portfolios (multi-select; remembered names) ---------- */
function normalizePortfolioName(name) {
  return String(name == null ? "" : name).trim().replace(/\s+/g, " ");
}
function projectPortfolios(proj) {
  if (!proj) return [];
  const raw = Array.isArray(proj.portfolios) ? proj.portfolios : [];
  const seen = new Set();
  const out = [];
  raw.forEach((name) => {
    const n = normalizePortfolioName(name);
    if (!n || seen.has(n.toLowerCase())) return;
    seen.add(n.toLowerCase());
    out.push(n);
  });
  return out;
}

/* Project deliverables are stored as a small ordered list. Older project
   records used a free-form Objectives field; read it once as a backwards-
   compatible source so existing text appears in the new editor. */
function projectDeliverables(proj) {
  if (!proj) return [];
  const raw = Array.isArray(proj.deliverables)
    ? proj.deliverables
    : String(proj.deliverables || proj.objectives || "").split(/\r?\n/);
  return raw
    .map((item) => String(item == null ? "" : item).replace(/^\s*[-*•]\s*/, "").trim())
    .filter(Boolean);
}
function getKnownPortfolioNames() {
  const names = new Set();
  const cfg = getLocationConfig();
  (cfg.portfolios || []).forEach((n) => {
    const name = normalizePortfolioName(n);
    if (name) names.add(name);
  });
  const db = window.AEWTTR && window.AEWTTR.db;
  ((db && db.projects) || []).forEach((p) => {
    projectPortfolios(p).forEach((n) => names.add(n));
  });
  return [...names].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
}
function rememberPortfolioNames(names) {
  const cfg = getLocationConfig();
  let changed = false;
  (names || []).forEach((raw) => {
    const name = normalizePortfolioName(raw);
    if (!name) return;
    const exists = (cfg.portfolios || []).some((n) => normalizePortfolioName(n).toLowerCase() === name.toLowerCase());
    if (exists) return;
    cfg.portfolios.push(name);
    changed = true;
  });
  if (!changed) return;
  cfg.portfolios = getKnownPortfolioNames();
  if (typeof Repo !== "undefined" && Repo && typeof Repo.save === "function") {
    try { Repo.save("locationConfig", cfg); } catch (_) { /* local / pre-init */ }
  }
}

/* ---------- project locations (multi-select; remembered like portfolios) ---------- */
function normalizeLocationName(name) {
  return String(name == null ? "" : name).trim().replace(/\s+/g, " ");
}
function projectLocations(proj) {
  if (!proj) return [];
  const raw = Array.isArray(proj.locations) ? proj.locations : [];
  const seen = new Set();
  const out = [];
  raw.forEach((name) => {
    const n = normalizeLocationName(name);
    if (!n || seen.has(n.toLowerCase())) return;
    seen.add(n.toLowerCase());
    out.push(n);
  });
  return out;
}
function getKnownLocationNames() {
  const names = new Set();
  const cfg = getLocationConfig();
  (cfg.locations || []).forEach((n) => {
    const name = normalizeLocationName(n);
    if (name) names.add(name);
  });
  const db = window.AEWTTR && window.AEWTTR.db;
  ((db && db.projects) || []).forEach((p) => {
    projectLocations(p).forEach((n) => names.add(n));
  });
  return [...names].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
}
function rememberLocationNames(names) {
  const cfg = getLocationConfig();
  let changed = false;
  (names || []).forEach((raw) => {
    const name = normalizeLocationName(raw);
    if (!name) return;
    const exists = (cfg.locations || []).some((n) => normalizeLocationName(n).toLowerCase() === name.toLowerCase());
    if (exists) return;
    cfg.locations.push(name);
    changed = true;
  });
  if (!changed) return;
  cfg.locations = getKnownLocationNames();
  if (typeof Repo !== "undefined" && Repo && typeof Repo.save === "function") {
    try { Repo.save("locationConfig", cfg); } catch (_) { /* local / pre-init */ }
  }
}

/* ---------- contractor catalog (remembered like portfolios / locations) ---------- */
function normalizeContractorName(name) {
  return String(name == null ? "" : name).trim().replace(/\s+/g, " ");
}
function parseContractorList(value) {
  if (Array.isArray(value)) {
    return value.map(normalizeContractorName).filter(Boolean);
  }
  const raw = String(value == null ? "" : value).trim();
  if (!raw) return [];
  // Prefer semicolon splits so company names may contain commas.
  const parts = raw.includes(";") ? raw.split(";") : [raw];
  const seen = new Set();
  const out = [];
  parts.forEach((part) => {
    const n = normalizeContractorName(part);
    if (!n || seen.has(n.toLowerCase())) return;
    seen.add(n.toLowerCase());
    out.push(n);
  });
  return out;
}
function getKnownContractorNames() {
  const names = new Set();
  const cfg = getLocationConfig();
  (cfg.contractors || []).forEach((n) => {
    const name = normalizeContractorName(n);
    if (name) names.add(name);
  });
  const db = window.AEWTTR && window.AEWTTR.db;
  ((db && db.projects) || []).forEach((p) => {
    parseContractorList(p.contractor).forEach((n) => names.add(n));
    const finance = db.projectExtra && db.projectExtra[p.id] && db.projectExtra[p.id].finance;
    if (finance && finance.summary) parseContractorList(finance.summary.contractor).forEach((n) => names.add(n));
  });
  Object.keys((db && db.ganttTasks) || {}).forEach((pid) => {
    (db.ganttTasks[pid] || []).forEach((task) => {
      if (!(task && (task.itemType === "divider" || task.workItemLevel === "Divider"))) return;
      parseContractorList(task.contractor || (task.metadata && task.metadata.contractor)).forEach((n) => names.add(n));
    });
  });
  ((db && db.documents) || []).forEach((doc) => {
    parseContractorList(doc.contractorName).forEach((n) => names.add(n));
  });
  return [...names].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
}
function rememberContractorNames(names) {
  const cfg = getLocationConfig();
  let changed = false;
  (names || []).forEach((raw) => {
    const name = normalizeContractorName(raw);
    if (!name) return;
    const exists = (cfg.contractors || []).some((n) => normalizeContractorName(n).toLowerCase() === name.toLowerCase());
    if (exists) return;
    cfg.contractors.push(name);
    changed = true;
  });
  if (!changed) return;
  cfg.contractors = getKnownContractorNames();
  if (typeof Repo !== "undefined" && Repo && typeof Repo.save === "function") {
    try { Repo.save("locationConfig", cfg); } catch (_) { /* local / pre-init */ }
  }
}

/* ---------- config end items (single primary value; catalog like portfolios) ---------- */
function normalizeConfigEndItemName(name) {
  return String(name == null ? "" : name).trim().replace(/\s+/g, " ");
}
function getKnownConfigEndItemNames() {
  const names = new Set();
  const cfg = getLocationConfig();
  (cfg.configEndItems || []).forEach((n) => {
    const name = normalizeConfigEndItemName(n);
    if (name) names.add(name);
  });
  const db = window.AEWTTR && window.AEWTTR.db;
  ((db && db.projects) || []).forEach((p) => {
    const name = normalizeConfigEndItemName(p && p.configEndItem);
    if (name) names.add(name);
  });
  Object.keys((db && db.ganttTasks) || {}).forEach((pid) => {
    (db.ganttTasks[pid] || []).forEach((task) => {
      if (!(task && (task.itemType === "divider" || task.workItemLevel === "Divider"))) return;
      const name = normalizeConfigEndItemName(task.configEndItem || (task.metadata && task.metadata.configEndItem));
      if (name) names.add(name);
    });
  });
  ((db && db.documents) || []).forEach((doc) => {
    (doc.configEndItems || []).forEach((n) => {
      const name = normalizeConfigEndItemName(n);
      if (name) names.add(name);
    });
    const legacy = normalizeConfigEndItemName(doc.portfolioOrConfigItem);
    if (legacy) names.add(legacy);
  });
  return [...names].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
}
function rememberConfigEndItemNames(names) {
  const cfg = getLocationConfig();
  let changed = false;
  (names || []).forEach((raw) => {
    const name = normalizeConfigEndItemName(raw);
    if (!name) return;
    const exists = (cfg.configEndItems || []).some((n) => normalizeConfigEndItemName(n).toLowerCase() === name.toLowerCase());
    if (exists) return;
    cfg.configEndItems.push(name);
    changed = true;
  });
  if (!changed) return;
  cfg.configEndItems = getKnownConfigEndItemNames();
  if (typeof Repo !== "undefined" && Repo && typeof Repo.save === "function") {
    try { Repo.save("locationConfig", cfg); } catch (_) { /* local / pre-init */ }
  }
}

/* ---- Tag-picker helpers for program/funding fields ---- */
function _makeFieldHelpers(cfgKey, projField) {
  const norm = (name) => String(name == null ? "" : name).trim().replace(/\s+/g, " ");
  const getKnown = () => {
    const names = new Set();
    const cfg = getLocationConfig();
    ((cfg[cfgKey]) || []).forEach(n => { const v = norm(n); if (v) names.add(v); });
    const db = window.AEWTTR && window.AEWTTR.db;
    ((db && db.projects) || []).forEach(p => { const v = norm(p && p[projField]); if (v) names.add(v); });
    return [...names].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
  };
  const remember = (names) => {
    const cfg = getLocationConfig();
    if (!cfg[cfgKey]) cfg[cfgKey] = [];
    let changed = false;
    (names || []).forEach(raw => {
      const v = norm(raw);
      if (!v) return;
      if (cfg[cfgKey].some(n => norm(n).toLowerCase() === v.toLowerCase())) return;
      cfg[cfgKey].push(v);
      changed = true;
    });
    if (!changed) return;
    cfg[cfgKey] = getKnown();
    if (typeof Repo !== "undefined" && Repo && typeof Repo.save === "function") {
      try { Repo.save("locationConfig", cfg); } catch (_) {}
    }
  };
  return { norm, getKnown, remember };
}
const _fundingTypeHelpers  = _makeFieldHelpers("fundingTypes",  "fundingType");
const _fiscalYearHelpers   = _makeFieldHelpers("fiscalYears",   "fiscalYear");
const _fundingStatusHelpers= _makeFieldHelpers("fundingStatuses","fundingStatus");
const _taskOrderHelpers    = _makeFieldHelpers("taskOrders",    "taskOrder");
const _programHelpers      = _makeFieldHelpers("programs",      "program");

function normalizeFundingTypeName(n)   { return _fundingTypeHelpers.norm(n); }
function getKnownFundingTypeNames()    { return _fundingTypeHelpers.getKnown(); }
function rememberFundingTypeNames(ns)  { return _fundingTypeHelpers.remember(ns); }

function normalizeFiscalYearName(n)    { return _fiscalYearHelpers.norm(n); }
function getKnownFiscalYearNames()     { return _fiscalYearHelpers.getKnown(); }
function rememberFiscalYearNames(ns)   { return _fiscalYearHelpers.remember(ns); }

function normalizeFundingStatusName(n) { return _fundingStatusHelpers.norm(n); }
function getKnownFundingStatusNames()  { return _fundingStatusHelpers.getKnown(); }
function rememberFundingStatusNames(ns){ return _fundingStatusHelpers.remember(ns); }

function normalizeTaskOrderName(n)     { return _taskOrderHelpers.norm(n); }
function getKnownTaskOrderNames()      { return _taskOrderHelpers.getKnown(); }
function rememberTaskOrderNames(ns)    { return _taskOrderHelpers.remember(ns); }

function normalizeProgramName(n)       { return _programHelpers.norm(n); }
function getKnownProgramNames()        { return _programHelpers.getKnown(); }
function rememberProgramNames(ns)      { return _programHelpers.remember(ns); }

const _contractHelpers     = _makeFieldHelpers("contracts",     "contract");
function normalizeContractName(n)      { return _contractHelpers.norm(n); }
function getKnownContractNames()       { return _contractHelpers.getKnown(); }
function rememberContractNames(ns)     { return _contractHelpers.remember(ns); }

const _atoHelpers = _makeFieldHelpers("atos", "ato");
function normalizeAtoName(n)           { return _atoHelpers.norm(n); }
function getKnownAtoNames()            { return _atoHelpers.getKnown(); }
function rememberAtoNames(ns)          { return _atoHelpers.remember(ns); }

/* % of a project's tracker tasks marked Done. Null (not 0) when the
   project has no tasks yet, so callers can distinguish "no data" from
   "0% complete" instead of treating an empty tracker as a red flag. */
function computeProjectCompletionPct(proj) {
  const db = window.AEWTTR.db;
  const tasks = (db.ganttTasks && db.ganttTasks[proj.id]) || [];
  if (!tasks.length) return null;
  const doneCount = tasks.filter(t => t.status === "Done").length;
  return Math.round((doneCount / tasks.length) * 100);
}
function computeProjectRag(proj) {
  const db = window.AEWTTR.db;
  const tasks = (db.ganttTasks && db.ganttTasks[proj.id]) || [];
  if (!tasks.length) return proj.rag || "Green";
  const cfg = getRagConfig();
  const behindCount = tasks.filter(t => t.health === "At Risk" || t.health === "Off Track").length;
  const completionPct = computeProjectCompletionPct(proj);

  if (behindCount >= cfg.behindRed || completionPct < cfg.completionRedBelow) return "Red";
  if (behindCount >= cfg.behindAmber || completionPct < cfg.completionAmberBelow) return "Amber";
  return "Green";
}

function flattenTaskSubitems(subs) {
  const out = [];
  function walk(list) {
    (list || []).forEach((s) => {
      out.push(s);
      walk(s.subtasks || s.children || []);
    });
  }
  walk(subs);
  return out;
}

function taskProgressPct(task) {
  const subs = flattenTaskSubitems((task && task.subtasks) || []);
  if (!subs.length) return 0;
  return Math.round((subs.filter((s) => s.done).length / subs.length) * 100);
}

function syncTaskStatusFromSubtasks(task) {
  if (!task) return;
  if (task.itemType === "divider" || task.workItemLevel === "Divider") return;
  const subs = flattenTaskSubitems(task.subtasks || []);
  if (!subs.length) {
    if (!task.status) task.status = "Not Started";
    return;
  }
  const pct = taskProgressPct(task);
  if (pct === 100) task.status = "Done";
  else if (pct > 0) task.status = "In Progress";
  else task.status = "Not Started";
}

function taskProgressBarHtml(task) {
  return "";
}

function ganttBarClassForTask(task) {
  const pct = taskProgressPct(task);
  if (pct === 100) return "status-Done";
  if (pct > 0) return "status-In-Progress";
  return "status-Not-Started";
}

/* Project "Status" (Not Started/In Progress/Blocked/Done) used to be a
   manually-set field on the project (extra.status) — separate from RAG
   (red/amber/green health, already automatic above). It's derived from the
   tracker tasks instead now, so it can never drift from what's actually
   tracked. Same four values the old manual dropdown offered, so every
   existing display (stat card, kv-row, statusPill) keeps working unchanged. */
function computeProjectStatus(proj) {
  const db = window.AEWTTR.db;
  const tasks = (db.ganttTasks[proj.id] || []);
  if (!tasks.length) return "Not Started";
  if (tasks.every((t) => taskProgressPct(t) === 100)) return "Done";
  if (tasks.some((t) => t.health === "Off Track")) return "Blocked";
  if (tasks.some((t) => taskProgressPct(t) > 0)) return "In Progress";
  return "Not Started";
}

/* ---------- toast ---------- */
function toast(msg, type) {
  let stack = $(".aewttr-toast-stack");
  if (!stack) {
    stack = el(`<div class="aewttr-toast-stack"></div>`);
    document.body.appendChild(stack);
  }
  const t = el(`<div class="aewttr-toast ${type || ""}">${escapeHtml(msg)}</div>`);
  stack.appendChild(t);
  setTimeout(() => { t.style.opacity = "0"; t.style.transition = "opacity .25s"; setTimeout(() => t.remove(), 250); }, 3500);
}

/* ---------- context menu ---------- */
let _contextMenuEl = null;
function hideContextMenu() {
  if (_contextMenuEl) {
    _contextMenuEl.remove();
    _contextMenuEl = null;
  }
  document.removeEventListener("click", hideContextMenu, true);
  document.removeEventListener("contextmenu", _hideContextMenuOnNew, true);
  document.removeEventListener("keydown", _contextMenuEsc, true);
}
function _hideContextMenuOnNew() { hideContextMenu(); }
function _contextMenuEsc(e) { if (e.key === "Escape") hideContextMenu(); }
function showContextMenu(x, y, items) {
  hideContextMenu();
  if (!items || !items.length) return;
  const menu = el(`<div class="aewttr-context-menu" role="menu"></div>`);
  items.forEach((item) => {
    if (item.separator) {
      menu.appendChild(el(`<div class="aewttr-context-menu-sep" role="separator"></div>`));
      return;
    }
    const btn = el(`<button type="button" class="aewttr-context-menu-item${item.danger ? " is-danger" : ""}" role="menuitem">${item.icon ? `<i class="bx ${item.icon}"></i>` : ""}<span>${escapeHtml(item.label)}</span></button>`);
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      hideContextMenu();
      item.action();
    });
    menu.appendChild(btn);
  });
  document.body.appendChild(menu);
  _contextMenuEl = menu;
  const rect = menu.getBoundingClientRect();
  const left = Math.min(x, window.innerWidth - rect.width - 8);
  const top = Math.min(y, window.innerHeight - rect.height - 8);
  menu.style.left = `${Math.max(8, left)}px`;
  menu.style.top = `${Math.max(8, top)}px`;
  setTimeout(() => {
    document.addEventListener("click", hideContextMenu, true);
    document.addEventListener("contextmenu", _hideContextMenuOnNew, true);
    document.addEventListener("keydown", _contextMenuEsc, true);
  }, 0);
}

/* ---------- modal ---------- */
let _activeModalDismiss = null;
function openModal(innerHtml, opts) {
  closeModal();
  opts = opts || {};
  const backdrop = el(`<div class="aewttr-modal-backdrop"></div>`);
  const modalClasses = [
    opts.xwide ? "xwide" : (opts.wide ? "wide" : ""),
    opts.docreview ? "docreview-modal" : "",
    opts.className || ""
  ].filter(Boolean).join(" ");
  const modal = el(`<div class="aewttr-modal ${modalClasses}">${innerHtml}</div>`);
  backdrop.appendChild(modal);
  document.body.appendChild(backdrop);
  _activeModalDismiss = typeof opts.onDismiss === "function" ? opts.onDismiss : null;
  backdrop.addEventListener("mousedown", (e) => {
    if (e.target !== backdrop) return;
    if (_activeModalDismiss) _activeModalDismiss();
    else closeModal();
  });
  document.addEventListener("keydown", escCloseHandler);
  return modal;
}
function escCloseHandler(e) {
  if (e.key !== "Escape") return;
  if (_activeModalDismiss) _activeModalDismiss();
  else closeModal();
}
function closeModal() {
  _activeModalDismiss = null;
  document.removeEventListener("keydown", escCloseHandler);
  const b = $(".aewttr-modal-backdrop");
  if (b) b.remove();
}

/* In-app confirm — replaces window.confirm across PULSE */
function confirmDialog(opts) {
  opts = Object.assign({
    title: "Confirm",
    message: "Are you sure?",
    confirmLabel: "Confirm",
    cancelLabel: "Cancel",
    danger: null
  }, opts || {});
  if (opts.danger == null) {
    opts.danger = /delete|remove|end/i.test(`${opts.confirmLabel} ${opts.title}`);
  }
  return new Promise((resolve) => {
    let settled = false;
    // Removes THIS dialog's own backdrop by direct reference, not a global
    // ".aewttr-modal-backdrop" lookup — confirmDialog is routinely called
    // from inside an already-open modal (e.g. a delete button in a chat
    // popup), which stacks a second backdrop on top. A global querySelector
    // returns the FIRST match in document order — the outer modal, opened
    // first — so finishing the confirm dialog was silently closing the
    // modal underneath it instead of itself.
    const finish = (result) => {
      if (settled) return;
      settled = true;
      document.removeEventListener("keydown", onKey);
      backdrop.remove();
      resolve(result);
    };
    const onKey = (e) => { if (e.key === "Escape") finish(false); };
    const backdrop = el(`<div class="aewttr-modal-backdrop"></div>`);
    const modal = el(`
      <div class="aewttr-modal aewttr-confirm-modal">
        <div class="aewttr-modal-head">
          <h3>${escapeHtml(opts.title)}</h3>
          <button class="aewttr-modal-close" type="button" aria-label="Close">&times;</button>
        </div>
        <div class="aewttr-modal-body">
          <p class="aewttr-confirm-message">${escapeHtml(opts.message)}</p>
        </div>
        <div class="aewttr-modal-foot">
          <button class="btn-aewttr-ghost" type="button" id="aewttr-confirm-cancel">${escapeHtml(opts.cancelLabel)}</button>
          <button class="${opts.danger ? "btn-danger-outline" : "btn-aewttr"}" type="button" id="aewttr-confirm-ok">${escapeHtml(opts.confirmLabel)}</button>
        </div>
      </div>
    `);
    backdrop.appendChild(modal);
    document.body.appendChild(backdrop);
    backdrop.addEventListener("mousedown", (e) => { if (e.target === backdrop) finish(false); });
    document.addEventListener("keydown", onKey);
    $(".aewttr-modal-close", modal).addEventListener("click", () => finish(false));
    $("#aewttr-confirm-cancel", modal).addEventListener("click", () => finish(false));
    $("#aewttr-confirm-ok", modal).addEventListener("click", () => finish(true));
    const okBtn = $("#aewttr-confirm-ok", modal);
    if (okBtn) okBtn.focus();
  });
}

/* Single-text-field prompt, styled like confirmDialog. Resolves the trimmed
   input string on Save, or null on Cancel/Escape/backdrop click. */
function promptDialog(opts) {
  opts = Object.assign({ title: "Enter a value", label: "", placeholder: "", value: "", confirmLabel: "Save", cancelLabel: "Cancel" }, opts || {});
  return new Promise((resolve) => {
    let settled = false;
    // See confirmDialog's comment above — removes this dialog's own
    // backdrop by reference, not a global lookup that could grab a modal
    // this dialog happens to be nested inside of.
    const finish = (result) => {
      if (settled) return;
      settled = true;
      document.removeEventListener("keydown", onKey);
      backdrop.remove();
      resolve(result);
    };
    const onKey = (e) => { if (e.key === "Escape") finish(null); };
    const backdrop = el(`<div class="aewttr-modal-backdrop"></div>`);
    const modal = el(`
      <div class="aewttr-modal aewttr-confirm-modal">
        <div class="aewttr-modal-head">
          <h3>${escapeHtml(opts.title)}</h3>
          <button class="aewttr-modal-close" type="button" aria-label="Close">&times;</button>
        </div>
        <div class="aewttr-modal-body">
          ${opts.label ? `<div class="form-row" style="margin:0;"><label>${escapeHtml(opts.label)}</label><input class="input-aewttr" id="aewttr-prompt-input" placeholder="${escapeHtml(opts.placeholder)}" value="${escapeHtml(opts.value)}"></div>` : ""}
        </div>
        <div class="aewttr-modal-foot">
          <button class="btn-aewttr-ghost" type="button" id="aewttr-prompt-cancel">${escapeHtml(opts.cancelLabel)}</button>
          <button class="btn-aewttr" type="button" id="aewttr-prompt-ok">${escapeHtml(opts.confirmLabel)}</button>
        </div>
      </div>
    `);
    backdrop.appendChild(modal);
    document.body.appendChild(backdrop);
    backdrop.addEventListener("mousedown", (e) => { if (e.target === backdrop) finish(null); });
    document.addEventListener("keydown", onKey);
    const input = $("#aewttr-prompt-input", modal);
    const submit = () => finish(input ? input.value.trim() : "");
    $(".aewttr-modal-close", modal).addEventListener("click", () => finish(null));
    $("#aewttr-prompt-cancel", modal).addEventListener("click", () => finish(null));
    $("#aewttr-prompt-ok", modal).addEventListener("click", submit);
    if (input) {
      input.addEventListener("keydown", (e) => { if (e.key === "Enter") submit(); });
      input.focus();
      input.select();
    }
  });
}

/* Indents a chat-composer textarea: Ctrl/Cmd+Enter inserts a new line one
   bullet level deeper than the current line instead of posting — lets
   someone build a nested outline inside one message before sending it.
   Shared by every message-style tool (task/subtask notes, meeting notes). */
function insertComposerIndentLine(input) {
  const start = input.selectionStart;
  const end = input.selectionEnd;
  const value = input.value;
  const before = value.slice(0, start);
  const lineStart = before.lastIndexOf("\n") + 1;
  const currentLine = before.slice(lineStart);
  const indentMatch = currentLine.match(/^(\s*)(?:[•\-]\s?)?/);
  const baseIndent = (indentMatch && indentMatch[1]) || "";
  const insert = `\n${baseIndent}  • `;
  input.value = value.slice(0, start) + insert + value.slice(end);
  const pos = start + insert.length;
  input.selectionStart = input.selectionEnd = pos;
}

/* Wires a chat-style composer textarea + send button: Enter posts,
   Shift+Enter adds a plain newline, Ctrl/Cmd+Enter adds an indented bullet
   sub-line (see insertComposerIndentLine). `onPost(text)` is called with
   the trimmed text; the textarea is cleared/resized after. Shared by every
   message-style tool so they all behave identically. */
function wireChatComposer(input, sendBtn, onPost) {
  function autoSize() {
    input.style.height = "auto";
    input.style.height = Math.min(120, input.scrollHeight) + "px";
  }
  function submit() {
    const text = input.value.trim();
    if (!text) return;
    onPost(text);
    input.value = "";
    autoSize();
  }
  if (sendBtn) sendBtn.addEventListener("click", submit);
  input.addEventListener("input", autoSize);
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      insertComposerIndentLine(input);
      autoSize();
      return;
    }
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      submit();
    }
  });
  autoSize();
}

/* Chat-style popup for task/subtask status updates — replaces the old
   inline expand-in-place notes row plus a plain-textarea prompt for adding
   one. Messages render oldest-to-newest like a chat thread, with an input
   pinned to the bottom; Enter posts, Shift+Enter for a newline, Ctrl+Enter
   for an indented bullet sub-line — matching how every chat/outline app
   already works so nobody has to think about it. Own messages get inline
   Edit/Delete; `task.notes` stays newest-first in storage (unchanged —
   other call sites read notes[0] as "latest"), this just reverses it for
   display. */
function currentUserNoteIdentity() {
  const user = (window.AEWTTR && window.AEWTTR.db && window.AEWTTR.db.user) || {};
  return {
    name: user.name || "Unknown",
    id: user.id || "",
    email: String(user.email || "").trim()
  };
}
function isNoteAuthor(note) {
  if (!note) return false;
  const me = currentUserNoteIdentity();
  if (note.authorId && me.id && String(note.authorId) === String(me.id)) return true;
  if (note.authorEmail && me.email && String(note.authorEmail).toLowerCase() === me.email.toLowerCase()) return true;
  // Older notes predate authorId/authorEmail and only have a plain name —
  // match loosely (case/whitespace) so a display-name rename, casing
  // change, or stray space doesn't permanently lock the real author out.
  return !!(me.name && note.author && String(note.author).trim().toLowerCase() === me.name.trim().toLowerCase());
}
/* Notes an old, unmatchable author string can leave permanently stuck even
   after the loose match above — Admin can always delete as a backstop. */
function canDeleteNote(note) {
  if (isNoteAuthor(note)) return true;
  const user = window.AEWTTR && window.AEWTTR.db && window.AEWTTR.db.user;
  return !!(user && user.isAdmin);
}
function stampNoteAuthor(note) {
  const me = currentUserNoteIdentity();
  note.author = me.name;
  if (me.id) note.authorId = me.id;
  if (me.email) note.authorEmail = me.email;
  return note;
}
function touchNoteTimestamp(note) {
  const now = new Date();
  note.date = now.toISOString().slice(0, 10);
  note.time = now.toTimeString().slice(0, 5);
  note.editedAt = now.toISOString();
  return note;
}
function formatNoteTimestamp(note) {
  const base = `${note.date || ""}${note.time ? ` · ${note.time}` : ""}`.trim();
  return note.editedAt ? `${base}${base ? " · " : ""}edited` : base;
}

function openTaskNotesModal(task, taskLabel, onNotesChange) {
  const modal = openModal(`
    <div class="aewttr-modal-head">
      <h3>${escapeHtml(taskLabel || "Notes")}</h3>
      <button class="aewttr-modal-close" type="button" aria-label="Close">&times;</button>
    </div>
    <div class="task-notes-chat-body" id="task-notes-chat-body"></div>
    <div class="task-notes-input-row">
      <textarea class="task-notes-input" id="task-notes-input" placeholder="Type an update — Enter to post, Ctrl+Enter for a sub-line…" rows="1"></textarea>
      <button type="button" class="btn-aewttr btn-aewttr-sm task-notes-send" id="task-notes-send"${tip("Post (Enter)")}><i class="bx bx-send"></i></button>
    </div>
  `, { className: "task-notes-modal" });

  const chatBody = $("#task-notes-chat-body", modal);
  const input = $("#task-notes-input", modal);
  let editingId = null;

  function notify() {
    if (typeof onNotesChange === "function") onNotesChange(task.notes);
  }

  // Editing/cancelling/deleting an older message shouldn't yank the view —
  // only a brand-new post (or the initial open) should jump to the newest.
  function renderMessages(opts) {
    const stickToBottom = !!(opts && opts.scrollToBottom);
    const scrollTop = chatBody.scrollTop;
    const notes = (task.notes || []).slice().reverse();
    chatBody.innerHTML = notes.length
      ? notes.map((n) => {
          const isMine = isNoteAuthor(n);
          const canDelete = canDeleteNote(n);
          const isEditing = editingId === n.id;
          return `
          <div class="task-notes-bubble-row ${isMine ? "mine" : ""}">
            <div class="task-notes-bubble">
              <div class="task-notes-bubble-meta">
                <strong>${escapeHtml(n.author || "Unknown")}</strong><span>${escapeHtml(formatNoteTimestamp(n))}</span>
                ${!isEditing && (isMine || canDelete) ? `
                  <span class="task-notes-bubble-actions">
                    ${isMine ? `<button type="button" data-edit-note="${n.id}" aria-label="Edit"${tip("Edit")}><i class="bx bx-pencil"></i></button>` : ""}
                    ${canDelete ? `<button type="button" data-delete-note="${n.id}" aria-label="Delete"${tip("Delete")}><i class="bx bx-trash"></i></button>` : ""}
                  </span>` : ""}
              </div>
              ${isEditing
                ? `<textarea class="task-notes-edit-input" id="task-notes-edit-${n.id}">${escapeHtml(n.text)}</textarea>
                   <div class="task-notes-edit-actions">
                     <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-cancel-edit="${n.id}">Cancel</button>
                     <button type="button" class="btn-aewttr btn-aewttr-sm" data-save-edit="${n.id}">Save</button>
                   </div>`
                : `<div class="task-notes-bubble-text">${escapeHtml(n.text)}</div>`}
            </div>
          </div>
        `;
        }).join("")
      : `<div class="task-notes-empty">No updates yet. Say something below to start the thread.</div>`;
    chatBody.scrollTop = stickToBottom ? chatBody.scrollHeight : scrollTop;
    wireMessageActions();
  }

  function wireMessageActions() {
    $all("[data-edit-note]", chatBody).forEach((btn) => btn.addEventListener("click", () => {
      const note = (task.notes || []).find((n) => n.id === btn.dataset.editNote);
      if (!isNoteAuthor(note)) { toast("You can only edit your own notes.", "error"); return; }
      editingId = btn.dataset.editNote;
      renderMessages();
      const editInput = $(`#task-notes-edit-${editingId}`, chatBody);
      if (editInput) { editInput.focus(); editInput.setSelectionRange(editInput.value.length, editInput.value.length); }
    }));
    $all("[data-cancel-edit]", chatBody).forEach((btn) => btn.addEventListener("click", () => {
      editingId = null;
      renderMessages();
    }));
    $all("[data-save-edit]", chatBody).forEach((btn) => btn.addEventListener("click", () => {
      const id = btn.dataset.saveEdit;
      const editInput = $(`#task-notes-edit-${id}`, chatBody);
      const text = editInput ? editInput.value.trim() : "";
      if (!text) { toast("Note can't be empty.", "error"); return; }
      const note = (task.notes || []).find((n) => n.id === id);
      if (!note || !isNoteAuthor(note)) {
        toast("You can only edit your own notes.", "error");
        editingId = null;
        renderMessages();
        return;
      }
      note.text = text;
      touchNoteTimestamp(note);
      editingId = null;
      renderMessages();
      notify();
    }));
    $all("[data-delete-note]", chatBody).forEach((btn) => btn.addEventListener("click", async () => {
      const id = btn.dataset.deleteNote;
      const note = (task.notes || []).find((n) => n.id === id);
      if (!canDeleteNote(note)) { toast("You can only delete your own notes.", "error"); return; }
      const ok = await confirmDialog({ title: "Delete note", message: "Delete this note? This cannot be undone.", confirmLabel: "Delete", danger: true });
      if (!ok) return;
      task.notes = (task.notes || []).filter((n) => n.id !== id);
      renderMessages();
      notify();
    }));
  }

  $(".aewttr-modal-close", modal).addEventListener("click", () => closeModal());
  wireChatComposer(input, $("#task-notes-send", modal), (text) => {
    task.notes = task.notes || [];
    const note = stampNoteAuthor({
      id: uid("nt"),
      text
    });
    touchNoteTimestamp(note);
    delete note.editedAt;
    task.notes.unshift(note);
    renderMessages({ scrollToBottom: true });
    notify();
  });

  renderMessages({ scrollToBottom: true });
  input.focus();
}

function promptTextareaDialog(opts) {
  opts = Object.assign({ title: "Enter a value", label: "", placeholder: "", value: "", confirmLabel: "Save", cancelLabel: "Cancel" }, opts || {});
  return new Promise((resolve) => {
    let settled = false;
    // See confirmDialog's comment above — same own-reference fix.
    const finish = (result) => {
      if (settled) return;
      settled = true;
      document.removeEventListener("keydown", onKey);
      backdrop.remove();
      resolve(result);
    };
    const onKey = (e) => { if (e.key === "Escape") finish(null); };
    const backdrop = el(`<div class="aewttr-modal-backdrop"></div>`);
    const modal = el(`
      <div class="aewttr-modal aewttr-confirm-modal" style="max-width:500px;">
        <div class="aewttr-modal-head">
          <h3>${escapeHtml(opts.title)}</h3>
          <button class="aewttr-modal-close" type="button" aria-label="Close">&times;</button>
        </div>
        <div class="aewttr-modal-body">
          ${opts.label ? `<div class="form-row" style="margin:0;"><label>${escapeHtml(opts.label)}</label><textarea class="textarea-aewttr" id="aewttr-prompt-input" placeholder="${escapeHtml(opts.placeholder)}" style="min-height:100px;">${escapeHtml(opts.value)}</textarea></div>` : ""}
        </div>
        <div class="aewttr-modal-foot">
          <button class="btn-aewttr-ghost" type="button" id="aewttr-prompt-cancel">${escapeHtml(opts.cancelLabel)}</button>
          <button class="btn-aewttr" type="button" id="aewttr-prompt-ok">${escapeHtml(opts.confirmLabel)}</button>
        </div>
      </div>
    `);
    backdrop.appendChild(modal);
    document.body.appendChild(backdrop);
    backdrop.addEventListener("mousedown", (e) => { if (e.target === backdrop) finish(null); });
    document.addEventListener("keydown", onKey);
    const input = $("#aewttr-prompt-input", modal);
    const submit = () => finish(input ? input.value.trim() : "");
    $(".aewttr-modal-close", modal).addEventListener("click", () => finish(null));
    $("#aewttr-prompt-cancel", modal).addEventListener("click", () => finish(null));
    $("#aewttr-prompt-ok", modal).addEventListener("click", submit);
    if (input) {
      input.focus();
    }
  });
}

/* ---------- nav / role gating ---------- */
const NAV_ITEMS = [
  { route: "dashboard", icon: "bx-grid-alt", label: "Dashboard" },
  { route: "overview", icon: "bx-bar-chart-alt-2", label: "Overview" },
  { route: "projects", icon: "bx-folder-open", label: "Projects" },
  { route: "weekly", icon: "bx-conversation", label: "Weekly Meeting" },
  { route: "travel", icon: "bx-world", label: "Travel" },
  { route: "docreview", icon: "bx-file-blank", label: "Document Review" },
  { route: "admin", icon: "bx-shield-quarter", label: "Admin" }
];

/* Routing is hash-only (location.hash), on purpose — no history.pushState.
   This app is hosted by dropping its HTML/JS straight onto a real
   SharePoint .aspx page (a Firepit web part) or into a document library
   file served under that same page. pushState changes location.pathname/
   search, i.e. what page SharePoint itself thinks it's on — and in
   practice that got silently reverted/ignored by SharePoint's own page
   framework (or worse, threw a SecurityError in a sandboxed embed),
   making every click that called navigate() a no-op: "none of the
   buttons or links work." location.hash never changes what resource is
   being requested — it's a same-document, same-URL operation in every
   browser and every hosting context, sandboxed or not, .aspx or not — so
   it's the one navigation mechanism this app can rely on unconditionally.
   A `?page=...`-style link is still accepted for backward compatibility
   with anything already shared that way, but navigate() itself always
   writes the hash. */
function parseLegacyHashRoute() {
  const hash = location.hash.replace(/^#\/?/, "");
  const [pathPart, queryPart] = hash.split("?");
  const query = {};
  if (queryPart) {
    queryPart.split("&").forEach((pair) => {
      const [key, value] = pair.split("=");
      if (!key) return;
      query[decodeURIComponent(key)] = decodeURIComponent(value || "");
    });
  }
  return { path: pathPart || "", query };
}

function currentRoute() {
  const legacy = parseLegacyHashRoute();
  let pathPart = legacy.path;
  const query = Object.assign({}, legacy.query);
  if (!pathPart) {
    // Back-compat read path only — honors an old ?page=... link if one is
    // present and there's no hash yet. navigate() never writes this form.
    const url = new URL(location.href);
    const routePath = String(url.searchParams.get("page") || "").replace(/^#?\/?/, "");
    if (routePath) {
      pathPart = routePath;
      url.searchParams.forEach((value, key) => { if (key !== "page") query[key] = value; });
    }
  }
  const parts = String(pathPart || "").split("/").filter(Boolean);
  return { app: parts[0] || "dashboard", parts: parts.slice(1), query };
}

function navigate(path, query) {
  const route = String(path || "").replace(/^#?\/?/, "");
  const queryString = query && typeof query === "object"
    ? Object.keys(query)
      .filter((key) => query[key] != null && query[key] !== "")
      .map((key) => `${encodeURIComponent(key)}=${encodeURIComponent(String(query[key]))}`)
      .join("&")
    : "";
  const nextHash = `#/${route}${queryString ? `?${queryString}` : ""}`;
  if (location.hash === nextHash) { renderPage(); return; }
  location.hash = nextHash; // triggers the hashchange listener -> renderPage()
}
function consumePendingRouteAction() {
  const action = window.AEWTTR.pendingRouteAction || null;
  window.AEWTTR.pendingRouteAction = null;
  return action;
}
function queueRouteAction(action) {
  window.AEWTTR.pendingRouteAction = action;
}
function consumeRouteIntent(scopeKey) {
  const pending = consumePendingRouteAction();
  if (pending && typeof pending === "object") return pending;
  const query = currentRoute().query || {};
  if (!Object.keys(query).length) return null;
  if (!window.AEWTTR.routeIntentSeen) window.AEWTTR.routeIntentSeen = {};
  const seenKey = `${scopeKey || "route"}|${location.pathname}${location.search}${location.hash}`;
  if (window.AEWTTR.routeIntentSeen[seenKey]) return null;
  window.AEWTTR.routeIntentSeen[seenKey] = true;
  return query;
}
function docReviewsForCurrentUser() {
  const db = window.AEWTTR.db || {};
  const me = currentUserIdentity ? currentUserIdentity() : { name: db.user && db.user.name, email: db.user && db.user.email, isAdmin: !!(db.user && db.user.isAdmin) };
  const rows = [];
  Object.keys(db.docs || {}).forEach((column) => {
    (db.docs[column] || []).forEach((doc) => {
      if (doc.isArchived) return;
      const pending = (doc.reviewers || []).some((reviewer) => {
        if (reviewer.decision && reviewer.decision !== "Pending") return false;
        return typeof samePersonByNameOrEmail === "function" && samePersonByNameOrEmail(reviewer, me);
      });
      if (pending) rows.push({ doc, column });
    });
  });
  return rows.sort((a, b) => String(b.doc.deadline || b.doc.submittedAt || "").localeCompare(String(a.doc.deadline || a.doc.submittedAt || "")));
}
window.docReviewsForCurrentUser = docReviewsForCurrentUser;

/* ---- In-app action-item digest ---- */

function digestLsKey() {
  const db = window.AEWTTR && window.AEWTTR.db;
  const email = (db && db.user && db.user.email) || "anon";
  return `pulse-digest-last-${email}`;
}

function digestShouldShow(prefs) {
  if (!prefs || !prefs.digest || !prefs.digest.enabled) return false;
  const raw = lsGet(digestLsKey());
  if (!raw) return true;
  const last = new Date(raw);
  if (isNaN(last)) return true;
  const msNow = Date.now();
  const elapsed = msNow - last.getTime();
  if (prefs.digest.frequency === "weekly") return elapsed > 7 * 24 * 60 * 60 * 1000;
  const lastDay = last.toDateString();
  const nowDay = new Date(msNow).toDateString();
  return lastDay !== nowDay;
}

function buildDigestItems() {
  const db = window.AEWTTR && window.AEWTTR.db;
  if (!db) return {};
  const me = currentUserIdentity ? currentUserIdentity() : { name: db.user && db.user.name, email: db.user && db.user.email };
  const myName = (me && me.name) || "";

  const tasks = [];
  Object.keys(db.ganttTasks || {}).forEach((pid) => {
    (db.ganttTasks[pid] || []).forEach((task) => {
      if (task.done || task.status === "Complete") return;
      const assignee = (task.assignee || "").toLowerCase();
      if (assignee && myName && assignee.includes(myName.split(" ")[0].toLowerCase())) tasks.push({ task, pid });
    });
  });

  const docs = typeof docReviewsForCurrentUser === "function" ? docReviewsForCurrentUser() : [];

  const travel = [];
  (db.travelRequests || []).forEach((r) => {
    if (typeof isCurrentUserTravelRequester === "function" && isCurrentUserTravelRequester(r)) {
      if (r.status === "Pending" || r.status === "Pending Finance" || r.status === "Approved") travel.push(r);
    }
  });

  const toApprove = [];
  if (typeof canApproveTravelRequests === "function" && canApproveTravelRequests()) {
    (db.travelRequests || []).forEach((r) => {
      if (r.status === "Pending" || r.status === "Pending Finance") toApprove.push(r);
    });
  }

  return { tasks, docs, travel, toApprove };
}

function showDigestModal(opts) {
  opts = opts || {};
  const db = window.AEWTTR && window.AEWTTR.db;
  if (!db) return;
  const prefs = typeof normalizeNotificationPrefs === "function"
    ? normalizeNotificationPrefs(db.user && db.user.notificationPrefs)
    : null;
  if (!opts.force && !digestShouldShow(prefs)) return;

  lsSet(digestLsKey(), new Date().toISOString());

  const { tasks, docs, travel, toApprove } = buildDigestItems();
  const total = tasks.length + docs.length + travel.length + toApprove.length;
  const freq = prefs && prefs.digest && prefs.digest.frequency === "weekly" ? "Weekly" : "Daily";

  function sectionHtml(icon, title, items, emptyMsg) {
    if (!items.length) return `
      <div class="digest-section">
        <div class="digest-section-head"><i class="bx ${icon}"></i> ${title}</div>
        <p class="digest-empty">${emptyMsg}</p>
      </div>`;
    return `
      <div class="digest-section">
        <div class="digest-section-head"><i class="bx ${icon}"></i> ${title} <span class="digest-count">${items.length}</span></div>
        <ul class="digest-list">${items}</ul>
      </div>`;
  }

  const taskItems = tasks.slice(0, 10).map(({ task }) =>
    `<li class="digest-item"><button class="digest-link" data-digest-nav="projects">${escapeHtml(task.title || task.text || "Untitled task")}</button></li>`
  ).join("");
  const docItems = docs.slice(0, 10).map(({ doc }) =>
    `<li class="digest-item"><button class="digest-link" data-digest-nav="docreview">${escapeHtml(doc.title || "Untitled document")}</button></li>`
  ).join("");
  const travelItems = travel.slice(0, 5).map((r) =>
    `<li class="digest-item"><button class="digest-link" data-digest-nav="travel/list">${escapeHtml(r.tripTitle || r.id || "Travel request")} <span class="digest-status">${escapeHtml(r.status)}</span></button></li>`
  ).join("");
  const approveItems = toApprove.slice(0, 5).map((r) =>
    `<li class="digest-item"><button class="digest-link" data-digest-nav="travel/all">${escapeHtml(r.requester || r.id)} — ${escapeHtml(r.destination || r.requestType || "Travel")} <span class="digest-status">${escapeHtml(r.status)}</span></button></li>`
  ).join("");

  const html = `
    <div class="digest-modal-body">
      <div class="digest-header">
        <i class="bx bx-calendar-check digest-header-icon"></i>
        <div>
          <div class="digest-header-title">${freq} Action Digest</div>
          <div class="digest-header-sub">${total === 0 ? "You're all caught up — nothing needs your attention right now." : `${total} item${total === 1 ? "" : "s"} need${total === 1 ? "s" : ""} your attention.`}</div>
        </div>
      </div>
      ${sectionHtml("bx-task", "Your Tasks", taskItems, "No open tasks assigned to you.")}
      ${sectionHtml("bx-file-blank", "Documents to Review", docItems, "No documents pending your review.")}
      ${sectionHtml("bx-trip", "Your Travel", travelItems, "No active travel requests.")}
      ${toApprove.length > 0 ? sectionHtml("bx-check-shield", "Travel to Approve", approveItems, "") : ""}
      <div class="digest-footer">
        <button class="btn-aewttr-outline btn-sm" id="digest-close-btn">Dismiss</button>
        <button class="btn-aewttr btn-sm" data-digest-nav="notification-settings">Digest Settings</button>
      </div>
    </div>`;

  if (typeof openModal === "function") {
    const modal = openModal(html);
    const digestEl = modal && modal.querySelector(".digest-modal-body");
    if (digestEl) {
      digestEl.querySelectorAll("[data-digest-nav]").forEach((btn) => {
        btn.addEventListener("click", () => {
          if (typeof navigate === "function") navigate(btn.dataset.digestNav);
          if (typeof closeModal === "function") closeModal();
        });
      });
      const closeBtn = digestEl.querySelector("#digest-close-btn");
      if (closeBtn) closeBtn.addEventListener("click", () => { if (typeof closeModal === "function") closeModal(); });
    }
  }
}

window.showDigestModal = showDigestModal;

function currentAppRole() {
  const diagnosticsRole = window.AEWTTR && window.AEWTTR.spDiagnostics && window.AEWTTR.spDiagnostics.currentRole;
  return (diagnosticsRole && diagnosticsRole.role)
    || (window.AEWTTR.db && window.AEWTTR.db.user && window.AEWTTR.db.user.role)
    || "Member";
}

function isViewerRole() {
  const role = currentAppRole();
  return role === "Viewer" || role === "Guest";
}

function isMemberRole() {
  const role = currentAppRole();
  return role === "Member";
}

function canCurrentUserEdit() {
  const role = currentAppRole();
  if (role === "Member" || role === "Viewer" || role === "Guest") return false;
  return true;
}

function canCreateProject() {
  const user = window.AEWTTR.db && window.AEWTTR.db.user;
  if (!user) return false;
  if (user.isAdmin) return true;
  const role = currentAppRole();
  if (role === "Member" || role === "Viewer" || role === "Guest") return false;
  return true;
}

function canCurrentUserAccessAdmin() {
  const diagnosticsRole = window.AEWTTR && window.AEWTTR.spDiagnostics && window.AEWTTR.spDiagnostics.currentRole;
  if (diagnosticsRole && diagnosticsRole.isAdmin) return true;
  return !!(window.AEWTTR.db && window.AEWTTR.db.user && window.AEWTTR.db.user.isAdmin);
}

function canManageMeetings() {
  const diagnosticsRole = window.AEWTTR && window.AEWTTR.spDiagnostics && window.AEWTTR.spDiagnostics.currentRole;
  if (diagnosticsRole && (diagnosticsRole.isAdmin || diagnosticsRole.isMeetingAdmin)) return true;
  const user = window.AEWTTR.db && window.AEWTTR.db.user;
  if (!user) return false;
  if (user.isAdmin || user.isMeetingAdmin) return true;
  const cfg = window.AEWTTR.db && window.AEWTTR.db.locationConfig;
  const admins = (cfg && cfg.meetingAdminEmails) || [];
  const email = String(user.email || user.loginName || "").trim().toLowerCase();
  const name = String(user.name || user.displayName || "").trim().toLowerCase();
  return admins.some((a) => {
    const al = String(a || "").trim().toLowerCase();
    return al && (al === email || al === name);
  });
}

function canEditProject(proj) {
  const user = window.AEWTTR.db && window.AEWTTR.db.user;
  if (!user) return false;
  const role = currentAppRole();
  if (user.isAdmin || role === "PM Admin") return true;
  if (role === "Member" || role === "Viewer" || role === "Guest") return false;
  const db = window.AEWTTR.db;
  const roster = (db.projectPeople && db.projectPeople[proj && proj.id]) || [];
  const email = String(user.email || user.loginName || "").trim().toLowerCase();
  const name = String(user.name || user.displayName || "").trim().toLowerCase();
  return roster.some((p) => {
    if (!p.isProjectAdmin) return false;
    const pe = String(p.email || "").trim().toLowerCase();
    const pn = String(p.label || "").trim().toLowerCase();
    return (email && pe && email === pe) || (name && pn && name === pn);
  });
}

function canSubmitForms() {
  const role = currentAppRole();
  return role !== "Viewer" && role !== "Guest";
}

function canUseDocReview() {
  const role = currentAppRole();
  return role !== "Viewer" && role !== "Guest";
}

function canApproveTravelRequests() {
  const diagnosticsRole = window.AEWTTR && window.AEWTTR.spDiagnostics && window.AEWTTR.spDiagnostics.currentRole;
  if (diagnosticsRole && diagnosticsRole.isAdmin) return true;
  return !!(window.AEWTTR.db && window.AEWTTR.db.user && window.AEWTTR.db.user.isAdmin);
}

function canAssignTravelCo() {
  const diagnosticsRole = window.AEWTTR && window.AEWTTR.spDiagnostics && window.AEWTTR.spDiagnostics.currentRole;
  if (diagnosticsRole) return !!diagnosticsRole.isFinanceAdmin;
  const user = window.AEWTTR.db && window.AEWTTR.db.user;
  return !!(user && user.isFinanceAdmin);
}

function canRecordConcurrence() {
  return canApproveTravelRequests();
}

function canAccessTravelApprovals() {
  return canApproveTravelRequests() || canAssignTravelCo();
}

function requireAdmin(renderFn) {
  if (canCurrentUserAccessAdmin()) return renderFn();
  return `
    <div class="lock-block">
      <i class="bx bx-lock-alt"></i>
      <h3>Admin access required</h3>
      <p>SharePoint permissions control actual access. App roles control in-app tools and views.</p>
    </div>`;
}

/* ---------- shared searchable people picker (search, click, add — no ctrl-click) ----------
   Reused anywhere a multi-select of site users is needed (Document Review
   reviewers, etc). Travel's traveler picker predates this and keeps its own
   copy since it's already working; this is the shared one for everything
   else. `people` is mutated in place (push/splice) so the caller's array
   stays in sync. */
function normalizePersonKey(value) { return String(value == null ? "" : value).trim().toLowerCase(); }

function samePersonByNameOrEmail(a, b) {
  const nameA = String(a && a.name || "").trim().toLowerCase();
  const emailA = String(a && a.email || "").trim().toLowerCase();
  const nameB = String(b && b.name || "").trim().toLowerCase();
  const emailB = String(b && b.email || "").trim().toLowerCase();
  if (emailA && emailB) return emailA === emailB;
  return !!nameA && !!nameB && nameA === nameB;
}

/* Hide machine / service principals from directory UI + pickers without
   deleting App Role rows (unassigning those can wipe real privileges). */
function isHiddenServicePrincipal(person) {
  if (!person) return false;
  const name = String(person.displayName || person.name || person.Title || person.UserDisplayName || "").trim();
  const email = String(person.email || person.UserEmail || "").trim();
  const login = String(person.loginName || person.LoginName || "").trim();
  const hay = `${name} ${email} ${login}`.toLowerCase();
  if (!hay.trim()) return false;

  if (/\bnt\s*service\b/.test(hay) || /\bnt\s*authority\b/.test(hay)) return true;
  if (/\bservice\s*account\b/.test(hay)) return true;
  if (/\bnative\\/.test(hay) || /\bnative\//.test(hay)) return true;
  // Machine / computer accounts (SAM ends with $) and claim forms with $.
  if (/\w\$(\s|$|@)/.test(hay) || /\/\$[^a-z]|\\[^\\\s]+\$/.test(hay)) return true;
  if (name.endsWith("$") || email.split("@")[0].endsWith("$")) return true;
  // Common SharePoint / app-only / crawl principals.
  if (/spo[_-]?crawler|sharepoint\\system|app@sharepoint|c:0[!#.a]|i:0i\.t|i:0#\.f\|membership\|app@/.test(hay)) return true;
  if (/\biis\s*apppool\b|\bapp\s*pool\b/.test(hay)) return true;
  // Leading svc- / svc_ account names only (not john.svc@…).
  const compact = hay.replace(/\s+/g, "");
  if (/^(?:[^@]*[\\/|])?svc[-_.]/.test(compact) || /^svc[-_.]/.test(email.toLowerCase())) return true;
  return false;
}

function isPeoplePickerJunkValue(value) {
  const text = String(value == null ? "" : value).trim();
  if (!text) return true;
  if (/dispform\.aspx/i.test(text) || /\.aspx(\?|$)/i.test(text)) return true;
  if (text.startsWith("/") || text.startsWith("http://") || text.startsWith("https://")) return true;
  if (/^\d{1,4}[\/\-.]\d{1,2}[\/\-.]\d{1,4}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$/.test(text)) return true;
  if (!/[a-zA-Z]/.test(text)) return true;
  return false;
}

function sanitizeDirectoryEntry(entry) {
  if (!entry) return null;
  const name = String(entry.name || entry.displayName || "").trim();
  const email = String(entry.email || "").trim().toLowerCase();
  const id = String(entry.id || email || name).trim();
  if (!id || !name || isPeoplePickerJunkValue(name)) return null;
  if (!email && isPeoplePickerJunkValue(id)) return null;
  return { id, name, email };
}

function dedupeDirectoryEntries(entries) {
  const byKey = new Map();
  (entries || []).forEach((entry) => {
    const safe = sanitizeDirectoryEntry(entry);
    if (!safe) return;
    const key = safe.email || normalizePersonKey(safe.name);
    if (!key) return;
    if (!byKey.has(key)) {
      byKey.set(key, safe);
      return;
    }
    const existing = byKey.get(key);
    if (!existing.email && safe.email) byKey.set(key, safe);
  });
  return Array.from(byKey.values()).sort((a, b) => a.name.localeCompare(b.name));
}

/* A picker searchFn that searches the WHOLE Microsoft 365 tenant (via
   sharePointAdapter.searchTenantPeopleDirectory — Graph first, then the
   SharePoint People Picker/Search API as fallbacks), not just people with
   access to the current site. `loadLocalDirectory` (optional) is a
   zero-arg async function returning a smaller, always-available directory
   (e.g. current site users) — fetched once and cached, then merged in as
   an instant supplement so a live-search hiccup doesn't mean an empty box. */
function makeTenantWideSearchFn(loadLocalDirectory) {
  let localDirectoryPromise = null;
  return async function tenantWideSearchFn(query) {
    if (!localDirectoryPromise && typeof loadLocalDirectory === "function") {
      localDirectoryPromise = loadLocalDirectory().catch(() => []);
    }
    const q = normalizePersonKey(query);
    const [liveResults, localDirectory] = await Promise.all([
      (window.AEWTTR.mode === "sharepoint" && window.AEWTTR.siteUrl && typeof sharePointAdapter.searchTenantPeopleDirectory === "function")
        ? sharePointAdapter.searchTenantPeopleDirectory(window.AEWTTR.siteUrl, query, { maxResults: 12 }).catch((e) => {
            console.error("[tenantWideSearchFn] live tenant search failed:", e);
            return [];
          })
        : Promise.resolve([]),
      localDirectoryPromise || Promise.resolve([])
    ]);
    const localMatches = (localDirectory || []).filter((entry) =>
      normalizePersonKey(entry.name).includes(q) || normalizePersonKey(entry.email).includes(q));
    return dedupeDirectoryEntries([...liveResults, ...localMatches]);
  };
}

function getMemberDirectory() {
  // Person pickers use AEWTTR-associated people only (active App Roles / members).
  const fromMembers = dedupeDirectoryEntries((window.AEWTTR.db.members || []).map((m) => ({ id: m.id, name: m.name, email: m.email || "" })))
    .filter((entry) => !isHiddenServicePrincipal(entry));
  const fromSite = dedupeDirectoryEntries(window.AEWTTR.siteMemberDirectory || [])
    .filter((entry) => !isHiddenServicePrincipal(entry));
  if (fromSite.length) return fromSite;
  return fromMembers;
}

function getSiteMemberDirectory() {
  const directory = dedupeDirectoryEntries(window.AEWTTR.siteMemberDirectory || [])
    .filter((entry) => !isHiddenServicePrincipal(entry));
  if (directory.length) return directory;
  return getMemberDirectory();
}

/* ---------- global people groups (Admin-managed + project auto-sync) ----------
   Persisted in the SharePoint list formerly called "PULSE Doc Reviewer Groups"
   (same list, broader use). In-memory canonical store is db.groups; db.
   docReviewerGroups is kept as an alias for older call sites. */

function loadPulseGroups() {
  const db = window.AEWTTR.db;
  if (!Array.isArray(db.groups)) {
    db.groups = Array.isArray(db.docReviewerGroups) ? db.docReviewerGroups : [];
  }
  db.docReviewerGroups = db.groups;
  return db.groups;
}

function isPickerGroupEntry(entry) {
  return !!(entry && (entry.type === "group" || entry.isGroup) && (entry.groupId || entry.id));
}

function normalizeGroupMembers(members) {
  return (members || []).filter((m) => m && m.name).map((m) => ({
    id: m.id || m.memberId || "",
    name: String(m.name || "").trim(),
    email: String(m.email || "").trim()
  }));
}

function groupToPickerEntry(group) {
  if (!group || !group.id) return null;
  const members = normalizeGroupMembers(group.members);
  return {
    type: "group",
    isGroup: true,
    groupId: group.id,
    id: group.id,
    name: group.name || "Group",
    email: "",
    members,
    memberCount: members.length,
    source: group.source || "manual",
    projectId: group.projectId || ""
  };
}

function matchPulseGroups(query, limit) {
  const q = normalizePersonKey(query);
  if (!q) return [];
  return loadPulseGroups()
    .filter((g) => g && g.name && normalizePersonKey(g.name).includes(q))
    .slice(0, limit == null ? 6 : limit)
    .map(groupToPickerEntry)
    .filter(Boolean);
}

function expandPickerPeople(entries) {
  const out = [];
  (entries || []).forEach((entry) => {
    if (!entry) return;
    if (isPickerGroupEntry(entry)) {
      const group = loadPulseGroups().find((g) => g.id === (entry.groupId || entry.id));
      const members = normalizeGroupMembers((group && group.members) || entry.members || []);
      members.forEach((m) => {
        if (!m.name) return;
        if (out.some((p) => samePersonByNameOrEmail(p, m))) return;
        out.push({
          id: m.id || `gmem-${normalizePersonKey(m.email || m.name)}`,
          name: m.name,
          email: m.email || ""
        });
      });
      return;
    }
    if (!entry.name) return;
    if (out.some((p) => samePersonByNameOrEmail(p, entry))) return;
    out.push({
      id: entry.id || `p-${normalizePersonKey(entry.email || entry.name)}`,
      name: entry.name,
      email: entry.email || ""
    });
  });
  return out;
}

function projectPeopleAsGroupMembers(proj) {
  const db = window.AEWTTR.db;
  if (!proj || !proj.id) return [];
  const roster = (db.projectPeople && db.projectPeople[proj.id]) || [];
  const members = [];
  roster.forEach((person) => {
    if (!person) return;
    if (person.type === "company") return;
    const known = person.memberId ? (db.members || []).find((m) => m.id === person.memberId) : null;
    const name = (known && known.name) || person.label || "";
    const email = (known && known.email) || person.email || "";
    if (!name) return;
    if (members.some((m) => samePersonByNameOrEmail(m, { name, email }))) return;
    members.push({
      id: person.memberId || person.id || "",
      name,
      email
    });
  });
  return members;
}

async function upsertPulseGroup(group) {
  if (isSharePointMode() && typeof ensureDocReviewerGroupsList === "function") {
    try { await ensureDocReviewerGroupsList(currentSiteUrl()); } catch (e) {}
  }
  const groups = loadPulseGroups();
  const payload = {
    id: group.id || uid("GRP"),
    name: String(group.name || "").trim(),
    members: normalizeGroupMembers(group.members),
    source: group.source === "project" ? "project" : "manual",
    projectId: group.projectId || "",
    createdBy: group.createdBy || "",
    createdByEmail: group.createdByEmail || "",
    createdDate: group.createdDate || (typeof todayIsoDate === "function" ? todayIsoDate() : new Date().toISOString().slice(0, 10)),
    _spId: group._spId
  };
  const index = groups.findIndex((g) => g.id === payload.id
    || (payload.projectId && g.projectId === payload.projectId && g.source === "project"));
  if (index >= 0) {
    const prev = groups[index];
    Object.assign(prev, payload, { id: prev.id, _spId: prev._spId || payload._spId });
    await Repo.save("docReviewerGroup", prev);
    return prev;
  }
  groups.push(payload);
  await Repo.save("docReviewerGroup", payload);
  return payload;
}

async function deletePulseGroup(id) {
  const groups = loadPulseGroups();
  const index = groups.findIndex((g) => g.id === id);
  if (index < 0) return;
  const [group] = groups.splice(index, 1);
  await Repo.remove("docReviewerGroup", group);
}

/* Idempotent upsert: project roster → a group named after the project. */
async function syncProjectPulseGroup(proj) {
  if (!proj || !proj.id || !proj.name) return null;
  const groups = loadPulseGroups();
  const existing = groups.find((g) => g.source === "project" && g.projectId === proj.id)
    || groups.find((g) => g.source === "project" && normalizePersonKey(g.name) === normalizePersonKey(proj.name));
  const members = projectPeopleAsGroupMembers(proj);
  return upsertPulseGroup({
    id: (existing && existing.id) || uid("GRP"),
    name: proj.name,
    members,
    source: "project",
    projectId: proj.id,
    createdBy: (existing && existing.createdBy) || "",
    createdByEmail: (existing && existing.createdByEmail) || "",
    createdDate: (existing && existing.createdDate) || (typeof todayIsoDate === "function" ? todayIsoDate() : new Date().toISOString().slice(0, 10)),
    _spId: existing && existing._spId
  });
}

function renderPersonChips(mount, people, onRemove, opts) {
  opts = opts || {};
  const emptyLabel = opts.emptyLabel || "None selected.";
  mount.innerHTML = people.length ? people.map((p, i) => {
    if (isPickerGroupEntry(p)) {
      const count = Array.isArray(p.members) ? p.members.length : (p.memberCount || 0);
      return `
      <span class="traveler-chip traveler-chip--group">
        <i class="bx bx-group" aria-hidden="true"></i>
        <span>${escapeHtml(p.name)} <small>Group${count ? ` · ${count}` : ""}</small></span>
        <button type="button" data-remove-person="${i}" aria-label="Remove">&times;</button>
      </span>`;
    }
    return `
    <span class="traveler-chip">
      <span>${escapeHtml(p.name)}${p.email ? ` <small>${escapeHtml(p.email)}</small>` : ""}</span>
      <button type="button" data-remove-person="${i}" aria-label="Remove">&times;</button>
    </span>`;
  }).join("") : `<span class="traveler-empty">${escapeHtml(emptyLabel)}</span>`;
  $all("[data-remove-person]", mount).forEach((button) => button.addEventListener("click", () => onRemove(Number(button.dataset.removePerson))));
}

/* ids: { mount, input, suggestions } — three element ids already in the DOM.
   opts: { directory, allowManualEmail, includeGroups, expandGroups, singleSelect,
           searchFn, emptyLabel, maxResults } */
function wirePeoplePicker(body, people, ids, opts) {
  opts = opts || {};
  const directory = opts.directory || getMemberDirectory();
  const includeGroups = opts.includeGroups != null ? !!opts.includeGroups : !opts.singleSelect;
  const expandGroups = !!opts.expandGroups;
  let onChange = null;
  let searchToken = 0;
  const mount = $("#" + ids.mount, body);
  const input = $("#" + ids.input, body);
  const suggestions = $("#" + ids.suggestions, body);
  if (!mount || !input || !suggestions) return { refresh() {}, setOnChange() {} };

  function chipOpts() {
    return { emptyLabel: opts.emptyLabel };
  }
  function personExists(entry) {
    if (isPickerGroupEntry(entry)) {
      const gid = entry.groupId || entry.id;
      return people.some((p) => isPickerGroupEntry(p) && (p.groupId || p.id) === gid);
    }
    const emailKey = normalizePersonKey(entry.email);
    const nameKey = normalizePersonKey(entry.name);
    return people.some((p) => {
      if (isPickerGroupEntry(p)) return false;
      return (emailKey && normalizePersonKey(p.email) === emailKey) || (!emailKey && nameKey && normalizePersonKey(p.name) === nameKey);
    });
  }
  function addPerson(entry) {
    if (!entry || !entry.name) return;
    if (isPickerGroupEntry(entry) && expandGroups) {
      const expanded = expandPickerPeople([entry]);
      expanded.forEach((member) => {
        if (personExists(member)) return;
        if (opts.singleSelect) {
          people.length = 0;
          people.push(member);
        } else {
          people.push(member);
        }
      });
      renderPersonChips(mount, people, removeAt, chipOpts());
      input.value = "";
      suggestions.innerHTML = "";
      if (typeof onChange === "function") onChange();
      return;
    }
    // singleSelect: replace whoever's there instead of requiring a manual
    // remove-then-add — used for "who owns this" fields (one person, not a
    // roster) so picking a new owner is a single click.
    if (opts.singleSelect) {
      people.length = 0;
      people.push(entry);
      renderPersonChips(mount, people, removeAt, chipOpts());
      input.value = "";
      suggestions.innerHTML = "";
      if (typeof onChange === "function") onChange();
      return;
    }
    if (personExists(entry)) return;
    people.push(entry);
    renderPersonChips(mount, people, removeAt, chipOpts());
    input.value = "";
    suggestions.innerHTML = "";
    if (typeof onChange === "function") onChange();
  }
  function removeAt(index) {
    people.splice(index, 1);
    renderPersonChips(mount, people, removeAt, chipOpts());
    if (typeof onChange === "function") onChange();
  }
  function renderSuggestionButtons(matches, diagnosticsHtml) {
    suggestions.innerHTML = matches.length ? matches.map((entry, index) => {
      if (isPickerGroupEntry(entry)) {
        const count = Array.isArray(entry.members) ? entry.members.length : (entry.memberCount || 0);
        return `
      <button type="button" class="traveler-suggestion traveler-suggestion--group" data-suggestion-index="${index}">
        <i class="bx bx-group" aria-hidden="true"></i>
        <strong>${escapeHtml(entry.name)}</strong>
        <span>Group · ${count} member${count === 1 ? "" : "s"}${expandGroups ? " — adds each person" : ""}</span>
      </button>`;
      }
      return `
      <button type="button" class="traveler-suggestion" data-suggestion-index="${index}">
        <strong>${escapeHtml(entry.name)}</strong>
        <span>${escapeHtml(entry.email || "No email on file")}</span>
      </button>`;
    }).join("") : `<div class="traveler-suggestion-empty">No matches${opts.allowManualEmail === false ? "." : ". Press Enter to add an email manually."}${diagnosticsHtml || ""}</div>`;
    $all("[data-suggestion-index]", suggestions).forEach((button) => {
      button.addEventListener("click", () => addPerson(matches[Number(button.dataset.suggestionIndex)]));
    });
  }
  /* Turns sharePointAdapter's per-source search attempt log into a plain
     summary shown right in the "no matches" box — so a failure is visible
     and reportable without opening dev tools. Every attempted source
     (People Picker, Search API, Graph) shows either how many raw hits it
     got or its actual error, in order. */
  function tenantSearchDiagnosticsHtml() {
    if (typeof sharePointAdapter === "undefined" || typeof sharePointAdapter.getLastTenantSearchAttempts !== "function") return "";
    const attempts = sharePointAdapter.getLastTenantSearchAttempts();
    if (!attempts || !attempts.length) return "";
    const lines = attempts.map((a) => a.ok
      ? `${escapeHtml(a.source)}: ${a.rawCount != null ? a.rawCount : a.matchCount} raw hit${(a.rawCount != null ? a.rawCount : a.matchCount) === 1 ? "" : "s"}, ${a.matchCount} usable`
      : `${escapeHtml(a.source)}: ${escapeHtml(a.error)}`);
    return `<div style="margin-top:6px;padding-top:6px;border-top:1px solid var(--aewttr-border);font-size:10.5px;color:var(--aewttr-muted);text-align:left;">${lines.map((l) => `<div>${l}</div>`).join("")}</div>`;
  }
  function mergeGroupMatches(personMatches, query, maxTotal) {
    if (!includeGroups) return personMatches.slice(0, maxTotal);
    const groups = matchPulseGroups(query, 4).filter((g) => !personExists(g));
    const combined = [...groups, ...personMatches];
    return combined.slice(0, maxTotal);
  }
  function drawSuggestions() {
    const query = normalizePersonKey(input.value);
    if (!query) { suggestions.innerHTML = ""; return; }
    const personMatches = dedupeDirectoryEntries(directory).filter((entry) => !personExists(entry)
      && (normalizePersonKey(entry.name).includes(query) || normalizePersonKey(entry.email).includes(query))
    );
    renderSuggestionButtons(mergeGroupMatches(personMatches, query, opts.maxResults || 8));
  }
  async function drawAsyncSuggestions() {
    const query = input.value.trim();
    const minLen = opts.minQueryLength == null ? 2 : opts.minQueryLength;
    if (!query) { suggestions.innerHTML = ""; return; }
    if (query.length < minLen) {
      suggestions.innerHTML = `<div class="traveler-suggestion-empty">Type at least ${minLen} characters to search the organization.</div>`;
      return;
    }
    const token = ++searchToken;
    suggestions.innerHTML = `<div class="traveler-suggestion-empty">Searching Microsoft 365…</div>`;
    try {
      const rawResults = await opts.searchFn(query);
      const personMatches = dedupeDirectoryEntries(rawResults).filter((entry) => entry && entry.name && !personExists(entry));
      const matches = mergeGroupMatches(personMatches, query, opts.maxResults || 12);
      if (token !== searchToken) return;
      if (!matches.length && Array.isArray(rawResults) && rawResults.length) {
        // The search itself succeeded and returned raw hits, but every one
        // was filtered out as junk (no valid name/email/claims login) — say
        // so explicitly rather than showing the same empty state as "no
        // matches at all", since these need different fixes.
        suggestions.innerHTML = `<div class="traveler-suggestion-empty">Found ${rawResults.length} raw result${rawResults.length === 1 ? "" : "s"}, but none looked like a real person after filtering. Check the browser console for details.</div>`;
        console.warn("[wirePeoplePicker] search returned results but all were filtered as junk:", rawResults);
        return;
      }
      renderSuggestionButtons(matches, matches.length ? "" : tenantSearchDiagnosticsHtml());
    } catch (e) {
      if (token !== searchToken) return;
      const detail = (e && e.friendly) || (e && e.message) || String(e);
      suggestions.innerHTML = `<div class="traveler-suggestion-empty">Search failed: ${escapeHtml(detail)}</div>`;
      console.error("[wirePeoplePicker] searchFn threw:", e);
    }
  }
  if (typeof opts.searchFn === "function") {
    let debounceTimer = null;
    input.addEventListener("input", () => {
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(() => { drawAsyncSuggestions(); }, opts.debounceMs || 300);
    });
  } else {
    input.addEventListener("input", drawSuggestions);
  }
  input.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    const query = input.value.trim();
    if (!query) return;
    if (includeGroups) {
      const groupExact = matchPulseGroups(query, 8).find((g) =>
        !personExists(g) && normalizePersonKey(g.name) === normalizePersonKey(query)
      );
      if (groupExact) { addPerson(groupExact); return; }
    }
    const exact = directory.find((entry) => !personExists(entry) && (
      normalizePersonKey(entry.email) === normalizePersonKey(query) || normalizePersonKey(entry.name) === normalizePersonKey(query)
    ));
    if (exact) { addPerson(exact); return; }
    if (opts.allowManualEmail !== false) {
      const emailMatch = query.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
      if (emailMatch) {
        const email = emailMatch[0];
        addPerson({ id: `manual-${normalizePersonKey(email)}`, name: query.replace(email, "").replace(/[<>]/g, "").trim() || email, email });
        return;
      }
    }
    toast("Pick someone from search results, or enter a valid email.", "error");
  });

  function refresh() {
    renderPersonChips(mount, people, removeAt, chipOpts());
    if (typeof opts.searchFn === "function") drawAsyncSuggestions();
    else drawSuggestions();
  }
  renderPersonChips(mount, people, removeAt, chipOpts());
  return { refresh, setOnChange(fn) { onChange = fn; } };
}

/* ---------- shared single-assignee autocomplete ----------
   For every "who owns this" text field (task assignee, subtask owner, rock
   owner, meeting task owner, etc). Unlike wirePeoplePicker (chips, multi-
   select), this keeps a single plain text input as the source of truth —
   callers already read/write `input.value` as a name string everywhere, so
   this only adds a dropdown of matching site members on top; it never
   requires the caller to change their save path. Manual/free-text names
   still work (e.g. contractors not in the directory) since nothing blocks
   submitting whatever text is typed. */
function wireAssigneeAutocomplete(body, inputOrId, opts) {
  opts = opts || {};
  const input = typeof inputOrId === "string" ? $("#" + inputOrId, body) : inputOrId;
  if (!input) return { destroy() {} };
  const directory = opts.directory || getMemberDirectory();

  // Portal dropdown to body to avoid overflow:hidden clipping in toolbar containers
  const list = document.createElement("div");
  list.className = "assignee-suggestions assignee-suggestions--portal";
  document.body.appendChild(list);

  function positionPortal() {
    const rect = input.getBoundingClientRect();
    list.style.left = `${rect.left}px`;
    list.style.width = `${Math.max(rect.width, 220)}px`;
    const spaceBelow = window.innerHeight - rect.bottom - 8;
    const spaceAbove = rect.top - 8;
    if (spaceBelow >= 120 || spaceBelow >= spaceAbove) {
      list.style.top = `${rect.bottom + 4}px`;
      list.style.bottom = "auto";
    } else {
      list.style.bottom = `${window.innerHeight - rect.top + 4}px`;
      list.style.top = "auto";
    }
  }

  let suppressDraw = false;
  let activeIndex = -1;
  let currentMatches = [];

  function close() { list.innerHTML = ""; list.classList.remove("open"); activeIndex = -1; }
  function choose(entry) {
    input.value = entry.name;
    close();
    suppressDraw = true;
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
    suppressDraw = false;
    if (typeof opts.onSelect === "function") opts.onSelect(entry);
  }
  function draw() {
    if (suppressDraw) return;
    const query = normalizePersonKey(input.value);
    currentMatches = query
      ? directory.filter((entry) => normalizePersonKey(entry.name).includes(query)).slice(0, 8)
      : directory.slice(0, 8);
    if (!currentMatches.length) { close(); return; }
    list.innerHTML = currentMatches.map((entry, index) => `
      <button type="button" class="assignee-suggestion ${index === activeIndex ? 'active' : ''}" data-assignee-index="${index}">
        ${typeof userAvatarHtml === "function" ? userAvatarHtml(entry.name, entry.email, 28) : ""}
        <div class="assignee-suggestion-info">
          <strong>${escapeHtml(entry.name)}</strong>
          ${entry.email ? `<span>${escapeHtml(entry.email)}</span>` : ""}
        </div>
      </button>
    `).join("");
    list.classList.add("open");
    positionPortal();
    $all("[data-assignee-index]", list).forEach((button) => {
      button.addEventListener("mousedown", (event) => {
        event.preventDefault();
        choose(currentMatches[Number(button.dataset.assigneeIndex)]);
      });
    });
  }

  function onKeyDown(e) {
    if (!list.classList.contains("open") || !currentMatches.length) return;
    if (e.key === "ArrowDown") {
      e.preventDefault();
      activeIndex = (activeIndex + 1) % currentMatches.length;
      draw();
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      activeIndex = activeIndex <= 0 ? currentMatches.length - 1 : activeIndex - 1;
      draw();
    } else if (e.key === "Enter") {
      e.preventDefault();
      const match = activeIndex >= 0 ? currentMatches[activeIndex] : currentMatches[0];
      if (match) {
        choose(match);
        // Find next focusable element
        const focusableSelectors = 'input:not([type="hidden"]):not([disabled]), select:not([disabled]), textarea:not([disabled]), button:not([disabled]), [tabindex]:not([tabindex="-1"])';
        let container = input.closest('.aewttr-modal-body') || document.body;
        const focusables = Array.from(container.querySelectorAll(focusableSelectors));
        const currentIndex = focusables.indexOf(input);
        if (currentIndex !== -1 && currentIndex + 1 < focusables.length) {
          focusables[currentIndex + 1].focus();
        }
      }
    }
  }

  const onInput = () => { activeIndex = -1; draw(); };
  const onFocus = () => { activeIndex = -1; draw(); };
  const onBlur = () => setTimeout(close, 120);
  input.addEventListener("input", onInput);
  input.addEventListener("focus", onFocus);
  input.addEventListener("blur", onBlur);
  input.addEventListener("keydown", onKeyDown);
  input.setAttribute("autocomplete", "off");
  return {
    destroy() {
      input.removeEventListener("input", onInput);
      input.removeEventListener("focus", onFocus);
      input.removeEventListener("blur", onBlur);
      input.removeEventListener("keydown", onKeyDown);
      list.remove();
    }
  };
}

/* ---------- CapraGPT (OpenAI-compatible chat completions) ----------
   Config — including the API key — lives in APP_CONFIG.aiReview in
   app-config.js. See the comment there for where to put the key and what
   to check if CapraGPT's endpoint URL differs from the placeholder. */
async function callCapraGpt(userText, systemPrompt) {
  // Config always reads from APP_CONFIG.aiReview in app-config.js.
  const live = getAiConfig();
  const cfg = {
    enabled: live.enabled,
    apiKey: live.apiKey,
    endpoint: live.endpoint,
    model: live.model
  };
  if (!cfg.enabled) throw new Error("AI review is disabled — turn it on in Admin > AI Review Config.");
  if (!cfg.apiKey) throw new Error("No CapraGPT API key configured — set it in Admin > AI Review Config.");
  let response;
  try {
    response = await fetch(cfg.endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${cfg.apiKey}`
      },
      body: JSON.stringify({
        model: cfg.model || "gpt-4.1",
        messages: [
          { role: "system", content: systemPrompt || "You are a helpful assistant." },
          { role: "user", content: userText }
        ],
        temperature: 0.2
      })
    });
  } catch (networkError) {
    throw new Error(`Could not reach CapraGPT: ${String(networkError && networkError.message || networkError)}`);
  }
  const raw = await response.text().catch(() => "");
  if (!response.ok) {
    throw new Error(`CapraGPT request failed (HTTP ${response.status}): ${raw.slice(0, 300) || response.statusText}`);
  }
  let data;
  try { data = JSON.parse(raw); } catch (e) {
    throw new Error("CapraGPT returned a non-JSON response — check the Endpoint in Admin > AI Review Config is correct.");
  }
  const content = data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
  if (!content) throw new Error("CapraGPT response didn't include a message — check the model name and endpoint.");
  return content;
}

/* ---------- shell render ---------- */
function fmtRelativeTime(iso) {
  if (!iso) return "";
  const then = new Date(iso).getTime();
  if (!then || Number.isNaN(then)) return "";
  const diff = Date.now() - then;
  const mins = Math.floor(diff / 60000);
  if (mins < 1) return "Just now";
  if (mins < 60) return `${mins}m ago`;
  const hours = Math.floor(mins / 60);
  if (hours < 24) return `${hours}h ago`;
  const days = Math.floor(hours / 24);
  if (days < 7) return `${days}d ago`;
  return fmtDate(String(iso).slice(0, 10));
}

function collectUserNotificationEvents() {
  const events = [];
  if (typeof window.getDocReviewNotificationEvents === "function") {
    events.push(...window.getDocReviewNotificationEvents());
  }
  if (typeof window.getTravelNotificationEvents === "function") {
    events.push(...window.getTravelNotificationEvents());
  }
  return events.sort((a, b) => String(b.sortKey || "").localeCompare(String(a.sortKey || "")));
}

function getUserNotificationItems() {
  return collectUserNotificationEvents();
}

function getAppAlertCounts() {
  const docReview = typeof window.getDocReviewMetrics === "function"
    ? window.getDocReviewMetrics().needsMyReviewCount || 0
    : 0;
  const travelMetrics = typeof window.getTravelMetrics === "function" ? window.getTravelMetrics() : null;
  const travelAction = (travelMetrics && travelMetrics.needsMyActionCount) || 0;
  const travelApprove = (travelMetrics && travelMetrics.pendingApprovalCount) || 0;
  const travelStatus = typeof window.getTravelStatusUpdateCount === "function"
    ? window.getTravelStatusUpdateCount()
    : 0;
  return {
    docreview: docReview,
    travel: travelAction,
    travelMine: travelStatus,
    travelApprove,
    // Travel approvals are only actionable from the Travel tab now — Admin no
    // longer has an Approval Queue tool, so it shouldn't also badge the same
    // count a second time.
    admin: 0
  };
}

function renderAlertIndicator(count, options) {
  options = options || {};
  const total = Number(count) || 0;
  if (!total) return "";
  const label = options.label || `${total} item${total === 1 ? "" : "s"} need attention`;
  const display = total > 99 ? "99+" : String(total);
  if (options.variant === "bell") {
    return `<span class="aewttr-alert-bell" aria-label="${escapeHtml(label)}" title="${escapeHtml(label)}"><i class="bx bx-bell"></i><span class="aewttr-alert-bell-count">${display}</span></span>`;
  }
  if (options.variant === "tile") {
    return `<div class="app-tile-alert" aria-label="${escapeHtml(label)}" title="${escapeHtml(label)}"><i class="bx bx-bell"></i><span>${display}</span></div>`;
  }
  if (options.variant === "pill") {
    return `<span class="filter-pill-alert" aria-label="${escapeHtml(label)}" title="${escapeHtml(label)}"><i class="bx bx-bell"></i><span>${display}</span></span>`;
  }
  if (options.variant === "icon") {
    return `<span class="aewttr-alert-icon" aria-label="${escapeHtml(label)}" title="${escapeHtml(label)}"><i class="bx bx-bell"></i></span>`;
  }
  return `<span class="aewttr-nav-badge" aria-label="${escapeHtml(label)}" title="${escapeHtml(label)}">${display}</span>`;
}

window.getAppAlertCounts = getAppAlertCounts;
window.renderAlertIndicator = renderAlertIndicator;

function scheduleNotificationRefresh() {
  refreshUserNotifications();
  requestAnimationFrame(() => refreshUserNotifications());
  setTimeout(() => refreshUserNotifications(), 400);
  setTimeout(() => refreshUserNotifications(), 2000);
}

function refreshUserNotifications() {
  const btn = $("#user-notify-btn");
  if (!btn) return;
  const events = collectUserNotificationEvents();
  const count = events.length;
  const badge = $("#user-notify-badge");
  btn.classList.toggle("has-alerts", count > 0);
  btn.setAttribute("aria-label", count ? `${count} notification${count === 1 ? "" : "s"}` : "Notifications");
  if (badge) {
    badge.hidden = !count;
    badge.textContent = count > 99 ? "99+" : String(count);
  }
  const panel = $("#user-notify-panel");
  if (!panel) return;
  if (!window.AEWTTR.state.notifyPanelTab) window.AEWTTR.state.notifyPanelTab = "Document Review";
  const tabs = [
    { key: "Document Review", label: "Document Review" },
    { key: "Travel", label: "Travel" },
    { key: "All", label: "All" }
  ];
  const activeTab = window.AEWTTR.state.notifyPanelTab;
  const filtered = activeTab === "All"
    ? events
    : events.filter((item) => String(item.category || "") === activeTab);
  const tabCounts = {
    "Document Review": events.filter((item) => item.category === "Document Review").length,
    Travel: events.filter((item) => item.category === "Travel").length,
    All: events.length
  };
  const panelHeadActions = `<button type="button" class="aewttr-icon-btn aewttr-notify-prefs-btn" id="user-notify-prefs-btn"${tip("Choose which notification areas you receive")} aria-label="Notification preferences"><i class="bx bx-cog"></i></button>`;
  panel.innerHTML = `
    <div class="aewttr-notify-panel-head">
      <span style="display:flex;align-items:center;gap:8px;"><span>Notifications</span>${count ? `<span class="aewttr-notify-panel-count">${count}</span>` : ""}</span>
      ${panelHeadActions}
    </div>
    <div class="aewttr-notify-tabs" role="tablist" aria-label="Notification areas">
      ${tabs.map((tab) => `
        <button type="button" class="aewttr-notify-tab ${activeTab === tab.key ? "active" : ""}" data-notify-tab="${escapeHtml(tab.key)}" role="tab" aria-selected="${activeTab === tab.key ? "true" : "false"}">
          ${escapeHtml(tab.label)}${tabCounts[tab.key] ? `<em>${tabCounts[tab.key]}</em>` : ""}
        </button>`).join("")}
    </div>
    ${filtered.length ? `
    <div class="aewttr-notify-list">
      ${filtered.map((item) => `
        <button type="button" class="aewttr-notify-item" data-notify-route="${escapeHtml(item.route)}"${item.queryTr ? ` data-notify-tr="${escapeHtml(item.queryTr)}"` : ""}${item.queryDoc ? ` data-notify-doc="${escapeHtml(item.queryDoc)}"` : ""}${item.queryMode ? ` data-notify-mode="${escapeHtml(item.queryMode)}"` : ""}>
          <span class="aewttr-notify-item-icon tone-${escapeHtml(item.tone || "blue")}"><i class="bx ${escapeHtml(item.icon || "bx-bell")}"></i></span>
          <span class="aewttr-notify-item-body">
            <span class="aewttr-notify-item-top">
              <strong>${escapeHtml(item.title)}</strong>
              ${item.time ? `<time>${escapeHtml(item.time)}</time>` : ""}
            </span>
            <span class="aewttr-notify-preview">${escapeHtml(item.preview)}</span>
            ${item.category ? `<span class="aewttr-notify-category">${escapeHtml(item.category)}</span>` : ""}
          </span>
        </button>`).join("")}
    </div>` : `
    <div class="aewttr-notify-empty">${count && activeTab !== "All" ? `No ${escapeHtml(activeTab)} updates right now.` : "You're all caught up — no new updates."}</div>`}`;
  $all("[data-notify-tab]", panel).forEach((tabBtn) => {
    tabBtn.addEventListener("click", (event) => {
      event.stopPropagation();
      window.AEWTTR.state.notifyPanelTab = tabBtn.dataset.notifyTab;
      refreshUserNotifications();
    });
  });
  $all("[data-notify-route]", panel).forEach((node) => {
    node.addEventListener("click", () => {
      panel.hidden = true;
      const query = {};
      if (node.dataset.notifyTr) query.tr = node.dataset.notifyTr;
      if (node.dataset.notifyDoc) query.doc = node.dataset.notifyDoc;
      if (node.dataset.notifyMode) query.mode = node.dataset.notifyMode;
      if (Object.keys(query).length) queueRouteAction(query);
      navigate(node.dataset.notifyRoute, Object.keys(query).length ? query : undefined);
    });
  });
  const prefsBtn = $("#user-notify-prefs-btn", panel);
  if (prefsBtn) prefsBtn.addEventListener("click", (event) => {
    event.stopPropagation();
    panel.hidden = true;
    navigate("notification-settings");
  });
}

function wireUserNotificationPanel() {
  const btn = $("#user-notify-btn");
  const panel = $("#user-notify-panel");
  if (!btn || !panel) return;
  btn.addEventListener("click", (event) => {
    event.stopPropagation();
    panel.hidden = !panel.hidden;
    if (!panel.hidden) refreshUserNotifications();
  });
  document.addEventListener("click", (event) => {
    if (panel.hidden) return;
    if (event.target.closest(".aewttr-notify-wrap")) return;
    panel.hidden = true;
  });
}

function renderShell() {
  const db = window.AEWTTR.db;
  const root = $("#aewttr-root");
  const photoUrl = currentUserProfilePhotoUrl();
  const cuiTop = typeof cuiMarkingBarHtml === "function" ? cuiMarkingBarHtml("top") : "";
  const cuiBottom = typeof cuiMarkingBarHtml === "function" ? cuiMarkingBarHtml("bottom") : "";
  root.innerHTML = `
    <div class="aewttr-shell">
      ${cuiTop}
      <header class="aewttr-topnav">
        <button type="button" class="aewttr-topnav-brand" id="pulse-home-btn" aria-label="Go to home">
          <img class="pulse-nav-dots-img" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAACb0lEQVR42q2YTWtTQRSG33MTYtUgCCrdSBA3imCV6s9wI1rostD2V9Tf4j+w1C7FhRs3CgbBj40NrkqxdBH8aJPePG7OlPF6k9yvA2HuwNw5z8yd8545MbkBJklmBtCS9EDSbUkLkn5I6pvZNx/bMrNUTVlw7s/rQB9I+deGwAtgOUA05hxIgAvAdsZpCpxmYMbAZmMQQMshdtzBKGf1ABN3PvH+SpMQaz7pCfMtdYgj4JrDWx3nHeCzT5xSzMbePvM52lX9J5LuS7olybxfxEwSkh55P60DcCeasMx7JqkHdD10rSrAgjunwvsd/6nODhz6asqsIMAOJf2qC/BO0p+Sn2HiYz+a2YkrI5UAzGwg6ZUDlDlMJul59FxLCZemqF6ejbx94wqa1FZCbzcycR4EZxLBhfgfANeDjM9YWDGRiiBWgYM5O/Aa6Pn4ZIq0t3Jg2jNhIohFYAt4C+wDh8CeJ6qn0fgk836So7CXgW6en5kQUb/rk3TmOEui5ycOu+fw+76YLWCxCETuduVtbewc6PnnmWUHwGrhLDrvIHkUmB/IQXSAT3MO8TgC2WgklYcw9JCMQ3RWKg9wS3XTeNvbx5kUXTSVv2xKP3aj21IRC5/lN3AjqaqeZpYC5yTdrXiXOC/pYT0ZlS5KulQhH4T0f6UuwMh/ZS2k/+NKAOEGZGY/JX331UxKAiDpU50dCDG8W/Eu8VXSh7pp3PxqfhQJzjwLV/+1Rooab1emFC9ZEQpCtePwrSYhNjNaMO2Cs+1lYFJLCadALHsBO8xZfR9Y/68ab7DCPivZgZuS7km6KulY0hdJ7128zv4GkKS/w3S3ykLv268AAAAASUVORK5CYII=" width="32" height="32" alt="" aria-hidden="true">
          ${window.PULSE_PORT_CONFIG && window.PULSE_PORT_CONFIG.brandLabel ? `<span class="pulse-brand-port-tag">${escapeHtml(window.PULSE_PORT_CONFIG.brandLabel)}</span>` : ""}
        </button>
        <nav class="aewttr-nav" id="aewttr-nav"></nav>
        <div class="aewttr-topnav-right">
          <div class="save-indicator" id="save-indicator" aria-live="polite" aria-label="Saving changes" role="status">
            <span class="save-indicator__dot"></span><span class="save-indicator__dot"></span>
          </div>
          <button type="button" class="aewttr-report-issue" id="report-issue-btn"${tip("Report a problem with the current page")}>
            <i class="bx bx-message-error"></i><span>Report issue</span>
          </button>
          <button class="aewttr-icon-btn" id="theme-toggle"${tip("Switch between light and dark theme")} aria-label="Toggle light / dark theme">
            <i class="bx bx-moon" id="theme-toggle-icon"></i>
          </button>
          <div class="aewttr-notify-wrap">
            <button type="button" class="aewttr-notify-btn" id="user-notify-btn"${tip("View notifications for travel, documents, and more")} aria-label="Notifications">
              <i class="bx bx-bell"></i>
              <span class="aewttr-notify-badge" id="user-notify-badge" hidden>0</span>
            </button>
            <div class="aewttr-notify-panel" id="user-notify-panel" hidden></div>
          </div>
          <button type="button" class="aewttr-topnav-user" id="topnav-user-btn"${tip("Notification settings")} aria-label="${escapeHtml(db.user.name)}, ${escapeHtml(currentAppRole())}. Open notification settings">
            <div class="av ${photoUrl ? "has-photo" : ""}">
              ${photoUrl ? `<img class="av-photo" src="${escapeHtml(photoUrl)}" alt="${escapeHtml(db.user.name)} profile photo" onerror="this.remove(); this.parentElement.classList.remove('has-photo');">` : ""}
              <span class="av-fallback">${escapeHtml(initials(db.user.name))}</span>
            </div>
            <div class="meta">
              <div class="who">${escapeHtml(db.user.name)}</div>
              <div class="role">${escapeHtml(currentAppRole())}</div>
            </div>
          </button>
        </div>
      </header>
      <div class="aewttr-main">
        <div class="aewttr-page-header">
          <div>
            <h1 id="page-title">Dashboard</h1>
            <div class="subtitle" id="page-subtitle">Welcome back</div>
          </div>
          <div class="aewttr-topbar-actions" id="page-actions"></div>
        </div>
        <main class="aewttr-content" id="page-content"></main>
      </div>
      ${cuiBottom}
    </div>
  `;
  renderNav();
  scheduleNotificationRefresh();
  wireUserNotificationPanel();
  const topnavUserBtn = $("#topnav-user-btn", root);
  if (topnavUserBtn) topnavUserBtn.addEventListener("click", openUserSettingsModal);
  wirePulseBrandHome($("#pulse-home-btn", root));
  document.body.classList.remove("help-mode");
  initTheme();
  const reportIssueBtn = $("#report-issue-btn", root);
  if (reportIssueBtn) reportIssueBtn.addEventListener("click", () => {
    if (typeof openIssueReportModal === "function") openIssueReportModal();
  });
  $("#theme-toggle").addEventListener("click", toggleTheme);
  wirePortaledTooltips(document);
}

/* ---------- User settings modal ---------- */
function openUserSettingsModal() {
  const db = window.AEWTTR.db;
  const user = db.user || {};
  const prefs = typeof normalizeNotificationPrefs === "function" ? normalizeNotificationPrefs(user.notificationPrefs) : (user.notificationPrefs || {});
  const photoUrl = typeof currentUserProfilePhotoUrl === "function" ? currentUserProfilePhotoUrl() : "";
  const role = typeof currentAppRole === "function" ? currentAppRole() : "";
  const member = typeof currentUserMember === "function" ? currentUserMember() : null;
  const initStr = typeof initials === "function" ? initials(user.name) : (user.name || "?").slice(0, 2).toUpperCase();

  let editingSection = null;
  let saveTimer = null;

  function setSettingsSaveStatus(text, isErr) {
    const el = modal.querySelector("#uset-save-status");
    if (!el) return;
    el.textContent = text;
    el.style.color = isErr ? "var(--aewttr-red)" : "var(--aewttr-muted)";
  }

  function autoSave() {
    setSettingsSaveStatus("Saving…");
    clearTimeout(saveTimer);
    saveTimer = setTimeout(async () => {
      const liveDb = window.AEWTTR.db;
      liveDb.user.notificationPrefs = prefs;
      const liveMember = typeof currentUserMember === "function" ? currentUserMember() : null;
      if (liveMember) liveMember.notificationPrefs = prefs;
      if (typeof isSharePointMode === "function" && isSharePointMode() && typeof sharePointAdapter !== "undefined") {
        try {
          await sharePointAdapter.saveCurrentUserNotificationPrefs(window.AEWTTR.siteUrl, window.AEWTTR.currentSpUser, prefs);
        } catch (e) {
          setSettingsSaveStatus("Save failed", true);
          return;
        }
      }
      setSettingsSaveStatus("Saved");
      setTimeout(() => setSettingsSaveStatus(""), 2000);
    }, 500);
  }

  const CHANNEL_META = [
    { key: "email", label: "Email", icon: "bx-envelope" },
    { key: "teams", label: "Teams", icon: "bx-chat" },
    { key: "inApp", label: "In-app", icon: "bx-bell" }
  ];

  const AREA_META = typeof NOTIFICATION_AREA_META !== "undefined"
    ? Object.entries(NOTIFICATION_AREA_META).map(([key, v]) => ({ key, label: v.label || key }))
    : [
      { key: "Travel", label: "Travel" },
      { key: "Documents", label: "Documents" },
      { key: "Projects", label: "Projects" },
      { key: "Weekly", label: "Weekly meeting" }
    ];

  function channelViewHtml() {
    return CHANNEL_META.map((c) => {
      const on = prefs.channels && prefs.channels[c.key];
      return `<span class="uset-channel-chip ${on ? "uset-channel-chip--on" : ""}"><i class="bx ${c.icon}"></i> ${c.label}: <strong>${on ? "On" : "Off"}</strong></span>`;
    }).join("");
  }

  function channelEditHtml() {
    return CHANNEL_META.map((c) => {
      const on = prefs.channels && prefs.channels[c.key];
      return `<label class="uset-toggle-row"><span class="uset-toggle-label"><i class="bx ${c.icon}"></i> ${c.label}</span><input type="checkbox" class="uset-channel-toggle" data-channel="${c.key}" ${on ? "checked" : ""}></label>`;
    }).join("");
  }

  function areaViewHtml() {
    if (prefs.everything) return `<span class="uset-area-chip uset-area-chip--on">Everything</span>`;
    const areas = prefs.areas || [];
    return AREA_META.map((a) => {
      const on = areas.includes(a.key);
      return `<span class="uset-area-chip ${on ? "uset-area-chip--on" : ""}">${a.label}</span>`;
    }).join("");
  }

  function areaEditHtml() {
    const areas = prefs.areas || [];
    return `<label class="uset-toggle-row"><span class="uset-toggle-label"><strong>Everything</strong> <small>(all events)</small></span><input type="checkbox" id="uset-everything" ${prefs.everything ? "checked" : ""}></label>
    <div id="uset-areas-list" ${prefs.everything ? 'style="opacity:.4;pointer-events:none;"' : ""}>
      ${AREA_META.map((a) => `<label class="uset-toggle-row"><span class="uset-toggle-label">${a.label}</span><input type="checkbox" class="uset-area-cb" data-area="${a.key}" ${areas.includes(a.key) ? "checked" : ""}></label>`).join("")}
    </div>`;
  }

  function sectionHtml(id, title, viewContent, isEditing) {
    return `
      <div class="uset-section" id="uset-sec-${id}">
        <div class="uset-section-head">
          <span class="uset-section-title">${title}</span>
          <button type="button" class="uset-edit-btn" data-edit-section="${id}" title="Edit">
            <i class="bx ${isEditing ? "bx-check" : "bx-pencil"}"></i>
            ${isEditing ? "Done" : "Edit"}
          </button>
        </div>
        <div class="uset-section-body">
          ${isEditing ? viewContent.edit : viewContent.view}
        </div>
      </div>`;
  }

  function renderModal() {
    const content = modal.querySelector(".uset-content");
    if (!content) return;
    content.innerHTML = `
      <div class="uset-profile">
        <div class="av ${photoUrl ? "has-photo" : ""}" style="width:52px;height:52px;font-size:18px;">
          ${photoUrl ? `<img class="av-photo" src="${escapeHtml(photoUrl)}" alt="">` : ""}
          <span class="av-fallback">${escapeHtml(initStr)}</span>
        </div>
        <div>
          <div class="uset-profile-name">${escapeHtml(user.name || "")}</div>
          <div class="uset-profile-role">${escapeHtml(member ? member.email || "" : "")}${role ? ` · ${escapeHtml(role)}` : ""}</div>
        </div>
      </div>
      ${sectionHtml("channels", "Notification channels",
        { view: channelViewHtml(), edit: channelEditHtml() },
        editingSection === "channels")}
      ${sectionHtml("areas", "Notification areas",
        { view: areaViewHtml(), edit: areaEditHtml() },
        editingSection === "areas")}
      ${sectionHtml("delivery", "Document delivery",
        {
          view: `<span class="uset-area-chip uset-area-chip--on">${prefs.documents && prefs.documents.delivery === "digest" ? "Digest" : "Immediate"}</span>`,
          edit: `<div style="display:flex;gap:8px;flex-wrap:wrap;">
            <label class="uset-radio-row"><input type="radio" name="uset-delivery" value="immediate" ${(!prefs.documents || prefs.documents.delivery !== "digest") ? "checked" : ""}> Immediate</label>
            <label class="uset-radio-row"><input type="radio" name="uset-delivery" value="digest" ${prefs.documents && prefs.documents.delivery === "digest" ? "checked" : ""}> Digest</label>
          </div>`
        },
        editingSection === "delivery")}
      <div class="uset-actions">
        <a href="#" id="uset-full-settings">Open full notification settings →</a>
      </div>
    `;

    // Wire edit toggles
    content.querySelectorAll("[data-edit-section]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const sec = btn.dataset.editSection;
        editingSection = (editingSection === sec) ? null : sec;
        renderModal();
        // Focus first input in section
        const secEl = modal.querySelector(`#uset-sec-${sec} input, #uset-sec-${sec} select`);
        if (secEl) secEl.focus();
      });
    });

    // Wire channel toggles
    content.querySelectorAll(".uset-channel-toggle").forEach((cb) => {
      cb.addEventListener("change", () => {
        if (!prefs.channels) prefs.channels = {};
        prefs.channels[cb.dataset.channel] = cb.checked;
        autoSave();
      });
    });

    // Wire everything toggle
    const everythingCb = content.querySelector("#uset-everything");
    if (everythingCb) {
      everythingCb.addEventListener("change", () => {
        prefs.everything = everythingCb.checked;
        const areasList = content.querySelector("#uset-areas-list");
        if (areasList) { areasList.style.opacity = everythingCb.checked ? ".4" : ""; areasList.style.pointerEvents = everythingCb.checked ? "none" : ""; }
        autoSave();
      });
    }

    // Wire area checkboxes
    content.querySelectorAll(".uset-area-cb").forEach((cb) => {
      cb.addEventListener("change", () => {
        if (!prefs.areas) prefs.areas = [];
        if (cb.checked) { if (!prefs.areas.includes(cb.dataset.area)) prefs.areas.push(cb.dataset.area); }
        else { prefs.areas = prefs.areas.filter((a) => a !== cb.dataset.area); }
        autoSave();
      });
    });

    // Wire delivery radio
    content.querySelectorAll("[name='uset-delivery']").forEach((r) => {
      r.addEventListener("change", () => {
        if (!prefs.documents) prefs.documents = {};
        prefs.documents.delivery = r.value;
        autoSave();
      });
    });

    // Full settings link
    const fullLink = content.querySelector("#uset-full-settings");
    if (fullLink) {
      fullLink.addEventListener("click", (e) => {
        e.preventDefault();
        closeModal();
        if (typeof navigate === "function") navigate("notification-settings");
      });
    }
  }

  const modal = openModal(`
    <div class="aewttr-modal-head">
      <h3>Settings</h3>
      <div class="uset-save-status" id="uset-save-status"></div>
      <button class="aewttr-modal-close">&times;</button>
    </div>
    <div class="aewttr-modal-body uset-modal-body">
      <div class="uset-content"></div>
    </div>
  `, { wide: true });

  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  renderModal();
}

/* ---------- safe localStorage wrapper ---------- */
function lsGet(key) { try { return localStorage.getItem(key); } catch (e) { return null; } }
function lsSet(key, val) { try { localStorage.setItem(key, val); } catch (e) { /* sandboxed */ } }
function ssGet(key) { try { return sessionStorage.getItem(key); } catch (e) { return null; } }
function ssSet(key, val) { try { sessionStorage.setItem(key, val); } catch (e) { /* sandboxed */ } }

// Cache is refreshed on every write (see writeRecord/repoRemove) and every
// background poll, so in an active session it's rarely more than a few
// seconds old — this ceiling only bounds the worst case (someone else's
// change, seen from a tab that's been idle a while).
const SP_DB_CACHE_MAX_AGE_MS = 30 * 60 * 1000;

function spDbCacheKey(siteUrl) {
  return `pulse-sp-db::${String(siteUrl || "").replace(/\/$/, "").toLowerCase()}`;
}

function readSpDbCache(siteUrl) {
  // Try sessionStorage first (fastest — same tab), then localStorage (survives tab close).
  const ssRaw = ssGet(spDbCacheKey(siteUrl));
  const lsRaw = ssRaw ? null : lsGet(spDbCacheKey(siteUrl));
  const raw = ssRaw || lsRaw;
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || !parsed.data || !parsed.savedAt) return null;
    if (Date.now() - Number(parsed.savedAt) > SP_DB_CACHE_MAX_AGE_MS) return null;
    return parsed;
  } catch (e) {
    return null;
  }
}

function persistSpDbCache(siteUrl, db) {
  if (!siteUrl || !db) return;
  try {
    const payload = {
      savedAt: Date.now(),
      data: typeof getSerializableDb === "function" ? getSerializableDb(db) : JSON.parse(JSON.stringify(db))
    };
    const serialized = JSON.stringify(payload);
    ssSet(spDbCacheKey(siteUrl), serialized);
    // Mirror to localStorage so the cache survives tab close/reopen.
    lsSet(spDbCacheKey(siteUrl), serialized);
  } catch (e) { /* storage full or blocked */ }
}

function buildBootUserObject(currentUser, roleInfo) {
  return {
    id: "u" + (currentUser.spUserId || 0),
    name: currentUser.displayName || "Unknown",
    email: currentUser.email || "",
    role: roleInfo.role,
    isAdmin: roleInfo.isAdmin,
    isMeetingAdmin: roleInfo.isMeetingAdmin,
    isFinanceAdmin: roleInfo.isFinanceAdmin,
    isDocAdmin: roleInfo.isDocAdmin,
    themeMode: normalizeThemeValue(roleInfo.themeMode || getCurrentTheme()),
    notificationPrefs: roleInfo.notificationPrefs || null,
    spo: "—",
    powerbi: "—"
  };
}

/* ---------- theme (light / dark) ---------- */
const THEME_KEY = "aewttr_theme";
function normalizeThemeValue(value) {
  return String(value || "").trim().toLowerCase() === "dark" ? "dark" : "light";
}
function getCurrentTheme() {
  return normalizeThemeValue(lsGet(THEME_KEY) || document.documentElement.getAttribute("data-theme") || "light");
}
function setCurrentThemeValue(theme) {
  const normalized = normalizeThemeValue(theme);
  lsSet(THEME_KEY, normalized);
  document.documentElement.setAttribute("data-theme", normalized);
  updateThemeIcon();
  return normalized;
}
function initTheme() {
  setCurrentThemeValue(getCurrentTheme());
}
async function persistThemePreference(theme) {
  const db = window.AEWTTR && window.AEWTTR.db;
  if (db && db.user) db.user.themeMode = normalizeThemeValue(theme);
  if (!(window.AEWTTR && window.AEWTTR.mode === "sharepoint" && window.AEWTTR.siteUrl && window.AEWTTR.currentSpUser)) return;
  try {
    await sharePointAdapter.saveCurrentUserThemeMode(window.AEWTTR.siteUrl, window.AEWTTR.currentSpUser, theme);
  } catch (e) {
    console.warn("PULSE: failed saving theme preference to SharePoint role record.", e);
  }
}
async function toggleTheme() {
  const next = getCurrentTheme() === "dark" ? "light" : "dark";
  setCurrentThemeValue(next);
  persistThemePreference(next);
}
function updateThemeIcon() {
  const icon = $("#theme-toggle-icon");
  if (icon) icon.className = "bx " + (getCurrentTheme() === "dark" ? "bx-sun" : "bx-moon");
}

/* ---------- help mode ---------- */
const HELP_MODE_KEY = "aewttr_help_mode";
function isHelpModeOn() { return lsGet(HELP_MODE_KEY) === "1"; }
function setHelpMode(on) {
  lsSet(HELP_MODE_KEY, on ? "1" : "0");
  applyHelpMode();
}
function applyHelpMode() {
  document.body.classList.toggle("help-mode", isHelpModeOn());
  const label = $("#help-toggle-label");
  if (label) label.textContent = "Help Mode: " + (isHelpModeOn() ? "On" : "Off");
  const btn = $("#help-toggle");
  if (btn) btn.classList.toggle("active", isHelpModeOn());
}
function renderNav() {
  const { app } = currentRoute();
  const nav = $("#aewttr-nav");
  if (!nav) return;
  const navList = (window.PULSE_PORT_CONFIG && window.PULSE_PORT_CONFIG.navItems) ? window.PULSE_PORT_CONFIG.navItems : NAV_ITEMS;
  const visibleItems = navList.filter((i) => {
    if (i.adminOnly && !canCurrentUserAccessAdmin()) return false;
    if (i.route === "admin" && !canCurrentUserAccessAdmin()) return false;
    return true;
  });
  const alerts = typeof getAppAlertCounts === "function" ? getAppAlertCounts() : { docreview: 0, travel: 0 };
  const navAlertLabels = {
    travel: "Travel updates need your attention",
    docreview: "Documents need your review",
    admin: "Admin approvals need your attention"
  };

  const existing = $all(".aewttr-nav-item", nav);
  const canPatch = existing.length === visibleItems.length
    && existing.every((item, index) => item.dataset.route === visibleItems[index].route);

  if (canPatch) {
    existing.forEach((item, index) => {
      const route = visibleItems[index].route;
      item.classList.toggle("active", route === app);
      let badge = item.querySelector(".aewttr-alert-bell");
      const count = alerts[route] || 0;
      if (!count) {
        if (badge) badge.remove();
        return;
      }
      const label = navAlertLabels[route] || `${count} items need attention`;
      if (!badge) {
        item.insertAdjacentHTML("beforeend", renderAlertIndicator(count, { variant: "bell", label }));
        return;
      }
      badge.setAttribute("aria-label", label);
      badge.setAttribute("title", label);
      const countEl = badge.querySelector(".aewttr-alert-bell-count");
      if (countEl) countEl.textContent = count > 99 ? "99+" : String(count);
    });
    return;
  }

  nav.innerHTML = visibleItems.map(i => `
    <div class="aewttr-nav-item ${i.route === app ? "active" : ""}" data-route="${i.route}">
      <span class="aewttr-nav-label">${i.label}</span>
      ${alerts[i.route] ? renderAlertIndicator(alerts[i.route], { variant: "bell", label: navAlertLabels[i.route] || `${alerts[i.route]} items need attention` }) : ""}
    </div>`).join("");
  $all(".aewttr-nav-item", nav).forEach(item => {
    item.addEventListener("click", () => navigate(item.dataset.route));
  });
}
function setTopbar(title, subtitle, actionsHtml) {
  $("#page-title").textContent = title;
  $("#page-subtitle").textContent = subtitle || "";
  $("#page-actions").innerHTML = actionsHtml || "";
}

/* ---------- router dispatch ---------- */
const PAGE_RENDERERS = {}; // populated by page modules: PAGE_RENDERERS.dashboard = function(parts){...}

/* Polling every 30s alone means "I just navigated to Projects right after
   someone else added one" waits up to 30s for stale data to catch up. Since
   every navigation already runs through here, piggyback a background
   refresh on it too — throttled so rapid clicking around the app doesn't
   hammer SharePoint, and skipped entirely by refreshSharePointData's own
   guards (pending local writes, mid-wizard, already in flight). Fire-and-
   forget: navigation must stay synchronous/instant. */
function maybeProactiveRefreshOnNavigate() {
  if (window.AEWTTR.mode !== "sharepoint") return;
  // Leaving a page: push any debounced edits to SharePoint right away so
  // quick edit → navigate → come back never loses work.
  if (typeof Repo !== "undefined" && Repo.hasPendingChanges && Repo.hasPendingChanges() && typeof Repo.flush === "function") {
    Repo.flush().catch(() => {});
    return;
  }
  const lastAt = window.AEWTTR.lastSharePointRefreshAt || 0;
  if (Date.now() - lastAt < 8000) return;
  if (typeof refreshSharePointData !== "function") return;
  refreshSharePointData("navigate");
}

function showSaveIndicator() {
  const el = document.getElementById("save-indicator");
  if (el) el.classList.add("is-active");
}
function hideSaveIndicator() {
  const el = document.getElementById("save-indicator");
  if (el) el.classList.remove("is-active");
}

/* Registry of "flush my pending debounced save right now" callbacks. Any page
   with a local setTimeout-before-Repo.save autosave pattern (Project Settings,
   Notification Settings, meeting agenda fields, etc.) registers one here on
   mount and unregisters on its own clean teardown. Without this, navigating
   away while a local timer is still pending silently drops the edit — the
   save call was never even made, so Repo.flush() (below) has nothing to
   rescue, since it only knows about writes already inside Repo's own queue. */
window.AEWTTR._pendingAutosaveFlushers = window.AEWTTR._pendingAutosaveFlushers || new Set();
function registerAutosaveFlusher(fn) { window.AEWTTR._pendingAutosaveFlushers.add(fn); }
function unregisterAutosaveFlusher(fn) { window.AEWTTR._pendingAutosaveFlushers.delete(fn); }
async function flushPendingAutosaves() {
  const flushers = Array.from(window.AEWTTR._pendingAutosaveFlushers);
  await Promise.all(flushers.map((fn) => {
    try { return Promise.resolve(fn()); } catch (e) { return null; }
  }));
}

async function renderPage() {
  // If there are pending saves (local autosave timers or repo queue), flush
  // them now and show a two-dot indicator. Only shown when the user navigates away.
  const hasPendingLocal = window.AEWTTR._pendingAutosaveFlushers.size > 0;
  const hasPendingRepo = typeof Repo !== "undefined" && Repo.hasPendingChanges();
  if (hasPendingLocal || hasPendingRepo) {
    showSaveIndicator();
    try { if (hasPendingLocal) await flushPendingAutosaves(); } catch (e) {}
    try { if (typeof Repo !== "undefined") await Repo.flush(); } catch (e) {}
    hideSaveIndicator();
  }
  try {
    const pc = document.getElementById("page-content");
    if (pc) pc.classList.remove("meeting-app--live");
    if (typeof applyCuiMarking === "function") applyCuiMarking();
    renderNav();
    scheduleNotificationRefresh();
    maybeProactiveRefreshOnNavigate();
    const { app, parts } = currentRoute();
    if ((app === "admin" || app === "logs") && !canCurrentUserAccessAdmin()) {
      navigate("dashboard");
      toast("That page is only available to app admins.", "error");
      return;
    }
    const fn = PAGE_RENDERERS[app] || (window.PULSE_PORT_CONFIG && PAGE_RENDERERS.travel) || PAGE_RENDERERS.dashboard;
    if (typeof fn !== "function") {
      console.error("PULSE: missing page renderer for", app);
      toast("This page is not available yet.", "error");
      return;
    }
    fn(parts);
  } catch (e) {
    console.error("PULSE: page render failed", e);
    toast("This page failed to load. Try again or choose another section.", "error");
    const mount = document.getElementById("page-content");
    if (mount) {
      mount.innerHTML = `<div class="empty-state" style="padding:40px 24px;max-width:520px;margin:0 auto;">
        <h3 style="margin:0 0 8px;">This page hit an error</h3>
        <p style="color:var(--aewttr-muted);margin:0 0 16px;">You can go back to the dashboard and try again. If this keeps happening, report the message below to your app admin.</p>
        <pre style="background:var(--aewttr-surface-2);padding:12px;border-radius:8px;font-size:12px;white-space:pre-wrap;word-break:break-word;">${escapeHtml(String(e && e.message || e))}</pre>
        <button type="button" class="btn-aewttr" style="margin-top:16px;" onclick="navigate('dashboard')">Go to dashboard</button>
      </div>`;
    }
  }
}

/* ---------- generic kanban board ---------- */
function boardPriorityPill(priority) {
  if (!priority) return "";
  const styles = {
    Low: "#22A06B",
    Medium: "#579BFC",
    High: "#E56910",
    Critical: "#6E5DC6"
  };
  const bg = styles[priority] || styles.Medium;
  return `<span class="monday-label-pill" style="background:${bg}">${escapeHtml(priority)}</span>`;
}

/* opts: { mount, columns:[names], data:{col:[task]}, ownerField:'assignee'|'owner', showPriority:bool, onSave:fn() } */
function buildKanban(opts) {
  const canEditBoard = canCurrentUserEdit();
  function countOf(col) { return (opts.data[col] || []).length; }
  function subtaskBadge(task) {
    if (!task.subtasks || !task.subtasks.length) return "";
    const done = task.subtasks.filter(s => s.done).length;
    return `<span class="kc-badge">${done}/${task.subtasks.length} subtasks</span>`;
  }
  function docsBadge(task) {
    if (!opts.showRelatedDocs || !task.relatedDocs || !task.relatedDocs.length) return "";
    return `<span class="kc-badge"><i class="bx bx-link-alt"></i> ${task.relatedDocs.length} doc${task.relatedDocs.length === 1 ? "" : "s"}</span>`;
  }
  function cardHtml(task, col, idx) {
    const owner = task[opts.ownerField] || "Unassigned";
    const hasDrawer = (task.subtasks && task.subtasks.length) || (opts.showRelatedDocs && task.relatedDocs && task.relatedDocs.length);
    const subCount = (task.subtasks || []).length;
    const monday = !!opts.mondayStyle;
    const customTags = (opts.customFields || [])
      .filter((f) => task[f.key])
      .map((f) => `<span class="kanban-card-field-tag" title="${escapeHtml(f.label)}">${escapeHtml(String(task[f.key]))}</span>`)
      .join("");
    return `
      <div class="kanban-card${monday ? " kanban-card--monday" : ""}" draggable="${canEditBoard ? "true" : "false"}" data-col="${escapeHtml(col)}" data-id="${task.id}">
        ${monday && (task.priority || customTags) ? `<div class="kanban-card-tags">${task.priority ? boardPriorityPill(task.priority) : ""}${customTags}</div>` : ""}
        <div class="ctitle-row">${issueTypeIcon(task)}<span class="ctitle">${escapeHtml(task.title)}</span></div>
        <div class="cmeta">
          <span>${userAvatarHtml(owner, memberEmailForPerson(owner), 21)}${monday ? "" : escapeHtml(owner)}</span>
          ${task.due ? `<span>${fmtDate(task.due)}</span>` : ""}
        </div>
        <div class="cmeta card-footer-row" style="margin-top:8px;">
          ${monday ? `
            <span class="kanban-card-foot-icons">
              ${subCount ? `<span title="${subCount} subitem${subCount === 1 ? "" : "s"}"><i class="bx bx-list-ul"></i> ${subCount}</span>` : ""}
              ${task.linkedTaskId ? `<span title="Linked to tracker task"><i class="bx bx-link"></i></span>` : ""}
            </span>
            <span>${userAvatarHtml(owner, memberEmailForPerson(owner), 24)}</span>
          ` : `
          <span class="issue-key">${(String(opts.keyPrefix || "PULSE").replace(/[^A-Za-z0-9]/g, "").toUpperCase() || "PULSE")}-${issueSeq(task.id)}</span>
          <span style="display:flex;align-items:center;gap:6px;">
            ${opts.showPriority && task.priority ? issuePriority(task.priority) : ""}
            ${subtaskBadge(task)}${docsBadge(task)}
          </span>`}
        </div>
        ${hasDrawer ? `
          <div class="kc-subtasks">
            ${(task.subtasks || []).map((s, si) => `
              <label><input type="checkbox" data-sub="${si}" ${s.done ? "checked" : ""}> <span style="${s.done ? "text-decoration:line-through;color:var(--aewttr-muted);" : ""}">${escapeHtml(s.text)}</span></label>
            `).join("")}
            ${opts.showRelatedDocs && task.relatedDocs && task.relatedDocs.length ? `
              <div class="kc-docs">
                ${task.relatedDocs.map(d => `<div class="kc-doc-item"><i class="bx bx-file"></i> ${escapeHtml(d.fileName)} <span class="kc-doc-loc">${escapeHtml(d.location)}</span></div>`).join("")}
              </div>` : ""}
          </div>` : ""}
        <div class="kc-actions">
          <div class="kc-arrows">
            ${canEditBoard && idx > 0 ? `<button data-act="left"${tip("Move column left")}><i class="bx bx-chevron-left"></i></button>` : ""}
            ${canEditBoard && idx < opts.columns.length - 1 ? `<button data-act="right"${tip("Move column right")}><i class="bx bx-chevron-right"></i></button>` : ""}
          </div>
          ${canEditBoard ? `<button data-act="edit"${tip("Edit this task")}><i class="bx bx-edit"></i></button>
          <button data-act="delete"${tip("Delete this task")}><i class="bx bx-trash"></i></button>` : ""}
        </div>
      </div>`;
  }
  function render() {
    const colColor = (col) => (opts.columnColor ? opts.columnColor(col) : "");
    opts.mount.innerHTML = `
      <div class="kanban-wrap">
        ${opts.columns.map(col => {
          const color = colColor(col);
          const count = countOf(col);
          return `
          <div class="kanban-col${opts.mondayStyle ? " kanban-col--monday" : ""}" data-col="${escapeHtml(col)}" ${color ? `style="--kanban-col-accent:${color}"` : ""}>
            <div class="kanban-col-head${opts.mondayStyle ? " kanban-col-head--monday" : ""}" data-col-head="${escapeHtml(col)}">
              ${color ? `<span class="kanban-col-dot" style="background:${color}"></span>` : ""}
              <h4 contenteditable="${canEditBoard && opts.allowColumnEdit ? "true" : "false"}" spellcheck="false" data-col-rename="${escapeHtml(col)}">${escapeHtml(col)}</h4>
              <span class="kanban-col-count">${count}</span>
              ${canEditBoard && opts.allowColumnEdit ? `<button type="button" class="kanban-col-menu" data-col-menu="${escapeHtml(col)}"${tip("Column options")} aria-label="Column options"><i class="bx bx-dots-horizontal-rounded"></i></button>` : ""}
            </div>
            <div class="kanban-cards" data-col-drop="${escapeHtml(col)}">
              ${(opts.data[col] || []).map((t, i) => cardHtml(t, col, opts.columns.indexOf(col))).join("")}
            </div>
            ${canEditBoard ? `<button class="kanban-add-btn" data-add="${escapeHtml(col)}"${tip(`Add a task to ${col}`)}><i class="bx bx-plus"></i> ${escapeHtml(opts.addButtonLabel || "Add task")}</button>` : ""}
          </div>`;
        }).join("")}
        ${canEditBoard && opts.allowColumnEdit ? `
        <div class="kanban-col kanban-col-add">
          <button type="button" class="kanban-col-add-btn" id="kanban-inline-add-col"><i class="bx bx-plus"></i><span>Add column</span></button>
        </div>` : ""}
      </div>`;
    wire();
  }
  function save(task, col) { aewttrSaveStore(); if (opts.onSave) opts.onSave(task, col); }
  function findTask(col, id) {
    const list = opts.data[col] || [];
    return list.find(t => t.id === id);
  }
  function moveTask(fromCol, id, toCol, toIndex) {
    const list = opts.data[fromCol] || [];
    const i = list.findIndex(t => t.id === id);
    if (i < 0) return;
    const [task] = list.splice(i, 1);
    if (!opts.data[toCol]) opts.data[toCol] = [];
    if (toIndex == null) opts.data[toCol].push(task); else opts.data[toCol].splice(toIndex, 0, task);
  }
  function relatedDocRowHtml(doc) {
    doc = doc || { fileName: "", location: "Windchill" };
    return `<div class="related-doc-row">
      <input class="input-aewttr rd-filename" placeholder="File name" value="${escapeHtml(doc.fileName)}">
      <select class="select-aewttr rd-location">
        <option ${doc.location === "Windchill" ? "selected" : ""}>Windchill</option>
        <option ${doc.location === "SPO" ? "selected" : ""}>SPO</option>
      </select>
      <button type="button" class="btn-aewttr-ghost btn-aewttr-sm rd-remove">&times;</button>
    </div>`;
  }
  function openTaskForm(existing, col) {
    const isEdit = !!existing;
    const ownerLabel = opts.ownerField === "owner" ? "Owner" : "Assignee";
    const modal = openModal(`
      <div class="aewttr-modal-head"><h3>${isEdit ? "Edit Task" : "Add Task"}</h3><button class="aewttr-modal-close">&times;</button></div>
      <div class="aewttr-modal-body">
        <div class="form-row"><label>Title</label><input class="input-aewttr" id="kf-title" value="${existing ? escapeHtml(existing.title) : ""}"></div>
        <div class="form-grid-2">
          <div class="form-row"><label>${ownerLabel}</label><input class="input-aewttr" id="kf-owner" value="${existing ? escapeHtml(existing[opts.ownerField] || "") : ""}"></div>
          <div class="form-row"><label>Due date</label><input type="date" class="input-aewttr" id="kf-due" value="${existing && existing.due ? existing.due : ""}"></div>
        </div>
        ${opts.showPriority ? `
        <div class="form-row"><label>Priority</label>
          <select class="select-aewttr" id="kf-priority">
            ${["Low", "Medium", "High"].map(p => `<option ${existing && existing.priority === p ? "selected" : ""}>${p}</option>`).join("")}
          </select>
        </div>` : ""}
        <div class="form-row"><label>Subtasks (one per line)</label>
          <textarea class="textarea-aewttr" id="kf-subtasks">${existing && existing.subtasks ? existing.subtasks.map(s => s.text).join("\n") : ""}</textarea>
        </div>
        ${opts.showRelatedDocs ? `
        <div class="form-row"><label>Related Document(s)</label>
          <div id="kf-related-docs">${(existing && existing.relatedDocs || []).map(relatedDocRowHtml).join("")}</div>
          <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="kf-add-doc"><i class="bx bx-plus"></i> Add Document</button>
        </div>` : ""}
        ${isEdit ? "" : `<div class="form-row"><label>Column</label><input class="input-aewttr" value="${escapeHtml(col)}" disabled></div>`}
      </div>
      <div class="aewttr-modal-foot">
        <button class="btn-aewttr-ghost" id="kf-cancel">Cancel</button>
        <button class="btn-aewttr" id="kf-save">${isEdit ? "Save Changes" : "Add Task"}</button>
      </div>
    `);
    $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
    $("#kf-cancel", modal).addEventListener("click", closeModal);
    function wireDocRemovers() {
      $all(".rd-remove", modal).forEach(b => b.addEventListener("click", () => b.closest(".related-doc-row").remove()));
    }
    if (opts.showRelatedDocs) {
      wireDocRemovers();
      $("#kf-add-doc", modal).addEventListener("click", () => {
        $("#kf-related-docs", modal).insertAdjacentHTML("beforeend", relatedDocRowHtml());
        wireDocRemovers();
      });
    }
    $("#kf-save", modal).addEventListener("click", () => {
      const title = $("#kf-title", modal).value.trim();
      if (!title) { toast("Title is required", "error"); return; }
      const subtasks = $("#kf-subtasks", modal).value.split("\n").map(s => s.trim()).filter(Boolean)
        .map(text => ({ text, done: existing && existing.subtasks ? (existing.subtasks.find(s => s.text === text) || {}).done || false : false }));
      const payload = {
        title,
        due: $("#kf-due", modal).value || "",
        subtasks
      };
      if (opts.showRelatedDocs) {
        payload.relatedDocs = $all(".related-doc-row", modal)
          .map(row => ({ fileName: $(".rd-filename", row).value.trim(), location: $(".rd-location", row).value }))
          .filter(d => d.fileName);
      }
      payload[opts.ownerField] = $("#kf-owner", modal).value.trim() || "Unassigned";
      if (opts.showPriority) payload.priority = $("#kf-priority", modal).value;
      if (isEdit) {
        Object.assign(existing, payload);
        toast("Task updated", "success");
      } else {
        payload.id = uid("t");
        if (!opts.data[col]) opts.data[col] = [];
        opts.data[col].push(payload);
        toast("Task added", "success");
      }
      save(isEdit ? existing : payload, col);
      closeModal();
      render();
    });
  }
  function wire() {
    if (!canEditBoard) return;
    $all(".kanban-card", opts.mount).forEach(card => {
      const col = card.dataset.col, id = card.dataset.id;
      card.addEventListener("click", (e) => {
        if (e.target.closest("button") || e.target.tagName === "INPUT") return;
        card.classList.toggle("expanded");
      });
      // Deferred via setTimeout: adding the low-opacity "dragging" class
      // synchronously in dragstart can make Chrome capture the native drag
      // ghost image already-faded (or blank) — the cursor then drags around
      // with no visible card under it. Applying it a tick later lets the
      // browser snapshot the full-opacity card first.
      card.addEventListener("dragstart", () => { setTimeout(() => card.classList.add("dragging"), 0); });
      card.addEventListener("dragend", () => { card.classList.remove("dragging"); render(); });
      // Right-click parity with the table view — cards already support
      // drag and the left/right arrow buttons to move between columns, but
      // had no contextmenu handler at all, so right-clicking here just
      // showed the browser's default menu instead of anything useful.
      card.addEventListener("contextmenu", (e) => {
        if (typeof openBoardContextMenu !== "function") return;
        e.preventDefault();
        const task = findTask(col, id);
        if (!task) return;
        const idx = opts.columns.indexOf(col);
        const items = [
          { label: "Edit task", onClick: () => openTaskForm(task, col) }
        ];
        opts.columns.forEach((targetCol, targetIdx) => {
          if (targetIdx === idx) return;
          items.push({ label: `Move to “${targetCol}”`, onClick: () => { moveTask(col, id, targetCol); save(task, targetCol); render(); } });
        });
        items.push({ separator: true });
        items.push({ label: "Delete task", danger: true, onClick: async () => {
          const ok = await confirmDialog({ title: "Delete task", message: "Delete this task?", confirmLabel: "Delete", danger: true });
          if (!ok) return;
          opts.data[col] = opts.data[col].filter((t) => t.id !== id);
          aewttrSaveStore();
          if (opts.onDelete) opts.onDelete(task, col);
          render();
          toast("Task deleted", "success");
        }});
        openBoardContextMenu(e.clientX, e.clientY, items);
      });
      $all("input[type=checkbox]", card).forEach(cb => {
        cb.addEventListener("click", (e) => e.stopPropagation());
        cb.addEventListener("change", () => {
          const task = findTask(col, id);
          task.subtasks[+cb.dataset.sub].done = cb.checked;
          save(task, col);
        });
      });
      const btn = (act) => card.querySelector(`[data-act="${act}"]`);
      if (btn("left")) btn("left").addEventListener("click", (e) => {
        e.stopPropagation();
        const idx = opts.columns.indexOf(col);
        const toCol = opts.columns[idx - 1];
        const task = findTask(col, id);
        moveTask(col, id, toCol);
        save(task, toCol); render();
        toast("Moved to " + toCol);
      });
      if (btn("right")) btn("right").addEventListener("click", (e) => {
        e.stopPropagation();
        const idx = opts.columns.indexOf(col);
        const toCol = opts.columns[idx + 1];
        const task = findTask(col, id);
        moveTask(col, id, toCol);
        save(task, toCol); render();
        toast("Moved to " + toCol);
      });
      btn("edit").addEventListener("click", (e) => { e.stopPropagation(); openTaskForm(findTask(col, id), col); });
      btn("delete").addEventListener("click", async (e) => {
        e.stopPropagation();
        const task = findTask(col, id);
        const ok = await confirmDialog({
          title: "Delete task",
          message: "Delete this task?",
          confirmLabel: "Delete",
          danger: true
        });
        if (!ok) return;
        opts.data[col] = opts.data[col].filter(t => t.id !== id);
        aewttrSaveStore();
        if (opts.onDelete) opts.onDelete(task, col);
        render();
        toast("Task deleted", "success");
      });
    });
    $all(".kanban-cards", opts.mount).forEach(dropzone => {
      dropzone.addEventListener("dragover", (e) => { e.preventDefault(); dropzone.parentElement.classList.add("drag-over"); });
      dropzone.addEventListener("dragleave", () => dropzone.parentElement.classList.remove("drag-over"));
      dropzone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropzone.parentElement.classList.remove("drag-over");
        const dragging = $(".dragging", opts.mount);
        if (!dragging) return;
        const fromCol = dragging.dataset.col, id = dragging.dataset.id;
        const toCol = dropzone.dataset.colDrop;
        const task = findTask(fromCol, id);
        moveTask(fromCol, id, toCol);
        save(task, toCol); render();
      });
    });
    $all("[data-add]", opts.mount).forEach(b => b.addEventListener("click", () => openTaskForm(null, b.dataset.add)));
    $all("[data-col-rename]", opts.mount).forEach((node) => {
      const oldName = node.dataset.colRename;
      node.addEventListener("keydown", (e) => {
        if (e.key === "Enter") { e.preventDefault(); node.blur(); }
      });
      node.addEventListener("blur", () => {
        const next = node.textContent.trim();
        if (!next || next === oldName) { node.textContent = oldName; return; }
        if (opts.onRenameColumn && opts.onRenameColumn(oldName, next)) {
          if (opts.onColumnsChange) opts.onColumnsChange();
          else render();
        } else {
          node.textContent = oldName;
        }
      });
    });
    $all("[data-col-menu]", opts.mount).forEach((btn) => btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const col = btn.dataset.colMenu;
      const rect = btn.getBoundingClientRect();
      if (typeof openBoardContextMenu !== "function") return;
      openBoardContextMenu(rect.left, rect.bottom + 4, [
        { label: "Rename column", onClick: async () => {
          const next = typeof promptBoardColumnName === "function"
            ? await promptBoardColumnName("Rename column", col)
            : prompt("Rename column", col);
          if (next && opts.onRenameColumn && opts.onRenameColumn(col, next)) {
            if (opts.onColumnsChange) opts.onColumnsChange();
            else render();
          }
        }},
        { label: "Delete column", danger: true, onClick: async () => {
          if (opts.onRemoveColumn && await opts.onRemoveColumn(col)) {
            if (opts.onColumnsChange) opts.onColumnsChange();
            else render();
          }
        }},
        { separator: true },
        { label: "Add column", onClick: async () => {
          const name = typeof promptBoardColumnName === "function"
            ? await promptBoardColumnName("Add column")
            : prompt("New column name");
          if (name && opts.onAddColumn && opts.onAddColumn(name)) {
            if (opts.onColumnsChange) opts.onColumnsChange();
            else render();
          }
        }}
      ]);
    }));
    opts.mount.addEventListener("contextmenu", (e) => {
      const head = e.target.closest("[data-col-head]");
      const card = e.target.closest(".kanban-card[data-id]");
      if (!head && !card) return;
      e.preventDefault();
      if (typeof openBoardContextMenu !== "function") return;
      if (card) {
        const fromCol = card.dataset.col;
        const id = card.dataset.id;
        const moveItems = (opts.columns || []).filter((c) => c !== fromCol).map((toCol) => ({
          label: `Move to ${toCol}`,
          onClick: () => {
            moveTask(fromCol, id, toCol);
            save(findTask(toCol, id), toCol);
            render();
          }
        }));
        openBoardContextMenu(e.clientX, e.clientY, [
          { label: "Edit item", onClick: () => {
            if (opts.onOpenItem) opts.onOpenItem(id);
            else openTaskForm(findTask(fromCol, id), fromCol);
          }},
          ...moveItems,
          { label: "Delete item", danger: true, onClick: async () => {
            const task = findTask(fromCol, id);
            const ok = await confirmDialog({ title: "Delete item", message: "Delete this item?", confirmLabel: "Delete", danger: true });
            if (!ok) return;
            opts.data[fromCol] = (opts.data[fromCol] || []).filter((t) => t.id !== id);
            save(task, fromCol);
            render();
          }}
        ]);
        return;
      }
      const col = head.dataset.colHead;
      openBoardContextMenu(e.clientX, e.clientY, [
        { label: "Rename column", onClick: async () => {
          const next = typeof promptBoardColumnName === "function"
            ? await promptBoardColumnName("Rename column", col)
            : prompt("Rename column", col);
          if (next && opts.onRenameColumn && opts.onRenameColumn(col, next)) {
            if (opts.onColumnsChange) opts.onColumnsChange();
            else render();
          }
        }},
        { label: "Delete column", danger: true, onClick: async () => {
          if (opts.onRemoveColumn && await opts.onRemoveColumn(col)) {
            if (opts.onColumnsChange) opts.onColumnsChange();
            else render();
          }
        }},
        { separator: true },
        { label: "Add column", onClick: async () => {
          const name = typeof promptBoardColumnName === "function"
            ? await promptBoardColumnName("Add column")
            : prompt("New column name");
          if (name && opts.onAddColumn && opts.onAddColumn(name)) {
            if (opts.onColumnsChange) opts.onColumnsChange();
            else render();
          }
        }}
      ]);
    });
    const addColBtn = $("#kanban-inline-add-col", opts.mount);
    if (addColBtn && opts.onAddColumn) {
      addColBtn.addEventListener("click", async () => {
        const name = typeof promptBoardColumnName === "function"
          ? await promptBoardColumnName("Add column")
          : prompt("New column name");
        if (!name) return;
        if (opts.onAddColumn(name.trim())) {
          if (opts.onColumnsChange) opts.onColumnsChange();
          else render();
        }
      });
    }
  }
  render();
  return { render };
}

/* ---------- Monday.com-style item table ----------
   Shared between the project Tracker's "Table" view and custom Board
   "Table" boards: a flat list of items with a colored status cell (and an
   optional second colored column — health for Tracker, priority for
   Boards), an owner avatar, a due date, and click-to-expand indented
   subitems. Reuses the same status-X / health-X / priority-X color classes
   already defined for pills elsewhere in the app, just applied as solid
   cell backgrounds instead of small tinted tags, so the look stays
   consistent with the rest of PULSE while reading like a Monday board. */
function renderMondayTable(mount, opts) {
  const canEdit = opts.canEdit !== false;
  const isTracker = opts.mode === "tracker";
  const extraColumns = opts.extraColumns || [];
  const showMilestone = isTracker && opts.milestoneColumn !== false;
  const colCount = isTracker
    ? (11 + extraColumns.length)
    : (4 + extraColumns.length + (opts.showSubitemsColumn === false ? 0 : 1));
  const memberListId = opts.memberListId || `monday-members-${Math.random().toString(36).slice(2, 8)}`;
  const members = (window.AEWTTR.db && window.AEWTTR.db.members) || [];

  function milestoneToggleHtml(divider) {
    if (!showMilestone || !divider) return "";
    const on = !!divider.isMilestone;
    return `<button type="button" class="monday-milestone-toggle${on ? " is-on" : ""}" data-milestone-toggle="${escapeHtml(divider.id)}"${tip(on ? "Clear milestone flag on this divider" : "Flag divider as milestone for status export")} aria-pressed="${on ? "true" : "false"}" aria-label="Milestone"${canEdit ? "" : " disabled"}>
      <i class="bx ${on ? "bxs-flag" : "bx-flag"}" aria-hidden="true"></i>
    </button>`;
  }

  function notesButtonHtml(item, parentId, subIndex) {
    if (!isTracker) return "";
    const notes = item.notes || [];
    const btnId = parentId != null ? `${parentId}:${subIndex}` : item.id;
    const count = notes.length;
    return `<button type="button" class="monday-notes-btn" data-open-notes="${escapeHtml(btnId)}"${tip(count ? `${count} note${count === 1 ? "" : "s"} — open thread` : "Add notes")}>
      <i class="bx bx-message-rounded-dots"></i>
      ${count ? `<span class="monday-notes-btn-count">${count}</span>` : ""}
    </button>`;
  }

  function inlineTextInput(task, field, subtask, subPath) {
    const value = subtask != null ? (subtask[field] || "") : (task[field] || "");
    const subAttrs = subtask != null
      ? ` data-sub-text-field="${field}" data-sub-path="${escapeHtml(String(subPath))}"`
      : ` data-text-field="${field}"`;
    const listAttr = (field === "assignee" || field === "owner") ? ` list="${memberListId}"` : "";
    return `<input type="text" class="monday-inline-text"${subAttrs} data-id="${task.id}" value="${escapeHtml(value)}"${listAttr} ${canEdit ? "" : "disabled"}>`;
  }
  function ownerCell(task) {
    if (isTracker || opts.inlineOwner !== false) return inlineTextInput(task, opts.ownerField || "assignee");
    const owner = task[opts.ownerField] || "Unassigned";
    return `<span class="monday-owner-cell">${userAvatarHtml(owner, memberEmailForPerson(owner), 22)}<span>${escapeHtml(owner)}</span></span>`;
  }
  function extraColumnCell(task, col) {
    // Boards' custom fields can be type "text" (a plain inline input, no
    // fixed option list) alongside the built-in select-type columns
    // (Priority, Status). Tracker/legacy callers never set col.type, so
    // this stays a no-op for them — falls straight through to selectCell.
    return col.type === "text" ? inlineTextInput(task, col.key) : selectCell(task, col);
  }
  function selectCell(task, col) {
    const value = task[col.key] || col.options[0];
    const colorClass = `${col.colorPrefix || col.key}-${String(value).replace(/\s+/g, "-")}`;
    const customColor = col.colorFor ? col.colorFor(value) : "";
    const style = customColor ? ` style="background-color:${customColor}"` : "";
    return `<select class="monday-status-select ${customColor ? "monday-status-custom" : colorClass}" data-field="${col.key}" data-color-prefix="${col.colorPrefix || col.key}" data-id="${task.id}" data-custom-color="${customColor ? "1" : ""}"${style} ${canEdit ? "" : "disabled"} aria-label="${escapeHtml(col.label)}">
      ${col.options.map(o => `<option value="${escapeHtml(o)}" ${o === value ? "selected" : ""}>${escapeHtml(o)}</option>`).join("")}
    </select>`;
  }
  function resolveTask(id) {
    const pools = [opts.allTasks, opts.tasks].filter(Boolean);
    for (let i = 0; i < pools.length; i++) {
      const found = pools[i].find((t) => t.id === id);
      if (found) return found;
    }
    return null;
  }
  function getSubByPath(task, pathStr) {
    if (!task || pathStr == null || pathStr === "") return null;
    if (typeof getSubtaskAtPath === "function") return getSubtaskAtPath(task, pathStr);
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
  function subSelectCell(task, col, subtask, subPath) {
    const value = (subtask && subtask[col.key]) || task[col.key] || col.options[0];
    const colorClass = `${col.colorPrefix || col.key}-${String(value).replace(/\s+/g, "-")}`;
    return `<select class="monday-status-select ${colorClass}" data-sub-field="${col.key}" data-id="${task.id}" data-sub-path="${escapeHtml(String(subPath))}" ${canEdit ? "" : "disabled"} aria-label="${escapeHtml(col.label)}">
      ${col.options.map(o => `<option value="${escapeHtml(o)}" ${o === value ? "selected" : ""}>${escapeHtml(o)}</option>`).join("")}
    </select>`;
  }
  function dateInputCell(task, field, subtask, subPath) {
    const value = subtask ? (subtask[field] || "") : (task[field] || "");
    if (subtask != null) {
      return `<input type="date" class="monday-date-input" data-date-field="${field}" data-id="${task.id}" data-sub-path="${escapeHtml(String(subPath))}" value="${value}" ${canEdit ? "" : "disabled"}>`;
    }
    return `<input type="date" class="monday-date-input" data-date-field="${field}" data-id="${task.id}" value="${value}" ${canEdit ? "" : "disabled"}>`;
  }
  function subitemsCell(task) {
    const flat = typeof flattenTaskSubitems === "function" ? flattenTaskSubitems(task.subtasks || []) : (task.subtasks || []);
    if (!flat.length) return `<span class="monday-subtask-empty">—</span>`;
    const done = flat.filter((s) => s.done).length;
    return `<span class="monday-subtask-chip ${done === flat.length ? "all-done" : ""}">${done}/${flat.length}</span>`;
  }
  function nestedSubRows(task, list, pathPrefix, depth) {
    const healthCol = opts.healthColumn;
    const expandedSubs = opts.subitemExpanded || {};
    let html = "";
    (list || []).forEach((s, si) => {
      const path = pathPrefix.length ? `${pathPrefix}.${si}` : String(si);
      const children = s.subtasks || [];
      const hasKids = children.length > 0;
      const open = hasKids && !!expandedSubs[`${task.id}:${path}`];
    const indent = 28 + depth * 22;
      if (isTracker) {
        // Inline rows matching main tracker column structure
        html += `
          <tr class="monday-row monday-row--tracker monday-sub-row-item monday-sub-depth-${Math.min(depth, 8)}${s.done ? " is-complete" : ""}" data-sub-row="1" data-task-id="${task.id}" data-sub-path="${escapeHtml(path)}" data-id="${task.id}">
            <td class="mgp-num-col"></td>
            <td class="monday-drag-cell">${canEdit ? `<button type="button" class="cl-drag-handle" aria-label="Drag to reorder"><i class="bx bx-grid-vertical"></i></button>` : ""}</td>
            <td class="monday-complete-col"><button type="button" class="tracker-complete-btn${s.done ? " is-done" : ""}" data-subtask-complete="${task.id}:${escapeHtml(path)}" aria-label="${s.done ? "Mark subtask as not done" : "Mark subtask done"}" aria-pressed="${s.done ? "true" : "false"}" ${canEdit ? "" : "disabled"}><i class="bx ${s.done ? "bxs-check-circle" : "bx-circle"}"></i></button></td>
            <td class="monday-expand-cell">${hasKids ? `<button type="button" class="monday-expand-btn monday-expand-btn--nested" data-toggle-sub="${task.id}:${escapeHtml(path)}" aria-label="${open ? "Collapse" : "Expand"}"><i class="bx bx-chevron-${open ? "down" : "right"}"></i></button>` : ""}</td>
            <td class="monday-item-cell monday-subitem-cell" style="padding-left:${indent}px"><span class="monday-sub-nest-arrow" aria-hidden="true"><i class="bx bx-subdirectory-right"></i></span>${inlineTextInput(task, "text", s, path)}</td>
            <td class="monday-owner-td">${inlineTextInput(task, "assignee", s, path)}</td>
            <td class="monday-date-td">${dateInputCell(task, opts.startField || "start", s, path)}</td>
            <td class="monday-date-td">${dateInputCell(task, opts.endField || "end", s, path)}</td>
            ${healthCol ? `<td class="monday-select-cell">${subSelectCell(task, healthCol, s, path)}</td>` : ""}
            ${extraColumns.map(() => `<td></td>`).join("")}
            <td class="monday-notes-td">
              <div class="monday-notes-cell">
                ${notesButtonHtml(s, task.id, path)}
                ${canEdit ? `<button type="button" class="monday-add-sub-btn" data-add-subtask="${task.id}" data-parent-path="${escapeHtml(path)}" aria-label="Add subtask" title="Add subtask"><i class="bx bx-list-plus"></i></button>` : ""}
                ${canEdit ? `<button type="button" class="monday-row-delete-btn" data-delete-subtask="${task.id}:${escapeHtml(path)}" aria-label="Delete subtask" title="Delete subtask"><i class="bx bx-trash"></i></button>` : ""}
              </div>
            </td>
          </tr>`;
        if (canEdit) html += `<tr class="subtask-insert-row" data-parent-task-id="${task.id}" data-parent-path="${escapeHtml(path)}"><td colspan="${colCount}"><div class="subtask-insert-wrap" style="--sub-insert-indent:${indent}px"><div class="subtask-insert-line"></div><button type="button" class="subtask-insert-btn" data-add-subtask-sibling="${task.id}" data-sibling-path="${escapeHtml(path)}" aria-label="Add a subtask at this level" title="Add a subtask at this level"><i class="bx bx-plus"></i></button></div></td></tr>`;
      } else {
        html += `
          <tr class="monday-sub-row-item monday-sub-depth-${Math.min(depth, 8)}" data-sub-row="1" data-task-id="${task.id}" data-sub-path="${escapeHtml(path)}" style="--sub-indent:${depth * 18}px;">
            <td class="monday-drag-cell">${canEdit ? `<button type="button" class="cl-drag-handle" aria-label="Drag to reorder"><i class="bx bx-grid-vertical"></i></button>` : ""}</td>
            <td class="monday-done-col"><button type="button" class="tracker-complete-btn${s.done ? " is-done" : ""}" data-subtask-complete="${task.id}:${escapeHtml(path)}" aria-label="${s.done ? "Mark subtask as not done" : "Mark subtask done"}" aria-pressed="${s.done ? "true" : "false"}" ${canEdit ? "" : "disabled"}><i class="bx ${s.done ? "bxs-check-circle" : "bx-circle"}"></i></button></td>
            <td class="monday-expand-cell">${hasKids ? `<button type="button" class="monday-expand-btn monday-expand-btn--nested" data-toggle-sub="${task.id}:${escapeHtml(path)}" aria-label="${open ? "Collapse" : "Expand"}"><i class="bx bx-chevron-${open ? "down" : "right"}"></i></button>` : (depth > 0 ? `<span class="monday-sub-nest-spacer" aria-hidden="true"></span>` : "")}</td>
            <td class="monday-sub-title-cell"><span class="monday-sub-nest-arrow" aria-hidden="true"><i class="bx bx-subdirectory-right"></i></span>${inlineTextInput(task, "text", s, path)}</td>
            <td>${inlineTextInput(task, "assignee", s, path)}</td>
            <td class="monday-date-td">${dateInputCell(task, opts.startField || "start", s, path)}</td>
            <td class="monday-date-td">${dateInputCell(task, opts.endField || "end", s, path)}</td>
            ${healthCol ? `<td class="monday-select-cell">${subSelectCell(task, healthCol, s, path)}</td>` : ""}
            <td class="monday-notes-td">
              <div class="monday-notes-cell">
                ${notesButtonHtml(s, task.id, path)}
                ${canEdit ? `<button type="button" class="monday-add-sub-btn" data-add-subtask="${task.id}" data-parent-path="${escapeHtml(path)}" aria-label="Add subtask" title="Add subtask"><i class="bx bx-list-plus"></i></button>` : ""}
                ${canEdit ? `<button type="button" class="monday-row-delete-btn" data-delete-subtask="${task.id}:${escapeHtml(path)}" aria-label="Delete subtask" title="Delete subtask"><i class="bx bx-trash"></i></button>` : ""}
              </div>
            </td>
          </tr>`;
        if (canEdit) html += `<tr class="subtask-insert-row" data-parent-task-id="${task.id}" data-parent-path="${escapeHtml(path)}"><td colspan="${colCount}"><div class="subtask-insert-wrap" style="--sub-insert-indent:${indent}px"><div class="subtask-insert-line"></div><button type="button" class="subtask-insert-btn" data-add-subtask-sibling="${task.id}" data-sibling-path="${escapeHtml(path)}" aria-label="Add a subtask at this level" title="Add a subtask at this level"><i class="bx bx-plus"></i></button></div></td></tr>`;
      }
      if (open) html += nestedSubRows(task, children, path, depth + 1);
    });
    return html;
  }
  let tableRowNum = 0;
  function rowHtml(task) {
    tableRowNum++;
    const rowNum = tableRowNum;
    const isExpanded = !!opts.expanded[task.id];
    const subtasks = task.subtasks || [];
    const flatCount = typeof flattenTaskSubitems === "function" ? flattenTaskSubitems(subtasks).length : subtasks.length;
    const trackerCells = isTracker ? `
        <td class="monday-date-td">${dateInputCell(task, opts.startField || "start")}</td>
        <td class="monday-date-td">${dateInputCell(task, opts.endField || "end")}</td>
        ${opts.healthColumn ? `<td class="monday-select-cell">${selectCell(task, opts.healthColumn)}</td>` : ""}
        ${extraColumns.map((col) => `<td class="monday-select-cell">${extraColumnCell(task, col)}</td>`).join("")}
        <td class="monday-notes-td">
          <div class="monday-notes-cell">
            ${notesButtonHtml(task)}
            ${canEdit ? `<button type="button" class="monday-add-sub-btn" data-add-subtask="${task.id}" data-parent-path="" aria-label="Add subtask" title="Add subtask"><i class="bx bx-list-plus"></i></button>` : ""}
            ${canEdit ? `<button type="button" class="monday-row-delete-btn" data-delete-task="${task.id}" aria-label="Delete task" title="Delete task"><i class="bx bx-trash"></i></button>` : ""}
          </div>
        </td>
      ` : `
        <td class="monday-select-cell">${selectCell(task, opts.statusColumn)}</td>
        ${extraColumns.map((col) => `<td class="monday-select-cell">${extraColumnCell(task, col)}</td>`).join("")}
        <td class="monday-due-cell">${task[opts.dueField] ? fmtDate(task[opts.dueField]) : "—"}</td>
        ${opts.showSubitemsColumn === false ? "" : `<td class="monday-subtask-cell">${subitemsCell(task)}</td>`}
      `;
    return `
      <tr class="monday-row${isTracker ? " monday-row--tracker" : ""}${isTracker && taskIsFullyDone(task) ? " is-complete" : ""}" data-id="${task.id}"${isTracker ? ` data-row-id="${task.id}"` : ""}>
        ${isTracker ? `<td class="mgp-num-col">${rowNum}</td><td class="monday-drag-cell">${canEdit ? `<button type="button" class="cl-drag-handle" aria-label="Drag to reorder"><i class="bx bx-grid-vertical"></i></button>` : ""}</td>` : ""}
        ${isTracker ? `<td class="monday-complete-col"><button type="button" class="tracker-complete-btn${taskIsFullyDone(task) ? " is-done" : ""}" data-task-complete="${task.id}" aria-label="${taskIsFullyDone(task) ? "Mark as not done" : "Mark task done"}" aria-pressed="${taskIsFullyDone(task) ? "true" : "false"}"><i class="bx ${taskIsFullyDone(task) ? "bxs-check-circle" : "bx-circle"}"></i></button></td>` : ""}
        <td class="monday-expand-cell">${flatCount ? `<button type="button" class="monday-expand-btn" data-toggle="${task.id}" aria-label="${isExpanded ? "Collapse subitems" : "Expand subitems"}"><i class="bx bx-chevron-${isExpanded ? "down" : "right"}"></i></button>` : ""}</td>
        <td class="monday-item-cell">${isTracker ? inlineTextInput(task, "title") : `<button type="button" class="monday-item-title" data-open="${task.id}">${escapeHtml(task.title)}</button>`}</td>
        <td class="monday-owner-td">${ownerCell(task)}</td>
        ${trackerCells}
      </tr>
      ${isTracker && isExpanded ? nestedSubRows(task, subtasks, "", 0) : ""}
      ${!isTracker && isExpanded ? `
      <tr class="monday-subrow"><td colspan="${colCount}">
        <table class="monday-subtable">
          <thead><tr><th></th><th>Owner</th><th>Status</th><th>Due</th><th></th></tr></thead>
          <tbody>
            ${subtasks.map((s, si) => `
              <tr>
                <td><input type="checkbox" data-subtask-done="${task.id}:${si}" ${s.done ? "checked" : ""} aria-label="Mark subitem done" ${canEdit ? "" : "disabled"}></td>
                <td>${escapeHtml(s.assignee || "Unassigned")}</td>
                <td><span class="status-pill ${s.done ? "status-Done" : "status-In-Progress"}">${s.done ? "Done" : "Open"}</span></td>
                <td>${(s.end || task[opts.dueField]) ? fmtDate(s.end || task[opts.dueField]) : "—"}</td>
                <td><button type="button" class="monday-subtask-edit" data-open-sub="${task.id}:${si}" aria-label="Full subitem editor"><i class="bx bx-edit"></i></button></td>
              </tr>`).join("")}
          </tbody>
        </table>
      </td></tr>` : ""}
    `;
  }

  function dividerSectionRow(section) {
    const d = section.divider;
    const collapsed = !!section.collapsed;
    const count = (section.tasks || []).length;
    const dividerTasks = section.tasks || [];
    const dividerUnits = dividerTasks.flatMap((task) => {
      const subs = flattenTaskSubitems(task.subtasks || []);
      return subs.length ? subs : [task];
    });
    const dividerDone = dividerUnits.filter((item) => item.done || item.status === "Done").length;
    const dividerPct = dividerUnits.length ? Math.round((dividerDone / dividerUnits.length) * 100) : 0;
    const milestoneClass = d.isMilestone ? " is-milestone" : "";
    const noteCount = Array.isArray(d.notes) ? d.notes.length : 0;
    return `
      <tr class="monday-group-row monday-divider-row${milestoneClass}" data-divider-id="${escapeHtml(d.id)}">
        <td colspan="${colCount}" class="tracker-divider-full-td">
          <div class="tracker-divider-block${milestoneClass}">
            <div class="monday-group-head tracker-divider-head${milestoneClass}">
              ${canEdit ? `<button type="button" class="cl-drag-handle tracker-divider-drag" aria-label="Drag to reorder divider"${tip("Drag to reorder divider")}><i class="bx bx-grid-vertical"></i></button>` : ""}
              <button type="button" class="tracker-divider-toggle" data-toggle-divider="${escapeHtml(d.id)}" aria-expanded="${collapsed ? "false" : "true"}"${tip(collapsed ? "Expand divider" : "Collapse divider")}>
                <i class="bx bx-chevron-${collapsed ? "right" : "down"} monday-group-chevron"></i>
              </button>
              <div class="tracker-divider-copy">
                <strong class="tracker-divider-title">${escapeHtml(d.title || "Untitled divider")}</strong>
                <span class="tracker-divider-progress">${dividerPct}% complete</span>
                ${d.isMilestone ? `<span class="tracker-divider-milestone tracker-divider-milestone--compact"><i class="bx bxs-flag" aria-hidden="true"></i></span>` : ""}
                <span class="monday-group-count">${count}</span>
              </div>
              <div class="tracker-divider-actions">
                ${showMilestone ? milestoneToggleHtml(d) : ""}
                ${canEdit ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-sm tracker-divider-add-task" data-add-under-divider="${escapeHtml(d.id)}"${tip("Add task under this divider")}><i class="bx bx-plus"></i></button>` : ""}
                ${canEdit ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-edit-divider="${escapeHtml(d.id)}"${tip("Edit divider fields")}><i class="bx bx-cog"></i></button>` : ""}
                ${canEdit && String(d.title || "").trim().toLowerCase() !== "general" ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-delete-divider="${escapeHtml(d.id)}"${tip("Delete divider (tasks move to General)")}><i class="bx bx-trash"></i></button>` : ""}
                <button type="button" class="monday-notes-btn tracker-divider-notes-btn${noteCount ? " has-notes" : ""}" data-open-notes="${escapeHtml(d.id)}"${tip(noteCount ? `${noteCount} note${noteCount === 1 ? "" : "s"} — open divider thread` : "Add notes for this divider")}>
                  <i class="bx bx-note"></i>
                  ${noteCount ? `<span class="monday-notes-btn-count">${noteCount}</span>` : ""}
                </button>
              </div>
            </div>
          </div>
        </td>
      </tr>`;
  }

  function rowInsertHtml(dividerId, afterTaskId) {
    if (!canEdit) return "";
    return `<tr class="row-insert-row">
      <td colspan="${colCount}">
        <div class="row-insert-wrap">
          <div class="row-insert-line"></div>
          <button type="button" class="row-insert-btn" data-add-under-divider="${escapeHtml(dividerId)}"${afterTaskId ? ` data-insert-after="${escapeHtml(afterTaskId)}"` : ""} aria-label="Add task here"><i class="bx bx-plus"></i></button>
        </div>
      </td>
    </tr>`;
  }

  function ungroupedSectionRow(section) {
    if (opts.hideGroupHeader) return "";
    const count = (section.tasks || []).length;
    return `
      <tr class="monday-group-row monday-ungrouped-row">
        <td colspan="${colCount}">
          <div class="monday-group-head tracker-divider-head tracker-divider-head--ungrouped tracker-divider-head--static">
            <div class="tracker-divider-copy">
              <strong>${escapeHtml(section.label || "General")}</strong>
              <span class="monday-group-count">${count} task${count === 1 ? "" : "s"}</span>
            </div>
          </div>
        </td>
      </tr>`;
  }

  const sections = Array.isArray(opts.sections) ? opts.sections : null;
  const flatTasks = sections
    ? sections.reduce((acc, s) => acc.concat(s.tasks || []), [])
    : (opts.tasks || []);
  if (!flatTasks.length && !(sections && sections.length) && !opts.inlineAddLabel && !opts.onAddDivider) {
    mount.innerHTML = `<div class="empty-state" style="padding:30px;">${opts.emptyMessage || "No items yet."}</div>`;
    return;
  }

  const inlineAddRow = opts.inlineAddLabel && canEdit && !sections ? `
    <tr class="monday-inline-add-row">
      <td colspan="${colCount}">
        <button type="button" class="monday-inline-add-btn" id="monday-inline-add">${escapeHtml(opts.inlineAddLabel)}</button>
      </td>
    </tr>` : "";

  const footerActions = isTracker ? "" : inlineAddRow;

  const groupRow = (!sections && opts.groupHeader) ? `
    <tr class="monday-group-row">
      <td colspan="${colCount}">
        <div class="monday-group-head">
          <span class="monday-group-chevron"><i class="bx bx-chevron-down"></i></span>
          <strong>${escapeHtml(opts.groupHeader.label)}</strong>
          <span class="monday-group-count">${opts.groupHeader.count || opts.tasks.length} item${(opts.groupHeader.count || opts.tasks.length) === 1 ? "" : "s"}</span>
        </div>
      </td>
    </tr>` : "";

  let bodyRows = "";
  if (sections && sections.length) {
    sections.forEach((section) => {
      if (section.kind === "divider") {
        bodyRows += dividerSectionRow(section);
        if (!section.collapsed) {
          const sTasks = section.tasks || [];
          sTasks.forEach((t) => { bodyRows += rowHtml(t); bodyRows += rowInsertHtml(section.divider.id, t.id); });
          if (!sTasks.length) bodyRows += rowInsertHtml(section.divider.id);
        }
      } else {
        bodyRows += ungroupedSectionRow(section);
        if (!section.collapsed) {
          const sTasks = section.tasks || [];
          sTasks.forEach((t) => { bodyRows += rowHtml(t); bodyRows += rowInsertHtml("", t.id); });
          if (!sTasks.length) bodyRows += rowInsertHtml("");
        }
      }
    });
  } else {
    bodyRows = flatTasks.map(rowHtml).join("");
  }

  const thead = isTracker ? `
          <th class="mgp-num-col mgp-num-col--head">#</th>
          <th class="monday-drag-cell" aria-label="Reorder"></th>
          <th class="monday-complete-col" aria-label="Done"></th>
          <th class="monday-expand-cell"></th>
          <th>Item</th>
          <th>Owner</th>
          <th>Start</th>
          <th>End</th>
          <th>${escapeHtml((opts.healthColumn && opts.healthColumn.label) || "Health")}</th>
          ${extraColumns.map((c) => `<th>${escapeHtml(c.label)}</th>`).join("")}
          <th>Actions</th>
        ` : `
          <th class="monday-expand-cell"></th>
          <th>Item</th>
          <th>Owner</th>
          <th>${escapeHtml(opts.statusColumn.label)}</th>
          ${extraColumns.map((c) => `<th>${escapeHtml(c.label)}</th>`).join("")}
          <th>Due</th>
          ${opts.showSubitemsColumn === false ? "" : "<th>Subitems</th>"}
        `;

  mount.innerHTML = `
    <div class="monday-table-wrap">
      ${isTracker && members.length ? `<datalist id="${memberListId}">${members.map((m) => `<option value="${escapeHtml(m.name)}">`).join("")}</datalist>` : ""}
      <table class="monday-table${isTracker ? " monday-table--tracker" : ""}">
        ${opts.hideHeader ? "" : `<thead><tr>${thead}</tr></thead>`}
        <tbody>${groupRow}${bodyRows}${isTracker ? footerActions : inlineAddRow}</tbody>
      </table>
      ${!flatTasks.length && opts.emptyMessage ? `<div class="monday-table-empty">${opts.emptyMessage}</div>` : ""}
    </div>
  `;

  // Keep opts.tasks usable for lookup when sections are used.
  if (sections) opts.tasks = flatTasks;

  $all("[data-toggle]", mount).forEach((btn) => btn.addEventListener("click", () => opts.onToggleExpand(btn.dataset.toggle)));
  $all("[data-toggle-sub]", mount).forEach((btn) => btn.addEventListener("click", () => {
    if (typeof opts.onToggleSubitemExpand === "function") opts.onToggleSubitemExpand(btn.dataset.toggleSub);
  }));
  $all("[data-toggle-divider]", mount).forEach((btn) => btn.addEventListener("click", () => {
    if (typeof opts.onToggleDivider === "function") opts.onToggleDivider(btn.dataset.toggleDivider);
  }));
  $all("[data-edit-divider]", mount).forEach((btn) => btn.addEventListener("click", () => {
    if (typeof opts.onEditDivider === "function") opts.onEditDivider(btn.dataset.editDivider);
  }));
  $all("[data-delete-divider]", mount).forEach((btn) => btn.addEventListener("click", () => {
    if (typeof opts.onDeleteDivider === "function") opts.onDeleteDivider(btn.dataset.deleteDivider);
  }));
  $all("[data-add-under-divider]", mount).forEach((btn) => btn.addEventListener("click", () => {
    if (typeof opts.onAddTaskUnderDivider !== "function") return;
    opts.onAddTaskUnderDivider(btn.dataset.addUnderDivider || "", btn.dataset.insertAfter || "");
  }));
  $all("[data-open]", mount).forEach((btn) => btn.addEventListener("click", () => opts.onOpenEditor(btn.dataset.open)));
  $all("[data-open-sub]", mount).forEach((btn) => {
    const raw = btn.dataset.openSub;
    const colon = raw.indexOf(":");
    const id = raw.slice(0, colon);
    const path = raw.slice(colon + 1);
    btn.addEventListener("click", () => opts.onOpenEditor(id, path));
  });
  $all("select[data-field]", mount).forEach((sel) => sel.addEventListener("change", () => {
    const task = resolveTask(sel.dataset.id);
    if (!task) return;
    const colorPrefix = sel.dataset.colorPrefix;
    const custom = sel.dataset.customColor === "1";
    if (custom && opts.statusColumn && opts.statusColumn.colorFor) {
      sel.style.backgroundColor = opts.statusColumn.colorFor(sel.value);
      sel.className = "monday-status-select monday-status-custom";
    } else {
      sel.className = `monday-status-select ${colorPrefix}-${sel.value.replace(/\s+/g, "-")}`;
      sel.style.backgroundColor = "";
    }
    opts.onFieldChange(task, sel.dataset.field, sel.value);
  }));
  $all("[data-milestone-toggle]", mount).forEach((btn) => btn.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (!canEdit || typeof opts.onFieldChange !== "function") return;
    const divider = resolveTask(btn.dataset.milestoneToggle);
    if (!divider || !(divider.itemType === "divider" || divider.workItemLevel === "Divider")) return;
    divider.isMilestone = !divider.isMilestone;
    if (divider.metadata && typeof divider.metadata === "object") {
      divider.metadata.isMilestone = !!divider.isMilestone;
    }
    opts.onFieldChange(divider, "isMilestone", !!divider.isMilestone);
    const on = !!divider.isMilestone;
    btn.classList.toggle("is-on", on);
    btn.setAttribute("aria-pressed", on ? "true" : "false");
    const icon = btn.querySelector("i");
    if (icon) icon.className = `bx ${on ? "bxs-flag" : "bx-flag"}`;
  }));
  $all("select[data-sub-field]", mount).forEach((sel) => sel.addEventListener("change", () => {
    const task = resolveTask(sel.dataset.id);
    const path = sel.dataset.subPath != null ? sel.dataset.subPath : sel.dataset.subIndex;
    const sub = getSubByPath(task, path);
    if (!task || !sub) return;
    sub[sel.dataset.subField] = sel.value;
    opts.onFieldChange(task, `subtask.${sel.dataset.subField}`, sel.value, path);
  }));
  $all("input[data-date-field]", mount).forEach((input) => {
    input.addEventListener("change", () => {
      const task = resolveTask(input.dataset.id);
      if (!task) return;
      const field = input.dataset.dateField;
      const path = input.dataset.subPath != null ? input.dataset.subPath : input.dataset.subIndex;
      if (path != null && path !== "") {
        const sub = getSubByPath(task, path);
        if (!sub) return;
        sub[field] = input.value;
        opts.onFieldChange(task, `subtask.${field}`, input.value, path);
      } else {
        task[field] = input.value;
        opts.onFieldChange(task, field, input.value);
      }
    });
  });
  function wireInlineText(input) {
    const commit = () => {
      const task = resolveTask(input.dataset.id);
      if (!task) return;
      const value = input.value.trim();
      if (input.dataset.subTextField != null) {
        const path = input.dataset.subPath != null ? input.dataset.subPath : input.dataset.subIndex;
        const sub = getSubByPath(task, path);
        if (!sub) return;
        const field = input.dataset.subTextField;
        if (field === "text" && !value) { input.value = sub.text || "Untitled subitem"; return; }
        sub[field] = value || (field === "assignee" ? "Unassigned" : sub[field]);
        opts.onFieldChange(task, `subtask.${field}`, sub[field], path);
        return;
      }
      const field = input.dataset.textField;
      if (field === "title" && !value) { input.value = task.title; return; }
      task[field] = value || (field === "assignee" ? "Unassigned" : task[field]);
      opts.onFieldChange(task, field, task[field]);
    };
    // Commit as the user types (debounced). Blur does NOT fire when the SPA
    // swaps the page DOM mid-edit, so waiting for blur alone silently lost
    // titles typed right before navigating away.
    let typeTimer = null;
    const commitSoon = () => {
      clearTimeout(typeTimer);
      typeTimer = setTimeout(() => { typeTimer = null; commit(); }, 500);
    };
    const commitNow = () => {
      clearTimeout(typeTimer);
      typeTimer = null;
      commit();
    };
    input.addEventListener("input", commitSoon);
    input.addEventListener("change", commitNow);
    input.addEventListener("blur", commitNow);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); input.blur(); }
    });
    if (isTracker && (input.dataset.textField === "assignee" || input.dataset.subTextField === "assignee")) {
      wireAssigneeAutocomplete(mount, input);
    }
  }
  $all("input[data-text-field], input[data-sub-text-field]", mount).forEach(wireInlineText);
  if (opts.focusTaskId) {
    const focusEl = $(`input[data-text-field="title"][data-id="${opts.focusTaskId}"]`, mount);
    if (focusEl) { focusEl.focus(); focusEl.select(); }
  }
  if (opts.focusSubtask && opts.focusSubtask.taskId && opts.focusSubtask.path != null) {
    const focusEl = $(`input[data-sub-text-field="text"][data-id="${opts.focusSubtask.taskId}"][data-sub-path="${opts.focusSubtask.path}"]`, mount);
    if (focusEl) { focusEl.focus(); focusEl.select(); }
  }
  const inlineAdd = $("#monday-inline-add", mount);
  if (inlineAdd && opts.onInlineAdd) inlineAdd.addEventListener("click", () => opts.onInlineAdd());
  const addDividerBtn = $("#monday-add-divider", mount);
  if (addDividerBtn && opts.onAddDivider) addDividerBtn.addEventListener("click", () => opts.onAddDivider());
  $all("[data-toggle-review]", mount).forEach((btn) => btn.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (typeof opts.onToggleReview === "function") opts.onToggleReview(btn.dataset.toggleReview);
  }));
  $all("[data-add-subtask]", mount).forEach((btn) => btn.addEventListener("click", () => {
    if (opts.onAddSubtask) opts.onAddSubtask(btn.dataset.addSubtask, btn.dataset.parentPath || "");
  }));
  $all("[data-add-subtask-sibling]", mount).forEach((btn) => btn.addEventListener("click", () => {
    if (opts.onAddSubtaskSibling) opts.onAddSubtaskSibling(btn.dataset.addSubtaskSibling, btn.dataset.siblingPath || "");
  }));
  $all("[data-delete-task]", mount).forEach((btn) => btn.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    const task = resolveTask(btn.dataset.deleteTask);
    if (task && typeof opts.onDeleteTask === "function") opts.onDeleteTask(task);
  }));
  $all("[data-delete-subtask]", mount).forEach((btn) => btn.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    const raw = btn.dataset.deleteSubtask || "";
    const colon = raw.indexOf(":");
    if (colon < 1) return;
    const task = resolveTask(raw.slice(0, colon));
    const path = raw.slice(colon + 1);
    if (task && path && typeof opts.onDeleteSubtask === "function") opts.onDeleteSubtask(task, path);
  }));

  $all("[data-open-notes]", mount).forEach(btn => btn.addEventListener("click", () => {
    const rawId = btn.dataset.openNotes;
    let targetTask = null;
    let isSub = false;
    let subPath = null;
    let parentTask = null;
    if (rawId.includes(":")) {
      const colon = rawId.indexOf(":");
      parentTask = resolveTask(rawId.slice(0, colon));
      subPath = rawId.slice(colon + 1);
      targetTask = getSubByPath(parentTask, subPath);
      isSub = true;
    } else {
      targetTask = resolveTask(rawId);
      parentTask = targetTask;
    }
    if (!targetTask) return;
    openTaskNotesModal(targetTask, targetTask.title || targetTask.text || "Notes", (notes) => {
      if (parentTask && typeof opts.onFieldChange === "function") {
        opts.onFieldChange(parentTask, "notes", notes, isSub ? subPath : undefined);
      }
      const count = Array.isArray(notes) ? notes.length : 0;
      btn.classList.toggle("has-notes", count > 0);
      let countEl = btn.querySelector(".monday-notes-btn-count");
      if (count) {
        if (!countEl) {
          countEl = document.createElement("span");
          countEl.className = "monday-notes-btn-count";
          btn.appendChild(countEl);
        }
        countEl.textContent = String(count);
      } else if (countEl) {
        countEl.remove();
      }
    });
  }));

  // Track mouse X so the row-insert "+" button floats under the cursor.
  if (canEdit) {
    mount.addEventListener("mousemove", (e) => {
      const insertRow = e.target.closest(".row-insert-row");
      if (!insertRow) return;
      const wrap = insertRow.querySelector(".row-insert-wrap");
      if (!wrap) return;
      const rect = wrap.getBoundingClientRect();
      const x = Math.max(14, Math.min(e.clientX - rect.left, rect.width - 14));
      wrap.style.setProperty("--insert-x", `${x}px`);
    });
  }

  // Ensure tooltips run if tip system exists
  if (typeof initTooltips === "function") initTooltips(mount);
  
  $all("[data-subtask-done]", mount).forEach((cb) => cb.addEventListener("change", () => {
    const raw = cb.dataset.subtaskDone;
    const colon = raw.indexOf(":");
    const id = raw.slice(0, colon);
    const path = raw.slice(colon + 1);
    const task = resolveTask(id);
    if (!task) return;
    opts.onToggleSubtaskDone(task, path, cb.checked);
  }));
  if (isTracker) {
    $all(".tracker-complete-btn[data-task-complete]", mount).forEach((btn) => btn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const task = resolveTask(btn.dataset.taskComplete);
      if (!task || typeof opts.onToggleComplete !== "function") return;
      opts.onToggleComplete(task);
      const done = taskIsFullyDone(task);
      btn.classList.toggle("is-done", done);
      btn.setAttribute("aria-pressed", done ? "true" : "false");
      btn.querySelector("i").className = `bx ${done ? "bxs-check-circle" : "bx-circle"}`;
      const row = btn.closest("tr");
      if (row) row.classList.toggle("is-complete", done);
    }));
    $all(".tracker-complete-btn[data-subtask-complete]", mount).forEach((btn) => btn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const raw = btn.dataset.subtaskComplete || "";
      const colon = raw.indexOf(":");
      if (colon < 1 || typeof opts.onToggleSubtaskDone !== "function") return;
      const task = resolveTask(raw.slice(0, colon));
      const path = raw.slice(colon + 1);
      const subtask = getSubByPath(task, path);
      if (!task || !subtask) return;
      opts.onToggleSubtaskDone(task, path, !subtask.done);
    }));
  }
  if (isTracker && canEdit && (opts.onDeleteTask || opts.onDeleteSubtask || opts.onEditDivider || opts.onAddTaskUnderDivider || opts.onDeleteDivider || opts.onFieldChange || opts.onAddSubtask)) {
    mount.addEventListener("contextmenu", (e) => {
      const dividerRow = e.target.closest("tr.monday-divider-row[data-divider-id]");
      const subRow = e.target.closest("tr[data-sub-row]");
      const taskRow = e.target.closest("tr.monday-row[data-row-id]");
      if (!dividerRow && !subRow && !taskRow) {
        if (opts.onAddDivider || opts.onAddTaskUnderDivider) {
          e.preventDefault();
          const emptyItems = [];
          if (opts.onAddDivider) emptyItems.push({ label: "New project divider", icon: "bx-layout", action: () => opts.onAddDivider() });
          if (opts.onAddTaskUnderDivider) emptyItems.push({ label: "Add task", icon: "bx-plus", action: () => opts.onAddTaskUnderDivider("") });
          if (emptyItems.length && typeof showContextMenu === "function") showContextMenu(e.clientX, e.clientY, emptyItems);
        }
        return;
      }
      e.preventDefault();
      if (typeof showContextMenu !== "function") return;

      if (dividerRow) {
        const dividerId = dividerRow.dataset.dividerId;
        const divider = (opts.allTasks || opts.tasks || []).find((t) => t.id === dividerId);
        if (!divider) return;
        const items = [];
        if (opts.onEditDivider) {
          items.push({ label: "Edit divider settings", icon: "bx-cog", action: () => opts.onEditDivider(dividerId) });
        }
        if (opts.onAddTaskUnderDivider) {
          items.push({ label: "Add task under divider", icon: "bx-plus", action: () => opts.onAddTaskUnderDivider(dividerId) });
        }
        if (opts.onFieldChange) {
          items.push({
            label: divider.isMilestone ? "Clear milestone" : "Mark as milestone",
            icon: divider.isMilestone ? "bxs-flag" : "bx-flag",
            action: () => {
              divider.isMilestone = !divider.isMilestone;
              if (divider.metadata) divider.metadata.isMilestone = !!divider.isMilestone;
              opts.onFieldChange(divider, "isMilestone", !!divider.isMilestone);
            }
          });
        }
        if (opts.onDeleteDivider && String(divider.title || "").trim().toLowerCase() !== "general") {
          items.push({ separator: true });
          items.push({ label: "Delete divider", icon: "bx-trash", danger: true, action: () => opts.onDeleteDivider(dividerId) });
        }
        if (items.length) showContextMenu(e.clientX, e.clientY, items);
        return;
      }

      const taskId = subRow ? subRow.dataset.taskId : taskRow.dataset.rowId;
      const task = resolveTask(taskId);
      if (!task) return;
      if (subRow && (opts.onDeleteSubtask || opts.onAddSubtask)) {
        const path = subRow.dataset.subPath != null ? subRow.dataset.subPath : subRow.dataset.subIndex;
        const menu = [];
        if (opts.onAddSubtask) menu.push({ label: "Add subtask", icon: "bx-list-plus", action: () => opts.onAddSubtask(task.id, path) });
        if (opts.onDeleteSubtask) {
          if (menu.length) menu.push({ separator: true });
          menu.push({ label: "Delete subtask", icon: "bx-trash", danger: true, action: () => opts.onDeleteSubtask(task, path) });
        }
        showContextMenu(e.clientX, e.clientY, menu);
        return;
      }
      const menuItems = [];
      if (opts.onOpenTask) menuItems.push({ label: "Edit task", icon: "bx-edit", action: () => opts.onOpenTask(task.id) });
      if (opts.onAddSubtask) menuItems.push({ label: "Add subtask", icon: "bx-list-plus", action: () => opts.onAddSubtask(task.id, "") });
      if (opts.meetingMode && opts.onToggleReview) {
        const reviewed = (task.reviewStatus || "Not Reviewed") !== "Not Reviewed";
        menuItems.push({
          label: reviewed ? "Mark not reviewed" : "Mark reviewed",
          icon: reviewed ? "bxs-check-circle" : "bx-circle",
          action: () => opts.onToggleReview(task.id)
        });
      }
      if (opts.onFieldChange) {
        menuItems.push({
          label: task.isMilestone ? "Clear milestone" : "Mark as milestone",
          icon: task.isMilestone ? "bxs-flag" : "bx-flag",
          action: () => {
            task.isMilestone = !task.isMilestone;
            opts.onFieldChange(task, "isMilestone", !!task.isMilestone);
          }
        });
      }
      if (opts.onDeleteTask) {
        menuItems.push({ separator: true });
        menuItems.push({ label: "Delete task", icon: "bx-trash", danger: true, action: () => opts.onDeleteTask(task) });
      }
      if (menuItems.length) showContextMenu(e.clientX, e.clientY, menuItems);
    });
  }
  if (opts.reorderable && opts.allTasks && opts.onReorder) {
    wireMondayTableRowReordering(mount, opts.allTasks, opts.onReorder, opts.onReorderDivider);
  }
}

function wireMondayTableRowReordering(mount, tasks, onReorder, onReorderDivider) {
  let draggedId = null;
  let dragKind = null; // "task" | "divider"

  function clearDragOver() {
    $all(".monday-row[data-row-id], .monday-divider-row[data-divider-id]", mount).forEach((el) => {
      el.classList.remove("drag-over");
    });
  }

  $all(".monday-row[data-row-id]", mount).forEach((row) => {
    const id = row.dataset.rowId;
    const handle = $(".cl-drag-handle", row);
    if (!handle) return;
    handle.addEventListener("mousedown", () => {
      row.draggable = true;
      const resetIfNoDrag = () => { if (!row.classList.contains("dragging")) row.draggable = false; };
      window.addEventListener("mouseup", resetIfNoDrag, { once: true });
    });
    row.addEventListener("dragstart", (e) => {
      draggedId = id;
      dragKind = "task";
      setTimeout(() => row.classList.add("dragging"), 0);
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", id);
    });
    row.addEventListener("dragend", () => {
      row.draggable = false;
      row.classList.remove("dragging");
      clearDragOver();
      draggedId = null;
      dragKind = null;
    });
    row.addEventListener("dragover", (e) => {
      if (!draggedId || draggedId === id || dragKind !== "task") return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      row.classList.add("drag-over");
    });
    row.addEventListener("dragleave", () => row.classList.remove("drag-over"));
    row.addEventListener("drop", (e) => {
      e.preventDefault();
      row.classList.remove("drag-over");
      if (!draggedId || draggedId === id || dragKind !== "task") return;
      const moved = typeof reorderTrackerTask === "function"
        ? reorderTrackerTask(tasks, draggedId, id)
        : (() => {
          const from = tasks.findIndex((t) => t.id === draggedId);
          const to = tasks.findIndex((t) => t.id === id);
          if (from < 0 || to < 0 || from === to) return null;
          const [item] = tasks.splice(from, 1);
          tasks.splice(to, 0, item);
          return item;
        })();
      draggedId = null;
      dragKind = null;
      if (moved && typeof onReorder === "function") onReorder(moved);
    });
  });

  $all(".monday-divider-row[data-divider-id]", mount).forEach((row) => {
    const id = row.dataset.dividerId;
    const handle = $(".cl-drag-handle", row);
    if (!handle || !id) return;
    handle.addEventListener("mousedown", () => {
      row.draggable = true;
      const resetIfNoDrag = () => { if (!row.classList.contains("dragging")) row.draggable = false; };
      window.addEventListener("mouseup", resetIfNoDrag, { once: true });
    });
    row.addEventListener("dragstart", (e) => {
      draggedId = id;
      dragKind = "divider";
      setTimeout(() => row.classList.add("dragging"), 0);
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", id);
    });
    row.addEventListener("dragend", () => {
      row.draggable = false;
      row.classList.remove("dragging");
      clearDragOver();
      draggedId = null;
      dragKind = null;
    });
    row.addEventListener("dragover", (e) => {
      if (!draggedId || draggedId === id || dragKind !== "divider") return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      row.classList.add("drag-over");
    });
    row.addEventListener("dragleave", () => row.classList.remove("drag-over"));
    row.addEventListener("drop", (e) => {
      e.preventDefault();
      row.classList.remove("drag-over");
      if (!draggedId || draggedId === id || dragKind !== "divider") return;
      const movedId = draggedId;
      const ok = typeof reorderTrackerDivider === "function"
        ? reorderTrackerDivider(tasks, movedId, id)
        : false;
      draggedId = null;
      dragKind = null;
      if (ok && typeof onReorderDivider === "function") onReorderDivider(movedId);
      else if (ok && typeof onReorder === "function") {
        const moved = tasks.find((t) => t.id === movedId);
        if (moved) onReorder(moved);
      }
    });
  });
}

window.addEventListener("hashchange", renderPage);

/* ---------- boot: SharePoint mode vs. local-dev fallback ---------- */

function minimalUnconfiguredDb(currentUser) {
  return normalizeStoreShape(buildEmptyStore({
    id: "u0",
    displayName: (currentUser && currentUser.displayName) || "Unknown",
    email: (currentUser && currentUser.email) || "",
    role: "Member",
    isAdmin: false,
    isMeetingAdmin: false
  }));
}

function syncSiteUsersInBackground(siteUrl, currentUser) {
  if (!siteUrl || !currentUser) return;
  // Refresh identity fields on existing AEWTTR people (does not auto-associate).
  if (typeof sharePointAdapter.syncSiteUsersToAppRoles === "function") {
    sharePointAdapter.syncSiteUsersToAppRoles(siteUrl, currentUser)
      .then(() => {
        if (typeof refreshMembersFromSharePoint === "function") return refreshMembersFromSharePoint(siteUrl);
        if (typeof rebuildMembersView === "function") rebuildMembersView();
      })
      .catch(() => {});
  }
  // Warm pickers from AEWTTR-associated App Roles.
  if (typeof sharePointAdapter.loadSiteMemberDirectory === "function") {
    sharePointAdapter.loadSiteMemberDirectory(siteUrl).catch(() => {});
  }
  // Warm Admin Users → All users directory if previously pulled.
  if (typeof sharePointAdapter.loadPeopleDirectory === "function") {
    sharePointAdapter.loadPeopleDirectory(siteUrl).catch(() => {});
  }
}

async function bootSharePointMode(siteUrl) {
  window.AEWTTR.mode = "sharepoint";
  window.AEWTTR.siteUrl = siteUrl;
  window.AEWTTR.spBootError = null;
  window.AEWTTR.setupStatus = null;
  bootLog(`Detected SharePoint site: ${siteUrl}`, "success");
  setBootProgress(10);

  let currentUser = null;
  try {
    bootLog("Loading current SharePoint user...");
    const restUser = await sharePointAdapter.getCurrentUser(siteUrl);
    const pageUser = sharePointAdapter.getCurrentUserFromPageContext();
    currentUser = Object.assign({}, pageUser, restUser);
    window.AEWTTR.currentSpUser = currentUser;
    bootLog(`Current user: ${currentUser.displayName || "Unknown"} (${currentUser.email || "no email"})`, "success");
    setBootProgress(25);
  } catch (e) {
    console.error("[PULSE boot] Loading current SharePoint user failed:", e);
    window.AEWTTR.currentSpUser = null;
    window.AEWTTR.spBootError = (e && e.friendly)
      ? e
      : sharePointAdapter.formatSpError(null, `${(e && e.stack) || (e && e.message) || String(e)} (this is not actually a network error — it's an unhandled JS exception during boot; see the browser console for the real cause)`);
    bootLog(`Current user lookup failed: ${window.AEWTTR.spBootError.friendly || "Unknown error"}`, "error");
    window.AEWTTR.db = minimalUnconfiguredDb(null);
    return;
  }

  try {
    try {
      bootLog("Preparing PULSE App Roles list core columns...");
      const rolesPrep = await ensureRolesListCore(siteUrl);
      bootLog(
        `Roles list ready${rolesPrep.listCreated ? " (list created)" : ""}.`
        + `${rolesPrep.ensuredFields.length ? ` Columns created: ${rolesPrep.ensuredFields.join(", ")}` : ""}`
        + `${rolesPrep.skippedFields.length ? ` Columns skipped: ${rolesPrep.skippedFields.join(", ")}` : ""}`,
        "success"
      );
    } catch (e) {
      bootLog(`Roles list setup warning: ${(e && e.friendly) || String(e)}`, "error");
    }
    setBootProgress(40);

    // Fire-and-forget: self-healing for "PULSE Document Review" — Submit
    // Document was 404ing with a generic error on any tenant where full
    // setup hadn't been run yet.
    ensureDocReviewList(siteUrl)
      .then((r) => bootLog(`PULSE Document Review ready${r.listCreated ? " (list created)" : ""}.`, "success"))
      .catch((e) => bootLog(`PULSE Document Review setup warning: ${(e && e.friendly) || String(e)}`, "error"));

    // Same pattern for global people groups (SharePoint list still titled
    // "PULSE Doc Reviewer Groups") and for PULSE Risks — "Save Risk" on the
    // Risk tab (and its compact copy on Project Home) hit the identical
    // generic 404 Document Review used to before this existed.
    ensureDocReviewerGroupsList(siteUrl)
      .then((r) => bootLog(`PULSE Groups list ready${r.listCreated ? " (list created)" : ""}.`, "success"))
      .catch((e) => bootLog(`PULSE Groups setup warning: ${(e && e.friendly) || String(e)}`, "error"));

    ensureRisksList(siteUrl)
      .then((r) => bootLog(`PULSE Risks ready${r.listCreated ? " (list created)" : ""}.`, "success"))
      .catch((e) => bootLog(`PULSE Risks setup warning: ${(e && e.friendly) || String(e)}`, "error"));

    let roleInfo = { role: APP_CONFIG.defaultUserRole, isAdmin: false, isMeetingAdmin: false, isFinanceAdmin: false, isDocAdmin: false, roleRecordId: null, associated: false, source: "Default" };
    try {
      bootLog("Resolving app role from PULSE App Roles...");
      // If no Admin exists yet, promote the current user so the site stays reachable.
      if (typeof sharePointAdapter.bootstrapAdminIfNeeded === "function") {
        try {
          const boot = await sharePointAdapter.bootstrapAdminIfNeeded(siteUrl, currentUser);
          if (boot && boot.bootstrapped) bootLog("Bootstrap Admin created for current user.", "success");
        } catch (e) {
          bootLog(`Bootstrap admin skipped: ${(e && e.friendly) || String(e)}`, "error");
        }
      }
      roleInfo = await sharePointAdapter.getCurrentUserRole(siteUrl, currentUser);
      bootLog(`Resolved role: ${roleInfo.role}${roleInfo.associated || roleInfo.isAdmin ? " (associated)" : " (not associated)"}`, "success");
      if (roleInfo.themeMode) setCurrentThemeValue(roleInfo.themeMode);
    } catch (e) {
      bootLog("Role lookup fell back to default Member.", "error");
    }
    setBootProgress(55);

    const bootUser = buildBootUserObject(currentUser, roleInfo);
    window.AEWTTR.pulseAssociated = !!(roleInfo.isAdmin || roleInfo.associated || (roleInfo.roleRecordId && sharePointAdapter.isUserPulseAssociated(roleInfo)));
    if (roleInfo.isAdmin) window.AEWTTR.pulseAssociated = true;
    const cachedDb = readSpDbCache(siteUrl);
    if (cachedDb && cachedDb.data) {
      const hydrated = normalizeStoreShape(Object.assign({}, cachedDb.data, { user: bootUser }));
      window.AEWTTR.db = hydrated;
      window.AEWTTR.dbHydratedFromCache = true;
      window.AEWTTR.dbCacheSavedAt = cachedDb.savedAt;
      bootLog(`Opened from cached site data (${fmtRelativeTime(new Date(cachedDb.savedAt).toISOString())}) — syncing fresh data in background...`, "success");
      setBootProgress(100);
      syncSiteUsersInBackground(siteUrl, currentUser);
      getSetupStatus(siteUrl)
        .then((setupStatus) => { window.AEWTTR.setupStatus = setupStatus; })
        .catch(() => {});
      return;
    }

    try {
      bootLog("Checking overall SharePoint setup status...");
      const setupStatus = await Promise.race([
        getSetupStatus(siteUrl),
        new Promise((_, reject) => setTimeout(() => reject(new Error("Setup status timed out")), 4000))
      ]);
      window.AEWTTR.setupStatus = setupStatus;
      bootLog(setupStatus.ready ? "Setup ready." : "Setup incomplete; loading available SharePoint data anyway.", setupStatus.ready ? "success" : "error");
      setBootProgress(70);

      bootLog("Loading SharePoint-backed app data...");
      const loadedDb = await loadAllFromSharePoint(siteUrl);
      loadedDb.user = bootUser;
      // Merge local-only fields (tickets) so they survive a fresh SP load.
      const localSnapshot = typeof aewttrLoadStore === "function" ? aewttrLoadStore() : null;
      if (localSnapshot && Array.isArray(localSnapshot.tickets) && localSnapshot.tickets.length > 0) {
        loadedDb.tickets = localSnapshot.tickets;
      }
      window.AEWTTR.db = normalizeStoreShape(loadedDb);
      window.AEWTTR.dbHydratedFromCache = false;
      persistSpDbCache(siteUrl, window.AEWTTR.db);
      bootLog("App data loaded. Opening PULSE...", "success");
      setBootProgress(100);
      syncSiteUsersInBackground(siteUrl, currentUser);
    } catch (setupOrLoadError) {
      bootLog(`Heavy setup/data load skipped: ${String(setupOrLoadError && setupOrLoadError.message || setupOrLoadError)}`, "error");
      const minDb = minimalUnconfiguredDb(currentUser);
      minDb.user.role = roleInfo.role;
      minDb.user.isAdmin = roleInfo.isAdmin;
      minDb.user.isMeetingAdmin = roleInfo.isMeetingAdmin;
      minDb.user.isFinanceAdmin = roleInfo.isFinanceAdmin;
      minDb.user.isDocAdmin = roleInfo.isDocAdmin;
      minDb.user.themeMode = normalizeThemeValue(roleInfo.themeMode || getCurrentTheme());
      window.AEWTTR.db = minDb;
      setBootProgress(100);
    }
  } catch (e) {
    console.error("[PULSE boot] Fatal error during SharePoint boot:", e);
    window.AEWTTR.spBootError = (e && e.friendly)
      ? e
      : sharePointAdapter.formatSpError(null, `${(e && e.stack) || (e && e.message) || String(e)} (this is not actually a network error — it's an unhandled JS exception during boot; see the browser console for the real cause)`);
    bootLog(`Boot failed: ${window.AEWTTR.spBootError.friendly || String(e)}`, "error");
    window.AEWTTR.db = minimalUnconfiguredDb(currentUser);
    setBootProgress(100);
  }
}

/* ---------- background refresh (SharePoint mode) ----------
   Quietly re-pull SharePoint lists on an interval, on tab focus, and on
   every in-app navigation (see maybeProactiveRefreshOnNavigate). Data lands
   in window.AEWTTR.db; nav badges / notification bell update, and when the UI
   is idle we soft-rebuild the active page (scroll preserved).

   Critical guard: do NOT replace window.AEWTTR.db while a modal is open or an
   autosave is mid-flight. Pages close over object references from the previous
   db; swapping underneath makes SharePoint writes succeed on orphaned objects
   while redraws read the live db — the classic "had to do it twice / only
   shows after refresh" failure. */
let _lastBgRefresh = 0;
let _bgRefreshTimer = null;
let _docReviewReminderTimer = null;

function isEditableFocused() {
  const active = document.activeElement;
  return !!(active && (active.tagName === "INPUT" || active.tagName === "TEXTAREA" || active.tagName === "SELECT" || active.isContentEditable));
}

/* True when the user is mid-edit (open editors, focused fields, drag, etc.). */
function isUserEditing() {
  if (isEditableFocused()) return true;
  if (document.body.classList.contains("gantt-dragging")) return true;
  // Side-panel task/subtask editors (no modal backdrop).
  if (document.querySelector(".task-side-panel, .task-side-backdrop, .task-editor-modal")) return true;
  return false;
}

/* True when swapping/redrawing would clobber in-progress UI work. */
function isBackgroundDataSwapBlocked() {
  if (document.querySelector(".aewttr-modal-backdrop")) return true;
  if (isUserEditing()) return true;
  if (document.querySelector('[data-state="pending"], [data-state="saving"]')) return true;
  // Debounced / in-flight Repo writes — swapping db behind them orphans the
  // objects those writes hold (classic "had to do it twice" failure).
  if (typeof Repo !== "undefined" && Repo.hasPendingChanges && Repo.hasPendingChanges()) return true;
  return false;
}

function isLivePageRedrawBlocked() {
  return isBackgroundDataSwapBlocked();
}

/* Soft-rebuild the current route, preserving scroll. Returns false when
   skipped because the user is mid-edit. */
function softRenderCurrentPage() {
  if (isLivePageRedrawBlocked()) return false;
  const route = currentRoute();
  // Meeting flows keep large in-closure UI state that a remount would wipe.
  if (route.app === "weekly") {
    const meeting = window.AEWTTR.db && window.AEWTTR.db.weeklyMeeting;
    if (meeting && (meeting.meetingStatus === "active" || meeting.meetingStatus === "in-progress")) return false;
    const projectMeetings = meeting && meeting.projectMeetings;
    if (projectMeetings && Object.keys(projectMeetings).some((pid) => {
      const pm = projectMeetings[pid];
      return pm && (pm.meetingStatus === "active" || pm.meetingStatus === "in-progress" || pm.currentSession);
    })) return false;
  }
  if (route.app === "travel" && route.parts[0] === "submit") return false;
  // Most routes use the page scroller, but Weekly Meeting's full task list
  // has a nested scroll panel. Preserve that panel first so a background
  // refresh cannot snap a user back to the top after expanding projects.
  const scrollerSelector = ".meeting-project-main-body--tasks";
  const scroller = document.querySelector(scrollerSelector) || document.querySelector(".aewttr-content");
  const scrollTop = scroller ? scroller.scrollTop : 0;
  try {
    renderPage();
    const refreshedScroller = document.querySelector(scrollerSelector) || document.querySelector(".aewttr-content");
    if (refreshedScroller) {
      refreshedScroller.scrollTop = scrollTop;
      requestAnimationFrame(() => { refreshedScroller.scrollTop = scrollTop; });
    }
    return true;
  } catch (e) {
    return false;
  }
}

/* Re-seat a project object into the LIVE db. If a background refresh replaced
   window.AEWTTR.db while a form still held a closed-over `proj`, subsequent
   saves would write the orphan to SharePoint but redraws would read the
   refreshed copy and look unchanged. */
function reanchorProject(proj) {
  if (!proj || !proj.id || !window.AEWTTR || !window.AEWTTR.db) return proj;
  const liveDb = window.AEWTTR.db;
  if (!Array.isArray(liveDb.projects)) liveDb.projects = [];
  const idx = liveDb.projects.findIndex((p) => p.id === proj.id);
  if (idx === -1) {
    liveDb.projects.unshift(proj);
  } else if (liveDb.projects[idx] !== proj) {
    const previous = liveDb.projects[idx];
    if (!proj._spId && previous && previous._spId) proj._spId = previous._spId;
    liveDb.projects[idx] = proj;
  }
  return proj;
}

/* Same idea for singleton config rows (location/rag/notification/ai). */
function reanchorConfig(key, obj) {
  if (!obj || !key || !window.AEWTTR || !window.AEWTTR.db) return obj;
  const liveDb = window.AEWTTR.db;
  const previous = liveDb[key];
  if (previous && previous !== obj) {
    if (!obj._spId && previous._spId) obj._spId = previous._spId;
  }
  liveDb[key] = obj;
  return obj;
}

function applyLiveRefreshFromBackground(options) {
  options = options || {};
  renderNav();
  refreshUserNotifications();
  window.dispatchEvent(new CustomEvent("pulse:data-refreshed", {
    detail: { route: currentRoute(), at: Date.now(), reason: options.reason || "" }
  }));
  // Soft remount only when explicitly requested (post-mutation). Background
  // polls rely on pulse:data-refreshed opt-in handlers so we don't wipe
  // in-progress meeting/board sessions that live in page closures.
  if (options.softRender === true) softRenderCurrentPage();
}

/* After a successful local mutation: redraw from current in-memory state
   without waiting for the next SharePoint poll. */
function notifyLocalDataChanged(reason) {
  applyLiveRefreshFromBackground({ reason: reason || "local-mutation", softRender: true });
}

async function refreshSharePointData(reason) {
  if (window.AEWTTR.mode !== "sharepoint" || !window.AEWTTR.siteUrl) return false;
  // Never poll / swap while the user is editing — even a successful fetch
  // would replace closed-over object refs and wipe in-progress form state.
  if (isBackgroundDataSwapBlocked()) return false;
  if (typeof Repo !== "undefined" && Repo.hasPendingChanges && Repo.hasPendingChanges()) return false;
  const route = currentRoute();
  if (route.app === "travel" && route.parts[0] === "submit") return false;
  // Weekly Meeting has a nested task-list scroller and holds live meeting UI
  // state in page closures. Swapping the backing database during its startup
  // or polling refresh can interrupt that scroll region, so leave this route
  // stable until the user navigates away or makes an explicit save.
  if (route.app === "weekly") return false;
  // Keep a wider quiet window after local writes so board/project JSON
  // saves are not overwritten by a poll that raced the SharePoint index.
  if (window.AEWTTR.lastLocalSpWriteAt && Date.now() - window.AEWTTR.lastLocalSpWriteAt < 15000) return false;
  if (window.AEWTTR.sharePointRefreshInFlight) return false;

  window.AEWTTR.sharePointRefreshInFlight = true;
  try {
    // Re-check after the network round-trip — a modal/editor may have opened
    // while we were fetching; swapping now would orphan closed-over refs.
    if (isBackgroundDataSwapBlocked()) return false;
    if (typeof Repo !== "undefined" && Repo.hasPendingChanges && Repo.hasPendingChanges()) return false;
    const previousDb = window.AEWTTR.db;
    const previousUser = previousDb && previousDb.user;
    const refreshed = await loadAllFromSharePoint(window.AEWTTR.siteUrl, { fast: true });
    if (isBackgroundDataSwapBlocked()) return false;
    if (typeof Repo !== "undefined" && Repo.hasPendingChanges && Repo.hasPendingChanges()) return false;
    refreshed.user = previousUser || refreshed.user;
    // Preserve local-only fields not backed by any SharePoint list so they
    // are not wiped by the background refresh (tickets live in pulse-local-db).
    if (previousDb && Array.isArray(previousDb.tickets) && previousDb.tickets.length > 0) {
      refreshed.tickets = previousDb.tickets;
    }
    window.AEWTTR.db = normalizeStoreShape(refreshed);
    window.AEWTTR.lastSharePointRefreshAt = Date.now();
    window.AEWTTR.dbHydratedFromCache = false;
    persistSpDbCache(window.AEWTTR.siteUrl, window.AEWTTR.db);
    applyLiveRefreshFromBackground({ reason: reason || "sharepoint-refresh" });
    return true;
  } catch (e) {
    return false;
  } finally {
    window.AEWTTR.sharePointRefreshInFlight = false;
  }
}

function scheduleDocReviewReminderSweep() {
  if (window.AEWTTR.mode !== "sharepoint") return;
  if (typeof window.runDocReviewReminderSweep !== "function") return;
  setTimeout(() => {
    window.runDocReviewReminderSweep().catch(() => {});
  }, 8000);
  if (_docReviewReminderTimer) clearInterval(_docReviewReminderTimer);
  _docReviewReminderTimer = setInterval(() => {
    if (document.visibilityState !== "visible") return;
    window.runDocReviewReminderSweep().catch(() => {});
  }, 60 * 60 * 1000);
}

function wireBackgroundRefresh() {
  if (window.AEWTTR.mode !== "sharepoint") return;
  scheduleDocReviewReminderSweep();
  setTimeout(() => {
    refreshSharePointData("startup").then((ok) => {
      if (ok) _lastBgRefresh = Date.now();
      scheduleNotificationRefresh();
    });
  }, window.AEWTTR.dbHydratedFromCache ? 50 : 1500);
  document.addEventListener("visibilitychange", async () => {
    if (document.visibilityState !== "visible") return;
    if (Date.now() - _lastBgRefresh < 15000) return;
    const ok = await refreshSharePointData("tab-focus");
    if (ok) _lastBgRefresh = Date.now();
  });
  if (_bgRefreshTimer) clearInterval(_bgRefreshTimer);
  // Tick often; success-gates _lastBgRefresh so a blocked attempt (modal open)
  // retries soon after the UI is idle instead of waiting a full quiet window.
  _bgRefreshTimer = setInterval(() => {
    if (document.visibilityState !== "visible") return;
    if (Date.now() - _lastBgRefresh < 5000) return;
    refreshSharePointData("interval").then((ok) => {
      if (ok) _lastBgRefresh = Date.now();
    });
  }, 5000);
}

function startPulseApplication() {
  // Firepit can add a flat FS package to a SharePoint page after that page
  // has already completed DOMContentLoaded.  Waiting only for the event in
  // that case leaves the initial loader on screen indefinitely.  Keep this
  // guard here as well as below so a host cannot start two boot sequences.
  if (window.AEWTTR && window.AEWTTR.bootStarted) return;
  (async () => {
    window.AEWTTR.bootStarted = true;
    window.AEWTTR.bootComplete = false;
    window.AEWTTR.bootFailed = false;
    const bootVisualStartedAt = performance.now();
    ensureBootLogPanel();
    bootLog("Starting PULSE boot sequence...");
    try {
      bootLog("Detecting SharePoint site...");
      const siteUrl = await sharePointAdapter.detectSharePointSite();
      if (siteUrl) {
        // Share the detected site URL with standalone tools (Tickets, My Travel, etc.)
        // so they can connect to the same SharePoint site without needing _spPageContextInfo.
        try { localStorage.setItem("pulse_sp_site_url", siteUrl); } catch (e) { /* sandboxed */ }
        await bootSharePointMode(siteUrl);
      } else {
        bootLog("SharePoint site not detected. Falling back to local mode.", "error");
        window.AEWTTR.mode = "local";
        window.AEWTTR.db = aewttrLoadStore();
        // Local builds are a complete product sandbox. Keep the local identity
        // administrative even when an older localStorage snapshot said Member.
        window.AEWTTR.db.user.role = "Admin";
        window.AEWTTR.db.user.isAdmin = true;
      }
      await waitForBootMinimum(bootVisualStartedAt);
      bootLog("Rendering application shell...");
      if (window.AEWTTR.mode === "sharepoint" && typeof isCurrentUserPulseAssociated === "function" && !isCurrentUserPulseAssociated()) {
        bootLog("Current user is not associated with PULSE — showing access blocked screen.", "error");
        renderPulseAccessBlockedScreen();
        if (window.PULSELoader) window.PULSELoader.finish();
      } else {
        renderShell();
        if (window.PULSELoader) window.PULSELoader.finish();
        const bootUrl = new URL(location.href);
        if (!bootUrl.searchParams.get("page") && !location.hash) {
          const defaultRoute = pulseComputeDefaultRoute();
          navigate(defaultRoute);
        } else {
          renderPage();
        }
        window.addEventListener("popstate", renderPage);
        window.addEventListener("hashchange", renderPage);
        if (typeof initAuditLogStore === "function") initAuditLogStore();
        if (typeof logUserAction === "function" && window.AEWTTR.db && window.AEWTTR.db.user) {
          logUserAction({
            action: "Login",
            area: "System",
            summary: `${window.AEWTTR.db.user.name} opened PULSE`,
            route: `${location.pathname}${location.search}${location.hash}` || "#/dashboard"
          });
        }
        wireBackgroundRefresh();
        scheduleNotificationRefresh();
        setTimeout(() => showDigestModal(), 1200);
        // Best-effort daily send: no server exists to run this on a real
        // schedule, so whoever's copy of PULSE is open at or after 8am
        // triggers the org-wide digest for everyone. Checking only once at
        // boot means a tab left open from before 8am would never notice the
        // clock passing 8am — so also re-check on a slow recurring timer;
        // the function itself no-ops immediately outside SharePoint mode or
        // before 8am or once today's send is already claimed, so this is
        // cheap on every tick that isn't the actual trigger moment.
        setTimeout(() => {
          if (typeof runDailyActionItemDigestIfDue === "function") runDailyActionItemDigestIfDue();
        }, 1500);
        setInterval(() => {
          if (document.visibilityState !== "visible") return;
          if (typeof runDailyActionItemDigestIfDue === "function") runDailyActionItemDigestIfDue();
        }, 5 * 60 * 1000);
      }
      window.AEWTTR.bootComplete = true;
    } catch (bootErr) {
      window.AEWTTR.bootFailed = true;
      bootLog(`Fatal startup error: ${String(bootErr && bootErr.message || bootErr)}`, "error");
      const root = document.getElementById("aewttr-root");
      if (root) root.innerHTML = `<div style="padding:32px;font-family:sans-serif;max-width:600px;">
        <h2 style="color:#c00;margin:0 0 12px;">PULSE failed to start</h2>
        <p style="margin:0 0 8px;">A JavaScript error occurred during boot. Copy this message and report it to your app admin:</p>
        <pre style="background:#f4f4f4;padding:12px;border-radius:4px;white-space:pre-wrap;word-break:break-all;font-size:13px;">${escapeHtml ? escapeHtml(String(bootErr)) : String(bootErr)}</pre>
      </div>`;
    }
  })();
}

// Normal page loads reach this listener.  Flat FS packages injected into an
// already-ready SharePoint/Firepit document start immediately instead.
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", startPulseApplication, { once: true });
} else {
  startPulseApplication();
}
