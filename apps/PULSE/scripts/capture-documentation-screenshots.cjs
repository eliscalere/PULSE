#!/usr/bin/env node

/* Captures the current PULSE interface for the documentation library.
   These images are the documentation source of truth for docs/current.

   Every record seeded here is synthetic and non-identifying: neutral role
   display names, example.mil addresses, and no real project or personal data.

   Usage:
     npx serve apps/PULSE -l 8743      (or any static server on the same port)
     node apps/PULSE/scripts/capture-documentation-screenshots.cjs

   PULSE_BASE_URL   overrides the served app URL
   PULSE_SHOT_OUT   overrides the output directory
*/

const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");

const root = path.resolve(__dirname, "..");
const baseUrl = process.env.PULSE_BASE_URL || "http://127.0.0.1:8743/index.html";
const outDir = process.env.PULSE_SHOT_OUT || path.join(root, "docs", "current", "screenshots");
const viewport = { width: 1600, height: 1000 };

/* `before` names a state hook applied after navigation but before the shot. */
const shots = [
  { file: "01-dashboard.png", route: "dashboard", waitFor: ".dashboard-v2" },
  { file: "02-overview-portfolio.png", route: "overview", waitFor: ".overview-toolbar", before: "overviewTeam" },
  { file: "03-projects-workspaces.png", route: "projects", waitFor: ".pgroup-top-nav" },
  { file: "04-project-workspace.png", route: "projects/PULSE-01", waitFor: ".pgroup-workspace" },
  { file: "05-project-tracker.png", route: "projects/PULSE-01/tracker", waitFor: ".pgroup-workspace" },
  { file: "06-project-checklist.png", route: "projects/PULSE-01/checklist", waitFor: ".pgroup-workspace" },
  { file: "07-project-documents.png", route: "projects/PULSE-01/documents", waitFor: ".pgroup-workspace" },
  { file: "08-project-people.png", route: "projects/PULSE-01/people", waitFor: ".pgroup-workspace" },
  { file: "09-weekly-meeting.png", route: "weekly", waitFor: "#page-content", before: "meeting" },
  { file: "10-travel-request-form.png", route: "travel/submit", waitFor: ".travel-spo-layout" },
  { file: "11-travel-my-travel.png", route: "travel/mytravel", waitFor: ".travel-spo-layout" },
  { file: "12-travel-calendar.png", route: "travel/calendar", waitFor: ".travel-spo-layout" },
  { file: "13-travel-debrief.png", route: "travel/debrief", waitFor: ".travel-spo-layout" },
  { file: "14-document-review.png", route: "docreview", waitFor: ".docreview-page" },
  { file: "15-tickets.png", route: "tickets", waitFor: ".tickets-page" },
  { file: "16-admin.png", route: "admin", waitFor: ".tile-grid" },
  { file: "17-notification-settings.png", route: "notification-settings", waitFor: ".notif-settings-page" },
  { file: "18-logs.png", route: "logs", waitFor: "#page-content" },
];

async function seed(page) {
  await page.evaluate(async () => {
    const iso = (offset) => {
      const d = new Date();
      d.setDate(d.getDate() + offset);
      return d.toISOString().slice(0, 10);
    };
    const today = iso(0);
    const db = window.AEWTTR.db;

    /* Name goes on before seeding so generated task owners inherit the neutral
       identity, and the full user object goes on after, because the portfolio
       seeder replaces db.user and db.members wholesale. */
    db.user = { ...db.user, name: "Program Lead", email: "program.lead@example.mil" };
    if (typeof seedDemoPortfolio === "function") await seedDemoPortfolio(document.createElement("div"));

    db.user = {
      ...db.user,
      id: "u-admin",
      name: "Program Lead",
      email: "program.lead@example.mil",
      role: "Admin",
      isAdmin: true,
      notificationPrefs: {
        areas: ["Weekly", "Travel", "Document Review", "Projects", "Admin"],
        tone: "direct",
        channels: { email: true, teams: true },
      },
    };
    db.members = [
      { id: "u-admin", name: "Program Lead", email: "program.lead@example.mil", role: "Admin", spo: "Site Owner", isAdmin: true },
      { id: "u-pm", name: "Program Manager", email: "program.manager@example.mil", role: "Member", spo: "Member", isAdmin: false },
      { id: "u-eng", name: "Systems Engineer", email: "systems.engineer@example.mil", role: "Member", spo: "Member", isAdmin: false },
      { id: "u-ops", name: "Operations Lead", email: "operations.lead@example.mil", role: "Member", spo: "Member", isAdmin: false },
    ];

    db.projectPeople = db.projectPeople || {};
    db.projectPeople["PULSE-01"] = [
      { id: "pp1", type: "member", memberId: "u-admin", label: "Program Lead", role: "Lead", email: "program.lead@example.mil" },
      { id: "pp2", type: "member", memberId: "u-eng", label: "Systems Engineer", role: "Engineer", email: "systems.engineer@example.mil" },
      { id: "pp3", type: "member", memberId: "u-ops", label: "Operations Lead", role: "Operations", email: "operations.lead@example.mil" },
    ];

    /* Weekly meeting: an active session with attendance and a note feed. */
    db.weeklyMeeting = db.weeklyMeeting || { roster: [], rocks: [], sessions: [], currentSession: null, meetingStatus: "idle", projectMeetings: {} };
    db.weeklyMeeting.roster = ["u-admin", "u-pm", "u-eng", "u-ops"];
    db.weeklyMeeting.currentSession = {
      id: "wm-current",
      date: today,
      notesFeed: [
        { id: "n1", author: "Program Lead", date: today, time: "09:02", text: "Reviewed portfolio status; two workspaces moved to In Progress." },
        { id: "n2", author: "Systems Engineer", date: today, time: "09:14", text: "Integration checklist complete pending final review package." },
      ],
      guests: [],
      attendance: { "u-admin": "Here", "u-pm": "Here", "u-eng": "Here", "u-ops": "Here" },
      activity: [],
      sessionStatus: "active",
    };
    db.weeklyMeeting.sessions = [db.weeklyMeeting.currentSession];
    db.weeklyMeeting.meetingStatus = "active";

    /* Document review: one package in review with a pending reviewer. */
    db.docs = db.docs || { "Not Started": [], "In Review": [], Concurrence: [], "Final / Signed": [], Archive: [] };
    db.docs["In Review"] = [{
      id: "doc-1", title: "Integration Review Package", projectCode: "PULSE-01",
      submitter: "Program Lead", submitterEmail: "program.lead@example.mil",
      date: today, deadline: iso(5), type: "Review Package",
      reviewers: [
        { id: "rv1", name: "Systems Engineer", email: "systems.engineer@example.mil", decision: "Pending", note: "" },
        { id: "rv2", name: "Operations Lead", email: "operations.lead@example.mil", decision: "Concur", note: "No issues identified." },
      ],
      comments: [], reviewActivity: [], revisions: [], _column: "In Review", isArchived: false,
    }];

    /* Travel: the capture user owns one upcoming and one completed trip so both
       My Travel tabs have content, plus a teammate's trip for All Travel and the
       calendar. */
    db.travelRequests = [
      {
        id: "TR-0001", requester: "Program Lead", requesterEmail: "program.lead@example.mil",
        formMode: "Standard", requestType: "Standard", tripTitle: "Integration Working Group",
        travelers: [{ name: "Program Lead", email: "program.lead@example.mil" }],
        projectIds: ["PULSE-01"], destination: "Patuxent River, MD",
        start: iso(21), end: iso(24), allDay: true, startTime: "", endTime: "",
        purpose: "Attend the quarterly integration working group and close open interface actions.",
        impactIfNotApproved: "Interface actions remain open through the next reporting cycle.",
        alternatives: [], engineeringForm: null, cost: 1850, type: "TDY",
        notes: "", status: "Submitted", chargeObject: "", chargeObjectStatus: "Pending",
        customerConcurrenceStatus: "Pending", requiresConcurrence: true, updates: [],
      },
      {
        id: "TR-0002", requester: "Program Lead", requesterEmail: "program.lead@example.mil",
        formMode: "Standard", requestType: "Standard", tripTitle: "Program Review",
        travelers: [{ name: "Program Lead", email: "program.lead@example.mil" }],
        projectIds: ["PULSE-02"], destination: "Norfolk, VA",
        start: iso(-14), end: iso(-12), allDay: true, startTime: "", endTime: "",
        purpose: "Present program status and confirm next-quarter priorities.",
        impactIfNotApproved: "", alternatives: [], engineeringForm: null,
        cost: 1200, type: "TDY", notes: "", status: "Approved",
        chargeObject: "CO-2026-0142", chargeObjectStatus: "Assigned",
        customerConcurrenceStatus: "Concurred", requiresConcurrence: false, updates: [],
      },
      {
        id: "TR-0003", requester: "Operations Lead", requesterEmail: "operations.lead@example.mil",
        formMode: "Standard", requestType: "Standard", tripTitle: "Site Survey",
        travelers: [{ name: "Operations Lead", email: "operations.lead@example.mil" }],
        projectIds: ["PULSE-03"], destination: "San Diego, CA",
        start: iso(10), end: iso(13), allDay: true, startTime: "", endTime: "",
        purpose: "Conduct the pre-installation site survey and confirm facility readiness.",
        impactIfNotApproved: "", alternatives: [], engineeringForm: null,
        cost: 2100, type: "TDY", notes: "", status: "Approved",
        chargeObject: "CO-2026-0151", chargeObjectStatus: "Assigned",
        customerConcurrenceStatus: "Concurred", requiresConcurrence: false, updates: [],
      },
    ];
    db.debriefs = [];

    /* Support tickets across a couple of states. */
    db.tickets = [
      {
        id: "ISS-0001", title: "Tracker column filter resets on refresh",
        project: "PULSE-01", type: "Bug", status: "In Progress",
        opened: iso(-3), reporter: "Systems Engineer", reporterEmail: "systems.engineer@example.mil",
        detail: "The tracker status filter returns to All after a background refresh.",
        workaround: "Re-apply the filter after the refresh completes.",
        esdp: "", affected: ["PULSE-01"],
        updates: [{ date: iso(-1), author: "Program Lead", text: "Reproduced; scheduled for the next maintenance pass." }],
      },
      {
        id: "ISS-0002", title: "Request access to the travel calendar",
        project: "", type: "Access", status: "Open",
        opened: iso(-1), reporter: "Operations Lead", reporterEmail: "operations.lead@example.mil",
        detail: "Need read access to the shared travel calendar for planning.",
        workaround: "", esdp: "", affected: [], updates: [],
      },
    ];

    /* Belt and braces: rewrite any default identity the seeder may still have
       stamped onto generated records, so no capture shows a placeholder name. */
    const neutralize = (value) => (value === "Local User" || value === "Demo User" ? "Program Lead" : value);
    Object.values(db.ganttTasks || {}).forEach((tasks) => (tasks || []).forEach((task) => {
      if (!task) return;
      task.owner = neutralize(task.owner);
      task.assignee = neutralize(task.assignee);
    }));
    Object.values(db.engChecklists || {}).forEach((columns) => Object.values(columns || {}).forEach((tasks) => (tasks || []).forEach((task) => {
      if (!task) return;
      task.owner = neutralize(task.owner);
      task.assignee = neutralize(task.assignee);
    })));
    if (typeof normalizeStoreShape === "function") normalizeStoreShape(db);

    /* Boot-time audit entries were written before the capture identity existed,
       and normalizeStoreShape reloads them from localStorage, so replace the log
       outright with a representative neutral set and drop the stored copy. */
    try { localStorage.removeItem("aewttr_audit_log"); } catch (e) { /* sandboxed */ }
    const auditAt = (minutesAgo) => new Date(Date.now() - minutesAgo * 60000).toISOString();
    db.auditLog = [
      { id: "LOG-6", ts: auditAt(4), action: "Update", area: "Projects", summary: "Updated task \"Finalize interface spec\" on Radar Integration", detail: null, route: "projects/PULSE-01/tracker", recordId: "PULSE-01", actorEmail: "program.lead@example.mil", actorName: "Program Lead", actorRole: "Admin" },
      { id: "LOG-5", ts: auditAt(12), action: "Create", area: "Travel", summary: "Submitted travel request TR-0001 (Integration Working Group)", detail: null, route: "travel/submit", recordId: "TR-0001", actorEmail: "program.lead@example.mil", actorName: "Program Lead", actorRole: "Admin" },
      { id: "LOG-4", ts: auditAt(38), action: "Update", area: "Document Review", summary: "Recorded concurrence on Integration Review Package", detail: null, route: "docreview", recordId: "doc-1", actorEmail: "operations.lead@example.mil", actorName: "Operations Lead", actorRole: "Member" },
      { id: "LOG-3", ts: auditAt(75), action: "Create", area: "Weekly", summary: "Started the weekly meeting session", detail: null, route: "weekly", recordId: "wm-current", actorEmail: "program.lead@example.mil", actorName: "Program Lead", actorRole: "Admin" },
      { id: "LOG-2", ts: auditAt(190), action: "Update", area: "Admin", summary: "Assigned charge object CO-2026-0151 to TR-0003", detail: null, route: "travel", recordId: "TR-0003", actorEmail: "program.lead@example.mil", actorName: "Program Lead", actorRole: "Admin" },
      { id: "LOG-1", ts: auditAt(320), action: "Create", area: "Projects", summary: "Created project workspace Radar Integration", detail: null, route: "projects", recordId: "PULSE-01", actorEmail: "program.lead@example.mil", actorName: "Program Lead", actorRole: "Admin" },
    ];

    window.AEWTTR.state = window.AEWTTR.state || {};
    window.AEWTTR.state.overviewView = "Team";
    window.AEWTTR.state.overviewTeamTab = "portfolio";
    window.AEWTTR.state.meetingView = { global: "live" };
    window.AEWTTR.state.meetingLiveTab = { global: "room" };
    window.AEWTTR.state.meetingActive = { global: "u-admin" };
    /* The header (identity chip, badges) is built once by renderShell at boot,
       so it still shows the pre-seed user until the shell is rebuilt. */
    if (typeof renderShell === "function") renderShell();
    if (typeof renderPage === "function") renderPage();
  });
}

async function prepare(page, mode) {
  if (mode === "overviewTeam") {
    await page.evaluate(() => {
      window.AEWTTR.state.overviewView = "Team";
      window.AEWTTR.state.overviewTeamTab = "portfolio";
      if (typeof renderPage === "function") renderPage();
    });
  }
  if (mode === "meeting") {
    await page.evaluate(() => {
      window.AEWTTR.state.meetingView.global = "live";
      window.AEWTTR.state.meetingLiveTab.global = "room";
      if (typeof renderPage === "function") renderPage();
    });
  }
}

async function main() {
  fs.mkdirSync(outDir, { recursive: true });
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport, deviceScaleFactor: 2 });
  const failures = [];
  page.on("pageerror", (error) => failures.push(`pageerror: ${error.message}`));

  await page.goto(baseUrl, { waitUntil: "networkidle" });
  await page.waitForFunction(() => window.AEWTTR && window.AEWTTR.bootComplete === true, null, { timeout: 30000 });
  await seed(page);
  await page.waitForTimeout(600);

  const captured = [];
  for (const shot of shots) {
    await page.evaluate((route) => { window.location.hash = `#/${route}`; }, shot.route);
    await page.waitForTimeout(450);
    await prepare(page, shot.before);
    try {
      await page.waitForSelector(shot.waitFor, { timeout: 15000 });
    } catch {
      failures.push(`${shot.file}: selector "${shot.waitFor}" never appeared`);
      continue;
    }
    await page.waitForTimeout(350);
    /* Transient save/seed toasts would otherwise sit in the corner of the shot. */
    await page.evaluate(() => {
      document.querySelectorAll(".aewttr-toast-stack, .aewttr-toast").forEach((node) => node.remove());
    });
    await page.locator("#aewttr-root").screenshot({ path: path.join(outDir, shot.file) });
    captured.push(shot.file);
    console.log(`captured ${shot.file}  (#/${shot.route})`);
  }

  await browser.close();
  console.log(`\n${captured.length}/${shots.length} screenshots written to ${outDir}`);
  if (failures.length) {
    console.error(`\n${failures.length} problem(s):`);
    for (const failure of [...new Set(failures)]) console.error(`  - ${failure}`);
    process.exitCode = 1;
  }
}

module.exports = { seed, prepare, shots, viewport, baseUrl };

if (require.main === module) {
  main().catch((error) => { console.error(error); process.exit(1); });
}
