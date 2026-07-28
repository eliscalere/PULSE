#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");

const BASE_URL = process.env.PULSE_BASE_URL || "http://127.0.0.1:8743/";
const OUT_DIR = "/Users/eliscalere/Downloads/AEWTTR PAS/presentation-assets/screenshots/closing";
const VIEWPORT = { width: 1600, height: 1000 };

const SCREENSHOTS = [
  {
    file: "closing-weekly-meetings.png",
    route: "#/weekly",
    waitFor: ".meeting-live-grid",
    beforeShot: async (page) => {
      await page.evaluate(() => {
        window.AEWTTR.state.meetingLiveTab = window.AEWTTR.state.meetingLiveTab || {};
        window.AEWTTR.state.meetingLiveTab.global = "project";
        if (typeof renderPage === "function") renderPage();
      });
      await page.waitForSelector('[data-live-tab="project"].active');
    },
    note: "Weekly Meeting — live project review view with active meeting controls, project queue, and embedded tracker panel."
  },
  {
    file: "closing-project-tracker.png",
    route: "#/projects/PULSE-01/tracker",
    waitFor: ".tracker-shell",
    note: "Project workspace tracker tab for PULSE-01 showing task health, owners, and schedule detail."
  },
  {
    file: "closing-document-review.png",
    route: "#/docreview",
    waitFor: ".docreview-page",
    note: "Document Review workflow board with status columns, reviewer state, and review-package card."
  },
  {
    file: "closing-travel-request.png",
    route: "#/travel/submit",
    waitFor: ".travel-wizard",
    note: "Travel request workflow with multi-step submission shell and supported request types."
  },
  {
    file: "closing-notification-settings.png",
    route: "#/notification-settings",
    waitFor: ".notif-settings-grid",
    note: "Notification Settings screen for delivery channels, areas, tone, and test notification flow."
  }
];

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

async function seedDemoState(page) {
  await page.evaluate(async () => {
    const today = new Date().toISOString().slice(0, 10);
    const db = window.AEWTTR.db;
    const placeholderMembers = [
      { id: "u-admin", name: "Program Lead", email: "program.lead@example.mil", role: "Admin", spo: "Site Owner", powerbi: "Workspace Owner", isAdmin: true },
      { id: "u-pm", name: "Program Manager", email: "program.manager@example.mil", role: "Member", spo: "Member", powerbi: "Viewer", isAdmin: false },
      { id: "u-eng", name: "Systems Engineer", email: "systems.engineer@example.mil", role: "Member", spo: "Member", powerbi: "Viewer", isAdmin: false },
      { id: "u-ops", name: "Operations Lead", email: "operations.lead@example.mil", role: "Member", spo: "Member", powerbi: "Viewer", isAdmin: false },
      { id: "u-fin", name: "Finance Analyst", email: "finance.analyst@example.mil", role: "Finance Admin", spo: "Member", powerbi: "Viewer", isAdmin: false }
    ];

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
        channels: { email: true, teams: true }
      }
    };
    db.members = placeholderMembers.map((member) => ({ ...member, hiddenFromMeetings: false }));
    db.notificationConfig = { teamsEnabled: true, alsoSendEmail: true };

    if (typeof seedDemoPortfolio === "function") {
      await seedDemoPortfolio(document.createElement("div"));
    }

    if (!db.projectPeople) db.projectPeople = {};
    db.projectPeople["PULSE-01"] = [
      { id: "pp1", type: "member", memberId: "u-admin", label: "Program Lead", role: "Lead", company: "AEWTTR", email: "program.lead@example.mil" },
      { id: "pp2", type: "member", memberId: "u-eng", label: "Systems Engineer", role: "Engineer", company: "AEWTTR", email: "systems.engineer@example.mil" },
      { id: "pp3", type: "member", memberId: "u-pm", label: "Program Manager", role: "PM", company: "AEWTTR", email: "program.manager@example.mil" }
    ];

    if (!db.weeklyMeeting) db.weeklyMeeting = { roster: [], rocks: [], sessions: [], currentSession: null, meetingStatus: "idle", projectMeetings: {} };
    db.weeklyMeeting.roster = ["u-admin", "u-pm", "u-eng", "u-ops"];
    db.weeklyMeeting.rocks = [
      { id: "wr1", ownerId: "u-admin", ownerName: "Program Lead", title: "Finalize Monday executive snapshot", projectId: "PULSE-05", due: today, status: "On Track", notes: "", checkups: [] },
      { id: "wr2", ownerId: "u-eng", ownerName: "Systems Engineer", title: "Close open radar integration actions", projectId: "PULSE-01", due: today, status: "Off Track", notes: "", checkups: [{ id: "wrc1", authorName: "Systems Engineer", date: today, time: "09:15", note: "Waiting on interface validation notes." }] },
      { id: "wr3", ownerId: "u-pm", ownerName: "Program Manager", title: "Route document package for concurrence", projectId: "PULSE-04", due: today, status: "On Track", notes: "", checkups: [] }
    ];
    db.weeklyMeeting.currentSession = {
      id: "wm1",
      date: today,
      notesFeed: [
        { id: "note1", author: "Program Lead", date: today, time: "08:30", text: "Weekly session started and project queue is live." }
      ],
      guests: [],
      attendance: { "u-admin": "Here", "u-pm": "Here", "u-eng": "Here", "u-ops": "Here" },
      activity: [],
      sessionStatus: "active"
    };
    db.weeklyMeeting.sessions = [db.weeklyMeeting.currentSession];
    db.weeklyMeeting.meetingStatus = "active";

    if (!db.docs) db.docs = { "Not Started": [], "In Review": [], Concurrence: [], "Final / Signed": [], Archive: [] };
    db.docs["In Review"] = [{
      id: "doc-demo-1",
      title: "AEWTTR Workflow Review Package",
      projectCode: "PULSE-04",
      submitter: "Quality Lead",
      submitterEmail: "quality.lead@example.mil",
      date: today,
      deadline: new Date(Date.now() + 5 * 86400000).toISOString().slice(0, 10),
      type: "Review Package",
      reviewers: [
        { id: "rev1", name: "Program Lead", email: "program.lead@example.mil", decision: "Pending", note: "" },
        { id: "rev2", name: "Systems Engineer", email: "systems.engineer@example.mil", decision: "Approved", note: "Ready for routing." }
      ],
      comments: [],
      reviewActivity: [
        { date: today, author: "Quality Lead", action: "Submitted", note: "Initial review package uploaded." }
      ],
      revisions: [
        { id: "r1", fileName: "workflow-review-package-v1.pdf", uploadedBy: "Quality Lead", uploadedOn: today, inlineText: "Document preview placeholder." }
      ],
      _column: "In Review",
      isArchived: false
    }];

    if (typeof normalizeStoreShape === "function") {
      normalizeStoreShape(db);
    }

    window.AEWTTR.state = window.AEWTTR.state || {};
    window.AEWTTR.state.meetingView = window.AEWTTR.state.meetingView || {};
    window.AEWTTR.state.meetingView.global = "live";
    window.AEWTTR.state.meetingLiveTab = window.AEWTTR.state.meetingLiveTab || {};
    window.AEWTTR.state.meetingLiveTab.global = "room";
    window.AEWTTR.state.meetingActive = window.AEWTTR.state.meetingActive || {};
    window.AEWTTR.state.meetingActive.global = "u-admin";
    window.AEWTTR.state.meetingActiveProject = window.AEWTTR.state.meetingActiveProject || {};
    window.AEWTTR.state.meetingActiveProject.global = "PULSE-01";

    if (typeof renderPage === "function") renderPage();
  });
}

async function main() {
  ensureDir(OUT_DIR);

  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: VIEWPORT, deviceScaleFactor: 1 });

  await page.goto(BASE_URL, { waitUntil: "networkidle" });
  await page.waitForFunction(() => window.AEWTTR && window.AEWTTR.bootComplete === true, null, { timeout: 20000 });
  await page.waitForTimeout(300);
  await seedDemoState(page);

  const notes = [
    "PULSE Closing Collage Screenshots",
    "===================================",
    `Captured: ${new Date().toISOString()}`,
    `Source app: ${BASE_URL}`,
    `Viewport: ${VIEWPORT.width}x${VIEWPORT.height}`,
    "",
    "All screenshots were captured from the real local PULSE app shell using Playwright with browser chrome removed.",
    "Demo portfolio data was created through the app's own built-in demo seed path, and any added member identities use non-identifying placeholder names and example.mil addresses.",
    ""
  ];

  for (const shot of SCREENSHOTS) {
    await page.evaluate(() => {
      if (typeof closeModal === "function") {
        try { closeModal(); } catch (e) {}
      }
      document.querySelectorAll(".aewttr-modal-backdrop").forEach((node) => node.remove());
      document.body.classList.remove("aewttr-modal-open");
      document.documentElement.classList.remove("aewttr-modal-open");
    });
    await page.goto(`${BASE_URL}${shot.route}`, { waitUntil: "networkidle" });
    await page.waitForTimeout(350);
    if (shot.beforeShot) await shot.beforeShot(page);
    await page.waitForSelector(shot.waitFor, { timeout: 15000 });
    await page.waitForTimeout(250);
    const target = page.locator("#aewttr-root");
    await target.screenshot({
      path: path.join(OUT_DIR, shot.file)
    });
    notes.push(`${shot.file}`);
    notes.push(`  ${shot.note}`);
    notes.push("");
  }

  fs.writeFileSync(path.join(OUT_DIR, "README.txt"), notes.join("\n"));
  await browser.close();
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
