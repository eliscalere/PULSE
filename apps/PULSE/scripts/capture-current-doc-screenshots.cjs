#!/usr/bin/env node

/* Captures the current PULSE interface with non-identifying demo records.
   These images are the documentation source of truth for docs/current. */
const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");

const root = path.resolve(__dirname, "..");
const baseUrl = process.env.PULSE_BASE_URL || "http://127.0.0.1:8743/";
const outDir = path.join(root, "docs", "current", "screenshots");
const viewport = { width: 1600, height: 1000 };

const shots = [
  { file: "01-dashboard-current.png", route: "#/dashboard", waitFor: ".dashboard-v2" },
  { file: "02-overview-current.png", route: "#/overview", waitFor: ".overview-team-overview", before: "team" },
  { file: "03-project-tracker-current.png", route: "#/projects/PULSE-01/tracker", waitFor: ".tracker-shell" },
  { file: "04-weekly-meeting-current.png", route: "#/weekly", waitFor: ".meeting-toolbar" },
  { file: "05-travel-request-current.png", route: "#/travel/submit", waitFor: ".travel-wizard" },
  { file: "06-document-review-current.png", route: "#/docreview", waitFor: ".docreview-page" },
  { file: "07-admin-current.png", route: "#/admin", waitFor: ".tile-grid" },
  { file: "08-notifications-current.png", route: "#/notification-settings", waitFor: ".notif-settings-grid" }
];

async function seed(page) {
  await page.evaluate(async () => {
    const today = new Date().toISOString().slice(0, 10);
    const db = window.AEWTTR.db;
    db.user = {
      ...db.user, id: "u-admin", name: "Program Lead", email: "program.lead@example.mil", role: "Admin", isAdmin: true,
      notificationPrefs: { areas: ["Weekly", "Travel", "Document Review", "Projects", "Admin"], tone: "direct", channels: { email: true, teams: true } }
    };
    db.members = [
      { id: "u-admin", name: "Program Lead", email: "program.lead@example.mil", role: "Admin", spo: "Site Owner", isAdmin: true },
      { id: "u-pm", name: "Program Manager", email: "program.manager@example.mil", role: "Member", spo: "Member", isAdmin: false },
      { id: "u-eng", name: "Systems Engineer", email: "systems.engineer@example.mil", role: "Member", spo: "Member", isAdmin: false },
      { id: "u-ops", name: "Operations Lead", email: "operations.lead@example.mil", role: "Member", spo: "Member", isAdmin: false }
    ];
    if (typeof seedDemoPortfolio === "function") await seedDemoPortfolio(document.createElement("div"));
    db.projectPeople = db.projectPeople || {};
    db.projectPeople["PULSE-01"] = [
      { id: "pp1", type: "member", memberId: "u-admin", label: "Program Lead", role: "Lead", email: "program.lead@example.mil" },
      { id: "pp2", type: "member", memberId: "u-eng", label: "Systems Engineer", role: "Engineer", email: "systems.engineer@example.mil" }
    ];
    db.weeklyMeeting = db.weeklyMeeting || { roster: [], rocks: [], sessions: [], currentSession: null, meetingStatus: "idle", projectMeetings: {} };
    db.weeklyMeeting.roster = ["u-admin", "u-pm", "u-eng", "u-ops"];
    db.weeklyMeeting.currentSession = { id: "wm-current", date: today, notesFeed: [{ id: "note-current", author: "Program Lead", date: today, time: "09:00", text: "Current documentation capture session." }], guests: [], attendance: { "u-admin": "Here", "u-pm": "Here", "u-eng": "Here", "u-ops": "Here" }, activity: [], sessionStatus: "active" };
    db.weeklyMeeting.sessions = [db.weeklyMeeting.currentSession];
    db.weeklyMeeting.meetingStatus = "active";
    db.docs = db.docs || { "Not Started": [], "In Review": [], Concurrence: [], "Final / Signed": [], Archive: [] };
    db.docs["In Review"] = [{ id: "doc-current", title: "Current PULSE Review Package", projectCode: "PULSE-01", submitter: "Program Lead", submitterEmail: "program.lead@example.mil", date: today, deadline: today, type: "Review Package", reviewers: [{ id: "rev-current", name: "Systems Engineer", email: "systems.engineer@example.mil", decision: "Pending", note: "" }], comments: [], reviewActivity: [], revisions: [], _column: "In Review", isArchived: false }];
    if (typeof normalizeStoreShape === "function") normalizeStoreShape(db);
    window.AEWTTR.state = window.AEWTTR.state || {};
    window.AEWTTR.state.overviewView = "Team";
    window.AEWTTR.state.overviewTeamTab = "portfolio";
    window.AEWTTR.state.meetingView = { global: "live" };
    window.AEWTTR.state.meetingLiveTab = { global: "room" };
    window.AEWTTR.state.meetingActive = { global: "u-admin" };
    if (typeof renderPage === "function") renderPage();
  });
}

async function prepare(page, mode) {
  if (mode === "team") await page.evaluate(() => { window.AEWTTR.state.overviewView = "Team"; window.AEWTTR.state.overviewTeamTab = "portfolio"; if (typeof renderPage === "function") renderPage(); });
  if (mode === "meeting") await page.evaluate(() => { window.AEWTTR.state.meetingLiveTab.global = "project"; if (typeof renderPage === "function") renderPage(); });
}

async function main() {
  fs.mkdirSync(outDir, { recursive: true });
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport, deviceScaleFactor: 1 });
  await page.goto(baseUrl, { waitUntil: "networkidle" });
  await page.waitForFunction(() => window.AEWTTR && window.AEWTTR.bootComplete === true, null, { timeout: 20000 });
  await seed(page);
  for (const shot of shots) {
    await page.evaluate((route) => { window.location.hash = route; }, shot.route);
    await page.waitForTimeout(350);
    await prepare(page, shot.before);
    await page.waitForSelector(shot.waitFor, { timeout: 15000 });
    await page.locator("#aewttr-root").screenshot({ path: path.join(outDir, shot.file) });
  }
  await browser.close();
}

main().catch((error) => { console.error(error); process.exit(1); });
