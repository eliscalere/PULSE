/* Source content for document 11, "Interface Reference".

   The figures are the capture set produced by
   apps/PULSE/scripts/capture-documentation-screenshots.cjs. Every record shown
   is synthetic: neutral role display names, example.mil addresses, and no real
   project or personal data.

     node apps/pulse-documentation/scripts/build-source-document.mjs 11

   Regenerate this document whenever the capture set is regenerated, so the PDF,
   the searchable text, and the figures the reader shows stay in agreement. */

export const meta = {
  number: "11",
  slug: "11_PULSE_Interface_Reference",
  title: "Interface Reference",
  runningHeader: "INTERFACE REFERENCE",
  footer: "PULSE INTERFACE REFERENCE | VERSION 1.0 | JULY 2026",
  orientation: "portrait",
};

export const cover = {
  kicker: "VISUAL REFERENCE",
  title: "Interface Reference",
  standfirst:
    "Eighteen captures of the shipping build, grouped the way the application is navigated. Use these to confirm that a written step matches what a user actually sees.",
  callout: {
    label: "NOTE",
    text: "Every record shown is synthetic. These captures confirm routes, labels, and rendered state only — they are not evidence of a completed SharePoint transaction, of administrator-only control, or of notification delivery.",
  },
  spine: [
    ["Captures", "18 routes"],
    ["Identity", "Neutral role names"],
    ["Data", "Synthetic, example.mil"],
    ["Captured", "July 2026"],
  ],
};

export const pages = [
  {
    kicker: "SITUATIONAL AWARENESS",
    title: "01 / SITUATIONAL AWARENESS",
    blocks: [
      { kind: "p", text: "Where a session starts: personal work, then the portfolio-wide picture." },
      { kind: "figure", file: "/screenshots/01-dashboard.webp", caption: "Dashboard", meta: "#/dashboard", description: "Counts, quick access into each area, and assigned work across projects." },
      { kind: "figure", file: "/screenshots/02-overview-portfolio.webp", caption: "Overview — portfolio", meta: "#/overview", description: "Team and portfolio roll-up for status, workload, and approvals." },
    ],
  },
  {
    kicker: "PROJECT WORKSPACES",
    title: "02 / PROJECT WORKSPACES",
    blocks: [
      { kind: "p", text: "The project is the front door: people, files, work items, and reporting all sit in project context." },
      { kind: "figure", file: "/screenshots/03-projects-workspaces.webp", caption: "All workspaces", meta: "#/projects", description: "Project list with health and entry points into each workspace." },
      { kind: "figure", file: "/screenshots/04-project-workspace.webp", caption: "Workspace home", meta: "#/projects/<code>", description: "Project landing view with the workspace section rail." },
      { kind: "figure", file: "/screenshots/05-project-tracker.webp", caption: "Tracker", meta: "#/projects/<code>/tracker", description: "Tasks and milestones with owner, dates, health, and timeline and risk tabs." },
      { kind: "figure", file: "/screenshots/06-project-checklist.webp", caption: "Checklist", meta: "#/projects/<code>/checklist", description: "Column-based checklist for repeatable delivery steps." },
      { kind: "figure", file: "/screenshots/07-project-documents.webp", caption: "Documents", meta: "#/projects/<code>/documents", description: "Project document surface backed by the SharePoint library." },
      { kind: "figure", file: "/screenshots/08-project-people.webp", caption: "People", meta: "#/projects/<code>/people", description: "Assigned members and roles for the workspace." },
    ],
  },
  {
    kicker: "RECURRING OPERATIONS",
    title: "03 / RECURRING OPERATIONS",
    blocks: [
      { kind: "p", text: "The weekly rhythm and the formal review path." },
      { kind: "figure", file: "/screenshots/09-weekly-meeting.webp", caption: "Weekly meeting", meta: "#/weekly", description: "Live session with attendance, minutes, project updates, and around-the-room." },
      { kind: "figure", file: "/screenshots/14-document-review.webp", caption: "Document review", meta: "#/docreview", description: "Review packages moving through concurrence and signature." },
    ],
  },
  {
    kicker: "TRAVEL",
    title: "04 / TRAVEL",
    blocks: [
      { kind: "p", text: "Request, track, and close out travel, including the standalone Firepit tools." },
      { kind: "figure", file: "/screenshots/10-travel-request-form.webp", caption: "Travel request form", meta: "#/travel/submit", description: "Guided TDY, conference, training, and leave request intake." },
      { kind: "figure", file: "/screenshots/11-travel-my-travel.webp", caption: "My travel", meta: "#/travel/mytravel", description: "A traveler's own requests by state, with concurrence and charge-object status." },
      { kind: "figure", file: "/screenshots/12-travel-calendar.webp", caption: "Travel calendar", meta: "#/travel/calendar", description: "Team travel and events on a shared calendar." },
      { kind: "figure", file: "/screenshots/13-travel-debrief.webp", caption: "Travel debrief", meta: "#/travel/debrief", description: "Post-trip debrief capture against an approved request." },
    ],
  },
  {
    kicker: "SUPPORT AND ADMINISTRATION",
    title: "05 / SUPPORT AND ADMINISTRATION",
    blocks: [
      { kind: "p", text: "Intake, preferences, configuration, and the activity record." },
      { kind: "figure", file: "/screenshots/15-tickets.webp", caption: "Tickets", meta: "#/tickets", description: "Blockers, bugs, access needs, and questions tracked to resolution." },
      { kind: "figure", file: "/screenshots/16-admin.webp", caption: "Admin", meta: "#/admin", description: "SharePoint setup, users and roles, configuration, and diagnostics." },
      { kind: "figure", file: "/screenshots/17-notification-settings.webp", caption: "Notification settings", meta: "#/notification-settings", description: "Per-user areas, tone, and channel preferences." },
      { kind: "figure", file: "/screenshots/18-logs.webp", caption: "Activity log", meta: "#/logs", description: "Audit record of actions by actor, area, and time." },
    ],
  },
  {
    kicker: "USING THESE CAPTURES",
    title: "06 / HOW TO USE THIS REFERENCE",
    blocks: [
      { kind: "p", text: "These captures exist so a procedure can be checked against the product. When a written step and a screen disagree, the screen is the fact and the step needs correcting — or the application has changed and the capture set is stale." },
      { kind: "h4", text: "WHEN TO REGENERATE" },
      {
        kind: "ul",
        items: [
          "A route, tab, or label changes.",
          "A status vocabulary changes, since the captures show live values.",
          "A new area ships that a reader would look for here.",
        ],
      },
      { kind: "p", text: "Serve the application locally, run the capture script, and rebuild this document. The script exits non-zero if a route's expected selector never appears, so a blank capture is not written as a success." },
      {
        kind: "callout",
        label: "CAUTION",
        text: "Do not capture against a site holding real records. The capture identity and every seeded record must remain synthetic, because these figures are published in a document that circulates.",
      },
      { kind: "h4", text: "RELATED MATERIAL" },
      {
        kind: "ul",
        items: [
          "Internal User Guide, for the procedure each screen supports.",
          "Document 12, Process Flows, for how records move between the states these screens display.",
          "Screenshot Manifest in the documentation source, for the per-file capture record.",
        ],
      },
    ],
  },
];
