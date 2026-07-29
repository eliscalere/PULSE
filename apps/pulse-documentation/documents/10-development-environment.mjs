/* Source content for document 10, "Development Environment & Scale Validation".

   Single source for both the controlled PDF and the searchable extraction:

     node apps/pulse-documentation/scripts/build-source-document.mjs 10

   Every ceiling and interval quoted here was read from the source in
   apps/PULSE/assets/js at the time of writing. If that code changes, update this
   document and regenerate; do not let the numbers drift. */

export const meta = {
  number: "10",
  slug: "10_PULSE_Development_Environment_and_Scale_Validation",
  title: "Development Environment & Scale Validation",
  runningHeader: "DEVELOPMENT ENVIRONMENT & SCALE VALIDATION",
  footer: "PULSE DEV & SCALE | VERSION 1.0 | JULY 2026",
  orientation: "portrait",
};

export const cover = {
  kicker: "FOLLOW-ON WORK",
  title: "Development Environment & Scale Validation",
  standfirst:
    "PULSE has no development environment. This document states the gap plainly, sets out how to stand up a development SharePoint site, how to load it heavily enough to find the breaking points, and how to evaluate Dataverse as an alternative store.",
  callout: {
    label: "IMPORTANT",
    text: "This is planned work, not a description of something that exists. Nothing in this document should be read as a capability PULSE has today.",
  },
  spine: [
    ["Status", "Not yet started"],
    ["Owner", "Unassigned"],
    ["Blocks", "Confident scale claims"],
    ["Output", "A written decision record"],
  ],
};

export const pages = [
  {
    kicker: "THE GAP",
    title: "01 / THERE IS NO DEVELOPMENT ENVIRONMENT",
    blocks: [
      { kind: "p", text: "PULSE is currently developed against one of two things: browser-local fallback state, which is not backed by SharePoint at all, or a real SharePoint site holding real records. There is no third option — no isolated site whose data can be destroyed and rebuilt freely, and no dataset large enough to show where the application stops behaving well." },
      { kind: "h4", text: "WHAT THIS MEANS TODAY" },
      {
        kind: "table",
        rows: [
          ["AREA", "CURRENT SITUATION"],
          ["Isolated environment", "None. No development or staging site is defined in configuration or documentation."],
          ["Automated tests", "None. There is no test directory and no test script."],
          ["Schema provisioning", "Exercised only against sites that already matter."],
          ["Scale evidence", "None. Behaviour past a few hundred records per list is unmeasured."],
          ["Rollback rehearsal", "Documented as a procedure but never practised against a disposable site."],
        ],
      },
      {
        kind: "callout",
        label: "CAUTION",
        text: "Because provisioning and destructive operations have only ever run against sites in use, a schema mistake currently has no safe place to be discovered. That is the most immediate reason to do this work, ahead of any performance question.",
      },
      { kind: "p", text: "Local fallback mode is useful for interface work and is what the screenshot captures use. It cannot substitute for a SharePoint site, because it exercises none of the REST transport, provisioning, permission, paging, or identity behaviour that carries the real risk." },
    ],
  },
  {
    kicker: "STEP ONE",
    title: "02 / STAND UP A DEVELOPMENT SITE",
    blocks: [
      { kind: "p", text: "The goal is a site nobody depends on, holding nothing real, that can be provisioned from empty and thrown away. Treat reaching that state as the first deliverable, before any load testing." },
      {
        kind: "ol",
        items: [
          "Request a separate SharePoint site collection for PULSE development, with the requesting team as owners. Confirm it is outside any retention or records scope that would make deletion difficult.",
          "Publish the current PULSE package to a page on that site.",
          "Run SharePoint Setup and let it provision the schema into an empty site. Record every list and column it creates, and every failure, because this is the first time that path is observed from a clean start.",
          "Confirm identity and role resolution behave as expected for at least one administrator and one ordinary member.",
          "Seed synthetic records only. Use neutral role display names and example.mil addresses, matching the convention already used for documentation captures.",
          "Write down how to rebuild the site from empty, so it can be reset between experiments rather than repaired.",
        ],
      },
      {
        kind: "callout",
        label: "NOTE",
        text: "The application already supports pointing at an explicit site through the manualSharePointSiteUrl value in app-config.js, which is empty today. That is the intended hook for aiming a local build at a development site.",
      },
      { kind: "p", text: "Do not copy production records into the development site. Generate synthetic data instead; the volumes needed for scale testing are far larger than anything worth copying, and copying imports real personal and project information into a site with weaker controls." },
    ],
  },
  {
    kicker: "KNOWN LIMITS",
    title: "03 / SCALE CEILINGS ALREADY IN THE CODE",
    blocks: [
      { kind: "p", text: "These are not hypotheses. They were read from the current source and are the specific things a large dataset would exercise first." },
      { kind: "h4", text: "PAGING AND THE REQUEST GUARD" },
      { kind: "p", text: "The REST adapter's item fetch follows SharePoint's continuation link in a loop bounded by a guard of 50 requests. When the guard is reached the loop simply stops and returns what it has. There is no error, no warning, and no marker on the result." },
      {
        kind: "table",
        rows: [
          ["LIST LOAD STYLE", "PAGE SIZE", "CEILING BEFORE SILENT TRUNCATION"],
          ["No explicit page size", "100 per request", "about 5,000 items"],
          ["Explicit page size of 500", "500 per request", "about 25,000 items"],
        ],
      },
      {
        kind: "callout",
        label: "IMPORTANT",
        text: "Silent truncation is the most dangerous behaviour listed in this document. Past the guard, the application shows a complete-looking view of an incomplete dataset. A scale test must check record counts against the list, not just confirm the page renders.",
      },
      { kind: "h4", text: "THE LIST VIEW THRESHOLD" },
      { kind: "p", text: "Three loads sort server-side on a date column: issues and tickets on their occurrence date, and the audit log on its action time. These are exactly the lists that grow without bound. Once a list passes SharePoint's 5,000-item view threshold, a sorted or filtered query against an unindexed column fails rather than degrading. Adding the right indexed columns is likely to be a prerequisite, not an optimisation." },
    ],
  },
  {
    kicker: "KNOWN LIMITS",
    title: "04 / THE REFRESH AND MEMORY MODEL",
    blocks: [
      { kind: "p", text: "PULSE holds the whole workspace in memory. Boot reads every PULSE list, plus a field-definition probe per list, and assembles a single object graph. Background refresh does not fetch deltas; it repeats that load and replaces the graph wholesale." },
      { kind: "h4", text: "HOW OFTEN THAT HAPPENS" },
      {
        kind: "ul",
        items: [
          "On a timer, gated so a successful refresh runs no more often than every five seconds, and only while the tab is visible.",
          "On navigation between areas.",
          "On the tab regaining focus.",
        ],
      },
      { kind: "p", text: "At small data volumes this is invisible and gives the application its live feel. The open question is what it costs once each list holds thousands of items: the request count multiplies, the parse and mapping work repeats, and each refresh allocates a fresh copy of the entire workspace." },
      {
        kind: "callout",
        label: "NOTE",
        text: "Whether the answer is delta loading, a longer interval, refreshing only the active area, or a different store is exactly what this evaluation should decide. It should not be decided from intuition, which is the current situation.",
      },
    ],
  },
  {
    kicker: "STEP TWO",
    title: "05 / BUILD A LARGE DATA SET",
    blocks: [
      { kind: "p", text: "Load the development site past the ceilings above, not up to comfortable numbers. The point is to find the failure, so the volumes should be deliberately unreasonable and then reduced until behaviour is acceptable." },
      {
        kind: "table",
        rows: [
          ["RECORD TYPE", "SUGGESTED VOLUME", "WHY THIS NUMBER"],
          ["Projects", "500 and 2,000", "Drives dashboard, overview, and every project roll-up"],
          ["Action items", "50,000 across projects", "Crosses both paging ceilings and the view threshold"],
          ["Risks", "5,000", "Loaded in full and grouped per project on every load"],
          ["Travel requests", "10,000", "Feeds My Travel filtering and the calendar"],
          ["Audit log", "100,000", "Sorted server-side; passes the threshold by a wide margin"],
          ["Issues and tickets", "20,000", "Sorted server-side and paged at 500"],
          ["Document review items", "5,000", "Bucketed into review columns during load"],
        ],
      },
      { kind: "p", text: "Generate the data with a script against the development site's REST endpoints so volumes are reproducible and the site can be rebuilt to a known state. Keep the generator in the repository alongside the capture script, so a later maintainer can recreate the conditions rather than trust a past result." },
      {
        kind: "callout",
        label: "CAUTION",
        text: "Create this data only on the development site. A generator pointed at the wrong site by an empty or stale configuration value is a realistic accident; require the target site to be passed explicitly rather than defaulted.",
      },
    ],
  },
  {
    kicker: "STEP THREE",
    title: "06 / WHAT TO MEASURE AND PROVOKE",
    blocks: [
      { kind: "h4", text: "MEASURE" },
      {
        kind: "ul",
        items: [
          "Time from page load to a usable interface, and the number of REST requests it took.",
          "Time for one background refresh, and whether refreshes begin to overlap or queue.",
          "Time to render the heaviest views: the tracker for a large project, the overview roll-up, the travel calendar, and the audit log.",
          "Memory held after boot and after an hour of idling with refreshes running.",
          "Time for a single save to complete under load, since saves compete with refreshes.",
        ],
      },
      { kind: "h4", text: "PROVOKE" },
      {
        kind: "ol",
        items: [
          "Truncation: compare counts shown in the application against counts in the list, on every list, past the guard.",
          "Threshold failures: sort and filter the audit log and issue lists past 5,000 items with and without indexed columns.",
          "Render stalls: open the largest project's tracker and confirm whether the interface remains responsive.",
          "Save behaviour under load: edit records rapidly while refreshes run, and confirm nothing is lost or reverted.",
          "Reload the page repeatedly under load, which is how a user reacts to slowness and doubles the request pressure.",
        ],
      },
      {
        kind: "callout",
        label: "FINAL CHECK",
        text: "Record the numbers, the site state, and the package version with every result. A performance claim without those three is not reproducible and will not survive the next change.",
      },
    ],
  },
  {
    kicker: "STEP FOUR",
    title: "07 / EVALUATING DATAVERSE",
    blocks: [
      { kind: "p", text: "Dataverse is worth evaluating as a store because the pressures above are the ones SharePoint Lists handle least well: server-side querying at volume, relational integrity between records, and predictable behaviour past the view threshold. Nothing in the current source references Dataverse; this would be new work with no existing commitment." },
      { kind: "h4", text: "WHAT TO COMPARE, USING THE SAME DATA SET" },
      {
        kind: "table",
        rows: [
          ["DIMENSION", "QUESTION TO ANSWER"],
          ["Query at volume", "Does filtering and sorting stay predictable past the volumes where lists fail?"],
          ["Paging", "Can a page of data be fetched without the application holding the whole workspace?"],
          ["Relationships", "Do project, task, and risk relationships become enforceable rather than mapped by code?"],
          ["Identity", "Does the existing page session still supply identity, or is a new authentication path required?"],
          ["Licensing", "What does each user cost, and who approves that?"],
          ["Migration", "Can existing records move without loss, and can it be reversed?"],
          ["Hosting", "Does the application still ship as a single-file package on a page?"],
        ],
      },
      {
        kind: "callout",
        label: "IMPORTANT",
        text: "Evaluate against the loaded development site, not a small proof of concept. A store change that is not measured under the volumes that motivated it proves nothing.",
      },
      { kind: "p", text: "Note the architectural consequence honestly: PULSE currently has no server, no database, and no separate sign-in, and its simplicity follows from that. Introducing Dataverse trades that simplicity for capability, and the evaluation should state the cost as clearly as the benefit." },
    ],
  },
  {
    kicker: "OUTPUT",
    title: "08 / WHAT A FINISHED EVALUATION PRODUCES",
    blocks: [
      { kind: "p", text: "The work is complete when the following exist and are recorded, not when the site has been built." },
      {
        kind: "ol",
        items: [
          "A development site that can be provisioned from empty and rebuilt to a known state, with the procedure written down.",
          "A data generator held in the repository, with the volumes it produces.",
          "Measured results for boot, refresh, render, memory, and save under load, each stamped with package version and site state.",
          "A list of confirmed failure points, with the volume at which each appears.",
          "A written recommendation on the refresh and paging model, with the evidence behind it.",
          "A written recommendation on Dataverse: adopt, reject, or revisit at a stated threshold, with the comparison behind it.",
        ],
      },
      {
        kind: "callout",
        label: "OPERATING STANDARD",
        text: "Until these results exist, PULSE should make no claim about the data volumes it supports. Stating a limit that has never been measured is worse than stating that the limit is unknown.",
      },
      { kind: "h4", text: "RELATED MATERIAL" },
      {
        kind: "ul",
        items: [
          "Standard Operating Procedures, SOP 10, for the controlled form of this procedure.",
          "Technical Reference, for the data model, refresh behaviour, and packaging this document measures.",
          "Document 09, Focused Tools & Package Delivery, since every focused package shares this data layer and inherits these ceilings.",
        ],
      },
    ],
  },
];
