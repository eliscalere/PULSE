/* Sample project generator — Admin → SharePoint Setup.
   Creates one richly filled project via existing Repo.save paths so
   SharePoint mode gets real list rows. Projects only (no doc-review). */

(function () {
  "use strict";

  function isoDate(offsetDays) {
    const d = new Date();
    d.setHours(12, 0, 0, 0);
    d.setDate(d.getDate() + (offsetDays || 0));
    return d.toISOString().slice(0, 10);
  }

  function stamp() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
  }

  function makeId(prefix) {
    return typeof uid === "function" ? uid(prefix) : `${prefix}-${Math.random().toString(36).slice(2, 9)}`;
  }

  function progress(cb, message, pct) {
    if (typeof cb === "function") {
      try { cb({ message, pct }); } catch (_) { /* ignore */ }
    }
  }

  function pickNames(db) {
    const me = (db.user && (db.user.name || db.user.displayName)) || "Demo User";
    const members = (db.members || []).filter((m) => m && m.name);
    const extras = ["Alex Rivers", "Jordan Blake", "Casey Nguyen", "Morgan Patel"];
    const names = [me];
    members.forEach((m) => {
      if (names.length >= 4) return;
      if (!names.some((n) => n.toLowerCase() === m.name.toLowerCase())) names.push(m.name);
    });
    extras.forEach((n) => {
      if (names.length >= 4) return;
      if (!names.some((x) => x.toLowerCase() === n.toLowerCase())) names.push(n);
    });
    while (names.length < 4) names.push(extras[names.length % extras.length]);
    return { me, pm: names[0], engineer: names[1], isso: names[2], rangePoc: names[3] };
  }

  function note(author, text, daysAgo) {
    const base = {
      id: makeId("pn"),
      author: author || "SAMPLE",
      date: isoDate(-(daysAgo || 0)),
      time: "09:15",
      text
    };
    return typeof stampNoteAuthor === "function" ? stampNoteAuthor(base) : base;
  }

  function taskNote(author, text, daysAgo) {
    return {
      id: makeId("nt"),
      author: author || "SAMPLE",
      date: isoDate(-(daysAgo || 0)),
      time: "10:30",
      text
    };
  }

  function nextProjectCode(db) {
    const stampPart = stamp().replace(/-/g, "").slice(0, 12);
    let code = `SMP${stampPart}`;
    let n = 0;
    while ((db.projects || []).some((p) => p.id === code)) {
      n += 1;
      code = `SMP${stampPart}x${n}`;
    }
    return code;
  }

  async function saveKind(kind, obj, extra) {
    if (typeof Repo === "undefined" || !Repo || typeof Repo.save !== "function") {
      if (typeof aewttrSaveStore === "function") aewttrSaveStore();
      return;
    }
    await Repo.save(kind, obj, extra);
  }

  async function flushSaves() {
    if (typeof Repo !== "undefined" && Repo && typeof Repo.flush === "function") {
      await Repo.flush();
    }
  }

  /**
   * @param {{ onProgress?: (p:{message:string,pct?:number}) => void, navigateAfter?: boolean }} [opts]
   * @returns {Promise<{ project: object, id: string, name: string }>}
   */
  async function generateSampleProject(opts) {
    opts = opts || {};
    const onProgress = opts.onProgress;
    const db = window.AEWTTR && window.AEWTTR.db;
    if (!db) throw new Error("PULSE data store is not ready.");

    progress(onProgress, "Building sample project metadata…", 5);

    const people = pickNames(db);
    const today = isoDate(0);
    const start = isoDate(-21);
    const due = isoDate(45);
    const tsLabel = stamp();
    const projectId = nextProjectCode(db);
    const portfolios = ["SAMPLE — AEWTTR Demo Portfolio", "Surface Combatants"];
    const locations = ["SAMPLE — Pacific Test Range", "San Diego / Pt. Loma"];
    const contractorName = "SAMPLE Harbor Systems LLC";

    const proj = {
      id: projectId,
      name: `[SAMPLE] Radar Fit-Check IPT — ${tsLabel}`,
      team: "AEWTTR Systems",
      priority: "High",
      effort: "L",
      rag: "Amber",
      lifecycleStatus: "Active",
      updated: today,
      tools: "Tracker, Boards, Meeting, Risks",
      stakeholders: "Program Office; Ship's Force; Range Safety; SAMPLE contractors",
      sharepointFolderUrl: "",
      coverImage: "",
      images: [],
      description:
        "SAMPLE demo project for AEWTTR PULSE. Fictional radar fit-check / interface verification program — safe to edit or delete. Generated from Admin → SharePoint Setup.",
      scope:
        "Verify antenna interface, power/data harnesses, and ship-check readiness for a fictional next-port availability window. SAMPLE data only.",
      objectives:
        "1) Freeze interface ICD package\n2) Complete bench + ship-side fit check\n3) Close open risks to Amber/Green before sail gate",
      linkSubtaskDates: true,
      projectType: "Project",
      parentProjectId: "",
      pm: "",
      engineer: "",
      isso: "",
      rangePoc: "",
      contract: "SAMPLE-N00024-26-C-0440",
      contractor: contractorName,
      taskOrder: "TO-SAMPLE-07",
      fundingType: "RDT&E",
      fiscalYear: "FY26",
      fundingStatus: "Funded — partial",
      changeRequestRequired: true,
      configEndItem: "AN/SPY-SAMPLE Array Seat",
      program: "SAMPLE Surface Radar Modernization",
      product: "Multi-Band Fit Package",
      locations: locations.slice(),
      portfolios: portfolios.slice(),
      mainFolderUrl: "",
      exportSelections: {},
      startDate: start,
      dueDate: due,
      completionDate: ""
    };

    if (!db.projects) db.projects = [];
    db.projects.unshift(proj);

    const extra = typeof ensureProjectExtra === "function"
      ? ensureProjectExtra(projectId)
      : (db.projectExtra[projectId] = { status: "In Progress", history: [], risks: [], handoff: "", meetingNotes: "", notes: [] });

    extra.status = "In Progress";
    extra.handoff = "SAMPLE handoff brief: coordinate ship access with Range POC; hold harness drawing package at Rev C before ship call.";
    extra.meetingNotes = "SAMPLE standup note: ICD freeze targeted for next IPT; range window contingent on Amber→Green risk burn-down.";
    extra.history = [
      { date: isoDate(-14), author: people.me, note: "SAMPLE project kickoff — IPT baseline established." },
      { date: isoDate(-7), author: people.pm, note: "Interface design review complete; open items tracked under divider A." },
      { date: today, author: people.me, note: "Generated via Admin → SharePoint Setup → Generate sample project." }
    ];
    extra.notes = [
      note(people.me, "SAMPLE: Prefer this project for demos of tracker dividers, risks, and boards.", 0),
      note(people.engineer, "Bench power supply calibration logged — see tracker subitem notes.", 3),
      note(people.isso, "SAMPLE cyber: temporary lab VLAN ACL approved through end of month.", 5)
    ];

    // People roster first so roles can point at entry ids
    if (!db.projectPeople) db.projectPeople = {};
    const roster = [];
    function addPerson(label, role, email) {
      const known = (db.members || []).find((m) => m.name && m.name.toLowerCase() === String(label).toLowerCase());
      const entry = {
        id: makeId("ppj"),
        type: known ? "member" : "person",
        memberId: known ? known.id : "",
        label: known ? known.name : label,
        role: role || "",
        company: known ? "" : (role === "Contractor" ? contractorName : ""),
        email: known ? (known.email || "") : (email || "")
      };
      roster.push(entry);
      return entry;
    }
    const pmEntry = addPerson(people.pm, "Product Manager");
    const engEntry = addPerson(people.engineer, "Engineer");
    const issoEntry = addPerson(people.isso, "ISSO");
    const rangeEntry = addPerson(people.rangePoc, "Range POC");
    const contractorEntry = addPerson(contractorName, "Contractor", "sample.contracts@harbor-systems.example");
    if (typeof currentUserAsProjectPersonEntry === "function") {
      const me = currentUserAsProjectPersonEntry();
      if (me && !roster.some((p) => (me.memberId && p.memberId === me.memberId) || (p.label && me.label && p.label.toLowerCase() === me.label.toLowerCase()))) {
        roster.push(me);
      }
    }
    db.projectPeople[projectId] = roster;
    proj.pm = pmEntry.id;
    proj.engineer = engEntry.id;
    proj.isso = issoEntry.id;
    proj.rangePoc = rangeEntry.id;
    // Role pickers store people-entry ids; catalog remember uses the display name separately.
    proj.contractor = contractorEntry.id;

    // Optional local finance seed (not a separate SP list)
    if (typeof projectFinanceDb === "function") {
      const finance = projectFinanceDb(projectId);
      finance.summary = {
        fundingType: proj.fundingType,
        contract: proj.contract,
        taskOrder: proj.taskOrder,
        fiscalYear: proj.fiscalYear
      };
      finance.snapshot = {
        totalCeiling: 4200000,
        fundedToDate: 1850000,
        obligated: 920000,
        accrued: 410000,
        actuals: 365000,
        forecastAtComplete: 1980000
      };
      finance.clins = [{
        id: makeId("clin"),
        code: "0001AA",
        title: "SAMPLE — Fit-check engineering support",
        type: "Labor",
        funded: 1200000,
        obligated: 600000,
        spent: 240000,
        status: "Active"
      }];
      finance.watchlist = [{
        id: makeId("fw"),
        title: "SAMPLE — Confirm FY26 incremental funding drop",
        owner: people.pm,
        due: isoDate(14),
        status: "Open",
        note: "Demo watch item only."
      }];
    }

    progress(onProgress, "Saving project to SharePoint…", 20);
    await saveKind("project", proj);
    await flushSaves();

    // Catalogs used by pickers
    if (typeof rememberPortfolioNames === "function") rememberPortfolioNames(portfolios);
    if (typeof rememberLocationNames === "function") rememberLocationNames(locations);
    if (typeof rememberContractorNames === "function") rememberContractorNames([contractorName]);
    await flushSaves();

    progress(onProgress, "Seeding tracker (divider, tasks, subitems)…", 40);

    const dividerId = makeId("div");
    const dividerPartial = {
      id: dividerId,
      title: "A — Interface & Fit-Check Work Package (SAMPLE)",
      assignee: people.pm,
      health: "At Risk",
      status: "In Progress",
      isMilestone: true,
      pm: people.pm,
      engineer: people.engineer,
      isso: people.isso,
      rangePoc: people.rangePoc,
      contract: proj.contract,
      contractor: contractorName,
      taskOrder: "TO-SAMPLE-07A",
      fundingType: "RDT&E",
      fiscalYear: "FY26",
      fundingStatus: "Funded",
      changeRequestRequired: true,
      configEndItem: proj.configEndItem,
      program: proj.program,
      product: "ICD Rev C Package",
      locations: [locations[0]],
      portfolios: [portfolios[0]],
      priority: "High",
      rag: "Amber",
      lifecycleStatus: "Active",
      scope: "SAMPLE mini-project: interface freeze + range fitness gate.",
      objectives: "Deliver signed ICD and complete fit-check dry run.",
      description: "Project divider used as a milestone-flagged mini-project for demos.",
      startDate: start,
      dueDate: isoDate(30),
      completionDate: ""
    };
    const divider = typeof createTrackerDivider === "function"
      ? createTrackerDivider(dividerPartial)
      : Object.assign({ itemType: "divider", workItemLevel: "Divider", subtasks: [], notes: [], parentDividerId: "" }, dividerPartial);
    if (typeof syncDividerMetadata === "function") syncDividerMetadata(divider);

    const task1Id = makeId("g");
    const task2Id = makeId("g");
    const task3Id = makeId("g");

    const nestedSub = typeof normalizeTaskSubtask === "function"
      ? normalizeTaskSubtask
      : function (parent, sub) {
          return Object.assign({
            id: makeId("st"),
            text: "",
            assignee: "",
            done: false,
            health: "On Track",
            start: parent && parent.start,
            end: parent && parent.end,
            linked: true,
            notes: [],
            relatedDocs: [],
            subtasks: []
          }, sub || {});
        };

    const task1 = {
      id: task1Id,
      itemType: "task",
      workItemLevel: "Task",
      title: "Finalize interface control document (SAMPLE)",
      assignee: people.engineer,
      start: isoDate(-10),
      end: isoDate(7),
      status: "In Progress",
      health: "On Track",
      parentDividerId: dividerId,
      notes: [taskNote(people.engineer, "SAMPLE: Waiting on harness pinout from Harbor Systems.", 1)],
      subtasks: []
    };
    task1.subtasks.push(nestedSub(task1, {
      id: makeId("st"),
      text: "Collect as-built harness drawings",
      assignee: people.engineer,
      done: true,
      health: "On Track",
      start: isoDate(-10),
      end: isoDate(-5),
      notes: [taskNote(people.engineer, "Drawings received Rev B.", 5)]
    }));
    const nestedParent = nestedSub(task1, {
      id: makeId("st"),
      text: "Resolve open ICD RFCs",
      assignee: people.pm,
      done: false,
      health: "At Risk",
      start: isoDate(-3),
      end: isoDate(5),
      notes: [],
      subtasks: []
    });
    nestedParent.subtasks.push(nestedSub(nestedParent, {
      id: makeId("st"),
      text: "RFC-SAMPLE-12 power margin reply",
      assignee: people.engineer,
      done: false,
      health: "At Risk",
      start: isoDate(-2),
      end: isoDate(3)
    }));
    task1.subtasks.push(nestedParent);

    const task2 = {
      id: task2Id,
      itemType: "task",
      workItemLevel: "Task",
      title: "Bench electrical fit-check (SAMPLE)",
      assignee: people.engineer,
      start: isoDate(0),
      end: isoDate(14),
      status: "Not Started",
      health: "On Track",
      parentDividerId: dividerId,
      notes: [taskNote(people.isso, "SAMPLE: Confirm lab badge list before power-up.", 0)],
      subtasks: [
        nestedSub({ start: isoDate(0), end: isoDate(7), health: "On Track" }, {
          id: makeId("st"),
          text: "Stage test cables and loads",
          assignee: people.rangePoc,
          done: false,
          start: isoDate(0),
          end: isoDate(5)
        })
      ]
    };

    const task3 = {
      id: task3Id,
      itemType: "task",
      workItemLevel: "Task",
      title: "Ship-check logistics briefing (ungrouped SAMPLE)",
      assignee: people.pm,
      start: isoDate(10),
      end: isoDate(21),
      status: "Not Started",
      health: "On Track",
      parentDividerId: "",
      notes: [],
      subtasks: []
    };

    const normalizedTasks = [divider, task1, task2, task3].map((t) =>
      typeof normalizeGanttTask === "function" ? normalizeGanttTask(t) : t
    );
    if (!db.ganttTasks) db.ganttTasks = {};
    db.ganttTasks[projectId] = normalizedTasks;

    // Divider first so ParentDividerId can resolve _spId
    await saveKind("actionItem", normalizedTasks[0], { projectCode: projectId, source: "Tracker" });
    await flushSaves();
    for (let i = 1; i < normalizedTasks.length; i++) {
      await saveKind("actionItem", normalizedTasks[i], { projectCode: projectId, source: "Tracker" });
    }
    await flushSaves();

    if (typeof ensureAssigneesFromTask === "function") {
      for (const t of normalizedTasks) {
        try { await ensureAssigneesFromTask(proj, t); } catch (_) { /* ignore */ }
      }
    }
    if (typeof ensureDividerRolesOnProjectPeople === "function") {
      try { await ensureDividerRolesOnProjectPeople(proj, divider); } catch (_) { /* ignore */ }
    }

    progress(onProgress, "Seeding risks…", 65);

    const riskDefs = [
      {
        name: "SAMPLE — Range window slip",
        description: "Pacific Test Range availability may slip if Amber risks remain open past sail gate.",
        owner: people.pm,
        likelihood: 3,
        impact: 4,
        category: "Schedule",
        mitigationPlan: "Burn down ICD RFCs; hold contingency week in July.",
        responseStrategy: "Mitigate",
        due: isoDate(20),
        status: "Mitigating",
        portfolio: portfolios[0]
      },
      {
        name: "SAMPLE — Harness lead-time",
        description: "Contractor cable kit may miss bench window.",
        owner: people.engineer,
        likelihood: 2,
        impact: 3,
        category: "Contract",
        mitigationPlan: "Partial kit for electrical check; full kit before ship call.",
        responseStrategy: "Mitigate",
        due: isoDate(12),
        status: "Open",
        portfolio: portfolios[0]
      },
      {
        name: "SAMPLE — Lab VLAN ACL expiry",
        description: "Temporary cyber ACL expires end of month.",
        owner: people.isso,
        likelihood: 2,
        impact: 2,
        category: "Cyber",
        mitigationPlan: "Request 30-day extension or migrate to permanent lab VLAN.",
        responseStrategy: "Accept",
        due: isoDate(18),
        status: "Monitoring",
        portfolio: portfolios[1]
      }
    ];

    extra.risks = riskDefs.map((def) => {
      const risk = typeof normalizeRiskRecord === "function"
        ? normalizeRiskRecord(def, projectId)
        : Object.assign({ id: makeId("risk"), projectId }, def);
      risk.lastReviewedDate = isoDate(-2);
      risk.reviewNotes = [{
        id: makeId("rrn"),
        author: people.me,
        date: isoDate(-2),
        note: "SAMPLE seed review — leave open for demo."
      }];
      return risk;
    });

    for (const risk of extra.risks) {
      await saveKind("risk", risk);
    }
    await flushSaves();

    progress(onProgress, "Seeding project board…", 80);

    if (!db.checklistBoards) db.checklistBoards = [];
    if (!db.checklistTasks) db.checklistTasks = {};
    const boardId = makeId("CB");
    const statuses = typeof defaultBoardStatuses === "function"
      ? defaultBoardStatuses()
      : ["To Do", "In Progress", "Complete"];
    const board = {
      id: boardId,
      name: "SAMPLE Readiness Board",
      projectId,
      team: "IPT",
      type: "kanban",
      statuses: statuses.slice(),
      columnMeta: {},
      customFields: []
    };
    if (typeof ensureBoardColumnMeta === "function") ensureBoardColumnMeta(board);
    db.checklistBoards.unshift(board);
    db.checklistTasks[boardId] = {};
    statuses.forEach((s) => { db.checklistTasks[boardId][s] = []; });
    db.checklistTasks[boardId][statuses[0] || "To Do"].push({
      id: makeId("t"),
      title: "SAMPLE — Publish IPT briefing slide pack",
      owner: people.pm,
      due: isoDate(10),
      priority: "High",
      subtasks: [{ text: "Pull tracker status export", done: false }]
    });
    db.checklistTasks[boardId][statuses[1] || "In Progress"].push({
      id: makeId("t"),
      title: "SAMPLE — Confirm badge list for lab",
      owner: people.isso,
      due: isoDate(5),
      priority: "Medium",
      subtasks: []
    });
    db.checklistTasks[boardId][statuses[2] || "Complete"].push({
      id: makeId("t"),
      title: "SAMPLE — Kickoff agenda approved",
      owner: people.me,
      due: isoDate(-7),
      priority: "Low",
      subtasks: [{ text: "Send calendar hold", done: true }]
    });

    // Final project save persists PeopleJson, RisksJson, NotesJson, BoardsJson
    progress(onProgress, "Finalizing project (people, notes, boards)…", 90);
    proj.updated = today;
    await saveKind("project", proj);
    await flushSaves();

    if (typeof syncProjectPulseGroup === "function") {
      try { await syncProjectPulseGroup(proj); } catch (e) { console.warn("sample project group sync", e); }
    }

    if (typeof notifyLocalDataChanged === "function") notifyLocalDataChanged("sample-project");
    if (typeof logUserAction === "function") {
      try {
        logUserAction({
          action: "Create",
          area: "Admin",
          summary: `Generated sample project ${projectId} — ${proj.name}`
        });
      } catch (_) { /* ignore */ }
    }

    progress(onProgress, "Sample project ready.", 100);

    if (opts.navigateAfter !== false && typeof navigate === "function") {
      navigate("projects/" + projectId);
    }

    return { project: proj, id: projectId, name: proj.name };
  }

  window.generateSampleProject = generateSampleProject;
})();
