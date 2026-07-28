/* ---------- shared meeting engine (global + project meetings) ---------- */

PAGE_RENDERERS.weekly = function () {
  renderMeetingApp($("#page-content"), { type: "global" });
};

function weeklyMeetingDb() {
  if (!window.AEWTTR.db.weeklyMeeting) window.AEWTTR.db.weeklyMeeting = { roster: [], sessions: [], currentSession: null, meetingStatus: "idle" };
  return window.AEWTTR.db.weeklyMeeting;
}
function weeklyMemberById(id) { return window.AEWTTR.db.members.find(m => m.id === id); }
function weeklyProjectById(id) { return window.AEWTTR.db.projects.find(p => p.id === id); }

function weeklyMemberMatchesName(member, name) {
  if (!member || !name) return false;
  const a = member.name.toLowerCase();
  const b = String(name).toLowerCase();
  return a === b || a.startsWith(b + " ") || a.split(" ")[0] === b;
}
function weeklyRosterMembers() {
  return weeklyMeetingDb().roster.map(id => weeklyMemberById(id)).filter(m => m && !m.hiddenFromMeetings);
}
/* Renders the meeting-notes “post” feed — newest first, one card per post.
   editable=true wires Edit/Delete actions (used for the live meeting view);
   read-only history views (past meetings) pass editable=false. */
function meetingNoteCardHtml(note, editable) {
  return `
    <div class="meeting-note-card" data-note-id="${escapeHtml(note.id || "")}">
      <div class="meeting-note-head">
        <span>${userAvatarHtml(note.author, memberEmailForPerson(note.author), 22)}</span>
        <strong>${escapeHtml(note.author)}</strong>
        <span class="meeting-note-time">${fmtDate(note.date)}${note.time ? ` · ${escapeHtml(note.time)}` : ""}</span>
        ${editable ? `
          <button class="meeting-note-action" data-edit-note="${escapeHtml(note.id || "")}" title="Edit"><i class="bx bx-edit"></i></button>
          <button class="meeting-note-action" data-delete-note="${escapeHtml(note.id || "")}" title="Delete"><i class="bx bx-trash"></i></button>
        ` : ""}
      </div>
      <div class="meeting-note-text" data-note-text>${escapeHtml(note.text)}</div>
    </div>`;
}
function renderMeetingNotesFeed(mount, session, scope, onChange) {
  const feed = session.notesFeed || [];
  mount.innerHTML = feed.length
    ? feed.map(note => meetingNoteCardHtml(note, true)).join("")
    : `<div class="empty-state" style="padding:14px 4px;">No notes posted yet — be the first.</div>`;

  $all("[data-edit-note]", mount).forEach(btn => btn.addEventListener("click", () => {
    const card = btn.closest(".meeting-note-card");
    const note = feed.find(n => n.id === btn.dataset.editNote);
    if (!note) return;
    const textEl = $("[data-note-text]", card);
    textEl.outerHTML = `
      <div data-note-text>
        <textarea class="textarea-aewttr" data-note-edit-input style="min-height:50px;margin-bottom:6px;">${escapeHtml(note.text)}</textarea>
        <div style="display:flex;gap:6px;">
          <button class="btn-aewttr btn-aewttr-sm" data-save-note="${escapeHtml(note.id)}">Save</button>
          <button class="btn-aewttr-ghost btn-aewttr-sm" data-cancel-note>Cancel</button>
        </div>
      </div>`;
    $("[data-cancel-note]", card).addEventListener("click", () => renderMeetingNotesFeed(mount, session, scope, onChange));
    $("[data-save-note]", card).addEventListener("click", async () => {
      const newText = $("[data-note-edit-input]", card).value.trim();
      if (!newText) { toast("Note can't be empty — delete it instead if you want it gone.", "error"); return; }
      note.text = newText;
      await saveMeetingSession(scope);
      renderMeetingNotesFeed(mount, session, scope, onChange);
      toast("Note updated", "success");
    });
  }));
  $all("[data-delete-note]", mount).forEach(btn => btn.addEventListener("click", async () => {
    const ok = await confirmDialog({ title: "Delete note", message: "Delete this note?", confirmLabel: "Delete", danger: true });
    if (!ok) return;
    session.notesFeed = feed.filter(n => n.id !== btn.dataset.deleteNote);
    await saveMeetingSession(scope);
    renderMeetingNotesFeed(mount, session, scope, onChange);
    toast("Note deleted", "success");
  }));
}

function participantEmail(participant) {
  if (!participant) return "";
  if (participant.memberId) {
    const member = weeklyMemberById(participant.memberId);
    if (member && member.email) return member.email;
  }
  return memberEmailForPerson(participant.name, "");
}

function projectParticipantLabel(participant) {
  if (!participant) return "Unknown participant";
  if (participant.type === "member" && participant.memberId) {
    const member = weeklyMemberById(participant.memberId);
    return member ? member.name : participant.label;
  }
  return participant.label;
}
function projectParticipantRole(participant) {
  const base = participant.role || (participant.type === "company" ? "Company" : participant.type === "person" ? "Contractor" : "Team Member");
  return participant.company ? `${base} · ${participant.company}` : base;
}

/* ---------- scope helpers ---------- */
function meetingScopeKey(scope) { return scope.type === "global" ? "global" : "project:" + scope.project.id; }
function currentUserName() { return window.AEWTTR.db.user.name; }
function currentMeetingUser() {
  const user = window.AEWTTR.db.user || {};
  return {
    name: String(user.name || "").trim(),
    email: String(user.email || "").trim().toLowerCase()
  };
}
function participantMatchesCurrentUser(participant) {
  const me = currentMeetingUser();
  if (!me.name && !me.email) return false;
  if (participant.memberId) {
    const member = weeklyMemberById(participant.memberId);
    if (member) {
      if (me.email && member.email && String(member.email).trim().toLowerCase() === me.email) return true;
      if (me.name && member.name === me.name) return true;
    }
  }
  return !!(me.name && participant.name === me.name);
}
function currentMeetingParticipant(scope) {
  return meetingParticipants(scope).find(participantMatchesCurrentUser) || null;
}
function initializeMeetingAttendance(session, scope) {
  if (!session) return session;
  session.attendance = session.attendance || {};
  if (Object.keys(session.attendance).length > 0) return session;
  const participants = meetingParticipants(scope);
  const starter = currentMeetingParticipant(scope);
  participants.forEach((participant) => {
    session.attendance[participant.id] = starter && participant.id === starter.id ? "Here" : "Out";
  });
  if (starter) session.startedBy = starter.name;
  return session;
}
function nowStamp() {
  const d = new Date();
  const pad = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function logMeetingActivity(scope, text, meta) {
  const session = activeMeetingSession(scope);
  if (!session) return;
  session.activity = session.activity || [];
  const entry = {
    id: (meta && meta.id) || (typeof uid === "function" ? uid("act") : `act-${Date.now()}`),
    time: nowStamp(),
    at: Date.now(),
    text: String(text || "").trim(),
    actor: (meta && meta.actor) || currentUserName(),
    type: (meta && meta.type) || "info",
    projectId: (meta && meta.projectId) || "",
    taskId: (meta && meta.taskId) || "",
    taskTitle: (meta && meta.taskTitle) || "",
    subitemText: (meta && meta.subitemText) || "",
    detail: (meta && meta.detail) || ""
  };
  if (!entry.text) return;
  session.activity.unshift(entry);
}
function findActiveMeetingScopeForProject(projectId) {
  const db = window.AEWTTR.db;
  if (!db || !db.weeklyMeeting) return null;
  if (projectId) {
    const pm = db.weeklyMeeting.projectMeetings && db.weeklyMeeting.projectMeetings[projectId];
    if (pm && pm.meetingStatus === "active" && pm.currentSession && pm.currentSession.sessionStatus !== "ended") {
      const project = weeklyProjectById(projectId) || { id: projectId, name: projectId };
      return { type: "project", project };
    }
  }
  if (db.weeklyMeeting.meetingStatus === "active" && db.weeklyMeeting.currentSession && db.weeklyMeeting.currentSession.sessionStatus !== "ended") {
    return { type: "global" };
  }
  return null;
}
function recordActiveMeetingTaskChange(projectId, task, field, value, opts) {
  const scope = findActiveMeetingScopeForProject(projectId);
  if (!scope || !task) return;
  const actor = currentUserName();
  const title = (task.title || "Untitled").trim() || "Untitled";
  const projectLabel = projectId || (scope.type === "project" ? scope.project.id : "");
  const where = projectLabel ? ` in ${projectLabel}` : "";
  const prev = opts && opts.prev;
  if (field === "reviewStatus") {
    if (value === "Updated") {
      logMeetingActivity(scope, `${actor} updated the status of ${title}${where}.`, {
        type: "status", actor, projectId: projectLabel, taskId: task.id, taskTitle: title, detail: value
      });
    } else if (value === "Reviewed - No Change") {
      logMeetingActivity(scope, `${actor} reviewed ${title}${where} — no change.`, {
        type: "review", actor, projectId: projectLabel, taskId: task.id, taskTitle: title, detail: value
      });
    }
    return;
  }
  if (field === "health" && value && value !== prev) {
    logMeetingActivity(scope, `${actor} marked ${title}${where} as ${value}.`, {
      type: "status", actor, projectId: projectLabel, taskId: task.id, taskTitle: title, detail: value
    });
    return;
  }
  if (field === "status" && value && value !== prev) {
    logMeetingActivity(scope, `${actor} updated the status of ${title}${where} to ${value}.`, {
      type: "status", actor, projectId: projectLabel, taskId: task.id, taskTitle: title, detail: value
    });
  }
}
function recordActiveMeetingNotePosted(projectId, noteText, context) {
  const scope = findActiveMeetingScopeForProject(projectId || (context && context.projectId));
  if (!scope) return;
  const session = activeMeetingSession(scope);
  const actor = currentUserName();
  const snippet = String(noteText || "").trim().replace(/\s+/g, " ");
  const short = snippet.length > 120 ? `${snippet.slice(0, 117)}…` : snippet;
  if (session && session.activity && short) {
    const already = session.activity.some((a) => a && a.type === "note" && a.detail === short);
    if (already) return;
  }
  const kind = (context && context.kind) || "meeting";
  const taskTitle = (context && context.taskTitle) || "";
  const subitemText = (context && context.subitemText) || "";
  const pid = projectId || (context && context.projectId) || "";
  let text;
  if (kind === "subitem" && taskTitle) {
    text = `${actor} added a note to ${taskTitle} · Subitem · ${subitemText || "Untitled"}${short ? ` saying “${short}”` : "."}`;
  } else if (kind === "task" && taskTitle) {
    text = `${actor} added a note to ${taskTitle}${short ? ` saying “${short}”` : "."}`;
  } else if (kind === "project" && pid) {
    text = `${actor} added a note to project ${pid}${short ? ` saying “${short}”` : "."}`;
  } else {
    text = `${actor} added a meeting note${short ? ` saying “${short}”` : "."}`;
  }
  logMeetingActivity(scope, text, {
    type: "note",
    actor,
    projectId: pid,
    taskId: (context && context.taskId) || "",
    taskTitle,
    subitemText,
    detail: short
  });
}
function meetingProjectCode(scope) { return scope.type === "project" ? scope.project.id : undefined; }
async function saveMeetingSession(scope) {
  const session = meetingData(scope).currentSession;
  if (!session) return;
  return Repo.save("meetingSession", session, { projectCode: meetingProjectCode(scope) });
}
async function saveMeetingTask(scope, task, pid) {
  if (typeof ensureAssigneesFromTask === "function") {
    try { await ensureAssigneesFromTask(pid, task); } catch (e) { console.warn("project people sync", e); }
  }
  await Promise.all([
    Repo.save("actionItem", task, { projectCode: pid, source: "Tracker" }),
    saveMeetingSession(scope)
  ]);
}
function meetingTrackerStateKey(scope, projectId) {
  return `meeting:${meetingScopeKey(scope)}:${projectId}`;
}

/* Passed to renderTrackerWorkspace as config.onOpenEditor, which otherwise
   defaults to opening the real Gantt task/subtask side panel itself — this
   override existed to route the redraw through the meeting page instead,
   but never actually opened anything (it only flipped an expanded-state
   flag nothing reads), so the "Edit" button in the meeting's tracker view
   silently did nothing. Mirror renderTrackerWorkspace's own default here:
   open the same side panel Tracker uses, just redrawing the meeting view
   (not the Tracker page) on save. */
function openMeetingTaskEditor(scope, pid, taskId, redraw, subIndex) {
  const db = window.AEWTTR.db;
  const tasks = db.ganttTasks[pid] || [];
  const task = tasks.find((t) => t.id === taskId);
  if (!task) return;
  const editorOpts = { presentation: "modal" };
  if (subIndex != null && subIndex !== "") {
    // Paths like "0.1.2" or numeric indexes — resolve nested when possible.
    let resolvedIndex = subIndex;
    if (typeof subIndex === "string" && subIndex.includes(".")) {
      const top = Number(String(subIndex).split(".")[0]);
      if (!Number.isNaN(top)) resolvedIndex = top;
    } else if (typeof subIndex === "string") {
      resolvedIndex = Number(subIndex);
    }
    if (!Number.isNaN(resolvedIndex) && typeof openSubtaskSidePanel === "function") {
      openSubtaskSidePanel(task, tasks, redraw, resolvedIndex, false, editorOpts);
      return;
    }
  }
  openTaskSidePanel(task, tasks, redraw, editorOpts);
}
function ensureProjectMeeting(pid) {
  const db = window.AEWTTR.db;
  if (!db.weeklyMeeting.projectMeetings) db.weeklyMeeting.projectMeetings = {};
  if (!db.weeklyMeeting.projectMeetings[pid]) {
    db.weeklyMeeting.projectMeetings[pid] = { sessions: [], currentSession: null, meetingStatus: "idle" };
  }
  return db.weeklyMeeting.projectMeetings[pid];
}
function meetingData(scope) {
  return scope.type === "global" ? weeklyMeetingDb() : ensureProjectMeeting(scope.project.id);
}
function meetingStatus(scope) {
  const data = meetingData(scope);
  if (data.meetingStatus) return data.meetingStatus;
  if (data.currentSession) return data.currentSession.sessionStatus === "ended" ? "ended" : "active";
  return (data.sessions && data.sessions.length) ? "ended" : "idle";
}
function meetingIsActive(scope) {
  const data = meetingData(scope);
  return meetingStatus(scope) === "active" && !!data.currentSession;
}
function activeMeetingSession(scope) {
  return meetingIsActive(scope) ? meetingData(scope).currentSession : null;
}
function meetingSession(scope) {
  return activeMeetingSession(scope);
}
function syncMeetingAttendance(session, scope) {
  if (!session) return session;
  session.attendance = session.attendance || {};
  meetingParticipants(scope).forEach((participant) => {
    if (!session.attendance[participant.id]) session.attendance[participant.id] = "Out";
  });
  return session;
}
async function startMeeting(scope, onDone) {
  if (!canManageMeetings()) {
    toast("Only Meeting Admins can start meetings.", "error");
    return;
  }
  const data = meetingData(scope);
  const planned = data.currentSession && data.currentSession.sessionStatus === "planned" ? data.currentSession : null;
  data.meetingStatus = "active";
  data.currentSession = planned || {
    id: uid(scope.type === "global" ? "wm" : "pmc"),
    date: new Date().toISOString().slice(0, 10),
    startedAt: Date.now(),
    notesFeed: [],
    guests: [],
    attendance: {},
    activity: [],
    sessionStatus: "active"
  };
  data.currentSession.sessionStatus = "active";
  data.currentSession.startedAt = data.currentSession.startedAt || Date.now();
  initializeMeetingAttendance(data.currentSession, scope);
  logMeetingActivity(scope, `${currentUserName()} started the meeting.`, { type: "lifecycle" });
  await Repo.save("meetingSession", data.currentSession, { projectCode: meetingProjectCode(scope), immediate: true });
  if (Repo.flush) await Repo.flush();
  toast("Meeting started", "success");
  if (onDone) onDone();
}
async function ensurePlanningSession(scope) {
  const data = meetingData(scope);
  if (data.currentSession && data.currentSession.sessionStatus === "planned") return data.currentSession;
  data.meetingStatus = "planned";
  data.currentSession = { id: uid(scope.type === "global" ? "wm" : "pmc"), date: new Date().toISOString().slice(0, 10), agenda: [], notesFeed: [], guests: [], attendance: {}, activity: [], sessionStatus: "planned" };
  await Repo.save("meetingSession", data.currentSession, { projectCode: meetingProjectCode(scope), immediate: true });
  return data.currentSession;
}
function meetingParticipants(scope) {
  const db = window.AEWTTR.db;
  if (scope.type === "global") {
    const roster = weeklyRosterMembers().map(m => ({ id: m.id, memberId: m.id, name: m.name, role: m.role, isMember: true, isGuest: false }));
    const session = activeMeetingSession(scope);
    const guests = session ? meetingSessionGuests(session).map(g => ({
      id: g.id,
      memberId: "",
      name: g.name,
      role: g.email ? `Guest · ${g.email}` : "Guest",
      isMember: false,
      isGuest: true
    })) : [];
    return roster.concat(guests);
  }
  const people = db.projectPeople[scope.project.id] || [];
  return people.map(p => ({
    id: p.id,
    memberId: p.memberId || "",
    name: projectParticipantLabel(p),
    role: projectParticipantRole(p),
    isMember: p.type === "member",
    isGuest: false
  }));
}
function ensureMeetingSession(scope) {
  const session = activeMeetingSession(scope);
  if (!session) return null;
  migrateMeetingSessionNotes(session);
  initializeMeetingAttendance(session, scope);
  return syncMeetingAttendance(session, scope);
}
function meetingTaskGroups(scope, participant) {
  const db = window.AEWTTR.db;
  const member = participant.memberId ? weeklyMemberById(participant.memberId) : null;
  const projectIds = scope.type === "project" ? [scope.project.id] : db.projects.map(p => p.id);
  const groups = [];
  projectIds.forEach(pid => {
    const tasks = member ? (db.ganttTasks[pid] || []).filter(t => weeklyMemberMatchesName(member, t.assignee)) : [];
    const isCurrentProject = scope.type === "project" && pid === scope.project.id;
    if (tasks.length || isCurrentProject) {
      const project = weeklyProjectById(pid);
      if (project) groups.push({ project, tasks });
    }
  });
  return groups;
}
function syncMeetingPresenceChrome(body, session, opts) {
  const { canFacilitate, allParticipants, me } = opts;
  const hereCount = allParticipants.filter(p => session.attendance[p.id] === "Here").length;
  const headSpan = $(".meeting-nav-head span", body);
  if (headSpan) {
    if (canFacilitate) headSpan.textContent = `${hereCount}/${allParticipants.length} present`;
    else if (me) headSpan.textContent = session.attendance[me.id] === "Here" ? "You're in the meeting" : "Join when ready";
  }
  const presentLabel = $(".meeting-projects-present", body);
  if (presentLabel) {
    if (canFacilitate) presentLabel.textContent = `${hereCount}/${allParticipants.length} present`;
    else if (me) presentLabel.textContent = session.attendance[me.id] === "Here" ? "You're in the meeting" : "Join when ready";
  }
  if (me) {
    const myHere = session.attendance[me.id] === "Here";
    const joinBtn = $("#mtg-join-self", body);
    if (joinBtn) joinBtn.style.display = myHere ? "none" : "";
  }
}
function setPresencePillState(btn, status) {
  const here = status === "Here";
  btn.classList.toggle("here", here);
  btn.classList.toggle("out", !here);
  btn.textContent = here ? "Here" : "Out";
  btn.setAttribute("aria-pressed", here ? "true" : "false");
}
async function toggleMeetingPresence(scope, body, session, participantId, opts) {
  const { canFacilitate, me, allParticipants } = opts;
  if (!canFacilitate && me && participantId !== me.id) return;
  const btn = safeQueryByAttr(body, "data-presence", participantId);
  if (btn && btn.disabled) return;
  const previous = session.attendance[participantId] === "Here" ? "Here" : "Out";
  const next = previous === "Here" ? "Out" : "Here";
  session.attendance[participantId] = next;
  if (btn) {
    setPresencePillState(btn, next);
    btn.classList.add("meeting-presence-pill--pressed");
    setTimeout(() => btn.classList.remove("meeting-presence-pill--pressed"), 140);
  }
  syncMeetingPresenceChrome(body, session, { canFacilitate, allParticipants, me });
  try {
    await saveMeetingSession(scope);
  } catch (err) {
    session.attendance[participantId] = previous;
    if (btn) setPresencePillState(btn, previous);
    syncMeetingPresenceChrome(body, session, { canFacilitate, allParticipants, me });
    toast("Couldn't save attendance — try again.", "error");
  }
}
function openMeetingAttendanceModal(scope, onDone) {
  const session = ensureMeetingSession(scope);
  const participants = meetingParticipants(scope);
  const canFacilitate = canManageMeetings();
  const me = currentMeetingParticipant(scope);
  const modal = openModal(`
    <div class="aewttr-modal-head">
      <div><h3 style="margin:0;">Attendance</h3><p style="margin:4px 0 0;font-size:12px;color:var(--aewttr-muted);">Mark who is in the meeting.</p></div>
      <div style="display:flex;align-items:center;gap:8px;">
        ${canFacilitate && scope.type === "global" ? `<button class="btn-aewttr-ghost btn-aewttr-sm" id="mtg-attendance-add-guest"${tip("Add a guest or outside attendee")}><i class="bx bx-user-plus"></i> Guest</button>` : ""}
        <button class="aewttr-modal-close" type="button" aria-label="Close">&times;</button>
      </div>
    </div>
    <div class="aewttr-modal-body" id="mtg-attendance-list"></div>
  `, "aewttr-modal--sm");
  const list = $("#mtg-attendance-list", modal);

  function renderList() {
    const hereCount = participants.filter((participant) => session.attendance[participant.id] === "Here").length;
    list.innerHTML = `
      <div class="meeting-attendance-summary">${hereCount}/${participants.length} present</div>
      <div class="meeting-attendance-list">
        ${participants.length ? participants.map((participant) => {
          const here = session.attendance[participant.id] === "Here";
          const editable = canFacilitate || (me && participant.id === me.id);
          return `<div class="meeting-attendance-row">
            ${userAvatarHtml({ name: participant.name, email: participantEmail(participant), size: 30 })}
            <div class="meeting-attendance-person"><strong>${escapeHtml(participant.name)}</strong><span>${escapeHtml(participant.role || "Participant")}</span></div>
            <button type="button" class="meeting-presence-pill ${here ? "here" : "out"}" data-attendance-person="${escapeHtml(participant.id)}" aria-pressed="${here}" ${editable ? "" : "disabled"}>${here ? "Here" : "Out"}</button>
          </div>`;
        }).join("") : `<div class="empty-state" style="padding:20px 0;">No meeting participants yet.</div>`}
      </div>`;

    $all("[data-attendance-person]", list).forEach((button) => button.addEventListener("click", async () => {
      const participantId = button.dataset.attendancePerson;
      const participant = participants.find((person) => person.id === participantId);
      if (!participant || (!canFacilitate && (!me || participant.id !== me.id))) return;
      const previous = session.attendance[participantId] === "Here" ? "Here" : "Out";
      session.attendance[participantId] = previous === "Here" ? "Out" : "Here";
      button.disabled = true;
      try {
        logMeetingActivity(scope, `${currentUserName()} marked ${participant.name} ${session.attendance[participantId] === "Here" ? "present" : "out"}.`, { type: "attendance" });
        await saveMeetingSession(scope);
        renderList();
      } catch (error) {
        session.attendance[participantId] = previous;
        toast("Couldn't save attendance — try again.", "error");
        renderList();
      }
    }));
  }

  renderList();
  const close = () => { closeModal(); if (typeof onDone === "function") onDone(); };
  $(".aewttr-modal-close", modal).addEventListener("click", close);
  const addGuestBtn = $("#mtg-attendance-add-guest", modal);
  if (addGuestBtn) addGuestBtn.addEventListener("click", () => openAddMeetingGuestModal(scope, () => renderList()));
}
function meetingViewState(scope) {
  if (!window.AEWTTR.state.meetingView) window.AEWTTR.state.meetingView = {};
  const key = meetingScopeKey(scope);
  if (!window.AEWTTR.state.meetingView[key]) window.AEWTTR.state.meetingView[key] = "live";
  return window.AEWTTR.state.meetingView[key];
}
function meetingQueueMode(scope) {
  return meetingLiveTab(scope) === "project" ? "project" : "person";
}
function meetingLiveTab(scope) {
  if (!window.AEWTTR.state.meetingLiveTab) window.AEWTTR.state.meetingLiveTab = {};
  const key = meetingScopeKey(scope);
  const tab = window.AEWTTR.state.meetingLiveTab[key];
  if (tab === "room" || tab === "project") return tab;
  // "notes" and legacy "agenda" both map to the combined doc editor
  window.AEWTTR.state.meetingLiveTab[key] = "notes";
  return "notes";
}
function migrateMeetingSessionNotes(session) {
  if (!session) return session;
  session.notesFeed = (session.notesFeed || [])
    .filter((note) => note && String(note.text || "").trim())
    .map((note) => ({
      ...note,
      author: note.author === "Legacy note" ? "Meeting notes" : note.author
    }));
  if (!session.notesFeed.length && session.notes && String(session.notes).trim()) {
    session.notesFeed = [{
      id: uid("note"),
      author: "Meeting notes",
      date: session.date || new Date().toISOString().slice(0, 10),
      time: "",
      text: String(session.notes).trim()
    }];
  }
  delete session.notes;
  return session;
}
function meetingSessionGuests(session) {
  if (!session) return [];
  session.guests = session.guests || [];
  return session.guests;
}
function meetingProjects(scope) {
  const db = window.AEWTTR.db;
  if (scope.type === "project") return [scope.project];
  return (db.projects || []).slice();
}
function meetingProjectPortfolios(project) {
  const list = typeof projectPortfolios === "function" ? projectPortfolios(project) : (Array.isArray(project && project.portfolios) ? project.portfolios : []);
  return list.length ? list : ["Unassigned"];
}
function meetingProjectPortfolioLabel(project) {
  const list = meetingProjectPortfolios(project).filter((p) => p !== "Unassigned");
  if (!list.length) return "Unassigned";
  if (list.length === 1) return list[0];
  if (list.length === 2) return list.join(", ");
  return `${list[0]} +${list.length - 1}`;
}
function meetingProjectFilterStatus(project) {
  // Meeting status is a live tracker rollup, never a manually maintained
  // project field. This keeps Started / Not Started / Blocked / Done aligned
  // with the actual work items people are updating.
  return typeof computeProjectStatus === "function" ? computeProjectStatus(project) : "Not Started";
}
function meetingProjectViewFilters(scope) {
  const key = meetingScopeKey(scope);
  if (!window.AEWTTR.state.meetingProjectFilters) window.AEWTTR.state.meetingProjectFilters = {};
  if (!window.AEWTTR.state.meetingProjectFilters[key]) {
    window.AEWTTR.state.meetingProjectFilters[key] = { portfolios: [], statuses: [] };
  }
  return window.AEWTTR.state.meetingProjectFilters[key];
}
function safeQueryByAttr(root, attr, value) {
  if (!root || !attr) return null;
  const want = String(value == null ? "" : value);
  try {
    return Array.from(root.querySelectorAll(`[${attr}]`)).find((node) => node.getAttribute(attr) === want) || null;
  } catch (e) {
    return null;
  }
}

const MEETING_ALL_TASKS_ID = "__all_tasks__";

/* Single selected project in the "By Project" sidebar. Falls back to the
   first visible project when nothing is selected yet or the current
   selection was filtered out by the portfolio/status filters.
   Special value MEETING_ALL_TASKS_ID = full nested task list across projects. */
function meetingProjectSelectedId(scope, visible) {
  const key = meetingScopeKey(scope);
  if (!window.AEWTTR.state.meetingProjectSelected) window.AEWTTR.state.meetingProjectSelected = {};
  const current = window.AEWTTR.state.meetingProjectSelected[key];
  if (current === MEETING_ALL_TASKS_ID) return MEETING_ALL_TASKS_ID;
  if (!current || !visible.some((p) => p.id === current)) {
    window.AEWTTR.state.meetingProjectSelected[key] = visible[0] ? visible[0].id : null;
  }
  return window.AEWTTR.state.meetingProjectSelected[key];
}
function projectMeetingSummary(project) {
  const tasks = (window.AEWTTR.db.ganttTasks[project.id] || []);
  const openTasks = tasks.filter((t) => taskProgressPct(t) < 100);
  const offTrack = openTasks.filter((t) => t.health === "Off Track").length;
  const atRisk = openTasks.filter((t) => t.health === "At Risk").length;
  return {
    tasks,
    openTasks,
    offTrack,
    atRisk,
    status: meetingProjectFilterStatus(project),
    portfolios: meetingProjectPortfolios(project),
    portfolioLabel: meetingProjectPortfolioLabel(project),
    rag: typeof computeProjectRag === "function" ? computeProjectRag(project) : (project.rag || "Green")
  };
}

/* ---------- top-level meeting shell ---------- */
function renderMeetingApp(mount, scope) {
  setTopbar(
    scope.type === "global" ? "Weekly Meeting" : `${scope.project.name} — Meeting`,
    scope.type === "global" ? "Cross-project pulse with a live participant work queue." : "Project-specific operating meeting with a live task queue."
  );
  const view = meetingViewState(scope);
  const key = meetingScopeKey(scope);
  const active = meetingIsActive(scope);
  const canFacilitate = canManageMeetings();
  const routeIntent = typeof consumeRouteIntent === "function" ? consumeRouteIntent(`meeting:${key}`) : null;
  mount.innerHTML = `
    <div class="meeting-toolbar">
      <div class="segmented">
        <button class="segmented-opt ${view === "live" ? "active" : ""}" data-view="live"${tip("Live meeting view")}>Meeting</button>
        <button class="segmented-opt ${view === "history" ? "active" : ""}" data-view="history"${tip("Past meeting history")}>History</button>
      </div>
      ${view === "live" && active ? `
        <div class="segmented meeting-live-tabs">
          <button class="segmented-opt ${meetingLiveTab(scope) === "notes" ? "active" : ""}" data-live-tab="notes"${tip("Meeting Minutes — Word document editor")}><i class="bx bx-file"></i> Meeting Minutes</button>
          <button class="segmented-opt ${meetingLiveTab(scope) === "project" ? "active" : ""}" data-live-tab="project"${tip("Major updates, project health, and status changes")}><i class="bx bx-list-check"></i> Major Updates & Status</button>
          <button class="segmented-opt ${meetingLiveTab(scope) === "room" ? "active" : ""}" data-live-tab="room"${tip("Walk the room person by person")}><i class="bx bx-group"></i> Around the Room</button>
        </div>
      ` : ""}
      ${active ? `<button class="btn-aewttr-outline btn-aewttr-sm" id="mtg-attendance"${tip("Review and update meeting attendance")}><i class="bx bx-check-square"></i> Attendance</button>` : ""}
      ${view === "live" && canFacilitate && active ? `<button class="btn-aewttr-outline btn-aewttr-sm" id="mtg-end"${tip("End the current meeting session")}><i class="bx bx-archive"></i> End Meeting</button>` : ""}
      ${view === "live" && canFacilitate && !active ? `<button class="btn-aewttr btn-aewttr-sm" id="mtg-start-toolbar"${tip("Start a new meeting session")}><i class="bx bx-play"></i> Start Meeting</button>` : ""}
      ${view === "live" && active ? `<span class="kc-badge" style="margin-left:auto;">Meeting active</span>` : ""}
      ${view === "live" && !active && meetingStatus(scope) === "ended" ? `<span class="kc-badge" style="margin-left:auto;opacity:.85;">Meeting ended</span>` : ""}
    </div>
    <div id="mtg-body"></div>
  `;
  $all("[data-view]", mount).forEach(b => b.addEventListener("click", () => {
    window.AEWTTR.state.meetingView[key] = b.dataset.view;
    renderMeetingApp(mount, scope);
  }));
  $all("[data-live-tab]", mount).forEach(b => b.addEventListener("click", () => {
    window.AEWTTR.state.meetingLiveTab[key] = b.dataset.liveTab;
    renderMeetingApp(mount, scope);
  }));
  const body = $("#mtg-body", mount);
  mount.classList.toggle("meeting-app--live", view === "live");
  const attendanceBtn = $("#mtg-attendance", mount);
  if (attendanceBtn) attendanceBtn.addEventListener("click", () => openMeetingAttendanceModal(scope, () => renderMeetingApp(mount, scope)));
  if (view === "history") return renderMeetingHistory(body, scope);
  renderMeetingLive(body, scope);
  const endBtn = $("#mtg-end", mount);
  if (endBtn) endBtn.addEventListener("click", () => openEndMeetingModal(scope, () => renderMeetingApp(mount, scope)));
  const startToolbarBtn = $("#mtg-start-toolbar", mount);
  if (startToolbarBtn) startToolbarBtn.addEventListener("click", () => startMeeting(scope, () => renderMeetingApp(mount, scope)));
}

function wireMeetingNotesComposer(mount, session, scope) {
  renderMeetingNotesFeed($("#mtg-notes-feed", mount), session, scope);
  const postBtn = $("#mtg-note-post", mount);
  const input = $("#mtg-note-input", mount);
  if (!postBtn || !input) return;
  // Same composer behavior as the task/subtask notes popup: Enter posts,
  // Shift+Enter for a plain newline, Ctrl+Enter for an indented sub-line —
  // one consistent chat interaction across every message-style tool in PULSE.
  wireChatComposer(input, postBtn, async (text) => {
    const now = new Date();
    session.notesFeed = session.notesFeed || [];
    session.notesFeed.unshift({
      id: uid("note"),
      author: currentUserName(),
      date: now.toISOString().slice(0, 10),
      time: now.toTimeString().slice(0, 5),
      text
    });
    delete session.notes;
    recordActiveMeetingNotePosted(
      scope.type === "project" ? scope.project.id : "",
      text,
      { kind: "meeting", projectId: scope.type === "project" ? scope.project.id : "" }
    );
    await saveMeetingSession(scope);
    renderMeetingNotesFeed($("#mtg-notes-feed", mount), session, scope);
  });
}

function renderMeetingAgendaTab(body, scope, session) {
  session.agenda = Array.isArray(session.agenda) ? session.agenda : [];
  body.className = "";

  const doneCount = session.agenda.filter((i) => i.done).length;
  const total = session.agenda.length;

  function agendaItemHtml(item, idx) {
    return `
      <div class="agenda-item${item.done ? " agenda-item--done" : ""}" data-agenda-id="${escapeHtml(item.id)}">
        <div class="agenda-item-header">
          <button type="button" class="agenda-check-btn${item.done ? " is-done" : ""}" data-agenda-done="${escapeHtml(item.id)}" aria-label="${item.done ? "Mark incomplete" : "Mark complete"}" aria-pressed="${item.done}">
            <i class="bx ${item.done ? "bxs-check-circle" : "bx-circle"}"></i>
          </button>
          <span class="agenda-item-num">${idx + 1}.</span>
          <input type="text" class="agenda-title-input" data-agenda-title="${escapeHtml(item.id)}" value="${escapeHtml(item.title || "")}" placeholder="Topic title…" aria-label="Agenda topic">
          <button type="button" class="agenda-delete-btn" data-agenda-delete="${escapeHtml(item.id)}" title="Remove item" aria-label="Remove agenda item">
            <i class="bx bx-x"></i>
          </button>
        </div>
        <div class="agenda-item-body">
          <textarea class="agenda-notes-input" data-agenda-notes="${escapeHtml(item.id)}" placeholder="Notes, decisions, action items… (use - for bullets)" rows="2">${escapeHtml(item.notes || "")}</textarea>
        </div>
      </div>`;
  }

  body.innerHTML = `
    <div class="meeting-agenda-tab">
      <div class="meeting-notes-tab-head">
        <div>
          <div class="side-panel-title">Meeting Agenda</div>
          <p style="font-size:12px;color:var(--aewttr-muted);margin:4px 0 0;">Add topics before or during the meeting. Check them off as you go.</p>
        </div>
        ${total ? `<span class="kc-badge">${doneCount}/${total} done</span>` : ""}
      </div>
      <div class="agenda-add-row">
        <input class="input-aewttr" id="mtg-agenda-input" placeholder="Add an agenda item…" aria-label="Agenda item" autocomplete="off">
        <button type="button" class="btn-aewttr btn-aewttr-sm" id="mtg-agenda-add-btn"><i class="bx bx-plus"></i> Add</button>
      </div>
      <div class="agenda-list" id="mtg-agenda-list">
        ${total
          ? session.agenda.map((item, idx) => agendaItemHtml(item, idx)).join("")
          : `<div class="agenda-empty"><i class="bx bx-list-ul"></i><p>No agenda items yet.</p><p class="agenda-empty-hint">Add the first topic above to get started.</p></div>`}
      </div>
    </div>`;

  function autoResize(ta) {
    ta.style.height = "auto";
    ta.style.height = Math.max(ta.scrollHeight, 52) + "px";
  }

  $all(".agenda-notes-input", body).forEach((ta) => autoResize(ta));

  function addItem() {
    const input = $("#mtg-agenda-input", body);
    const title = input && input.value.trim();
    if (!title) { if (input) input.focus(); return; }
    const item = { id: uid("agenda"), title, notes: "", done: false, author: currentUserName(), createdAt: new Date().toISOString() };
    session.agenda.push(item);
    input.value = "";
    logMeetingActivity(scope, `${currentUserName()} added agenda item: "${title}".`, { type: "agenda-add" });
    renderMeetingAgendaTab(body, scope, session);
    const newTextarea = body.querySelector(`[data-agenda-notes="${item.id}"]`);
    if (newTextarea) newTextarea.focus();
    saveMeetingSession(scope).catch(() => {
      const idx = session.agenda.findIndex((e) => e.id === item.id);
      if (idx >= 0) session.agenda.splice(idx, 1);
      renderMeetingAgendaTab(body, scope, session);
      toast("Couldn't save agenda item — try again.", "error");
    });
  }

  const addBtn = $("#mtg-agenda-add-btn", body);
  if (addBtn) addBtn.addEventListener("click", addItem);
  const agendaInput = $("#mtg-agenda-input", body);
  if (agendaInput) {
    agendaInput.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); addItem(); } });
  }

  $all("[data-agenda-done]", body).forEach((btn) => {
    btn.addEventListener("click", async () => {
      const item = session.agenda.find((e) => e.id === btn.dataset.agendaDone);
      if (!item) return;
      item.done = !item.done;
      logMeetingActivity(scope, `${currentUserName()} ${item.done ? "checked off" : "unchecked"} agenda item: "${item.title}".`, { type: "agenda-check" });
      renderMeetingAgendaTab(body, scope, session);
      saveMeetingSession(scope).catch(() => { item.done = !item.done; renderMeetingAgendaTab(body, scope, session); });
    });
  });

  $all("[data-agenda-title]", body).forEach((input) => {
    let titleTimer = null;
    input.addEventListener("input", () => {
      clearTimeout(titleTimer);
      titleTimer = setTimeout(async () => {
        const item = session.agenda.find((e) => e.id === input.dataset.agendaTitle);
        if (!item) return;
        item.title = input.value;
        await saveMeetingSession(scope);
      }, 600);
    });
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        const notesArea = input.closest(".agenda-item").querySelector(".agenda-notes-input");
        if (notesArea) notesArea.focus();
      }
    });
  });

  $all("[data-agenda-notes]", body).forEach((ta) => {
    let saveTimer = null;

    ta.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      const val = ta.value;
      const pos = ta.selectionStart;
      const lineStart = val.lastIndexOf("\n", pos - 1) + 1;
      const currentLine = val.slice(lineStart, pos);
      const bulletMatch = currentLine.match(/^([-•]\s)/);
      if (!bulletMatch) return;
      e.preventDefault();
      const prefix = bulletMatch[1];
      const content = currentLine.slice(prefix.length).trim();
      if (!content) {
        ta.value = val.slice(0, lineStart) + val.slice(pos);
        ta.selectionStart = ta.selectionEnd = lineStart;
      } else {
        const insert = "\n" + prefix;
        ta.value = val.slice(0, pos) + insert + val.slice(pos);
        ta.selectionStart = ta.selectionEnd = pos + insert.length;
      }
      ta.dispatchEvent(new Event("input"));
    });

    ta.addEventListener("input", () => {
      autoResize(ta);
      clearTimeout(saveTimer);
      saveTimer = setTimeout(async () => {
        const item = session.agenda.find((e) => e.id === ta.dataset.agendaNotes);
        if (!item) return;
        item.notes = ta.value;
        await saveMeetingSession(scope);
      }, 800);
    });
  });

  $all("[data-agenda-delete]", body).forEach((btn) => {
    btn.addEventListener("click", async () => {
      const id = btn.dataset.agendaDelete;
      const idx = session.agenda.findIndex((e) => e.id === id);
      if (idx < 0) return;
      const [removed] = session.agenda.splice(idx, 1);
      logMeetingActivity(scope, `${currentUserName()} removed agenda item: "${removed.title}".`, { type: "agenda-delete" });
      renderMeetingAgendaTab(body, scope, session);
      saveMeetingSession(scope).catch(() => {
        session.agenda.splice(idx, 0, removed);
        renderMeetingAgendaTab(body, scope, session);
        toast("Couldn't remove agenda item — try again.", "error");
      });
    });
  });
}

/* ---------- combined meeting document editor (single unified Word document) ---------- */

function getInitialDocHtml(session) {
  if (session.docHtml && typeof session.docHtml === "string" && session.docHtml.trim()) {
    return session.docHtml;
  }
  let html = "";
  if (Array.isArray(session.docBlocks) && session.docBlocks.length) {
    let inList = false;
    session.docBlocks.forEach(function(b) {
      if (b.type === "bullet" || b.type === "bullet-sub") {
        if (!inList) { html += "<ul>"; inList = true; }
        html += `<li>${b.content || ""}</li>`;
      } else {
        if (inList) { html += "</ul>"; inList = false; }
        if (b.type === "h2") html += `<h2>${b.content || ""}</h2>`;
        else if (b.type === "action" || b.type === "task") html += `<p class="doc-action-item"><label><input type="checkbox"${b.done ? " checked" : ""}> <span>${b.content || ""}</span></label></p>`;
        else html += `<p>${b.content || ""}</p>`;
      }
    });
    if (inList) html += "</ul>";
  } else if (session.agenda && session.agenda.length) {
    session.agenda.forEach(function(item) {
      if (item.title) html += `<h2>${escapeHtml(item.title)}</h2>`;
      if (item.notes) {
        String(item.notes).split("\n").forEach(function(line) {
          const tr = line.trim();
          if (tr) html += `<p>${escapeHtml(tr)}</p>`;
        });
      }
    });
  }
  if (!html.trim()) {
    html = "<p><br></p>";
  }
  session.docHtml = html;
  return html;
}

function renderMeetingDocEditor(body, scope, session, opts) {
  opts = opts || {};
  const initialHtml = getInitialDocHtml(session);

  const COMMANDS = [
    { id: "h1", label: "Heading 1", icon: "bx-heading", desc: "Big section heading" },
    { id: "h2", label: "Heading 2", icon: "bx-heading", desc: "Medium section heading" },
    { id: "bullet", label: "Bullet list", icon: "bx-list-ul", desc: "Simple bulleted list" },
    { id: "num", label: "Numbered list", icon: "bx-list-ol", desc: "Sequential numbered list" },
    { id: "action", label: "Checkbox item", icon: "bx-check-square", desc: "Single-line action item" },
    { id: "task", label: "Assigned task", icon: "bx-user-check", desc: "@Name — task sentence" },
    { id: "quote", label: "Quote", icon: "bx-block", desc: "Indented blockquote" },
    { id: "para", label: "Normal text", icon: "bx-text", desc: "Plain text paragraph" }
  ];

  // Left side context info
  const docTitle = opts.title || "Meeting Minutes";
  const ribbonLeftHtml = `
    <div class="doc-ribbon-context">
      <span class="doc-ribbon-context-title">${escapeHtml(docTitle)}</span>
      ${opts.subtitle ? `<span class="doc-ribbon-context-sub">${escapeHtml(opts.subtitle)}</span>` : ""}
    </div>
    ${opts.action ? `<button type="button" class="doc-ribbon-action doc-ribbon-action--primary" id="${escapeHtml(opts.action.id)}">${opts.action.icon ? `<i class="bx ${escapeHtml(opts.action.icon)}" aria-hidden="true"></i> ` : ""}${escapeHtml(opts.action.label)}</button>` : ""}
    <div class="doc-ribbon-sep" aria-hidden="true"></div>
  `;

  body.innerHTML = `
    <div class="meeting-doc-shell">
      <div class="meeting-doc-ribbon" role="toolbar" aria-label="Document formatting">
        ${ribbonLeftHtml}
        <!-- Undo / Redo group -->
        <div class="doc-ribbon-group" aria-label="Undo and Redo">
          <button type="button" class="doc-fmt-btn" id="doc-fmt-undo" title="Undo (Ctrl+Z)" aria-label="Undo"><i class="bx bx-undo" aria-hidden="true"></i></button>
          <button type="button" class="doc-fmt-btn" id="doc-fmt-redo" title="Redo (Ctrl+Y)" aria-label="Redo"><i class="bx bx-redo" aria-hidden="true"></i></button>
        </div>
        <div class="doc-ribbon-sep" aria-hidden="true"></div>
        <!-- Format group -->
        <div class="doc-ribbon-group" aria-label="Text formatting">
          <button type="button" class="doc-fmt-btn" id="doc-fmt-bold" title="Bold (Ctrl+B)" aria-label="Bold"><b>B</b></button>
          <button type="button" class="doc-fmt-btn" id="doc-fmt-italic" title="Italic (Ctrl+I)" aria-label="Italic"><i>I</i></button>
          <button type="button" class="doc-fmt-btn" id="doc-fmt-underline" title="Underline (Ctrl+U)" aria-label="Underline"><u>U</u></button>
          <button type="button" class="doc-fmt-btn" id="doc-fmt-strike" title="Strikethrough" aria-label="Strikethrough"><s>S</s></button>
        </div>
        <div class="doc-ribbon-sep" aria-hidden="true"></div>
        <!-- Style dropdown group -->
        <div class="doc-ribbon-group" aria-label="Style">
          <select class="doc-ribbon-select" id="doc-style-select" title="Paragraph Style" aria-label="Paragraph Style">
            <option value="p">Normal Text</option>
            <option value="h1">Heading 1</option>
            <option value="h2">Heading 2</option>
            <option value="blockquote">Quote</option>
          </select>
        </div>
        <div class="doc-ribbon-sep" aria-hidden="true"></div>
        <!-- Lists group -->
        <div class="doc-ribbon-group" aria-label="Lists">
          <button type="button" class="doc-fmt-btn" id="doc-list-bullet" title="Bullet List" aria-label="Bullet List"><i class="bx bx-list-ul" aria-hidden="true"></i></button>
          <button type="button" class="doc-fmt-btn" id="doc-list-num" title="Numbered List" aria-label="Numbered List"><i class="bx bx-list-ol" aria-hidden="true"></i></button>
        </div>
        <div class="doc-ribbon-sep" aria-hidden="true"></div>
        <!-- Insert group -->
        <div class="doc-ribbon-group" aria-label="Insert">
          <button type="button" class="doc-ribbon-action" id="doc-add-action"><i class="bx bx-check-square" aria-hidden="true"></i> Checkbox</button>
          <button type="button" class="doc-ribbon-action" id="doc-add-task-template"><i class="bx bx-user-check" aria-hidden="true"></i> Add Task</button>
        </div>
        <!-- Right side -->
        <div class="doc-ribbon-sep" aria-hidden="true"></div>
        <span class="doc-ribbon-meta" id="doc-word-count"></span>
        <span class="doc-save-status" id="doc-save-status" aria-live="polite"></span>
      </div>
      <div class="meeting-doc-canvas">
        <div class="meeting-doc-page">
          <div class="word-document-editor" id="meeting-doc-editor" contenteditable="true" spellcheck="true" data-placeholder="Type or paste your meeting document here… (use / for commands)"></div>
        </div>
      </div>
    </div>
  `;

  const editor = $("#meeting-doc-editor", body);
  if (editor) editor.innerHTML = initialHtml;

  // Slash command palette element
  let paletteEl = document.getElementById("doc-cmd-palette");
  if (paletteEl) paletteEl.remove();
  paletteEl = document.createElement("div");
  paletteEl.id = "doc-cmd-palette";
  paletteEl.className = "doc-cmd-palette";
  paletteEl.style.display = "none";
  document.body.appendChild(paletteEl);

  let cmdOpen = false;
  let cmdQuery = "";
  let cmdActiveIdx = 0;

  function filteredCmds() {
    if (!cmdQuery) return COMMANDS;
    const q = cmdQuery.toLowerCase();
    return COMMANDS.filter(c => c.label.toLowerCase().includes(q) || c.id.toLowerCase().includes(q));
  }

  function renderPalette() {
    const list = filteredCmds();
    if (!list.length) { closePalette(); return; }
    cmdActiveIdx = Math.max(0, Math.min(cmdActiveIdx, list.length - 1));
    paletteEl.innerHTML = list.map((c, i) => `
      <div class="doc-cmd-item${i === cmdActiveIdx ? " active" : ""}" data-cmd-id="${c.id}">
        <i class="bx ${c.icon} doc-cmd-icon"></i>
        <div class="doc-cmd-text">
          <span class="doc-cmd-label">${escapeHtml(c.label)}</span>
          <span class="doc-cmd-desc">${escapeHtml(c.desc)}</span>
        </div>
      </div>
    `).join("");

    $all(".doc-cmd-item", paletteEl).forEach((item, i) => {
      item.addEventListener("mousedown", (e) => {
        e.preventDefault();
        applySlashCommand(list[i]);
      });
    });
  }

  function openPalette() {
    cmdOpen = true;
    cmdActiveIdx = 0;
    const sel = window.getSelection();
    if (sel && sel.rangeCount) {
      const rect = sel.getRangeAt(0).getBoundingClientRect();
      paletteEl.style.top = Math.max(10, rect.bottom + window.scrollY + 4) + "px";
      paletteEl.style.left = Math.max(12, rect.left + window.scrollX) + "px";
    }
    paletteEl.style.display = "block";
    renderPalette();
  }

  function closePalette() {
    cmdOpen = false;
    cmdQuery = "";
    if (paletteEl) paletteEl.style.display = "none";
  }

  function applySlashCommand(cmd) {
    closePalette();
    const sel = window.getSelection();
    if (sel && sel.rangeCount) {
      const range = sel.getRangeAt(0);
      const node = range.startContainer;
      if (node && node.nodeType === Node.TEXT_NODE) {
        const txt = node.textContent;
        const slashIdx = txt.lastIndexOf("/");
        if (slashIdx >= 0) {
          node.textContent = txt.slice(0, slashIdx);
        }
      }
    }

    if (cmd.id === "h1") exec("formatBlock", "<h1>");
    else if (cmd.id === "h2") exec("formatBlock", "<h2>");
    else if (cmd.id === "bullet") exec("insertUnorderedList");
    else if (cmd.id === "num") exec("insertOrderedList");
    else if (cmd.id === "quote") exec("formatBlock", "<blockquote>");
    else if (cmd.id === "action") exec("insertHTML", '<p class="doc-action-item"><label><input type="checkbox"> <span contenteditable="true">Action item</span></label></p>');
    else if (cmd.id === "task") triggerAddTask();
    else if (cmd.id === "para") exec("formatBlock", "<p>");
  }

  function triggerAddTask() {
    const sel = window.getSelection();
    const anchor = sel && sel.anchorNode;
    const inLi = anchor ? (anchor.nodeType === 1 ? anchor.closest("li") : anchor.parentElement?.closest("li")) : null;

    openMeetingAddTaskModal(
      scope,
      currentMeetingParticipant(scope),
      (newTask, project) => {
        if (newTask) {
          const pName = project ? project.name : (scope.type === "project" ? scope.project.name : "Project");
          const taskContentHtml = `<span class="doc-task-inline"><strong>@${escapeHtml(newTask.assignee || "Unassigned")}</strong> — ${escapeHtml(newTask.title)} <em>(Project: ${escapeHtml(pName)})</em></span>`;

          if (inLi) {
            exec("insertHTML", taskContentHtml);
          } else {
            const sentenceHtml = `<p class="doc-task-sentence">${taskContentHtml}</p><p><br></p>`;
            exec("insertHTML", sentenceHtml);
          }
        }
      },
      scope.type === "project" ? scope.project.id : ""
    );
  }

  let saveTimer = null;
  function isEditorMounted() {
    return document.getElementById("meeting-doc-editor") === editor;
  }

  function updateWordCount() {
    const text = editor.textContent || "";
    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    const metaEl = document.getElementById("doc-word-count");
    if (metaEl) metaEl.textContent = `${words} ${words === 1 ? "word" : "words"}`;
  }

  function schedSave() {
    clearTimeout(saveTimer);
    updateWordCount();
    const statusEl = document.getElementById("doc-save-status");
    if (statusEl) { statusEl.textContent = "Saving…"; statusEl.className = "doc-save-status doc-saving"; }
    saveTimer = setTimeout(async function() {
      if (!isEditorMounted()) return;
      session.docHtml = editor.innerHTML;
      try {
        await saveMeetingSession(scope);
        const el = document.getElementById("doc-save-status");
        if (el) { el.textContent = "Saved"; el.className = "doc-save-status doc-saved"; }
        setTimeout(function() { const e2 = document.getElementById("doc-save-status"); if (e2 && e2.textContent === "Saved") e2.textContent = ""; }, 2000);
      } catch (error) {
        const el = document.getElementById("doc-save-status");
        if (el) { el.textContent = "Couldn’t save — try again"; el.className = "doc-save-status doc-save-failed"; }
      }
    }, 600);
  }

  function exec(cmd, arg = null) {
    if (!editor) return;
    editor.focus();
    document.execCommand(cmd, false, arg);
    schedSave();
  }

  // Ribbon button handlers
  $("#doc-fmt-undo", body)?.addEventListener("mousedown", (e) => { e.preventDefault(); exec("undo"); });
  $("#doc-fmt-redo", body)?.addEventListener("mousedown", (e) => { e.preventDefault(); exec("redo"); });
  $("#doc-fmt-bold", body)?.addEventListener("mousedown", (e) => { e.preventDefault(); exec("bold"); });
  $("#doc-fmt-italic", body)?.addEventListener("mousedown", (e) => { e.preventDefault(); exec("italic"); });
  $("#doc-fmt-underline", body)?.addEventListener("mousedown", (e) => { e.preventDefault(); exec("underline"); });
  $("#doc-fmt-strike", body)?.addEventListener("mousedown", (e) => { e.preventDefault(); exec("strikeThrough"); });
  $("#doc-list-bullet", body)?.addEventListener("mousedown", (e) => { e.preventDefault(); exec("insertUnorderedList"); });
  $("#doc-list-num", body)?.addEventListener("mousedown", (e) => { e.preventDefault(); exec("insertOrderedList"); });

  const styleSelect = $("#doc-style-select", body);
  if (styleSelect) {
    styleSelect.addEventListener("change", () => {
      const val = styleSelect.value;
      exec("formatBlock", `<${val}>`);
    });
  }

  $("#doc-add-action", body)?.addEventListener("click", () => {
    exec("insertHTML", '<p class="doc-action-item"><label><input type="checkbox"> <span>Action item</span></label></p>');
  });

  $("#doc-add-task-template", body)?.addEventListener("click", () => {
    triggerAddTask();
  });

  // Editor events
  editor.addEventListener("keydown", (e) => {
    if (cmdOpen) {
      if (e.key === "ArrowDown") { e.preventDefault(); cmdActiveIdx = Math.min(cmdActiveIdx + 1, filteredCmds().length - 1); renderPalette(); return; }
      if (e.key === "ArrowUp") { e.preventDefault(); cmdActiveIdx = Math.max(0, cmdActiveIdx - 1); renderPalette(); return; }
      if (e.key === "Enter" || e.key === "Tab") { e.preventDefault(); const list = filteredCmds(); if (list[cmdActiveIdx]) applySlashCommand(list[cmdActiveIdx]); return; }
      if (e.key === "Escape") { e.preventDefault(); closePalette(); return; }
      if (e.key === "Backspace" && !cmdQuery) { closePalette(); return; }
    }

    const sel = window.getSelection();
    const anchor = sel && sel.anchorNode;
    const li = anchor ? (anchor.nodeType === 1 ? anchor.closest("li") : anchor.parentElement?.closest("li")) : null;

    if (e.key === "Backspace") {
      if (li) {
        const isEmpty = !li.textContent.trim();
        const atStart = sel.anchorOffset === 0;
        if (isEmpty || atStart) {
          e.preventDefault();
          const isNested = li.parentElement && li.parentElement.closest("ul ul, ol ol, ul ol, ol ul");
          if (isNested) {
            exec("outdent");
          } else {
            // Root list item -> convert line back to normal paragraph text
            if (li.parentElement?.tagName === "OL") {
              exec("insertOrderedList");
            } else {
              exec("insertUnorderedList");
            }
            exec("formatBlock", "<p>");
          }
          return;
        }
      }
    }

    if (e.key === "Tab") {
      e.preventDefault();
      if (li) {
        if (e.shiftKey) exec("outdent");
        else exec("indent");
      } else {
        if (anchor) {
          const txt = anchor.textContent || "";
          const tr = txt.trim();
          if (tr === "-" || tr === "*" || tr === "•") {
            if (anchor.nodeType === 3) anchor.textContent = "";
            else anchor.innerHTML = "<br>";
            exec("insertUnorderedList");
            return;
          } else if (/^\d+\.?$/.test(tr)) {
            if (anchor.nodeType === 3) anchor.textContent = "";
            else anchor.innerHTML = "<br>";
            exec("insertOrderedList");
            return;
          }
        }
        if (!e.shiftKey) exec("insertUnorderedList");
      }
      return;
    }
  });

  editor.addEventListener("input", () => {
    const sel = window.getSelection();
    if (sel && sel.rangeCount) {
      const range = sel.getRangeAt(0);
      const node = range.startContainer;
      if (node && node.nodeType === Node.TEXT_NODE) {
        const textBefore = node.textContent.slice(0, range.startOffset);

        // Auto-bullet formatting on "- " or "* " or "1. "
        if (/^[-*•]\s$/.test(textBefore)) {
          node.textContent = node.textContent.slice(2);
          exec("insertUnorderedList");
          return;
        } else if (/^\d+\.\s$/.test(textBefore)) {
          node.textContent = node.textContent.replace(/^\d+\.\s/, "");
          exec("insertOrderedList");
          return;
        }

        const slashIdx = textBefore.lastIndexOf("/");
        if (slashIdx >= 0 && (slashIdx === 0 || /\s/.test(textBefore[slashIdx - 1]))) {
          cmdQuery = textBefore.slice(slashIdx + 1);
          openPalette();
        } else if (cmdOpen) {
          closePalette();
        }
      }
    }
    schedSave();
  });

  document.addEventListener("click", (e) => {
    if (cmdOpen && paletteEl && !paletteEl.contains(e.target) && !editor.contains(e.target)) {
      closePalette();
    }
  });

  updateWordCount();
}

function renderMeetingNotesTab(body, scope, session, opts) {
  const { canFacilitate, allParticipants, hereCount, myHere, me } = opts;
  body.innerHTML = `
    <div class="meeting-notes-tab">
      <div class="meeting-notes-tab-head">
        <div>
          <div class="side-panel-title">General Meeting Notes</div>
          <p style="font-size:12px;color:var(--aewttr-muted);margin:4px 0 0;">Post updates as they come up — everyone sees the feed, so nobody's note overwrites anyone else's.</p>
        </div>
        <span class="kc-badge">${hereCount}/${allParticipants.length} present</span>
      </div>
      <div class="aewttr-card aewttr-card-pad meeting-notes-tab-compose">
        <textarea class="textarea-aewttr" id="mtg-note-input" placeholder="Share an update, decision, or follow-up..." style="min-height:72px;"></textarea>
        <button class="btn-aewttr btn-aewttr-sm" id="mtg-note-post" style="margin-top:10px;"${tip("Post a note to the meeting log")}><i class="bx bx-send"></i> Post Note</button>
      </div>
      <div id="mtg-notes-feed" class="meeting-notes-tab-feed"></div>
    </div>
  `;
  wireMeetingNotesComposer(body, session, scope);
}

/* ---------- live meeting: nav + work queue + side panel ---------- */
function renderMeetingIdleState(body, scope) {
  const status = meetingStatus(scope);
  const canFacilitate = canManageMeetings();
  const me = currentMeetingParticipant(scope);
  const planning = meetingData(scope).currentSession && meetingData(scope).currentSession.sessionStatus === "planned" ? meetingData(scope).currentSession : null;
  if (planning) {
    body.className = "meeting-live-body meeting-live-body--doc";
    body.innerHTML = "";
    const prepHead = document.createElement("div");
    prepHead.className = "meeting-prep-head";
    prepHead.innerHTML = `<div><div class="side-panel-title">Upcoming meeting</div><p>Add notes and agenda items before the meeting starts.</p></div>${canFacilitate ? `<button class="btn-aewttr" id="mtg-start"${tip("Start the meeting with this agenda")}><i class="bx bx-play"></i> Start Meeting</button>` : ""}`;
    body.appendChild(prepHead);
    const editorMount = document.createElement("div");
    body.appendChild(editorMount);
    renderMeetingDocEditor(editorMount, scope, planning);
    const startBtn = $("#mtg-start", body);
    if (startBtn) startBtn.addEventListener("click", () => startMeeting(scope, () => renderMeetingApp($("#page-content"), scope)));
    return;
  }
  body.className = "meeting-live-body meeting-live-body--idle";
  body.innerHTML = `
    <div class="aewttr-card aewttr-card-pad" style="max-width:640px;margin:48px auto;text-align:center;">
      <div class="weekly-hero-kicker">${status === "ended" ? "Meeting ended" : "No meeting in progress"}</div>
      <h2 style="margin:8px 0 12px;">${status === "ended" ? "This session has been archived" : "Waiting to start"}</h2>
      <p style="color:var(--aewttr-muted);margin:0 0 20px;line-height:1.5;">
        ${status === "ended"
          ? "The meeting stays closed until a Meeting Admin starts a new session."
          : "A Meeting Admin can start the meeting when everyone is ready."}
      </p>
      <button class="btn-aewttr-outline" id="mtg-plan-agenda"><i class="bx bx-list-plus"></i> Plan Agenda</button>
      ${canFacilitate ? `<button class="btn-aewttr" id="mtg-start"${tip("Start the weekly meeting")}><i class="bx bx-play"></i> Start Meeting</button>` : ""}
      ${!canFacilitate && !me ? `<p style="font-size:12px;color:var(--aewttr-muted);margin:16px 0 0;">You're not on the ${scope.type === "global" ? "weekly roster" : "project people list"} yet — ask a Meeting Admin to add you.</p>` : ""}
    </div>
  `;
  const startBtn = $("#mtg-start", body);
  if (startBtn) {
    startBtn.addEventListener("click", () => startMeeting(scope, () => {
      const page = $("#page-content");
      if (page) renderMeetingApp(page, scope);
    }));
  }
  const planBtn = $("#mtg-plan-agenda", body);
  if (planBtn) planBtn.addEventListener("click", async () => { await ensurePlanningSession(scope); renderMeetingIdleState(body, scope); });
}

function renderMeetingLive(body, scope) {
  if (!meetingIsActive(scope)) return renderMeetingIdleState(body, scope);

  const session = ensureMeetingSession(scope);
  const allParticipants = meetingParticipants(scope);
  const canFacilitate = canManageMeetings();
  const me = currentMeetingParticipant(scope);
  const participants = canFacilitate ? allParticipants : (me ? [me] : []);
  if (!window.AEWTTR.state.meetingActive) window.AEWTTR.state.meetingActive = {};
  const key = meetingScopeKey(scope);
  if (canFacilitate) {
    if (!window.AEWTTR.state.meetingActive[key] || !participants.some(p => p.id === window.AEWTTR.state.meetingActive[key])) {
      window.AEWTTR.state.meetingActive[key] = participants[0] ? participants[0].id : null;
    }
  } else {
    window.AEWTTR.state.meetingActive[key] = me ? me.id : null;
  }
  const activeId = window.AEWTTR.state.meetingActive[key];
  const active = participants.find(p => p.id === activeId) || me;
  const hereCount = allParticipants.filter(p => session.attendance[p.id] === "Here").length;
  const myHere = me && session.attendance[me.id] === "Here";
  const liveTab = meetingLiveTab(scope);
  const queueMode = liveTab === "project" ? "project" : "person";

  if (liveTab === "notes") {
    body.className = "meeting-live-body--doc";
    return renderMeetingDocEditor(body, scope, session);
  }

  if (queueMode === "project") {
    return renderMeetingAllProjectsView(body, scope, {
      canFacilitate, allParticipants, hereCount, myHere, me, session
    });
  }

  body.className = "meeting-live-body";
  body.innerHTML = `
    <div class="meeting-live-grid">
      <div class="meeting-nav-col">
        <div class="meeting-nav-head">
          <span>${canFacilitate ? `${hereCount}/${allParticipants.length} present` : (myHere ? "You're in the meeting" : "Join when ready")}</span>
          <div style="display:flex;gap:4px;align-items:center;">
            ${canFacilitate && scope.type === "global" ? `
              <button class="btn-aewttr-ghost btn-aewttr-sm" id="mtg-add-guest"${tip("Add a guest or outside attendee")}><i class="bx bx-user-plus"></i></button>
              <button class="btn-aewttr-ghost btn-aewttr-sm" id="mtg-manage-roster"${tip("Manage meeting roster")}><i class="bx bx-cog"></i></button>
            ` : ""}
          </div>
        </div>
        <div class="meeting-nav-list">
          ${participants.length ? participants.map(p => `
            <div class="meeting-nav-item ${p.id === activeId ? "active" : ""}" ${canFacilitate ? `data-select="${p.id}"` : ""}>
              ${userAvatarHtml({ name: p.name, email: participantEmail(p), size: 30 })}
              <div class="meeting-nav-info">
                <div class="name">${escapeHtml(p.name)}${p.isGuest ? ` <span class="kc-badge">Guest</span>` : ""}</div>
                <div class="role">${escapeHtml(p.role)}</div>
              </div>
              <button type="button" class="meeting-presence-pill ${session.attendance[p.id] === "Here" ? "here" : "out"}" data-presence="${p.id}" aria-pressed="${session.attendance[p.id] === "Here" ? "true" : "false"}" aria-label="Toggle attendance for ${escapeHtml(p.name)}" ${!canFacilitate && me && p.id !== me.id ? "disabled" : ""}>${session.attendance[p.id] === "Here" ? "Here" : (canFacilitate || !me || p.id === me.id ? "Out" : "—")}</button>
            </div>`).join("") : `<div class="empty-state" style="padding:20px 10px;">${scope.type === "global" ? "No one is on the weekly roster yet." : "No one is on this project's People list yet."}</div>`}
        </div>
        ${!canFacilitate && me && !myHere ? `<button class="btn-aewttr btn-aewttr-sm" id="mtg-join-self" style="width:100%;margin-top:10px;"${tip("Mark yourself as present in the meeting")}><i class="bx bx-log-in"></i> Join Meeting</button>` : ""}
      </div>
      <div class="meeting-queue-col" id="mtg-queue"></div>
    </div>
  `;

  const peopleScroller = $(".meeting-nav-list", body);
  if (peopleScroller) {
    peopleScroller.addEventListener("wheel", (event) => {
      if (!event.deltaY || event.ctrlKey) return;
      const next = Math.max(0, Math.min(peopleScroller.scrollHeight - peopleScroller.clientHeight, peopleScroller.scrollTop + event.deltaY));
      if (next === peopleScroller.scrollTop) return;
      event.preventDefault();
      peopleScroller.scrollTop = next;
    }, { capture: true, passive: false });
  }

  if (canFacilitate) {
    $all("[data-select]", body).forEach(node => node.addEventListener("click", (e) => {
      if (e.target.closest("[data-presence]")) return;
      window.AEWTTR.state.meetingActive[key] = node.dataset.select;
      renderMeetingLive(body, scope);
    }));
  }
  $all("[data-presence]", body).forEach(btn => btn.addEventListener("click", (e) => {
    e.stopPropagation();
    e.preventDefault();
    toggleMeetingPresence(scope, body, session, btn.dataset.presence, { canFacilitate, me, allParticipants });
  }));
  const joinBtn = $("#mtg-join-self", body);
  if (joinBtn && me) {
    joinBtn.addEventListener("click", (e) => {
      e.preventDefault();
      toggleMeetingPresence(scope, body, session, me.id, { canFacilitate, me, allParticipants });
    });
  }
  const guestBtn = $("#mtg-add-guest", body);
  if (guestBtn) guestBtn.addEventListener("click", () => openAddMeetingGuestModal(scope, () => renderMeetingLive(body, scope)));
  const rosterBtn = $("#mtg-manage-roster", body);
  if (rosterBtn) rosterBtn.addEventListener("click", () => openWeeklyRosterModal(() => renderMeetingLive(body, scope)));

  if (active) renderMeetingQueue($("#mtg-queue", body), scope, active);
  else $("#mtg-queue", body).innerHTML = `<div class="empty-state" style="padding:40px;">${me ? "Join the meeting to update your tasks." : "You're not on the roster for this meeting."}</div>`;
}

/* ---------- project view: all projects in one collapsible list ---------- */
function openProjectMeetingNotesModal(project, onSaved) {
  const extra = ensureProjectExtra(project.id);
  if (!Array.isArray(extra.notes)) extra.notes = [];
  // Reuse the shared notes-chat modal (same UI as task/subtask notes).
  const notesHost = { notes: extra.notes };
  const knownIds = new Set((extra.notes || []).map((n) => n && n.id).filter(Boolean));
  openTaskNotesModal(notesHost, `Project notes — ${project.id}`, async () => {
    extra.notes = notesHost.notes || [];
    const newest = (extra.notes || []).find((n) => n && n.id && !knownIds.has(n.id));
    if (newest) {
      knownIds.add(newest.id);
      recordActiveMeetingNotePosted(project.id, newest.text, { kind: "project", projectId: project.id });
    }
    const proj = window.AEWTTR.db.projects.find((p) => p.id === project.id) || project;
    await Repo.save("project", proj);
    if (typeof onSaved === "function") onSaved();
  });
}

function openMeetingProjectDetailsModal(project, onSaved) {
  const db = window.AEWTTR.db;
  const proj = (db.projects || []).find((p) => p.id === project.id) || project;
  const extra = typeof ensureProjectExtra === "function" ? ensureProjectExtra(proj.id) : ((db.projectExtra && db.projectExtra[proj.id]) || { handoff: "" });
  const selectedPortfolios = new Set(typeof projectPortfolios === "function" ? projectPortfolios(proj) : (proj.portfolios || []));
  const selectedLocations = new Set(typeof projectLocations === "function" ? projectLocations(proj) : (proj.locations || []));
  const selectedConfigEnd = new Set();
  const initialConfigEnd = typeof normalizeConfigEndItemName === "function"
    ? normalizeConfigEndItemName(proj.configEndItem || "")
    : String(proj.configEndItem || "").trim();
  if (initialConfigEnd) selectedConfigEnd.add(initialConfigEnd);

  let saveTimer = null;
  let dirty = false;

  const modal = openModal(`
    <div class="aewttr-modal-head meeting-det-head">
      <div class="meeting-det-head-copy">
        <div class="meeting-det-kicker">${escapeHtml(proj.name || "Project")}</div>
        <h3>Project settings</h3>
      </div>
      <div class="meeting-det-head-actions">
        <span class="meeting-project-details-status" id="mtg-det-status" data-state="saved">Saved</span>
        <button class="aewttr-modal-close" type="button" aria-label="Close">&times;</button>
      </div>
    </div>
    <div class="aewttr-modal-body meeting-project-details">
      <div class="meeting-det-tabs" role="tablist">
        <button type="button" class="meeting-det-tab is-active" data-det-tab="general">General</button>
        <button type="button" class="meeting-det-tab" data-det-tab="roles">Roles</button>
        <button type="button" class="meeting-det-tab" data-det-tab="funding">Funding</button>
      </div>

      <div class="meeting-det-pane is-active" data-det-pane="general">
        <div class="meeting-det-stack">
          <div class="form-row"><label>Project name</label><input class="input-aewttr" id="mtg-det-name" value="${escapeHtml(proj.name || "")}"></div>
          <div class="form-row"><label>Description</label><textarea class="textarea-aewttr" id="mtg-det-desc" rows="2" placeholder="Short project description">${escapeHtml(proj.description || "")}</textarea></div>
          <div class="meeting-det-two">
            <div class="form-row"><label>Priority</label>
              <select class="select-aewttr" id="mtg-det-priority">
                <option value="" ${!proj.priority ? "selected" : ""}>None</option>
                <option value="High" ${proj.priority === "High" ? "selected" : ""}>High</option>
                <option value="Medium" ${proj.priority === "Medium" ? "selected" : ""}>Medium</option>
                <option value="Low" ${proj.priority === "Low" ? "selected" : ""}>Low</option>
              </select>
            </div>
            <div class="form-row"><label>Lifecycle</label>
              <select class="select-aewttr" id="mtg-det-lifecycle">
                <option value="" ${!proj.lifecycleStatus ? "selected" : ""}>None</option>
                <option value="Planning" ${proj.lifecycleStatus === "Planning" ? "selected" : ""}>Planning</option>
                <option value="Active" ${proj.lifecycleStatus === "Active" ? "selected" : ""}>Active</option>
                <option value="On Hold" ${proj.lifecycleStatus === "On Hold" ? "selected" : ""}>On Hold</option>
                <option value="Completed" ${proj.lifecycleStatus === "Completed" ? "selected" : ""}>Completed</option>
              </select>
            </div>
          </div>
          <div class="meeting-det-three">
            <div class="form-row"><label>Start</label><input type="date" class="input-aewttr" id="mtg-det-startdate" value="${escapeHtml(proj.startDate || "")}"></div>
            <div class="form-row"><label>Due</label><input type="date" class="input-aewttr" id="mtg-det-duedate" value="${escapeHtml(proj.dueDate || "")}"></div>
            <div class="form-row"><label>Completed</label><input type="date" class="input-aewttr" id="mtg-det-completiondate" value="${escapeHtml(proj.completionDate || "")}"></div>
          </div>
          <div class="form-row">
            <label>Change request</label>
            <div class="travel-choice-row">${typeof travelChoiceGroup === "function" ? travelChoiceGroup("mtg-det-crr", ["Yes", "No"], proj.changeRequestRequired ? "Yes" : "No") : ""}</div>
          </div>
          <div class="meeting-det-two">
            <div class="form-row">
              <label>Portfolios</label>
              ${typeof portfolioPickerHtml === "function" ? portfolioPickerHtml(Array.from(selectedPortfolios), "mtg-det-portfolios") : ""}
            </div>
            <div class="form-row">
              <label>Locations</label>
              ${typeof locationPickerHtml === "function" ? locationPickerHtml(Array.from(selectedLocations), "mtg-det-locations") : ""}
            </div>
          </div>
        </div>
      </div>

      <div class="meeting-det-pane" data-det-pane="roles" hidden>
        <p class="meeting-project-details-hint">Type at least 2 letters to search and replace someone.</p>
        <div class="project-role-picker-grid meeting-project-details-roles">
          ${(typeof ASSIGNABLE_PROJECT_ROLE_FIELDS !== "undefined" ? ASSIGNABLE_PROJECT_ROLE_FIELDS : [{ key: "pm", label: "Owner (PM)" }])
            .map((field) => typeof projectRolePickerHtml === "function" ? projectRolePickerHtml(proj, field.key === "pm" ? { ...field, label: "Owner (PM)" } : field) : "")
            .join("")}
        </div>
      </div>

      <div class="meeting-det-pane" data-det-pane="funding" hidden>
        <div class="meeting-det-stack">
          <div class="meeting-det-two">
            <div class="form-row"><label>Task Order</label><input class="input-aewttr" id="mtg-det-taskorder" value="${escapeHtml(proj.taskOrder || "")}"></div>
            <div class="form-row"><label>Funding Type</label><input class="input-aewttr" id="mtg-det-fundingtype" value="${escapeHtml(proj.fundingType || "")}"></div>
            <div class="form-row"><label>Fiscal Year</label><input class="input-aewttr" id="mtg-det-fiscalyear" value="${escapeHtml(proj.fiscalYear || "")}" placeholder="e.g. FY26"></div>
            <div class="form-row"><label>Funding Status</label><input class="input-aewttr" id="mtg-det-fundingstatus" value="${escapeHtml(proj.fundingStatus || "")}"></div>
          </div>
          <div class="form-row"><label>Configuration End Item</label>${typeof configEndItemPickerHtml === "function" ? configEndItemPickerHtml(Array.from(selectedConfigEnd), "mtg-det-config") : `<input class="input-aewttr" id="mtg-det-configenditem" value="${escapeHtml(proj.configEndItem || "")}">`}</div>
          <div class="form-row" style="margin-bottom:0;"><label>Handoff notes</label><textarea class="textarea-aewttr" id="mtg-det-handoff" rows="4">${escapeHtml(extra.handoff || "")}</textarea></div>
        </div>
      </div>

      <div class="meeting-project-details-foot">
        <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="mtg-details-settings">Open full settings page</button>
      </div>
    </div>
  `, { wide: true, className: "meeting-project-details-modal" });

  $all("[data-det-tab]", modal).forEach((tab) => tab.addEventListener("click", () => {
    const key = tab.dataset.detTab;
    $all("[data-det-tab]", modal).forEach((t) => t.classList.toggle("is-active", t === tab));
    $all("[data-det-pane]", modal).forEach((pane) => {
      const active = pane.dataset.detPane === key;
      pane.classList.toggle("is-active", active);
      pane.hidden = !active;
    });
  }));

  const statusEl = $("#mtg-det-status", modal);
  function setStatus(state, message) {
    if (!statusEl) return;
    statusEl.dataset.state = state;
    statusEl.textContent = message;
  }

  async function persistDetails() {
    if (typeof reanchorProject === "function") reanchorProject(proj);
    const name = ($("#mtg-det-name", modal).value || "").trim();
    if (!name) {
      setStatus("error", "Project name is required");
      return;
    }
    proj.name = name;
    proj.description = ($("#mtg-det-desc", modal).value || "").trim();
    proj.priority = $("#mtg-det-priority", modal).value;
    proj.lifecycleStatus = $("#mtg-det-lifecycle", modal).value;
    proj.startDate = $("#mtg-det-startdate", modal).value;
    proj.dueDate = $("#mtg-det-duedate", modal).value;
    proj.completionDate = $("#mtg-det-completiondate", modal).value;
    const crrChecked = $(`input[name="mtg-det-crr"]:checked`, modal);
    proj.changeRequestRequired = !!crrChecked && crrChecked.value === "Yes";
    proj.portfolios = Array.from(selectedPortfolios);
    if (typeof rememberPortfolioNames === "function") rememberPortfolioNames(proj.portfolios);
    proj.locations = Array.from(selectedLocations);
    proj.taskOrder = ($("#mtg-det-taskorder", modal).value || "").trim();
    proj.fundingType = ($("#mtg-det-fundingtype", modal).value || "").trim();
    proj.fiscalYear = ($("#mtg-det-fiscalyear", modal).value || "").trim();
    proj.fundingStatus = ($("#mtg-det-fundingstatus", modal).value || "").trim();
    proj.configEndItem = Array.from(selectedConfigEnd)[0] || (($("#mtg-det-configenditem", modal) || {}).value || "").trim();
    if (proj.configEndItem && typeof rememberConfigEndItemNames === "function") {
      rememberConfigEndItemNames([proj.configEndItem]);
    }
    extra.handoff = $("#mtg-det-handoff", modal).value || "";
    if (!db.projectExtra) db.projectExtra = {};
    db.projectExtra[proj.id] = extra;
    proj.updated = new Date().toISOString().slice(0, 10);
    setStatus("saving", "Saving…");
    try {
      await Repo.save("project", proj);
      dirty = false;
      setStatus("saved", "Saved");
    } catch (e) {
      setStatus("error", (e && e.friendly) || "Couldn’t save");
    }
  }

  function schedulePersist() {
    dirty = true;
    setStatus("pending", "Unsaved changes…");
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      saveTimer = null;
      persistDetails();
    }, 400);
  }

  ["mtg-det-name", "mtg-det-desc", "mtg-det-taskorder", "mtg-det-fundingtype", "mtg-det-fiscalyear", "mtg-det-fundingstatus", "mtg-det-handoff"].forEach((id) => {
    const field = $(`#${id}`, modal);
    if (!field) return;
    field.addEventListener("input", schedulePersist);
    field.addEventListener("change", schedulePersist);
  });
  ["mtg-det-priority", "mtg-det-lifecycle", "mtg-det-startdate", "mtg-det-duedate", "mtg-det-completiondate"].forEach((id) => {
    const field = $(`#${id}`, modal);
    if (field) field.addEventListener("change", schedulePersist);
  });
  $all(`input[name="mtg-det-crr"]`, modal).forEach((input) => input.addEventListener("change", schedulePersist));

  if (typeof wirePortfolioPicker === "function") {
    wirePortfolioPicker(modal, selectedPortfolios, "mtg-det-portfolios", schedulePersist);
  }
  if (typeof wireLocationPicker === "function") {
    wireLocationPicker(modal, selectedLocations, "mtg-det-locations", schedulePersist);
  }
  if (typeof wireConfigEndItemPicker === "function") {
    wireConfigEndItemPicker(modal, selectedConfigEnd, "mtg-det-config", schedulePersist);
  }
  if (typeof wireProjectRolePickers === "function") {
    wireProjectRolePickers(modal, proj, () => {
      // Role write already landed. Do not clear form dirty/pending — that
      // cancelled unfinished text autosaves when closing the modal.
      if (dirty || saveTimer) return;
      setStatus("saved", "Saved");
    });
  }

  async function finishClose() {
    clearTimeout(saveTimer);
    if (dirty) await persistDetails();
    closeModal();
    if (typeof onSaved === "function") onSaved();
  }

  $(".aewttr-modal-close", modal).addEventListener("click", () => { finishClose(); });
  $("#mtg-details-settings", modal).addEventListener("click", async () => {
    clearTimeout(saveTimer);
    if (dirty) await persistDetails();
    closeModal();
    navigate(`projects/${proj.id}/settings`);
  });
  const backdrop = modal.closest(".aewttr-modal-backdrop");
  if (backdrop) {
    backdrop.onmousedown = (e) => {
      if (e.target === backdrop) finishClose();
    };
  }
}

function meetingProjectFilterOptions(projects) {
  const names = new Set();
  projects.forEach((project) => meetingProjectPortfolios(project).forEach((p) => names.add(p)));
  const portfolios = [...names].sort((a, b) => a.localeCompare(b));
  const statuses = [...new Set(projects.map(meetingProjectFilterStatus))].sort((a, b) => a.localeCompare(b));
  return { portfolios, statuses };
}

function meetingProjectsMatchingFilters(projects, filters) {
  const portfolios = filters.portfolios || [];
  const statuses = filters.statuses || [];
  return projects.filter((project) => {
    if (portfolios.length) {
      const projectPorts = meetingProjectPortfolios(project);
      if (!projectPorts.some((p) => portfolios.includes(p))) return false;
    }
    if (statuses.length && !statuses.includes(meetingProjectFilterStatus(project))) return false;
    return true;
  });
}

function toggleMeetingFilterValue(list, value) {
  const next = (list || []).slice();
  const idx = next.indexOf(value);
  if (idx >= 0) next.splice(idx, 1);
  else next.push(value);
  return next;
}

function renderMeetingAllProjectsView(body, scope, opts) {
  // The full task list owns its own scroll panel. Preserve it when a task
  // action, filter change, or background refresh rebuilds this view; otherwise
  // replacing body.innerHTML puts the user back at the top of a long list.
  const previousTasksScroller = body.querySelector(".meeting-project-main-body--tasks");
  const previousTasksScrollTop = previousTasksScroller ? previousTasksScroller.scrollTop : 0;
  const { canFacilitate, allParticipants, hereCount, myHere, me, session } = opts;
  const projects = meetingProjects(scope);
  const filters = meetingProjectViewFilters(scope);
  const { portfolios, statuses } = meetingProjectFilterOptions(projects);
  const visible = meetingProjectsMatchingFilters(projects, filters);
  const defaultParticipant = meetingParticipants(scope).find((p) => p.memberId) || meetingParticipants(scope)[0];
  const selectedId = meetingProjectSelectedId(scope, visible);
  const showAllTasks = selectedId === MEETING_ALL_TASKS_ID;
  const selectedProject = showAllTasks ? null : (visible.find((p) => p.id === selectedId) || null);
  const selectedSummary = selectedProject ? projectMeetingSummary(selectedProject) : null;
  const selectedNoteCount = selectedProject ? (ensureProjectExtra(selectedProject.id).notes || []).length : 0;

  body.className = "meeting-live-body meeting-live-body--projects";
  body.innerHTML = `
    <div class="meeting-projects-shell meeting-projects-shell--split">
      <div class="meeting-projects-toolbar">
        <div class="meeting-projects-toolbar-meta">
          <span class="kc-badge">${visible.length} of ${projects.length}</span>
          <span class="meeting-projects-present">${canFacilitate ? `${hereCount}/${allParticipants.length} present` : (myHere ? "You're in the meeting" : "Join when ready")}</span>
        </div>
        <div class="meeting-projects-filters">
          <div class="meeting-filter-dropdown" data-filter-dropdown="portfolio">
            <button type="button" class="meeting-filter-trigger" data-filter-trigger="portfolio"${tip("Filter by portfolio")}>
              <span>Portfolio</span>
              ${filters.portfolios.length ? `<em>${filters.portfolios.length}</em>` : ""}
              <i class="bx bx-chevron-down"></i>
            </button>
            <div class="meeting-filter-menu" hidden>
              ${portfolios.length ? portfolios.map((p) => `
                <label class="meeting-filter-option">
                  <input type="checkbox" data-filter-portfolio="${escapeHtml(p)}" ${filters.portfolios.includes(p) ? "checked" : ""}>
                  <span>${escapeHtml(p)}</span>
                </label>`).join("") : `<div class="meeting-filter-empty">No portfolios</div>`}
            </div>
          </div>
          <div class="meeting-filter-dropdown" data-filter-dropdown="status">
            <button type="button" class="meeting-filter-trigger" data-filter-trigger="status"${tip("Filter by status")}>
              <span>Status</span>
              ${filters.statuses.length ? `<em>${filters.statuses.length}</em>` : ""}
              <i class="bx bx-chevron-down"></i>
            </button>
            <div class="meeting-filter-menu" hidden>
              ${statuses.length ? statuses.map((s) => `
                <label class="meeting-filter-option">
                  <input type="checkbox" data-filter-status="${escapeHtml(s)}" ${filters.statuses.includes(s) ? "checked" : ""}>
                  <span>${escapeHtml(s)}</span>
                </label>`).join("") : `<div class="meeting-filter-empty">No statuses</div>`}
            </div>
          </div>
          ${(filters.portfolios.length || filters.statuses.length) ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="mtg-clear-project-filters"${tip("Clear portfolio and status filters")}>Clear</button>` : ""}
        </div>
        <div class="meeting-projects-toolbar-actions">
          ${canFacilitate && scope.type === "global" ? `
            <button class="btn-aewttr-ghost btn-aewttr-sm" id="mtg-add-guest"${tip("Add a guest or outside attendee")}><i class="bx bx-user-plus"></i></button>
            <button class="btn-aewttr-ghost btn-aewttr-sm" id="mtg-manage-roster"${tip("Manage meeting roster")}><i class="bx bx-cog"></i></button>
          ` : ""}
          ${!canFacilitate && me && !myHere ? `<button class="btn-aewttr btn-aewttr-sm" id="mtg-join-self"${tip("Mark yourself as present in the meeting")}><i class="bx bx-log-in"></i> Join Meeting</button>` : ""}
        </div>
      </div>
      <div class="meeting-projects-split">
        <aside class="project-spo-nav meeting-projects-sidebar">
          <div class="meeting-projects-sidebar-list" id="mtg-projects-list">
            <button type="button" class="project-spo-link project-spo-link--meeting project-spo-link--tasks ${showAllTasks ? "active" : ""}" data-select-project="${MEETING_ALL_TASKS_ID}"${tip("Full nested task list across projects")}>
              <i class="bx bx-list-ul meeting-sidebar-tasks-icon" aria-hidden="true"></i>
              <span class="project-spo-link-copy">
                <strong>Full task list</strong>
                <span>Every project · task · subtask</span>
              </span>
            </button>
            <div class="meeting-sidebar-divider" aria-hidden="true"></div>
            ${visible.length ? visible.map((project) => {
              const summary = projectMeetingSummary(project);
              const isActive = project.id === selectedId;
              return `
                <button type="button" class="project-spo-link project-spo-link--meeting ${isActive ? "active" : ""}" data-select-project="${escapeHtml(project.id)}"${tip(project.name || project.id)}>
                  <span class="project-spo-link-copy">
                    <strong>${escapeHtml(project.name || project.id)}</strong>
                    <span>${escapeHtml(summary.portfolioLabel || "")}</span>
                  </span>
                  ${project.priority ? (typeof priorityTag === "function" ? priorityTag(project.priority) : "") : ""}
                </button>`;
            }).join("") : `<div class="empty-state" style="padding:24px 14px;">${projects.length ? "No projects match the selected filters." : "No projects yet."}</div>`}
          </div>
        </aside>
        <section class="meeting-project-main">
          ${showAllTasks ? `
            <div class="meeting-project-main-head">
              <div class="meeting-project-main-title">
                <strong>Full task list</strong>
                <span>${visible.length} project${visible.length === 1 ? "" : "s"} · nested tasks &amp; subtasks</span>
              </div>
            </div>
            <div class="meeting-project-main-body meeting-project-main-body--tasks" id="mtg-all-tasks-mount"></div>
          ` : selectedProject ? `
            <div class="meeting-project-main-head">
              <div class="meeting-project-main-title">
                <strong>${escapeHtml(selectedProject.name || selectedProject.id)}</strong>
                <span>${escapeHtml(selectedSummary.portfolioLabel || "")}</span>
              </div>
              <div class="meeting-project-main-actions">
                ${selectedProject.priority ? (typeof priorityTag === "function" ? priorityTag(selectedProject.priority) : "") : ""}
                <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="mtg-proj-expand-all"${tip("Expand all dividers")}><i class="bx bx-expand-alt"></i> Expand all</button>
                <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="mtg-proj-collapse-all"${tip("Collapse all dividers")}><i class="bx bx-collapse-alt"></i> Collapse all</button>
                <button type="button" class="btn-aewttr-outline btn-aewttr-sm meeting-project-details-btn" data-project-details="${escapeHtml(selectedProject.id)}"${tip("View portfolios, locations, priority, and owner")}>
                  <i class="bx bx-info-circle"></i> Details
                </button>
                <button type="button" class="btn-aewttr btn-aewttr-sm meeting-notes-btn ${selectedNoteCount ? "has-notes" : ""}" data-project-notes="${escapeHtml(selectedProject.id)}"${tip("Open project-wide notes")}>
                  <i class="bx bx-note"></i> Notes${selectedNoteCount ? ` (${selectedNoteCount})` : ""}
                </button>
              </div>
            </div>
            <div class="meeting-project-main-body" data-project-body="${escapeHtml(selectedProject.id)}"></div>
          ` : `<div class="empty-state" style="padding:60px 20px;">${projects.length ? "Select a project from the list." : "No projects yet."}</div>`}
        </section>
      </div>
    </div>
  `;

  function closeMeetingFilterMenus(except) {
    $all(".meeting-filter-menu", body).forEach((menu) => {
      if (except && menu === except) return;
      menu.hidden = true;
      const wrap = menu.closest(".meeting-filter-dropdown");
      if (wrap) wrap.classList.remove("is-open");
    });
  }
  $all("[data-filter-trigger]", body).forEach((btn) => btn.addEventListener("click", (e) => {
    e.stopPropagation();
    const wrap = btn.closest(".meeting-filter-dropdown");
    const menu = $(".meeting-filter-menu", wrap);
    const willOpen = menu.hidden;
    closeMeetingFilterMenus(willOpen ? menu : null);
    menu.hidden = !willOpen;
    wrap.classList.toggle("is-open", willOpen);
  }));
  $all("[data-filter-portfolio]", body).forEach((input) => input.addEventListener("change", () => {
    filters.portfolios = toggleMeetingFilterValue(filters.portfolios, input.dataset.filterPortfolio);
    renderMeetingAllProjectsView(body, scope, opts);
  }));
  $all("[data-filter-status]", body).forEach((input) => input.addEventListener("change", () => {
    filters.statuses = toggleMeetingFilterValue(filters.statuses, input.dataset.filterStatus);
    renderMeetingAllProjectsView(body, scope, opts);
  }));
  body.onclick = (e) => {
    if (!e.target.closest(".meeting-filter-dropdown")) closeMeetingFilterMenus();
  };
  $all(".meeting-filter-menu", body).forEach((menu) => menu.addEventListener("click", (e) => e.stopPropagation()));
  const clearBtn = $("#mtg-clear-project-filters", body);
  if (clearBtn) clearBtn.addEventListener("click", () => {
    filters.portfolios = [];
    filters.statuses = [];
    renderMeetingAllProjectsView(body, scope, opts);
  });

  $all("[data-select-project]", body).forEach((btn) => btn.addEventListener("click", () => {
    const key = meetingScopeKey(scope);
    window.AEWTTR.state.meetingProjectSelected[key] = btn.dataset.selectProject;
    renderMeetingAllProjectsView(body, scope, opts);
  }));

  $all("[data-project-notes]", body).forEach((btn) => btn.addEventListener("click", (e) => {
    e.stopPropagation();
    const project = projects.find((p) => p.id === btn.dataset.projectNotes);
    if (!project) return;
    openProjectMeetingNotesModal(project, async () => {
      await saveMeetingSession(scope);
      renderMeetingAllProjectsView(body, scope, opts);
    });
  }));
  $all("[data-project-details]", body).forEach((btn) => btn.addEventListener("click", (e) => {
    e.stopPropagation();
    const project = projects.find((p) => p.id === btn.dataset.projectDetails);
    if (!project) return;
    openMeetingProjectDetailsModal(project, async () => {
      await saveMeetingSession(scope);
      renderMeetingAllProjectsView(body, scope, opts);
    });
  }));

  const guestBtn = $("#mtg-add-guest", body);
  if (guestBtn) guestBtn.addEventListener("click", () => openAddMeetingGuestModal(scope, () => renderMeetingLive(body, scope)));
  const rosterBtn = $("#mtg-manage-roster", body);
  if (rosterBtn) rosterBtn.addEventListener("click", () => openWeeklyRosterModal(() => renderMeetingLive(body, scope)));
  const joinBtn = $("#mtg-join-self", body);
  if (joinBtn && me) {
    joinBtn.addEventListener("click", (e) => {
      e.preventDefault();
      toggleMeetingPresence(scope, body, session, me.id, { canFacilitate, me, allParticipants });
    });
  }

  if (showAllTasks) {
    const mount = $("#mtg-all-tasks-mount", body);
    if (mount) {
      renderMeetingFullTasksList(mount, scope, visible, () => renderMeetingAllProjectsView(body, scope, opts));
    }
  } else if (selectedProject) {
    const mount = safeQueryByAttr(body, "data-project-body", selectedProject.id);
    if (mount) {
      renderMeetingProjectTracker(mount, scope, selectedProject, defaultParticipant, () => renderMeetingAllProjectsView(body, scope, opts));
      const projStateKey = meetingTrackerStateKey(scope, selectedProject.id);
      const projTasks = (window.AEWTTR.db.ganttTasks && window.AEWTTR.db.ganttTasks[selectedProject.id]) || [];
      const projDividers = typeof trackerDividers === "function" ? trackerDividers(projTasks) : [];
      const onProjRedraw = () => renderMeetingProjectTracker(mount, scope, selectedProject, defaultParticipant, () => renderMeetingAllProjectsView(body, scope, opts));
      const projExpandBtn = $("#mtg-proj-expand-all", body);
      if (projExpandBtn) projExpandBtn.addEventListener("click", () => {
        const dc = typeof trackerDividerCollapsedStore === "function" ? trackerDividerCollapsedStore(projStateKey) : {};
        projDividers.forEach((d) => { dc[d.id] = false; });
        dc.__ungrouped__ = false;
        onProjRedraw();
      });
      const projCollapseBtn = $("#mtg-proj-collapse-all", body);
      if (projCollapseBtn) projCollapseBtn.addEventListener("click", () => {
        const dc = typeof trackerDividerCollapsedStore === "function" ? trackerDividerCollapsedStore(projStateKey) : {};
        projDividers.forEach((d) => { dc[d.id] = true; });
        dc.__ungrouped__ = true;
        onProjRedraw();
      });
    }
  }

  if (showAllTasks && previousTasksScrollTop > 0) {
    requestAnimationFrame(() => {
      const tasksScroller = body.querySelector(".meeting-project-main-body--tasks");
      if (tasksScroller) tasksScroller.scrollTop = previousTasksScrollTop;
    });
  }
}

function renderMeetingProjectTracker(mount, scope, project, defaultParticipant, onRedraw) {
  const stateKey = meetingTrackerStateKey(scope, project.id);
  if (typeof renderTrackerWorkspace !== "function") {
    mount.innerHTML = `<div class="empty-state">Tracker view is loading…</div>`;
    return;
  }
  renderTrackerWorkspace(mount, {
    proj: project,
    stateKey,
    defaultView: "table",
    meetingMode: true,
    saveSource: "Meeting",
    showHint: false,
    hideGroupHeader: true,
    collapseGanttGroupsByDefault: false,
    collapseDividersByDefault: true,
    flatGantt: true,
    defaultAssignee: defaultParticipant ? defaultParticipant.name : "Unassigned",
    onAddTask: () => openMeetingAddTaskModal(scope, defaultParticipant, onRedraw, project.id),
    onRedraw,
    onOpenEditor: (taskId, subIndex) => openMeetingTaskEditor(scope, project.id, taskId, onRedraw, subIndex),
    onAfterSave: () => saveMeetingSession(scope)
  });
}

function renderMeetingProjectQueue(mount, scope, project) {
  // Kept for Around-the-Room / legacy callers — single-project tracker panel.
  function redraw() { renderMeetingProjectQueue(mount, scope, project); }
  const summary = projectMeetingSummary(project);
  const extra = ensureProjectExtra(project.id);
  const noteCount = (extra.notes || []).length;
  const defaultParticipant = meetingParticipants(scope).find((p) => p.memberId) || meetingParticipants(scope)[0];

  mount.innerHTML = `
    <div class="meeting-tracker-panel">
      <div class="meeting-tracker-panel-head">
        <div class="meeting-tracker-panel-title">
          <strong>${escapeHtml(project.id)}</strong>
          <span>${escapeHtml(project.name)}</span>
        </div>
        <div class="meeting-tracker-panel-chips">
          <span class="meeting-project-meta-quiet" title="${escapeHtml(summary.portfolios.join(", "))}">${escapeHtml(summary.portfolioLabel)}</span>
          ${project.priority && typeof priorityTag === "function" ? priorityTag(project.priority) : ""}
        </div>
        <button type="button" class="btn-aewttr-outline btn-aewttr-sm meeting-project-details-btn" id="mtg-details-toggle"${tip("View portfolios, locations, priority, and owner")}><i class="bx bx-info-circle"></i> Details</button>
        <button type="button" class="btn-aewttr btn-aewttr-sm meeting-notes-btn ${noteCount ? "has-notes" : ""}" id="mtg-notes-toggle"${tip("Open project-wide notes")}><i class="bx bx-note"></i> Notes${noteCount ? ` (${noteCount})` : ""}</button>
      </div>
      <div class="meeting-tracker-panel-body" id="mtg-project-tracker-mount"></div>
    </div>
  `;

  renderMeetingProjectTracker($("#mtg-project-tracker-mount", mount), scope, project, defaultParticipant, redraw);

  const detailsToggle = $("#mtg-details-toggle", mount);
  if (detailsToggle) {
    detailsToggle.addEventListener("click", () => openMeetingProjectDetailsModal(project, redraw));
  }
  const notesToggle = $("#mtg-notes-toggle", mount);
  if (notesToggle) {
    notesToggle.addEventListener("click", () => {
      openProjectMeetingNotesModal(project, async () => {
        await saveMeetingSession(scope);
        redraw();
      });
    });
  }
}

function wireMeetingTaskCardHandlers(mount, scope, participant, redraw) {
  function taskFromKey(key) {
    const [pid, taskId] = key.split(":");
    return (window.AEWTTR.db.ganttTasks[pid] || []).find(t => t.id === taskId);
  }
  function toggleTaskExpanded(key) {
    const store = meetingExpandedTasks();
    store[key] = !store[key];
    redraw();
  }
  // Whole row toggles expand/collapse (not just the chevron button) — only
  // the title input and health select opt out, so typing a title or picking
  // a health status doesn't also collapse the card out from under you.
  $all("[data-task-row]", mount).forEach(row => row.addEventListener("click", (e) => {
    if (e.target.closest("input, select")) return;
    toggleTaskExpanded(row.dataset.taskRow);
  }));
  $all("[data-task-title]", mount).forEach(input => input.addEventListener("change", async () => {
    const key = input.dataset.taskTitle;
    const task = taskFromKey(key);
    const title = input.value.trim();
    if (!task || !title) { input.value = task ? task.title : ""; return; }
    task.title = title;
    await saveMeetingTask(scope, task, key.split(":")[0]);
    toast("Task title updated", "success");
    redraw();
  }));
  $all("[data-health]", mount).forEach(sel => sel.addEventListener("change", async () => {
    const key = sel.dataset.health;
    const task = taskFromKey(key);
    if (!task) return;
    const pid = key.split(":")[0];
    const prev = task.health;
    task.health = sel.value;
    recordActiveMeetingTaskChange(pid, task, "health", task.health, { prev });
    await saveMeetingTask(scope, task, pid);
    toast(`"${task.title}" marked ${task.health}`, "success");
    redraw();
  }));
  $all("[data-owner-wrap]", mount).forEach(wrap => {
    const key = wrap.dataset.ownerWrap;
    const task = taskFromKey(key);
    if (!task) return;
    const safeKey = meetingOwnerSafeKey(key);
    const people = task.assignee && task.assignee !== "Unassigned" ? [{ name: task.assignee, email: "" }] : [];
    const picker = wirePeoplePicker(mount, people, { mount: `mto-sel-${safeKey}`, input: `mto-input-${safeKey}`, suggestions: `mto-sugg-${safeKey}` }, { singleSelect: true, allowManualEmail: false });
    picker.setOnChange(async () => {
      // Deliberately no redraw() here: in "Around the Room" the task list is
      // grouped by assignee, so an immediate redraw would yank the task out
      // from under the person you're currently reviewing the instant the
      // owner field is cleared (before a replacement is even picked). Leave
      // it visible where it is — the new/empty owner is still saved — and it
      // naturally regroups next time the queue re-renders (switching people,
      // switching tabs, etc). This also keeps the just-picked person's email
      // showing in the chip instead of it vanishing on a forced re-render
      // (task.assignee only stores a name, so reconstructing the picker from
      // scratch loses the email until the directory is re-searched).
      task.assignee = people[0] ? people[0].name : "Unassigned";
      await saveMeetingTask(scope, task, key.split(":")[0]);
    });
  });
  $all("[data-due]", mount).forEach(input => input.addEventListener("change", async () => {
    const key = input.dataset.due;
    const task = taskFromKey(key);
    if (!task || !input.value) return;
    task.end = input.value;
    await saveMeetingTask(scope, task, key.split(":")[0]);
    toast("Due date updated", "success");
  }));
  $all("[data-open-meeting-task-notes]", mount).forEach(btn => btn.addEventListener("click", () => {
    const key = btn.dataset.openMeetingTaskNotes;
    const task = taskFromKey(key);
    if (!task) return;
    const pid = key.split(":")[0];
    const knownIds = new Set((task.notes || []).map((n) => n && n.id).filter(Boolean));
    openTaskNotesModal(task, task.title, () => {
      const newest = (task.notes || []).find((n) => n && n.id && !knownIds.has(n.id));
      if (newest) {
        knownIds.add(newest.id);
        recordActiveMeetingNotePosted(pid, newest.text, {
          kind: "task",
          projectId: pid,
          taskId: task.id,
          taskTitle: task.title || "Untitled"
        });
      }
      saveMeetingTask(scope, task, pid);
      redraw();
    });
  }));
  $all("[data-subtask-done]", mount).forEach(cb => cb.addEventListener("change", async () => {
    const [pid, taskId, idx] = cb.dataset.subtaskDone.split(":");
    const task = (window.AEWTTR.db.ganttTasks[pid] || []).find(t => t.id === taskId);
    if (!task || !task.subtasks[+idx]) return;
    task.subtasks[+idx].done = cb.checked;
    syncTaskStatusFromSubtasks(task);
    await saveMeetingTask(scope, task, pid);
    redraw();
  }));
  $all("[data-subtask-text]", mount).forEach(input => input.addEventListener("change", async () => {
    const [pid, taskId, idx] = input.dataset.subtaskText.split(":");
    const task = (window.AEWTTR.db.ganttTasks[pid] || []).find(t => t.id === taskId);
    const text = input.value.trim();
    if (!task || !task.subtasks[+idx] || !text) { if (task && task.subtasks[+idx]) input.value = task.subtasks[+idx].text; return; }
    task.subtasks[+idx].text = text;
    await saveMeetingTask(scope, task, pid);
    toast("Subtask updated", "success");
    redraw();
  }));
  // Start/end dates for a subtask are edited in the full subtask editor now —
  // the inline row was simplified to just checkbox + text + a compact
  // assignee/due summary, so there's no inline start/end input to wire here.
  $all("[data-open-task-editor]", mount).forEach(btn => btn.addEventListener("click", () => {
    const [pid, taskId] = btn.dataset.openTaskEditor.split(":");
    openMeetingTaskEditor(scope, pid, taskId, redraw);
  }));
  $all("[data-open-sub-editor]", mount).forEach(btn => btn.addEventListener("click", () => {
    const [pid, taskId, idx] = btn.dataset.openSubEditor.split(":");
    openMeetingTaskEditor(scope, pid, taskId, redraw, +idx);
  }));
  $all("[data-add-subtask]", mount).forEach(input => input.addEventListener("keydown", async (e) => {
    if (e.key !== "Enter") return;
    const text = input.value.trim();
    if (!text) return;
    const [pid, taskId] = input.dataset.addSubtask.split(":");
    const task = (window.AEWTTR.db.ganttTasks[pid] || []).find(t => t.id === taskId);
    if (!task) return;
    task.subtasks = task.subtasks || [];
    task.subtasks.push({
      text,
      assignee: participant.name,
      done: false,
      start: task.start,
      end: "",
      relatedDocs: []
    });
    logMeetingActivity(scope, `${currentUserName()} added a subitem to ${task.title || "Untitled"}: ${text}.`, {
      type: "create",
      projectId: pid,
      taskId: task.id,
      taskTitle: task.title || "Untitled",
      subitemText: text
    });
    await saveMeetingTask(scope, task, pid);
    toast("Subtask added", "success");
    redraw();
  }));
  $all("[data-delete-task]", mount).forEach(btn => btn.addEventListener("click", async () => {
    const ok = await confirmDialog({ title: "Delete task", message: "Delete this task?", confirmLabel: "Delete", danger: true });
    if (!ok) return;
    const [pid, taskId] = btn.dataset.deleteTask.split(":");
    const task = (window.AEWTTR.db.ganttTasks[pid] || []).find(t => t.id === taskId);
    window.AEWTTR.db.ganttTasks[pid] = (window.AEWTTR.db.ganttTasks[pid] || []).filter(t => t.id !== taskId);
    logMeetingActivity(scope, `${currentUserName()} removed task "${task ? task.title : taskId}".`);
    if (task) Repo.remove("actionItem", task);
    saveMeetingSession(scope);
    toast("Task deleted", "success");
    redraw();
  }));
}

function wireMeetingTaskInteractions(mount, scope, participant, redraw) {
  $all("[data-toggle-notes]", mount).forEach(btn => btn.addEventListener("click", () => {
    const panel = $(`#mtg-pnotes-${btn.dataset.toggleNotes}`, mount);
    panel.style.display = panel.style.display === "none" ? "block" : "none";
  }));
  wireMeetingTaskCardHandlers(mount, scope, participant, redraw);
}

function openAddMeetingGuestModal(scope, onDone) {
  if (!canManageMeetings()) {
    toast("Only Meeting Admins can add guests.", "error");
    return;
  }
  const session = activeMeetingSession(scope);
  if (!session) return;
  const existingGuests = meetingSessionGuests(session);
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>Add Guest Attendees</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body">
      <p style="font-size:12px;color:var(--aewttr-muted);margin-top:0;">Type guest names manually — one per line. Guests are not linked to SharePoint or Microsoft 365 accounts.</p>
      <div class="form-row"><label>Guest names</label><textarea class="textarea-aewttr" id="mg-names" placeholder="Jane Smith&#10;Alex Rivera" style="min-height:120px;"></textarea></div>
      ${existingGuests.length ? `<div style="margin-top:14px;"><div class="side-panel-title">Already in this session</div><div class="meeting-guest-existing">${existingGuests.map(g => `<span class="traveler-chip"><span>${escapeHtml(g.name)}</span></span>`).join("")}</div></div>` : ""}
    </div>
    <div class="aewttr-modal-foot">
      <button class="btn-aewttr-ghost" id="mg-cancel">Cancel</button>
      <button class="btn-aewttr" id="mg-save">Add Guests</button>
    </div>
  `);
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#mg-cancel", modal).addEventListener("click", closeModal);
  $("#mg-save", modal).addEventListener("click", async () => {
    const names = $("#mg-names", modal).value.split("\n").map((line) => line.trim()).filter(Boolean);
    if (!names.length) { toast("Enter at least one guest name.", "error"); return; }
    let added = 0;
    names.forEach((name) => {
      if (existingGuests.some((g) => g.name.toLowerCase() === name.toLowerCase())) return;
      const guest = { id: uid("guest"), name, email: "", addedBy: currentUserName(), addedAt: new Date().toISOString() };
      existingGuests.push(guest);
      session.attendance[guest.id] = "Out";
      logMeetingActivity(scope, `${currentUserName()} added guest ${name}.`);
      added += 1;
    });
    if (!added) { toast("Those guests are already in this meeting.", "error"); return; }
    await saveMeetingSession(scope);
    closeModal();
    toast(`${added} guest${added === 1 ? "" : "s"} added`, "success");
    if (onDone) onDone();
  });
}

/* ---------- work queue ---------- */
function renderMeetingQueue(mount, scope, participant) {
  function redraw() { renderMeetingQueue(mount, scope, participant); }
  // In Around the Room, tracker controls can catch the wheel event before the
  // scrollable task table receives it. Make the entire right-hand work area
  // scroll its active tracker, no matter which nested element is hovered.
  if (!mount.dataset.meetingQueueWheelRouting) {
    mount.dataset.meetingQueueWheelRouting = "true";
    mount.addEventListener("wheel", (event) => {
      if (!event.deltaY || event.ctrlKey) return;
      const targets = [
        ...mount.querySelectorAll(".monday-table-wrap, .monday-gantt, .tracker-view-mount--meeting")
      ];
      const scroller = targets.find((node) => node.scrollHeight > node.clientHeight + 1);
      if (!scroller) return;

      let delta = event.deltaY;
      if (event.deltaMode === 1) delta *= 16;
      if (event.deltaMode === 2) delta *= scroller.clientHeight;
      const maximumScrollTop = Math.max(0, scroller.scrollHeight - scroller.clientHeight);
      const nextScrollTop = Math.max(0, Math.min(maximumScrollTop, scroller.scrollTop + delta));
      if (nextScrollTop === scroller.scrollTop) return;

      event.preventDefault();
      event.stopPropagation();
      scroller.scrollTop = nextScrollTop;
    }, { capture: true, passive: false });
  }
  const groups = meetingTaskGroups(scope, participant);
  if (!groups.length) {
    mount.innerHTML = `<div class="meeting-tracker-panel meeting-tracker-panel--empty"><div class="empty-state" style="padding:30px;">No tasks assigned to ${escapeHtml(participant.name)} yet.</div></div>`;
    return;
  }

  const personKey = `meeting-person:${meetingScopeKey(scope)}:${participant.id || participant.name}`;
  if (!window.AEWTTR.state.meetingPersonProject) window.AEWTTR.state.meetingPersonProject = {};
  let activePid = window.AEWTTR.state.meetingPersonProject[personKey];
  if (!activePid || !groups.find((g) => g.project.id === activePid)) activePid = groups[0].project.id;
  const activeGroup = groups.find((g) => g.project.id === activePid) || groups[0];
  const notesKey = `${personKey}:${activePid}`;
  const projectExtra = ensureProjectExtra(activeGroup.project.id);
  const taskIds = new Set(activeGroup.tasks.map((t) => t.id));

  mount.innerHTML = `
    <div class="meeting-tracker-panel">
      <div class="meeting-tracker-panel-head meeting-tracker-panel-head--person">
        ${userAvatarHtml({ name: participant.name, email: participantEmail(participant), size: 28 })}
        <div class="meeting-tracker-panel-title">
          <strong>${escapeHtml(participant.name)}</strong>
          <span>${escapeHtml(participant.role)}</span>
        </div>
        <div class="meeting-project-chips" id="mtg-person-project-chips">
          ${groups.map((g) => `
            <button type="button" class="meeting-project-chip ${g.project.id === activePid ? "active" : ""}" data-person-project="${g.project.id}"${tip(g.project.name || g.project.id)}>
              <span>${escapeHtml(g.project.name || g.project.id)}</span> <em>${g.tasks.length}</em>
            </button>`).join("")}
        </div>
        <button type="button" class="btn-aewttr-outline btn-aewttr-sm meeting-project-details-btn" id="mtg-person-details-toggle"${tip("View portfolios, locations, priority, and owner")}><i class="bx bx-info-circle"></i> Details</button>
        <button type="button" class="btn-aewttr btn-aewttr-sm meeting-notes-btn ${projectExtra.notes && projectExtra.notes.length ? "has-notes" : ""}" id="mtg-person-notes-toggle"${tip("Open project-wide notes")}><i class="bx bx-note"></i> Notes${projectExtra.notes && projectExtra.notes.length ? ` (${projectExtra.notes.length})` : ""}</button>
      </div>
      <div class="meeting-tracker-panel-body" id="mtg-person-tracker-mount"></div>
    </div>
  `;

  $all("[data-person-project]", mount).forEach((chip) => chip.addEventListener("click", () => {
    window.AEWTTR.state.meetingPersonProject[personKey] = chip.dataset.personProject;
    redraw();
  }));

  const detailsToggle = $("#mtg-person-details-toggle", mount);
  if (detailsToggle) {
    detailsToggle.addEventListener("click", () => openMeetingProjectDetailsModal(activeGroup.project, redraw));
  }
  const notesToggle = $("#mtg-person-notes-toggle", mount);
  if (notesToggle) {
    notesToggle.addEventListener("click", () => {
      openProjectMeetingNotesModal(activeGroup.project, async () => {
        await saveMeetingSession(scope);
        redraw();
      });
    });
  }

  if (typeof renderTrackerWorkspace === "function") {
    renderTrackerWorkspace($("#mtg-person-tracker-mount", mount), {
      proj: activeGroup.project,
      stateKey: notesKey,
      defaultView: "table",
      meetingMode: true,
      saveSource: "Meeting",
      showHint: false,
      hideGroupHeader: true,
      collapseGanttGroupsByDefault: false,
      expandTasksByDefault: true,
      defaultAssignee: participant ? participant.name : "Unassigned",
      taskFilter: (task) => taskIds.has(task.id),
      onRedraw: redraw,
      onOpenEditor: (taskId, subIndex) => openMeetingTaskEditor(scope, activeGroup.project.id, taskId, redraw, subIndex),
      onAfterSave: () => saveMeetingSession(scope)
    });
  }
}

function meetingProjectGroupHtml(group) {
  const extra = ensureProjectExtra(group.project.id);
  return `
    <div class="meeting-project-group">
      <div class="meeting-project-head">
        <strong>${escapeHtml(group.project.id)} — ${escapeHtml(group.project.name)}</strong>
        <button class="btn-aewttr-ghost btn-aewttr-sm" data-toggle-notes="${group.project.id}"><i class="bx bx-note"></i> Project Notes</button>
      </div>
      <div class="meeting-project-notes" id="mtg-pnotes-${group.project.id}" style="display:none;">
        <textarea class="textarea-aewttr" data-project-notes="${group.project.id}" placeholder="Notes for this project, visible to anyone in the meeting.">${escapeHtml(extra.meetingNotes)}</textarea>
      </div>
      ${group.tasks.length ? group.tasks.map(t => meetingTaskCardHtml(group.project.id, t)).join("") : `<div class="empty-state" style="padding:14px;">No tasks yet for this project.</div>`}
    </div>`;
}

function meetingExpandedTasks() {
  if (!window.AEWTTR.state.meetingTaskExpanded) window.AEWTTR.state.meetingTaskExpanded = {};
  return window.AEWTTR.state.meetingTaskExpanded;
}
function meetingOwnerSafeKey(key) {
  return String(key).replace(/[^a-zA-Z0-9_-]/g, "-");
}

/* task.notes is the same array-of-message shape used by the Tracker's chat
   popup (openTaskNotesModal, app.js) — a task edited from both the Tracker
   and Weekly Meeting shares one notes thread. Older records may still carry
   a legacy plain string here; normalize once so both surfaces agree. */
function normalizeMeetingTaskNotes(task) {
  if (typeof task.notes === "string") {
    const legacy = task.notes.trim();
    task.notes = legacy ? [{ id: uid("nt"), author: "Meeting notes", date: new Date().toISOString().slice(0, 10), time: "", text: legacy }] : [];
  } else if (!Array.isArray(task.notes)) {
    task.notes = [];
  }
  return task.notes;
}

function meetingTaskNotesTriggerHtml(task, key) {
  const notes = normalizeMeetingTaskNotes(task);
  const latest = notes[0] || null;
  return `
    <button type="button" class="monday-notes-trigger" data-open-meeting-task-notes="${escapeHtml(key)}">
      <i class="bx bx-message-rounded-dots"></i>
      ${latest
        ? `<span class="monday-latest-note"><strong>${escapeHtml(latest.author)}</strong>: ${escapeHtml(latest.text)}</span>`
        : `<span class="monday-latest-note monday-latest-note--empty">No updates yet — click to start a thread</span>`}
      ${notes.length ? `<span class="monday-notes-count">${notes.length}</span>` : ""}
    </button>
  `;
}

function meetingTaskCardHtml(pid, task) {
  const key = `${pid}:${task.id}`;
  const subtasks = task.subtasks || [];
  const doneCount = subtasks.filter(s => s.done).length;
  const progressPct = subtasks.length ? Math.round((doneCount / subtasks.length) * 100) : 0;
  const isExpanded = !!meetingExpandedTasks()[key];
  return `
    <div class="meeting-task-card ${isExpanded ? "expanded" : ""}" data-task="${key}">
      <div class="meeting-task-row" data-task-row="${key}">
        <span class="health-dot health-${task.health.replace(/\s+/g, "-")}" title="${escapeHtml(task.health)}"></span>
        <input class="input-aewttr meeting-task-title-input" data-task-title="${key}" value="${escapeHtml(task.title)}" placeholder="Task title" aria-label="Task title">
        ${subtasks.length ? `<span class="meeting-task-row-chip" title="${doneCount} of ${subtasks.length} subtasks done">${doneCount}/${subtasks.length}</span>` : ""}
        <span class="meeting-task-row-owner" title="${escapeHtml(task.assignee || "Unassigned")}">${escapeHtml(task.assignee || "Unassigned")}</span>
        <span class="meeting-task-row-due">${task.end ? fmtDate(task.end) : "No due date"}</span>
        <select class="select-aewttr meeting-health-select health-${task.health.replace(/\s+/g, "-")}" data-health="${key}" aria-label="Task health">
          ${["On Track", "At Risk", "Off Track"].map(h => `<option ${task.health === h ? "selected" : ""}>${h}</option>`).join("")}
        </select>
        <button type="button" class="meeting-task-expand-btn" data-toggle-task="${key}" aria-label="${isExpanded ? "Collapse task" : "Expand task"}"><i class="bx bx-chevron-${isExpanded ? "up" : "down"}"></i></button>
      </div>
      <div class="meeting-task-detail">
        <div class="meeting-task-meta">
          <div class="form-row meeting-owner-field">
            <label>Owner</label>
            <div class="traveler-picker traveler-picker--inline meeting-owner-picker" data-owner-wrap="${key}">
              <div id="mto-sel-${meetingOwnerSafeKey(key)}" class="traveler-chip-list"></div>
              <input class="input-aewttr" id="mto-input-${meetingOwnerSafeKey(key)}" placeholder="Search owner…">
              <div id="mto-sugg-${meetingOwnerSafeKey(key)}" class="traveler-suggestions"></div>
            </div>
          </div>
          <div class="form-row"><label>Due</label><input type="date" class="input-aewttr" data-due="${key}" value="${task.end}"></div>
          ${subtasks.length ? `
          <div class="form-row meeting-subtask-progress">
            <label>Subtasks</label>
            <div class="meeting-subtask-progress-row">
              <div class="meeting-subtask-progress-bar" aria-hidden="true"><span style="width:${progressPct}%;"></span></div>
              <span class="meeting-subtask-progress-label">${doneCount}/${subtasks.length}</span>
            </div>
          </div>` : ""}
        </div>
        <div class="form-row"><label>Notes</label>${meetingTaskNotesTriggerHtml(task, key)}</div>
        <div class="form-row">
          <label>Subtasks</label>
          <div class="meeting-subtask-list">
            ${subtasks.length ? subtasks.map((s, si) => `
              <div class="meeting-subtask-row ${s.done ? "done" : ""}">
                <label class="meeting-subtask-check">
                  <input type="checkbox" data-subtask-done="${pid}:${task.id}:${si}" ${s.done ? "checked" : ""} aria-label="Mark subtask done">
                </label>
                <input class="input-aewttr meeting-subtask-text-row ${s.done ? "done-text" : ""}" data-subtask-text="${pid}:${task.id}:${si}" value="${escapeHtml(s.text)}" placeholder="Subtask description">
                <span class="meeting-subtask-row-meta">${escapeHtml(s.assignee || "Unassigned")}${(s.end || task.end) ? ` · ${fmtDate(s.end || task.end)}` : ""}</span>
                <button type="button" class="meeting-subtask-row-edit" data-open-sub-editor="${pid}:${task.id}:${si}" aria-label="Full subtask editor"${tip("Open full subtask editor")}><i class="bx bx-edit"></i></button>
              </div>`).join("") : `<div class="empty-state" style="padding:10px 0;">No subtasks yet — add one below or open the full task editor.</div>`}
          </div>
          <div class="meeting-add-subtask-row">
            <input type="text" class="input-aewttr" data-add-subtask="${key}" placeholder="Add a subtask and press Enter">
          </div>
        </div>
        <div class="meeting-task-actions">
          <button type="button" class="btn-aewttr-outline btn-aewttr-sm" data-open-task-editor="${pid}:${task.id}"><i class="bx bx-edit"></i> Full task editor</button>
          <button type="button" class="btn-danger-outline btn-aewttr-sm" data-delete-task="${key}"><i class="bx bx-trash"></i> Delete task</button>
        </div>
      </div>
    </div>`;
}

/* ---------- add task modal ---------- */
function openMeetingAddTaskModal(scope, defaultParticipant, onDone, forcedProjectId) {
  const canFacilitate = canManageMeetings();
  const me = currentMeetingParticipant(scope);
  const projectOptions = scope.type === "project" ? [scope.project] : (window.AEWTTR.db.projects || []);
  const today = new Date().toISOString().slice(0, 10);
  const defaultOwner = defaultParticipant ? defaultParticipant.name : (me ? me.name : "");
  const pendingOwner = defaultOwner ? [{ name: defaultOwner, email: "" }] : [];
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>Add Task</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body">
      <div class="form-row meeting-modal-inline-row">
        <label>Task</label>
        <div class="meeting-modal-inline-fields">
          <input class="input-aewttr" id="mt-title" placeholder="Task title">
          ${canFacilitate ? `
            <div class="traveler-picker traveler-picker--inline" data-owner-wrap="mt-owner">
              <div id="mt-owner-sel" class="traveler-chip-list"></div>
              <input class="input-aewttr" id="mt-owner-input" placeholder="Owner…">
              <div id="mt-owner-sugg" class="traveler-suggestions"></div>
            </div>` : ""}
        </div>
      </div>
      <div class="form-grid-2">
        <div class="form-row"><label>Project</label>
          <select class="select-aewttr" id="mt-project" ${scope.type === "project" ? "disabled" : ""}>
            ${projectOptions.map(p => `<option value="${p.id}">${escapeHtml(p.name || "Untitled project")}</option>`).join("")}
          </select>
        </div>
        <div class="form-row"><label>Due date</label><input type="date" class="input-aewttr" id="mt-due" value="${today}"></div>
      </div>
      <div class="form-row"><label>Health</label><select class="select-aewttr" id="mt-health">${["On Track", "At Risk", "Off Track"].map(h => `<option>${h}</option>`).join("")}</select></div>
    </div>
    <div class="aewttr-modal-foot">
      <button class="btn-aewttr-ghost" id="mt-cancel">Cancel</button>
      <button class="btn-aewttr" id="mt-save" ${projectOptions.length ? "" : "disabled"}>Add Task</button>
    </div>
  `);
  const projectSelect = $("#mt-project", modal);
  if (projectSelect && forcedProjectId && projectOptions.some((project) => project.id === forcedProjectId)) {
    projectSelect.value = forcedProjectId;
  }
  if (!projectOptions.length) {
    const projectRow = projectSelect && projectSelect.closest(".form-row");
    if (projectRow) projectRow.insertAdjacentHTML("beforeend", `<p class="form-error" role="alert">Create a project before adding meeting tasks.</p>`);
  }
  if (canFacilitate) wirePeoplePicker(modal, pendingOwner, { mount: "mt-owner-sel", input: "mt-owner-input", suggestions: "mt-owner-sugg" }, { singleSelect: true, allowManualEmail: false });
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#mt-cancel", modal).addEventListener("click", closeModal);
  async function submitMeetingTask(e) {
    if (e) e.preventDefault();
    const saveBtn = $("#mt-save", modal);
    if (saveBtn.disabled) return;
    const title = $("#mt-title", modal).value.trim();
    if (!title) { toast("Title is required", "error"); return; }
    const pid = forcedProjectId || (scope.type === "project" ? scope.project.id : $("#mt-project", modal).value);
    const liveDb = window.AEWTTR.db;
    const project = (liveDb.projects || []).find((candidate) => candidate.id === pid);
    if (!pid || !project) {
      toast("Choose a valid project.", "error");
      return;
    }
    saveBtn.disabled = true;
    saveBtn.innerHTML = `<i class="bx bx-loader-alt bx-spin" aria-hidden="true"></i> Adding…`;
    const due = $("#mt-due", modal).value || today;
    const owner = (canFacilitate ? (pendingOwner[0] && pendingOwner[0].name) : (me ? me.name : defaultOwner)) || "Unassigned";
    if (!liveDb.ganttTasks[pid]) liveDb.ganttTasks[pid] = [];
    const newTask = {
      id: uid("g"),
      title,
      assignee: owner,
      start: today,
      end: due,
      status: "Not Started",
      health: $("#mt-health", modal).value,
      notes: [],
      subtasks: []
    };
    liveDb.ganttTasks[pid].push(newTask);
    try {
      await Repo.save("actionItem", newTask, { projectCode: pid, source: "Tracker", immediate: true });
      logMeetingActivity(scope, `${currentUserName()} added a new task to ${project.name || "the project"}: ${title}.`, {
        type: "create",
        projectId: pid,
        taskId: newTask.id,
        taskTitle: title
      });
      try {
        await saveMeetingSession(scope);
      } catch (sessionError) {
        console.warn("meeting task activity save", sessionError);
        toast("Task added; the meeting activity log could not be updated.", "info");
      }
      if (typeof ensureAssigneesFromTask === "function") {
        try { await ensureAssigneesFromTask(pid, newTask); } catch (peopleError) { console.warn("project people sync", peopleError); }
      }
      if (typeof notifyTaskAssignee === "function") {
        try { await notifyTaskAssignee(newTask, pid); } catch (notifyError) { console.warn("task notification", notifyError); }
      }
      closeModal();
      toast("Task added", "success");
      if (onDone) onDone(newTask, project);
    } catch (err) {
      const taskIndex = liveDb.ganttTasks[pid].findIndex((task) => task.id === newTask.id);
      if (taskIndex >= 0) liveDb.ganttTasks[pid].splice(taskIndex, 1);
      saveBtn.disabled = false;
      saveBtn.innerHTML = "Add Task";
      toast((err && err.friendly) || "Could not add the task — try again.", "error");
    }
  }
  $("#mt-save", modal).addEventListener("click", submitMeetingTask);
  $("#mt-title", modal).addEventListener("keydown", (event) => {
    if (event.key === "Enter" && !event.shiftKey) submitMeetingTask(event);
  });
}

/* ---------- end meeting / archive ----------
   When a meeting ends we snapshot every note stream that was edited during
   the session (general meeting notes, project notes, task notes, and
   subitem/subtask notes) plus the tasks that were reviewed — so History can
   show a real minutes/notes/gantt wrap-up instead of only the free-form feed. */

function meetingNoteTimestampMs(note) {
  if (!note) return null;
  if (note.editedAt) {
    const edited = Date.parse(note.editedAt);
    if (!Number.isNaN(edited)) return edited;
  }
  if (note.date) {
    const time = (note.time && String(note.time).trim()) ? String(note.time).trim() : "00:00";
    const stamped = Date.parse(`${note.date}T${time.length === 5 ? `${time}:00` : time}`);
    if (!Number.isNaN(stamped)) return stamped;
  }
  return null;
}

function noteOccurredDuringMeeting(note, session) {
  const ts = meetingNoteTimestampMs(note);
  const startedAt = Number(session && session.startedAt) || null;
  const endedAt = Number(session && session.endedAt) || Date.now();
  if (ts != null && startedAt) {
    // Slight slack so notes posted right as the meeting opens still qualify.
    return ts >= (startedAt - 90 * 1000) && ts <= (endedAt + 90 * 1000);
  }
  return !!(session && note && note.date && note.date === session.date);
}

function pushCapturedMeetingNote(bucket, entry) {
  if (!entry || !String(entry.text || "").trim()) return;
  const dedupeKey = entry.id || `${entry.scopeKind}:${entry.scopeLabel}:${entry.date}:${entry.time}:${entry.text}`;
  if (bucket._seen.has(dedupeKey)) return;
  bucket._seen.add(dedupeKey);
  bucket.notes.push({
    id: entry.id || uid("mnt"),
    text: String(entry.text).trim(),
    author: entry.author || "Unknown",
    date: entry.date || "",
    time: entry.time || "",
    scopeKind: entry.scopeKind || "meeting",
    scopeLabel: entry.scopeLabel || "Meeting",
    projectId: entry.projectId || "",
    taskId: entry.taskId || "",
    taskTitle: entry.taskTitle || "",
    subitemText: entry.subitemText || "",
    parentTaskTitle: entry.parentTaskTitle || entry.taskTitle || ""
  });
}

function meetingNoteScopeLabel(entry) {
  const kind = entry.scopeKind || "meeting";
  const projectId = entry.projectId || "";
  const taskTitle = (entry.taskTitle || entry.parentTaskTitle || "").trim();
  const subitem = (entry.subitemText || "").trim();
  if (kind === "subitem") {
    const parent = taskTitle || "Task";
    const child = subitem || "Untitled subitem";
    return projectId ? `${projectId} · ${parent} · Subitem · ${child}` : `${parent} · Subitem · ${child}`;
  }
  if (kind === "task") {
    const title = taskTitle || "Untitled task";
    return projectId ? `${projectId} · ${title}` : title;
  }
  if (kind === "project") {
    return projectId ? `Project · ${projectId}` : "Project Notes";
  }
  return entry.scopeLabel || "Meeting";
}

function collectMeetingCapturedNotes(scope, session) {
  const db = window.AEWTTR.db;
  const bucket = { notes: [], _seen: new Set() };
  const pids = scope.type === "project" ? [scope.project.id] : (db.projects || []).map((p) => p.id);

  (session.notesFeed || []).forEach((note) => {
    pushCapturedMeetingNote(bucket, {
      ...note,
      scopeKind: "meeting",
      scopeLabel: "Meeting"
    });
  });

  pids.forEach((pid) => {
    const extra = (typeof ensureProjectExtra === "function" ? ensureProjectExtra(pid) : ((db.projectExtra && db.projectExtra[pid]) || {})) || {};
    (extra.notes || []).forEach((note) => {
      if (!noteOccurredDuringMeeting(note, session)) return;
      pushCapturedMeetingNote(bucket, {
        ...note,
        scopeKind: "project",
        scopeLabel: `Project · ${pid}`,
        projectId: pid
      });
    });
    (db.ganttTasks[pid] || []).forEach((task) => {
      const taskTitle = task.title || "Untitled";
      (task.notes || []).forEach((note) => {
        if (!noteOccurredDuringMeeting(note, session)) return;
        pushCapturedMeetingNote(bucket, {
          ...note,
          scopeKind: "task",
          scopeLabel: scope.type === "global" ? `${pid} · ${taskTitle}` : taskTitle,
          projectId: pid,
          taskId: task.id,
          taskTitle,
          parentTaskTitle: taskTitle
        });
      });
      (task.subtasks || []).forEach((sub) => {
        const subText = sub.text || "Untitled subitem";
        (sub.notes || []).forEach((note) => {
          if (!noteOccurredDuringMeeting(note, session)) return;
          pushCapturedMeetingNote(bucket, {
            ...note,
            scopeKind: "subitem",
            scopeLabel: scope.type === "global"
              ? `${pid} · ${taskTitle} · Subitem · ${subText}`
              : `${taskTitle} · Subitem · ${subText}`,
            projectId: pid,
            taskId: task.id,
            taskTitle,
            parentTaskTitle: taskTitle,
            subitemText: subText
          });
        });
      });
    });
  });

  bucket.notes.sort((a, b) => {
    const aKey = `${a.date || ""}T${a.time || "00:00"}`;
    const bKey = `${b.date || ""}T${b.time || "00:00"}`;
    return bKey.localeCompare(aKey);
  });
  return bucket.notes.map((n) => ({ ...n, scopeLabel: meetingNoteScopeLabel(n) || n.scopeLabel }));
}

function collectMeetingGanttChanges(scope) {
  const db = window.AEWTTR.db;
  const pids = scope.type === "project" ? [scope.project.id] : (db.projects || []).map((p) => p.id);
  const changes = [];
  let tasksReviewed = 0;
  let tasksUpdated = 0;
  const dirtyTasks = [];

  pids.forEach((pid) => {
    const project = weeklyProjectById(pid);
    (db.ganttTasks[pid] || []).forEach((task) => {
      const review = task.reviewStatus || "Not Reviewed";
      if (review === "Not Reviewed") return;
      tasksReviewed += 1;
      if (review === "Updated") tasksUpdated += 1;
      const subs = task.subtasks || [];
      const doneCount = subs.filter((s) => s.done).length;
      changes.push({
        projectId: pid,
        projectName: (project && project.name) || "",
        taskId: task.id,
        title: task.title || "Untitled",
        assignee: task.assignee || "Unassigned",
        start: task.start || "",
        end: task.end || "",
        status: task.status || "Not Started",
        health: task.health || "On Track",
        reviewStatus: review,
        progressPct: typeof taskProgressPct === "function" ? taskProgressPct(task) : (subs.length ? Math.round((doneCount / subs.length) * 100) : 0),
        subtaskCount: subs.length,
        doneCount
      });
      task.reviewStatus = "Not Reviewed";
      dirtyTasks.push({ task, projectId: pid });
    });
  });

  return { changes, tasksReviewed, tasksUpdated, dirtyTasks };
}

function meetingActivitySortMs(entry) {
  if (!entry) return 0;
  if (entry.at && Number(entry.at)) return Number(entry.at);
  if (entry.time) {
    const parsed = Date.parse(String(entry.time).replace(" ", "T"));
    if (!Number.isNaN(parsed)) return parsed;
  }
  if (entry.date) {
    const time = (entry.time && String(entry.time).trim()) ? String(entry.time).trim() : "00:00";
    const stamped = Date.parse(`${entry.date}T${time.length === 5 ? `${time}:00` : time}`);
    if (!Number.isNaN(stamped)) return stamped;
  }
  return 0;
}

function minutesLogEntry(partial) {
  return {
    id: partial.id || (typeof uid === "function" ? uid("ml") : `ml-${Date.now()}`),
    at: partial.at || meetingActivitySortMs(partial) || Date.now(),
    time: partial.time || "",
    text: String(partial.text || "").trim(),
    type: partial.type || "info",
    actor: partial.actor || ""
  };
}

function isSignificantMeetingActivity(entry) {
  if (!entry || !String(entry.text || "").trim()) return false;
  const type = entry.type || "";
  if (type && type !== "info") return true;
  const text = String(entry.text || "").toLowerCase();
  return /started the meeting|ended the meeting|added (a )?note|updated the status|marked |reviewed |added task|added a new task|added a subitem|removed task|added guest|added agenda|removed agenda|checked off agenda|unchecked agenda/.test(text);
}

function buildMeetingMinutesLog(scope, session, capturedNotes, ganttChanges) {
  const rows = [];
  const seen = new Set();
  function pushRow(partial) {
    const entry = minutesLogEntry(partial);
    if (!entry.text) return;
    const key = `${entry.type}:${entry.text}`;
    if (seen.has(key)) return;
    seen.add(key);
    rows.push(entry);
  }

  (session.activity || []).forEach((a) => {
    if (!isSignificantMeetingActivity(a)) return;
    pushRow({
      id: a.id,
      at: a.at || meetingActivitySortMs(a),
      time: a.time || "",
      text: a.text,
      type: a.type || "info",
      actor: a.actor || ""
    });
  });

  // Backfill human-readable lines from captured notes if live logging missed them
  // (older archives / notes posted outside tracked entry points).
  const noteTexts = new Set(
    rows.filter((r) => r.type === "note").map((r) => r.text.toLowerCase())
  );
  (capturedNotes || []).forEach((note) => {
    const author = note.author || "Someone";
    const snippet = String(note.text || "").trim().replace(/\s+/g, " ");
    const short = snippet.length > 120 ? `${snippet.slice(0, 117)}…` : snippet;
    let text;
    if (note.scopeKind === "subitem") {
      const parent = note.parentTaskTitle || note.taskTitle || "a task";
      const sub = note.subitemText || "Untitled subitem";
      text = `${author} added a note to ${parent} · Subitem · ${sub}${short ? ` saying “${short}”` : "."}`;
    } else if (note.scopeKind === "task") {
      text = `${author} added a note to ${note.taskTitle || "a task"}${short ? ` saying “${short}”` : "."}`;
    } else if (note.scopeKind === "project") {
      text = `${author} added a note to project ${note.projectId || ""}${short ? ` saying “${short}”` : "."}`;
    } else {
      text = `${author} added a meeting note${short ? ` saying “${short}”` : "."}`;
    }
    if (noteTexts.has(text.toLowerCase())) return;
    // Skip if activity already has a close match for this note author+snippet
    const already = rows.some((r) => r.type === "note" && r.text.includes(short.slice(0, 40)));
    if (already && short.length > 20) return;
    pushRow({
      at: meetingNoteTimestampMs(note) || meetingActivitySortMs(note) || (session.endedAt || Date.now()),
      time: note.time ? `${note.date || ""} ${note.time}`.trim() : (note.date || ""),
      text,
      type: "note",
      actor: author
    });
  });

  // Backfill status/review lines from gantt snapshot when activity lacked them
  const statusCovered = new Set(
    rows.filter((r) => r.type === "status" || r.type === "review").map((r) => r.text.toLowerCase())
  );
  (ganttChanges || []).forEach((change, idx) => {
    const title = change.title || "a task";
    const where = change.projectId ? ` in ${change.projectId}` : "";
    let text;
    if (change.reviewStatus === "Updated") {
      text = `Team updated the status of ${title}${where}${change.health ? ` (${change.health})` : ""}.`;
    } else {
      text = `Team reviewed ${title}${where} — no change.`;
    }
    if (statusCovered.has(text.toLowerCase())) return;
    const titleHit = rows.some((r) => (r.type === "status" || r.type === "review") && r.text.toLowerCase().includes(title.toLowerCase()));
    if (titleHit) return;
    pushRow({
      at: (session.endedAt || Date.now()) - ((ganttChanges.length - idx) * 1000),
      time: "",
      text,
      type: change.reviewStatus === "Updated" ? "status" : "review",
      actor: "Team"
    });
  });

  rows.sort((a, b) => (a.at || 0) - (b.at || 0));
  return rows;
}

function meetingHistoryMinutesLog(session) {
  if (session.minutesLog && session.minutesLog.length) return session.minutesLog;
  return buildMeetingMinutesLog(
    { type: "global" },
    session,
    meetingHistoryNotesForSession(session),
    session.ganttChanges || []
  );
}

function meetingMinutesTitle(scope, session) {
  const dateLabel = session && session.date ? fmtDate(session.date) : "Meeting";
  if (scope.type === "global") return `Weekly Pulse — ${dateLabel}`;
  const projectId = (scope.project && scope.project.id) || "Project";
  return `${projectId} Meeting — ${dateLabel}`;
}

function meetingMinutesDeepLink(scope) {
  if (typeof pulseAppRouteUrl !== "function") return typeof pulseAppUrl === "function" ? pulseAppUrl() : "";
  if (scope.type === "project" && scope.project && scope.project.id) {
    return pulseAppRouteUrl(`projects/${scope.project.id}/meeting`);
  }
  return pulseAppRouteUrl("weekly");
}

function collectMeetingMinutesRecipientEmails(scope, session) {
  const emails = new Set();
  meetingParticipants(scope).forEach((participant) => {
    const email = participantEmail(participant);
    if (email) emails.add(String(email).trim());
  });
  (session && session.guests ? session.guests : []).forEach((guest) => {
    const email = String((guest && guest.email) || "").trim();
    if (email) emails.add(email);
  });
  return Array.from(emails);
}

function buildMeetingMinutesEmailBody(title, log) {
  const lines = (log || []).map((entry) => String(entry.text || "").trim()).filter(Boolean);
  const list = lines.length
    ? `<ol style="padding-left:20px;margin:12px 0;">${lines.map((text) => `<li style="margin:6px 0;">${escapeHtml(text)}</li>`).join("")}</ol>`
    : `<p>No activity was logged during this meeting.</p>`;
  return `<p><strong>${escapeHtml(title)}</strong></p>${list}`;
}

function buildMeetingMinutesTeamsText(subject, preview, log, actionUrl) {
  const lines = (log || []).map((entry) => String(entry.text || "").trim()).filter(Boolean);
  const blocks = [
    `**${subject}**`,
    preview,
    lines.length ? lines.map((text, idx) => `${idx + 1}. ${text}`).join("\n") : "No activity was logged during this meeting."
  ];
  if (actionUrl) blocks.push(`[Open Meeting History](${actionUrl})`);
  return blocks.filter(Boolean).join("\n\n");
}

async function notifyMeetingMinutes(scope, session, recipientEmails) {
  if (typeof isSharePointMode !== "function" || !isSharePointMode()) return;
  if (typeof notifyUsers !== "function" || !session) return;
  const emails = Array.isArray(recipientEmails) && recipientEmails.length
    ? recipientEmails
    : collectMeetingMinutesRecipientEmails(scope, session);
  if (!emails.length) return;
  const title = meetingMinutesTitle(scope, session);
  const subject = `PULSE Meeting Minutes — ${title}`;
  const presentCount = (session.attendees || []).length + (session.guestAttendees || []).length;
  const noteCount = (session.capturedNotes || session.notesFeed || []).length;
  const log = session.minutesLog || [];
  const lines = log.map((entry) => String(entry.text || "").trim()).filter(Boolean);
  const preview = presentCount
    ? `Minutes are ready · ${presentCount} present · ${noteCount} note${noteCount === 1 ? "" : "s"}.`
    : `Minutes are ready · ${noteCount} note${noteCount === 1 ? "" : "s"}.`;
  const actionUrl = meetingMinutesDeepLink(scope);
  const minutesValue = lines.length
    ? lines.slice(0, 25).map((text, idx) => `${idx + 1}. ${text}`).join("\n") + (lines.length > 25 ? "\n…" : "")
    : "No activity logged";
  const facts = [
    { title: "Meeting", value: title },
    { title: "Present", value: String(presentCount) },
    { title: "Notes", value: String(noteCount) },
    { title: "Minutes", value: minutesValue }
  ];
  try {
    await notifyUsers({
      to: emails,
      subject,
      area: "Weekly",
      kind: "info",
      preview,
      facts,
      actionUrl,
      actionTitle: "Open Meeting History",
      body: buildMeetingMinutesEmailBody(title, log),
      teamsText: buildMeetingMinutesTeamsText(subject, preview, log, actionUrl)
    });
  } catch (e) {
    console.warn("PULSE: meeting minutes notification failed.", e);
  }
}

function openEndMeetingModal(scope, onDone) {
  if (!canManageMeetings()) {
    toast("Only Meeting Admins can end meetings.", "error");
    return;
  }
  const session = activeMeetingSession(scope);
  if (!session) return;
  const data = meetingData(scope);
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>End Meeting</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body">
      <p style="margin-top:0;">This archives the current session — meeting notes, project/task/subitem notes taken during the meeting, attendance, and tracker updates — and keeps the meeting closed until a Meeting Admin starts a new one.</p>
    </div>
    <div class="aewttr-modal-foot">
      <button class="btn-aewttr-ghost" id="em-cancel">Cancel</button>
      <button class="btn-aewttr" id="em-save">End &amp; Archive</button>
    </div>
  `);
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#em-cancel", modal).addEventListener("click", closeModal);
  $("#em-save", modal).addEventListener("click", () => {
    const saveBtn = $("#em-save", modal);
    if (saveBtn.disabled) return;
    saveBtn.disabled = true;
    saveBtn.textContent = "Archiving…";

    const participants = meetingParticipants(scope);
    const endedAt = Date.now();
    session.endedAt = endedAt;
    const capturedNotes = collectMeetingCapturedNotes(scope, session);
    const { changes, tasksReviewed, tasksUpdated, dirtyTasks } = collectMeetingGanttChanges(scope);

    logMeetingActivity(scope, `${currentUserName()} ended the meeting.`, { type: "lifecycle" });
    session.activity = session.activity || [];

    const attendees = participants.filter((p) => session.attendance[p.id] === "Here").map((p) => p.name);
    const guestAttendees = (session.guests || []).filter((g) => session.attendance[g.id] === "Here").map((g) => g.name);
    const minutesLog = buildMeetingMinutesLog(scope, session, capturedNotes, changes);
    const minutesRecipients = collectMeetingMinutesRecipientEmails(scope, session);

    Object.assign(session, {
      sessionStatus: "ended",
      startedAt: session.startedAt || null,
      endedAt,
      tasksUpdated,
      tasksReviewed,
      notesFeed: session.notesFeed || [],
      capturedNotes,
      ganttChanges: changes,
      minutesLog,
      guests: (session.guests || []).slice(),
      attendees,
      guestAttendees,
      activity: session.activity.slice()
    });

    data.sessions = data.sessions || [];
    data.sessions.unshift({ ...session });
    data.currentSession = null;
    data.meetingStatus = "ended";

    const taskSaves = dirtyTasks.map(({ task, projectId }) =>
      Repo.save("actionItem", task, { projectCode: projectId, source: "Meeting" })
    );

    Promise.all([
      ...taskSaves,
      Repo.save("meetingSession", session, { projectCode: meetingProjectCode(scope) })
    ]).then(() => (Repo.flush ? Repo.flush() : Promise.resolve())).then(async () => {
      await notifyMeetingMinutes(scope, session, minutesRecipients);
      if (!window.AEWTTR.state.meetingView) window.AEWTTR.state.meetingView = {};
      window.AEWTTR.state.meetingView[meetingScopeKey(scope)] = "history";
      closeModal();
      toast("Meeting ended and archived", "success");
      if (onDone) onDone();
    }).catch((err) => {
      saveBtn.disabled = false;
      saveBtn.textContent = "End & Archive";
      toast((err && err.friendly) || "Could not archive the meeting — try again.", "error");
    });
  });
}

/* ---------- history ---------- */
function meetingHistoryAttendeeHtml(names) {
  if (!names || !names.length) return `<span class="meeting-history-empty-chip">No one marked present</span>`;
  return names.map((name) => `<span class="meeting-history-person">${userAvatarHtml({ name, email: memberEmailForPerson(name), size: 22 })}<span>${escapeHtml(name)}</span></span>`).join("");
}

function meetingHistoryNotesForSession(session) {
  // New format: docBlocks
  if (session.docBlocks && session.docBlocks.length) {
    const docNotes = session.docBlocks
      .filter(function(b) { return b && b.content && String(b.content).trim(); })
      .map(function(b) {
        const prefix = b.type === "action" ? (b.done ? "[x] " : "[ ] ") : b.type === "bullet" ? "• " : b.type === "h2" ? "## " : "";
        return {
          id: b.id,
          author: "Meeting doc",
          date: session.date || "",
          time: "",
          text: prefix + String(b.content).trim(),
          scopeKind: "meeting",
          scopeLabel: "Meeting",
          parentTaskTitle: "",
          subitemText: ""
        };
      });
    if (docNotes.length) return docNotes;
  }
  const notes = (session.capturedNotes && session.capturedNotes.length)
    ? session.capturedNotes
    : (session.notesFeed || []).map((note) => ({
      ...note,
      scopeKind: "meeting",
      scopeLabel: "Meeting"
    }));
  return notes.map((n) => ({
    ...n,
    scopeLabel: meetingNoteScopeLabel(n) || n.scopeLabel || "Meeting",
    parentTaskTitle: n.parentTaskTitle || n.taskTitle || "",
    subitemText: n.subitemText || ""
  }));
}

function meetingHistoryCapturedNoteCardHtml(note) {
  const isMine = typeof isNoteAuthor === "function" ? isNoteAuthor(note) : false;
  const stamp = typeof formatNoteTimestamp === "function"
    ? formatNoteTimestamp(note)
    : `${note.date ? fmtDate(note.date) : ""}${note.time ? ` · ${note.time}` : ""}`.trim();
  const assigned = meetingNoteScopeLabel(note) || note.scopeLabel || "Meeting";
  return `
    <article class="proj-notes-card meeting-history-note-card ${isMine ? "mine" : ""}">
      <header class="proj-notes-card-head">
        <div class="proj-notes-card-assigned"><i class="bx bx-link-alt"></i> ${escapeHtml(assigned)}</div>
        <div class="proj-notes-card-meta">
          <strong>${escapeHtml(note.author || "Unknown")}</strong>
          <span>${escapeHtml(stamp)}</span>
        </div>
      </header>
      <div class="proj-notes-card-text">${escapeHtml(note.text || "")}</div>
    </article>`;
}

function meetingHistoryMinutesTimeLabel(entry) {
  if (entry.time && /\d{1,2}:\d{2}/.test(String(entry.time))) {
    const m = String(entry.time).match(/(\d{1,2}:\d{2})/);
    return m ? m[1] : String(entry.time);
  }
  if (entry.at) {
    const d = new Date(entry.at);
    if (!Number.isNaN(d.getTime())) {
      const pad = (n) => String(n).padStart(2, "0");
      return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
    }
  }
  return "";
}

function renderAgendaNotesHtml(notes) {
  if (!notes || !notes.trim()) return "";
  const lines = notes.split("\n");
  let html = "";
  let inList = false;
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) { if (inList) { html += "</ul>"; inList = false; } continue; }
    if (/^[-•]\s+/.test(trimmed)) {
      if (!inList) { html += '<ul class="agenda-history-bullets">'; inList = true; }
      html += `<li>${escapeHtml(trimmed.replace(/^[-•]\s+/, ""))}</li>`;
    } else {
      if (inList) { html += "</ul>"; inList = false; }
      html += `<p class="agenda-history-text">${escapeHtml(trimmed)}</p>`;
    }
  }
  if (inList) html += "</ul>";
  return html;
}

function meetingHistoryMinutesHtml(session, scope) {
  const notes = meetingHistoryNotesForSession(session);
  const changes = session.ganttChanges || [];
  const log = meetingHistoryMinutesLog(session);
  const attendees = session.attendees || [];
  const guests = session.guestAttendees || [];
  const updated = changes.filter((c) => c.reviewStatus === "Updated");

  return `
    <div class="meeting-minutes meeting-minutes--log">
      <header class="meeting-minutes-head">
        <div>
          <div class="meeting-minutes-kicker">${scope.type === "global" ? "Weekly Pulse" : escapeHtml(scope.project.id)}</div>
          <h4>${escapeHtml(fmtDate(session.date))}</h4>
          <p>${attendees.length} present${guests.length ? ` · ${guests.length} guest${guests.length === 1 ? "" : "s"}` : ""} · ${session.tasksUpdated || updated.length || 0} updated · ${notes.length} note${notes.length === 1 ? "" : "s"}</p>
        </div>
      </header>

      <section class="meeting-minutes-attendance">
        <div class="meeting-history-people">${meetingHistoryAttendeeHtml(attendees)}</div>
        ${guests.length ? `<div class="meeting-minutes-subhead">Guests</div><div class="meeting-history-people">${meetingHistoryAttendeeHtml(guests)}</div>` : ""}
      </section>

      ${(session.agenda || []).length ? `<section class="meeting-minutes-attendance"><div class="meeting-minutes-subhead">Agenda</div>${session.agenda.map((item) => `<div class="meeting-agenda-history-item"><i class="bx ${item.done ? "bxs-check-circle" : "bx-circle"}"></i><div class="meeting-agenda-history-body"><strong>${escapeHtml(item.title || "Untitled")}</strong>${item.notes ? renderAgendaNotesHtml(item.notes) : ""}</div></div>`).join("")}</section>` : ""}

      <section class="meeting-minutes-feed" aria-label="Meeting activity">
        ${log.length ? log.map((entry) => {
          const timeLabel = meetingHistoryMinutesTimeLabel(entry);
          return `
            <div class="meeting-minutes-line meeting-minutes-line--${escapeHtml(entry.type || "info")}">
              ${timeLabel ? `<time class="meeting-minutes-line-time">${escapeHtml(timeLabel)}</time>` : `<span class="meeting-minutes-line-time" aria-hidden="true"></span>`}
              <p class="meeting-minutes-line-text">${escapeHtml(entry.text)}</p>
            </div>`;
        }).join("") : `<div class="empty-state">No activity was logged during this meeting.</div>`}
      </section>
    </div>`;
}

function meetingHistoryDetailHtml(session, scope, sessionKey, activeTab) {
  const notes = meetingHistoryNotesForSession(session);
  const changes = session.ganttChanges || [];
  const tab = activeTab === "notes" || activeTab === "updates" ? activeTab : "minutes";
  const updatedCount = changes.filter((c) => c.reviewStatus === "Updated").length;
  return `
    <div class="meeting-history-tabs" role="tablist">
      <button type="button" class="meeting-history-tab ${tab === "minutes" ? "is-active" : ""}" data-history-tab="minutes" data-history-for="${escapeHtml(sessionKey)}">Meeting Minutes</button>
      <button type="button" class="meeting-history-tab ${tab === "notes" ? "is-active" : ""}" data-history-tab="notes" data-history-for="${escapeHtml(sessionKey)}">Notes <em>${notes.length}</em></button>
      <button type="button" class="meeting-history-tab ${tab === "updates" ? "is-active" : ""}" data-history-tab="updates" data-history-for="${escapeHtml(sessionKey)}">Project Updates <em>${changes.length}</em></button>
    </div>
    <div class="meeting-history-tab-panel" data-history-panel="${escapeHtml(sessionKey)}">
      ${tab === "minutes" ? meetingHistoryMinutesHtml(session, scope) : ""}
      ${tab === "notes" ? (notes.length
        ? `<div class="meeting-history-notes-hub proj-notes-hub"><div class="meeting-history-notes-list proj-notes-main-body">${notes.map(meetingHistoryCapturedNoteCardHtml).join("")}</div></div>`
        : `<div class="empty-state">No notes were captured during this meeting.</div>`) : ""}
      ${tab === "updates" ? `<div class="meeting-history-gantt">${meetingHistoryProjectUpdatesHtml(changes)}</div>` : ""}
    </div>`;
}

function meetingHistoryProjectUpdatesHtml(changes) {
  if (!changes || !changes.length) {
    return `<div class="empty-state">No tracker updates were marked during this meeting.</div>`;
  }
  const dated = changes.filter((c) => c.start && c.end);
  let rangeStart = null;
  let rangeEnd = null;
  if (dated.length && typeof ganttParseDate === "function") {
    rangeStart = new Date(Math.min(...dated.map((c) => ganttParseDate(c.start).getTime())));
    rangeEnd = new Date(Math.max(...dated.map((c) => ganttParseDate(c.end).getTime())));
  }
  const totalDays = (rangeStart && rangeEnd)
    ? Math.max((typeof ganttDaysBetween === "function" ? ganttDaysBetween(
        (typeof ganttIsoDate === "function" ? ganttIsoDate(rangeStart) : rangeStart.toISOString().slice(0, 10)),
        (typeof ganttIsoDate === "function" ? ganttIsoDate(rangeEnd) : rangeEnd.toISOString().slice(0, 10))
      ) : Math.round((rangeEnd - rangeStart) / 86400000)) + 1, 1)
    : 1;

  const byProject = {};
  changes.forEach((change) => {
    const pid = change.projectId || "—";
    if (!byProject[pid]) byProject[pid] = { projectId: pid, projectName: change.projectName || "", rows: [] };
    byProject[pid].rows.push(change);
  });

  return Object.keys(byProject).map((pid) => {
    const group = byProject[pid];
    return `
      <section class="meeting-history-gantt-group">
        <div class="meeting-history-gantt-group-head">
          <strong>${escapeHtml(group.projectId)}</strong>
          ${group.projectName ? `<span>${escapeHtml(group.projectName)}</span>` : ""}
        </div>
        <div class="meeting-history-gantt-rows">
          ${group.rows.map((row) => {
            let leftPct = 0;
            let widthPct = 100;
            if (rangeStart && row.start && row.end && typeof ganttParseDate === "function") {
              const startOff = Math.max(0, Math.round((ganttParseDate(row.start) - rangeStart) / 86400000));
              const span = Math.max(1, Math.round((ganttParseDate(row.end) - ganttParseDate(row.start)) / 86400000) + 1);
              leftPct = (startOff / totalDays) * 100;
              widthPct = Math.max((span / totalDays) * 100, 4);
            }
            const healthClass = `health-${String(row.health || "On Track").replace(/\s+/g, "-")}`;
            const reviewLabel = row.reviewStatus === "Updated" ? "Updated" : "Reviewed";
            return `
              <div class="meeting-history-gantt-row">
                <div class="meeting-history-gantt-label">
                  <span class="health-dot ${healthClass}" title="${escapeHtml(row.health || "")}"></span>
                  <div>
                    <strong>${escapeHtml(row.title)}</strong>
                    <span>${escapeHtml(row.assignee || "Unassigned")}${row.start && row.end ? ` · ${fmtDate(row.start)} – ${fmtDate(row.end)}` : ""}</span>
                  </div>
                  <em class="meeting-history-gantt-review ${row.reviewStatus === "Updated" ? "is-updated" : ""}">${escapeHtml(reviewLabel)}</em>
                </div>
                <div class="meeting-history-gantt-track" aria-hidden="true">
                  <div class="meeting-history-gantt-bar ${healthClass}" style="left:${leftPct.toFixed(2)}%;width:${widthPct.toFixed(2)}%;">
                    <span style="width:${Math.max(0, Math.min(100, Number(row.progressPct) || 0))}%;"></span>
                  </div>
                </div>
              </div>`;
          }).join("")}
        </div>
      </section>`;
  }).join("");
}

function renderMeetingHistory(body, scope) {
  if (!window.AEWTTR.state.meetingHistorySearch) window.AEWTTR.state.meetingHistorySearch = {};
  if (!window.AEWTTR.state.meetingHistoryExpanded) window.AEWTTR.state.meetingHistoryExpanded = {};
  if (!window.AEWTTR.state.meetingHistoryTab) window.AEWTTR.state.meetingHistoryTab = {};
  const key = meetingScopeKey(scope);
  const search = window.AEWTTR.state.meetingHistorySearch[key] || "";
  const expanded = window.AEWTTR.state.meetingHistoryExpanded;
  const tabs = window.AEWTTR.state.meetingHistoryTab;
  const data = meetingData(scope);
  const sessions = (data.sessions || []).slice().sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")));
  const q = search.toLowerCase();
  const filtered = sessions.filter((s) => {
    if (!q) return true;
    const captured = meetingHistoryNotesForSession(s);
    return s.date.includes(q)
      || captured.some((n) => `${n.text || ""} ${n.scopeLabel || ""} ${n.author || ""}`.toLowerCase().includes(q))
      || (s.attendees || []).join(" ").toLowerCase().includes(q)
      || (s.guestAttendees || []).join(" ").toLowerCase().includes(q)
      || (s.activity || []).some((a) => (a.text || "").toLowerCase().includes(q))
      || (s.minutesLog || []).some((a) => (a.text || "").toLowerCase().includes(q))
      || (s.ganttChanges || []).some((c) => `${c.title || ""} ${c.projectId || ""}`.toLowerCase().includes(q));
  });

  body.innerHTML = `
    <div class="meeting-history-toolbar">
      <div class="search-box meeting-history-search"><i class="bx bx-search"></i><input id="mtg-history-search" placeholder="Search date, people, notes, or activity…" value="${escapeHtml(search)}"></div>
      <div class="meeting-history-summary">${filtered.length} session${filtered.length === 1 ? "" : "s"}</div>
    </div>
    <div class="meeting-history-timeline">
      ${filtered.length ? filtered.map((session, index) => {
        const sessionKey = `${key}:${session.id || session.date}:${index}`;
        const isOpen = expanded[sessionKey] !== false;
        const notes = meetingHistoryNotesForSession(session);
        const noteCount = notes.length;
        const tasksUpdated = session.tasksUpdated || (session.ganttChanges || []).filter((c) => c.reviewStatus === "Updated").length;
        const tasksReviewed = session.tasksReviewed || (session.ganttChanges || []).length;
        const previewNote = notes[0];
        const activeTab = tabs[sessionKey] || "minutes";
        return `
        <article class="meeting-history-session ${isOpen ? "is-open" : ""}" data-history-session="${escapeHtml(sessionKey)}">
          <button type="button" class="meeting-history-session-head" data-toggle-history="${escapeHtml(sessionKey)}">
            <div class="meeting-history-date-rail">
              <span class="meeting-history-day">${escapeHtml(fmtDate(session.date))}</span>
              <span class="meeting-history-scope">${scope.type === "global" ? "Weekly Pulse" : escapeHtml(scope.project.id)}</span>
            </div>
            <div class="meeting-history-head-main">
              <div class="meeting-history-stat-row">
                <span class="meeting-history-stat"><i class="bx bx-check-circle"></i> ${tasksReviewed} reviewed</span>
                <span class="meeting-history-stat"><i class="bx bx-edit"></i> ${tasksUpdated} updated</span>
                <span class="meeting-history-stat"><i class="bx bx-note"></i> ${noteCount} note${noteCount === 1 ? "" : "s"}</span>
              </div>
              <div class="meeting-history-preview">${previewNote ? escapeHtml((previewNote.text || "").slice(0, 120)) : "No notes recorded."}</div>
            </div>
            <i class="bx ${isOpen ? "bx-chevron-up" : "bx-chevron-down"} meeting-history-chevron"></i>
          </button>
          <div class="meeting-history-session-body" ${isOpen ? "" : `hidden`}>
            ${meetingHistoryDetailHtml(session, scope, sessionKey, activeTab)}
          </div>
        </article>`;
      }).join("") : `<div class="empty-state" style="padding:40px;">No past meetings match your search.</div>`}
    </div>
  `;
  let _histSearchTimer = null;
  $("#mtg-history-search", body).addEventListener("input", (e) => {
    const cursor = e.target.selectionStart;
    window.AEWTTR.state.meetingHistorySearch[key] = e.target.value;
    clearTimeout(_histSearchTimer);
    _histSearchTimer = setTimeout(() => {
      renderMeetingHistory(body, scope);
      const el = $("#mtg-history-search", body);
      if (el) { el.focus(); try { el.setSelectionRange(cursor, cursor); } catch (_) {} }
    }, 150);
  });
  $all("[data-toggle-history]", body).forEach((btn) => btn.addEventListener("click", () => {
    const sessionKey = btn.dataset.toggleHistory;
    const isOpen = expanded[sessionKey] !== false;
    expanded[sessionKey] = !isOpen;
    renderMeetingHistory(body, scope);
  }));
  $all("[data-history-tab]", body).forEach((btn) => btn.addEventListener("click", (e) => {
    e.stopPropagation();
    tabs[btn.dataset.historyFor] = btn.dataset.historyTab;
    renderMeetingHistory(body, scope);
  }));
}

/* ---------- roster admin (used by Admin > User Access too) ---------- */
function openWeeklyRosterModal(onDone) {
  const db = weeklyMeetingDb();
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>Manage Weekly Meeting Roster</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body">
      <p style="font-size:12px;color:var(--aewttr-muted);margin-top:0;" data-help>Admins choose who participates in the room rotation.</p>
      <div class="weekly-roster-modal">
        ${window.AEWTTR.db.members.filter((member) => !member.hiddenFromMeetings).map(member => `
          <label class="weekly-roster-option">
            <input type="checkbox" value="${member.id}" ${db.roster.includes(member.id) ? "checked" : ""}>
            <div>
              <strong>${escapeHtml(member.name)}</strong>
              <span>${escapeHtml(member.role)}</span>
            </div>
          </label>`).join("")}
        ${window.AEWTTR.db.members.some((member) => member.hiddenFromMeetings) ? `<p style="font-size:11px;color:var(--aewttr-muted);margin:10px 0 0;" data-help>People hidden from meetings on the Users page do not appear here. Unhide them there to add them back to the rotation.</p>` : ""}
      </div>
    </div>
    <div class="aewttr-modal-foot">
      <button class="btn-aewttr-ghost" id="wro-cancel">Cancel</button>
      <button class="btn-aewttr" id="wro-save">Save Roster</button>
    </div>
  `, { wide: true });
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#wro-cancel", modal).addEventListener("click", closeModal);
  $("#wro-save", modal).addEventListener("click", () => {
    db.roster = $all("input[type=checkbox]:checked", modal).map(node => node.value);
    aewttrSaveStore();
    closeModal();
    toast("Weekly roster updated", "success");
    if (onDone) onDone();
  });
}

/* ---------- meeting project tasks list view ---------- */

/* Full nested task list across every visible meeting project (sidebar option).
   Uses the same Main table design as the project tracker, stacked per project. */
function renderMeetingFullTasksList(mount, scope, projects, onRedraw) {
  const list = projects || [];
  if (!list.length) {
    mount.innerHTML = `<div class="ptl-empty" style="padding:32px;text-align:center;color:var(--aewttr-muted);font-size:13px;">No projects in this meeting.</div>`;
    return;
  }

  const scopeKey = typeof meetingScopeKey === "function" ? meetingScopeKey(scope) : (scope.type || "global");
  if (!window.AEWTTR.state.mtfView) window.AEWTTR.state.mtfView = {};
  const globalView = window.AEWTTR.state.mtfView[scopeKey] || "table";
  const db = window.AEWTTR.db;
  const onChange = onRedraw || (() => {});

  // Nested tracker tables, controls, and drag surfaces can consume wheel
  // events. Route vertical wheel input from anywhere in the Full Task List to
  // the owning panel so the task area scrolls consistently with a mouse or
  // trackpad.
  if (!mount.dataset.fullTasksWheelRouting) {
    mount.dataset.fullTasksWheelRouting = "true";
    mount.addEventListener("wheel", (event) => {
      if (!event.deltaY || event.ctrlKey) return;

      const scroller = mount.closest(".meeting-project-main-body--tasks");
      if (!scroller) return;

      let delta = event.deltaY;
      if (event.deltaMode === 1) delta *= 16;
      if (event.deltaMode === 2) delta *= scroller.clientHeight;

      const maximumScrollTop = Math.max(0, scroller.scrollHeight - scroller.clientHeight);
      const nextScrollTop = Math.max(0, Math.min(maximumScrollTop, scroller.scrollTop + delta));
      if (nextScrollTop === scroller.scrollTop) return;

      event.preventDefault();
      event.stopPropagation();
      scroller.scrollTop = nextScrollTop;
    }, { capture: true, passive: false });
  }

  // Gantt: one continuous MS Project–style chart across every meeting project.
  if (globalView === "gantt") {
    mount.innerHTML = `
      <div class="mtf-global-head">
        <div class="monday-tracker-tabs mtf-global-tabs" id="mtf-global-tabs">
          <button type="button" class="monday-tracker-tab" data-mtf-view="table">Main table</button>
          <button type="button" class="monday-tracker-tab active" data-mtf-view="gantt">Timeline</button>
        </div>
      </div>
      <div class="mtf-combined-gantt" id="mtf-combined-gantt"></div>`;

    $all("[data-mtf-view]", mount).forEach((btn) => {
      btn.addEventListener("click", () => {
        window.AEWTTR.state.mtfView[scopeKey] = btn.dataset.mtfView;
        renderMeetingFullTasksList(mount, scope, projects, onRedraw);
      });
    });

    const ganttMount = $("#mtf-combined-gantt", mount);
    if (!ganttMount || typeof renderMondayGanttChart !== "function") return;

    const combined = [];
    list.forEach((proj) => {
      const tasks = (db.ganttTasks && db.ganttTasks[proj.id]) || [];
      tasks.forEach((t) => {
        if (!t) return;
        t._projectCode = proj.id;
        t._projectName = proj.name || "Untitled project";
        combined.push(t);
      });
    });

    const stateKey = `mtf-gantt-${scopeKey}`;
    const ui = typeof ensureTrackerUiState === "function"
      ? ensureTrackerUiState(stateKey, "gantt")
      : { expanded: {}, filterState: { assignee: "" } };
    const expanded = ui.expanded || {};
    const dividerCollapsed = typeof trackerDividerCollapsedStore === "function"
      ? trackerDividerCollapsedStore(stateKey)
      : {};

    function resolveProjectCode(task) {
      if (!task) return "";
      if (task._projectCode) return task._projectCode;
      const found = list.find((p) => ((db.ganttTasks && db.ganttTasks[p.id]) || []).some((x) => x.id === task.id));
      return found ? found.id : (list[0] && list[0].id) || "";
    }

    function redrawGantt() {
      renderMeetingFullTasksList(mount, scope, projects, onRedraw);
    }

    renderMondayGanttChart(
      ganttMount,
      combined,
      expanded,
      (taskId) => {
        expanded[taskId] = !expanded[taskId];
        redrawGantt();
      },
      (taskId, subPath) => {
        const task = combined.find((t) => t.id === taskId);
        const pid = resolveProjectCode(task);
        if (typeof openMeetingTaskEditor === "function") {
          openMeetingTaskEditor(scope, pid, taskId, onChange, subPath);
        }
      },
      redrawGantt,
      combined,
      {
        mondayStyle: true,
        meetingMode: true,
        flatGantt: true,
        groupByProject: true,
        saveSource: "Meeting",
        projectCode: (list[0] && list[0].id) || "",
        resolveProjectCode,
        dividerCollapsed,
        collapseDividersByDefault: false,
        collapseGanttGroupsByDefault: false,
        onDeleteTask: (task) => {
          const pid = resolveProjectCode(task);
          const source = (db.ganttTasks && db.ganttTasks[pid]) || [];
          const sourceIdx = source.findIndex((item) => item.id === task.id);
          const combinedIdx = combined.findIndex((item) => item.id === task.id);
          if (sourceIdx >= 0) source.splice(sourceIdx, 1);
          if (combinedIdx >= 0) combined.splice(combinedIdx, 1);
          return () => {
            if (sourceIdx >= 0 && !source.some((item) => item.id === task.id)) source.splice(sourceIdx, 0, task);
            if (combinedIdx >= 0 && !combined.some((item) => item.id === task.id)) combined.splice(combinedIdx, 0, task);
          };
        },
        onToggleReview: (taskId) => {
          const task = combined.find((t) => t.id === taskId);
          if (!task) return;
          const prev = task.reviewStatus || "Not Reviewed";
          const next = prev === "Not Reviewed" ? "Reviewed - No Change" : "Not Reviewed";
          task.reviewStatus = next;
          const pid = resolveProjectCode(task);
          if (typeof recordActiveMeetingTaskChange === "function") {
            recordActiveMeetingTaskChange(pid, task, "reviewStatus", next, { prev });
          }
          Repo.save("actionItem", task, { projectCode: pid, source: "Meeting" });
          redrawGantt();
        }
      }
    );
    return;
  }

  // Per-project collapse state (default: all expanded)
  if (!window.AEWTTR.state.mtfCollapsed) window.AEWTTR.state.mtfCollapsed = {};
  if (!window.AEWTTR.state.mtfCollapsed[scopeKey]) window.AEWTTR.state.mtfCollapsed[scopeKey] = {};
  const projCollapsed = window.AEWTTR.state.mtfCollapsed[scopeKey];

  // Build a combined, project-tagged task list for global search/filter
  if (!window.AEWTTR.state.mtfFilters) window.AEWTTR.state.mtfFilters = {};
  const filters = window.AEWTTR.state.mtfFilters[scopeKey] || { search: "", assignee: "", status: "" };
  window.AEWTTR.state.mtfFilters[scopeKey] = filters;

  // Collect all members across meeting projects for the assignee dropdown
  const allAssignees = new Set();
  list.forEach((proj) => {
    const tasks = (db.ganttTasks && db.ganttTasks[proj.id]) || [];
    tasks.forEach((t) => { if (t && t.assignee) allAssignees.add(t.assignee); });
  });

  function buildCombined(searchStr, assigneeStr, statusStr) {
    const q = searchStr.toLowerCase();
    const sections = [];
    list.forEach((proj) => {
      const allItems = (db.ganttTasks && db.ganttTasks[proj.id]) || [];
      const dividers = allItems.filter((t) => typeof isTrackerDivider === "function" ? isTrackerDivider(t) : t.itemType === "divider");
      const plainTasks = allItems.filter((t) => !(typeof isTrackerDivider === "function" ? isTrackerDivider(t) : t.itemType === "divider"));
      const byDivider = {};
      plainTasks.forEach((t) => {
        const key = t.parentDividerId || "__none__";
        if (!byDivider[key]) byDivider[key] = [];
        byDivider[key].push(t);
      });
      function matchTask(t) {
        if (q && !((t.title || "").toLowerCase().includes(q) || (t.assignee || "").toLowerCase().includes(q))) return false;
        if (assigneeStr && t.assignee !== assigneeStr) return false;
        if (statusStr && t.status !== statusStr) return false;
        return true;
      }
      dividers.forEach((divider) => {
        const divTasks = (byDivider[divider.id] || []).filter(matchTask);
        if (!divTasks.length && (q || assigneeStr || statusStr)) return;
        sections.push({ proj, divider, tasks: divTasks.length ? divTasks : (byDivider[divider.id] || []) });
      });
      const ungrouped = (byDivider.__none__ || []).filter(matchTask);
      if (!ungrouped.length && (q || assigneeStr || statusStr) && !(byDivider.__none__ || []).length) return;
      sections.push({ proj, divider: null, tasks: ungrouped.length ? ungrouped : (q || assigneeStr || statusStr ? [] : (byDivider.__none__ || [])) });
    });
    return sections.filter((s) => s.tasks.length);
  }

  function renderMtfBody(bodyEl) {
    if (!bodyEl) return;
    const combined = buildCombined(filters.search, filters.assignee, filters.status);

    if (!combined.length) {
      bodyEl.innerHTML = `<div class="ptl-empty" style="padding:32px;text-align:center;color:var(--aewttr-muted);font-size:13px;">${filters.search || filters.assignee || filters.status ? "No tasks match your filters." : "No tasks in this meeting."}</div>`;
      return;
    }

    bodyEl.innerHTML = "";

    // Group sections by project
    const byProject = [];
    const projOrder = [];
    combined.forEach(({ proj, divider, tasks }) => {
      let group = byProject.find((g) => g.proj.id === proj.id);
      if (!group) {
        group = { proj, sections: [] };
        byProject.push(group);
        projOrder.push(proj.id);
      }
      group.sections.push({ divider, tasks });
    });

    byProject.forEach(({ proj, sections }) => {
      const totalTasks = sections.reduce((n, s) => n + s.tasks.length, 0);
      const isCollapsed = !!projCollapsed[proj.id];

      const projHeader = document.createElement("div");
      projHeader.className = "mtf-proj-divider";
      projHeader.innerHTML = `
        <button type="button" class="mtf-proj-toggle" data-mtf-toggle-proj="${escapeHtml(proj.id)}" aria-expanded="${!isCollapsed}" aria-label="${isCollapsed ? "Expand" : "Collapse"} ${escapeHtml(proj.name || "Untitled project")}">
          <i class="bx bx-chevron-${isCollapsed ? "right" : "down"}"></i>
        </button>
        <strong class="mtf-proj-name">${escapeHtml(proj.name || "Untitled project")}</strong>
        <span class="mtf-proj-count">${totalTasks} task${totalTasks === 1 ? "" : "s"}</span>`;
      bodyEl.appendChild(projHeader);

      const projContent = document.createElement("div");
      projContent.className = "mtf-proj-content";
      if (isCollapsed) projContent.style.display = "none";
      bodyEl.appendChild(projContent);

      projHeader.querySelector("[data-mtf-toggle-proj]").addEventListener("click", () => {
        projCollapsed[proj.id] = !projCollapsed[proj.id];
        const nowCollapsed = projCollapsed[proj.id];
        projContent.style.display = nowCollapsed ? "none" : "";
        const btn = projHeader.querySelector("[data-mtf-toggle-proj]");
        if (btn) {
          btn.setAttribute("aria-expanded", String(!nowCollapsed));
          btn.querySelector("i").className = `bx bx-chevron-${nowCollapsed ? "right" : "down"}`;
        }
      });

      if (!isCollapsed) {
        const visibleTaskIds = new Set();
        sections.forEach(({ tasks }) => tasks.forEach((task) => visibleTaskIds.add(task.id)));
        const projectTasks = (db.ganttTasks && db.ganttTasks[proj.id]) || [];
        const displayTasks = projectTasks.filter((task) => !isTrackerDivider(task) && visibleTaskIds.has(task.id));
        const projectMount = document.createElement("div");
        projectMount.className = "mtf-section-mount";
        projContent.appendChild(projectMount);

        if (typeof renderTrackerWorkspace !== "function") return;
        const stateKey = `mtf-${scopeKey}-${proj.id}`;
        renderTrackerWorkspace(projectMount, {
          proj,
          tasks: displayTasks,
          stateKey,
          defaultView: "table",
          hideViewTabs: true,
          noToolbar: true,
          hideHeader: true,
          meetingMode: true,
          saveSource: "Meeting",
          showHint: false,
          collapseGanttGroupsByDefault: false,
          collapseDividersByDefault: false,
          flatGantt: true,
          onRedraw: onChange,
          onOpenEditor: (taskId, subIndex) => {
            if (typeof openMeetingTaskEditor === "function") {
              openMeetingTaskEditor(scope, proj.id, taskId, onChange, subIndex);
            }
          },
          onAfterSave: () => {
            if (typeof saveMeetingSession === "function") saveMeetingSession(scope);
          }
        });
      }
    });
  }

  function redrawFull() { renderMeetingFullTasksList(mount, scope, projects, onRedraw); }

  // Only rebuild the head HTML on a full redraw (not on search/filter changes)
  const existingHead = mount.querySelector(".mtf-global-head");
  if (!existingHead) {
    mount.innerHTML = `
      <div class="mtf-global-head">
        <div class="monday-tracker-tabs mtf-global-tabs" id="mtf-global-tabs">
          <button type="button" class="monday-tracker-tab active" data-mtf-view="table">Main table</button>
          <button type="button" class="monday-tracker-tab" data-mtf-view="gantt">Timeline</button>
        </div>
        <div class="mtf-global-filters">
          <div class="monday-tracker-search" style="flex:1;min-width:160px;">
            <i class="bx bx-search"></i>
            <input type="text" id="mtf-search" placeholder="Search tasks…" value="${escapeHtml(filters.search)}" autocomplete="off">
            <button type="button" class="gantt-search-clear" id="mtf-search-clear" style="${filters.search ? "" : "display:none;"}">&times;</button>
          </div>
          <select class="select-aewttr" id="mtf-assignee-filter" style="max-width:180px;">
            <option value="">All owners</option>
            ${Array.from(allAssignees).sort().map((a) => `<option value="${escapeHtml(a)}" ${filters.assignee === a ? "selected" : ""}>${escapeHtml(a)}</option>`).join("")}
          </select>
          <select class="select-aewttr" id="mtf-status-filter" style="max-width:150px;">
            <option value="">All statuses</option>
            ${["Not Started","In Progress","On Hold","Done"].map((s) => `<option value="${escapeHtml(s)}" ${filters.status === s ? "selected" : ""}>${escapeHtml(s)}</option>`).join("")}
          </select>
          <button type="button" class="btn-aewttr btn-aewttr-sm" id="mtf-add-task"><i class="bx bx-plus" aria-hidden="true"></i> Add task</button>
        </div>
      </div>
      <div id="mtf-unified-body"></div>
    `;

    $all("[data-mtf-view]", mount).forEach((btn) => {
      btn.addEventListener("click", () => {
        window.AEWTTR.state.mtfView[scopeKey] = btn.dataset.mtfView;
        renderMeetingFullTasksList(mount, scope, projects, onRedraw);
      });
    });
    const addTaskBtn = $("#mtf-add-task", mount);
    if (addTaskBtn) {
      addTaskBtn.addEventListener("click", () => openMeetingAddTaskModal(
        scope,
        currentMeetingParticipant(scope),
        redrawFull
      ));
    }

    // Inject the shared column header once above the unified body
    const headerMount = document.createElement("div");
    headerMount.className = "mtf-global-col-header";
    headerMount.innerHTML = `<table class="monday-table monday-table--tracker mtf-header-table"><thead><tr>
      <th class="mgp-num-col mgp-num-col--head">#</th>
      <th class="monday-drag-cell" aria-label="Reorder"></th>
      <th class="monday-complete-col" aria-label="Done"></th>
      <th class="monday-expand-cell"></th>
      <th>Item</th><th>Owner</th><th>Start</th><th>End</th><th>Health</th><th>Notes</th>
    </tr></thead></table>`;
    const unifiedBody = mount.querySelector("#mtf-unified-body");
    if (unifiedBody && unifiedBody.parentNode) unifiedBody.parentNode.insertBefore(headerMount, unifiedBody);

    let searchTimer = null;
    const searchInput = $("#mtf-search", mount);
    if (searchInput) {
      searchInput.addEventListener("input", () => {
        filters.search = searchInput.value;
        const clr = $("#mtf-search-clear", mount);
        if (clr) clr.style.display = filters.search ? "" : "none";
        clearTimeout(searchTimer);
        searchTimer = setTimeout(() => renderMtfBody($("#mtf-unified-body", mount)), 150);
      });
    }
    const searchClear = $("#mtf-search-clear", mount);
    if (searchClear) {
      searchClear.addEventListener("click", () => {
        filters.search = "";
        const si = $("#mtf-search", mount);
        if (si) si.value = "";
        renderMtfBody($("#mtf-unified-body", mount));
      });
    }
    const assigneeFilter = $("#mtf-assignee-filter", mount);
    if (assigneeFilter) assigneeFilter.addEventListener("change", () => { filters.assignee = assigneeFilter.value; renderMtfBody($("#mtf-unified-body", mount)); });
    const statusFilter = $("#mtf-status-filter", mount);
    if (statusFilter) statusFilter.addEventListener("change", () => { filters.status = statusFilter.value; renderMtfBody($("#mtf-unified-body", mount)); });
  }

  renderMtfBody($("#mtf-unified-body", mount));
}

function renderMeetingTasksList(mount, proj, allTasks, visibleTasks, onOpenEditor, onChange, stateKey) {
  const db = window.AEWTTR.db;
  const multi = Array.isArray(proj);
  const bundles = multi
    ? proj
    : [{ proj, tasks: (visibleTasks && visibleTasks.length ? visibleTasks : allTasks) || [] }];
  if (!window.AEWTTR.state.tasksListFilters) window.AEWTTR.state.tasksListFilters = {};
  const filters = window.AEWTTR.state.tasksListFilters[stateKey] || {
    status: "",
    health: "",
    assignee: "",
    endItem: "",
    portfolio: "",
    sortBy: "hierarchy",
    showMore: false
  };
  window.AEWTTR.state.tasksListFilters[stateKey] = filters;

  const HEALTH_RANK = { "Off Track": 3, "At Risk": 2, "On Track": 1 };
  function healthOf(item) {
    return item && item.health ? item.health : "On Track";
  }
  function healthClass(health) {
    if (health === "Off Track") return "off";
    if (health === "At Risk") return "at";
    return "on";
  }
  function healthPtlClass(health) {
    if (health === "Off Track") return "ptl-health--off";
    if (health === "At Risk") return "ptl-health--at";
    return "ptl-health--on";
  }
  function dividerMeta(divider) {
    return (divider && divider.metadata) || {};
  }
  function taskEndItem(task, divider, project) {
    return task.endItemConfig
      || task.configEndItem
      || (divider && (dividerMeta(divider).configEndItem || divider.configEndItem))
      || (project && project.configEndItem)
      || "";
  }
  function taskPortfolios(task, divider, project) {
    const fromDiv = divider
      ? (Array.isArray(dividerMeta(divider).portfolios) ? dividerMeta(divider).portfolios : (divider.portfolios || []))
      : [];
    const fromProj = project
      ? (typeof projectPortfolios === "function" ? projectPortfolios(project) : (project.portfolios || []))
      : [];
    return [...new Set([...(fromDiv || []), ...(fromProj || [])].filter(Boolean))];
  }
  function collectAssignees(task) {
    const names = new Set();
    if (task.assignee) names.add(task.assignee);
    if (typeof walkNestedSubtasks === "function") {
      walkNestedSubtasks(task.subtasks || [], (sub) => {
        if (sub.assignee) names.add(sub.assignee);
      });
    } else {
      (task.subtasks || []).forEach((sub) => { if (sub.assignee) names.add(sub.assignee); });
    }
    return names;
  }

  const allStatuses = new Set();
  const allAssignees = new Set(["Unassigned"]);
  const allEndItems = new Set();
  const allPortfolios = new Set();

  bundles.forEach(({ proj: project, tasks }) => {
    if (project && project.configEndItem) allEndItems.add(project.configEndItem);
    (typeof projectPortfolios === "function" ? projectPortfolios(project) : (project.portfolios || [])).forEach((p) => allPortfolios.add(p));
    const dividersById = new Map();
    tasks.forEach((t) => { if (isTrackerDivider(t)) dividersById.set(t.id, t); });
    tasks.forEach((t) => {
      if (isTrackerDivider(t)) {
        const meta = dividerMeta(t);
        if (meta.configEndItem) allEndItems.add(meta.configEndItem);
        (meta.portfolios || t.portfolios || []).forEach((p) => allPortfolios.add(p));
        return;
      }
      if (t.status) allStatuses.add(t.status);
      collectAssignees(t).forEach((n) => allAssignees.add(n));
      const parent = t.parentDividerId ? dividersById.get(t.parentDividerId) : null;
      const endItem = taskEndItem(t, parent, project);
      if (endItem) allEndItems.add(endItem);
      taskPortfolios(t, parent, project).forEach((p) => allPortfolios.add(p));
    });
  });

  function matchesFilters(task, project, dividersById) {
    if (isTrackerDivider(task)) return false;
    const parent = task.parentDividerId ? dividersById.get(task.parentDividerId) : null;
    const health = healthOf(task);
    if (filters.status && (task.status || "Not Started") !== filters.status) return false;
    if (filters.health === "high" && health === "On Track") return false;
    if (filters.health === "off" && health !== "Off Track") return false;
    if (filters.health === "at" && health !== "At Risk") return false;
    if (filters.health === "on" && health !== "On Track") return false;
    if (filters.assignee) {
      const names = collectAssignees(task);
      if (filters.assignee === "Unassigned") {
        if (task.assignee && task.assignee !== "Unassigned") return false;
      } else if (!names.has(filters.assignee)) {
        return false;
      }
    }
    if (filters.endItem && taskEndItem(task, parent, project) !== filters.endItem) return false;
    if (filters.portfolio) {
      const ports = taskPortfolios(task, parent, project);
      if (!ports.includes(filters.portfolio)) return false;
    }
    return true;
  }

  function sortTasks(list) {
    const sorted = list.slice();
    if (filters.sortBy === "health-desc") {
      sorted.sort((a, b) => (HEALTH_RANK[healthOf(b)] || 0) - (HEALTH_RANK[healthOf(a)] || 0)
        || String(a.title || "").localeCompare(String(b.title || "")));
    } else if (filters.sortBy === "health-asc") {
      sorted.sort((a, b) => (HEALTH_RANK[healthOf(a)] || 0) - (HEALTH_RANK[healthOf(b)] || 0)
        || String(a.title || "").localeCompare(String(b.title || "")));
    } else if (filters.sortBy === "due") {
      sorted.sort((a, b) => String(a.end || "9999").localeCompare(String(b.end || "9999"))
        || String(a.title || "").localeCompare(String(b.title || "")));
    } else if (filters.sortBy === "name") {
      sorted.sort((a, b) => String(a.title || "").localeCompare(String(b.title || "")));
    }
    return sorted;
  }

  const flatSort = filters.sortBy !== "hierarchy";
  const groups = [];
  /* projectForTask maps task id → project object for use in table cells */
  const projectForTask = new Map();
  const sectionForTask = new Map();

  bundles.forEach(({ proj: project, tasks }) => {
    const dividersById = new Map();
    tasks.forEach((t) => { if (isTrackerDivider(t)) dividersById.set(t.id, t); });
    const matched = sortTasks(tasks.filter((t) => matchesFilters(t, project, dividersById)));
    if (!matched.length && (filters.status || filters.health || filters.assignee || filters.endItem || filters.portfolio)) {
      return;
    }
    matched.forEach((t) => {
      projectForTask.set(t.id, project);
      if (t.parentDividerId) {
        const div = dividersById.get(t.parentDividerId);
        if (div) sectionForTask.set(t.id, div.title || "");
      }
    });
    if (flatSort) {
      if (!matched.length) return;
      groups.push({
        key: `flat-${project.id}`,
        label: project.name || project.id || "Project",
        projId: project.id,
        kind: "project",
        tasks: matched
      });
      return;
    }
    groups.push({
      key: `proj-${project.id}`,
      label: project.name || project.id || "Project",
      projId: project.id,
      kind: "project",
      meta: [
        project.priority ? `Priority ${project.priority}` : "",
        (typeof projectPortfolios === "function" ? projectPortfolios(project) : project.portfolios || []).join(", ")
      ].filter(Boolean).join(" · "),
      tasks: []
    });
    const used = new Set();
    tasks.filter(isTrackerDivider).forEach((divider) => {
      const children = matched.filter((t) => t.parentDividerId === divider.id);
      if (!children.length) return;
      children.forEach((t) => used.add(t.id));
      const meta = dividerMeta(divider);
      groups.push({
        key: `${project.id}:${divider.id}`,
        label: divider.title || "Divider",
        projId: project.id,
        kind: "divider",
        health: healthOf(divider),
        meta: [
          meta.configEndItem || divider.configEndItem || "",
          (meta.portfolios || []).join(", ")
        ].filter(Boolean).join(" · "),
        tasks: children
      });
    });
    const ungrouped = matched.filter((t) => !t.parentDividerId && !used.has(t.id));
    if (ungrouped.length) {
      groups.push({
        key: `${project.id}:ungrouped`,
        label: multi ? `${project.name || project.id} · Ungrouped` : "Ungrouped",
        projId: project.id,
        kind: "ungrouped",
        tasks: ungrouped
      });
    }
  });

  /* ---- build table rows ---- */
  const colSpan = multi ? 8 : 7;

  function renderSubTableRows(task) {
    const rows = [];
    function walk(subs, pathPrefix) {
      (subs || []).forEach((sub, idx) => {
        const path = pathPrefix ? `${pathPrefix}.${idx}` : String(idx);
        const h = healthOf(sub);
        const dueStr = sub.end ? fmtDate(sub.end) : "";
        rows.push(`<tr class="ptl-sub-row ptl-sub-hidden" data-subtask="${escapeHtml(task.id)}:${escapeHtml(path)}">
          <td class="ptl-td-expand"></td>
          <td class="ptl-td-health"><span class="ptl-health ${healthPtlClass(h)}" title="${escapeHtml(h)}"></span></td>
          <td class="ptl-td-title${multi ? "" : ""}" colspan="${multi ? 5 : 4}">
            <button type="button" class="ptl-title-btn ptl-title-btn--sub" data-open-subtask="${escapeHtml(task.id)}:${escapeHtml(path)}">
              <i class="bx bx-subdirectory-right ptl-sub-icon" aria-hidden="true"></i>
              ${escapeHtml(sub.text || sub.title || "Untitled")}
            </button>
          </td>
          <td class="ptl-td-due">${escapeHtml(dueStr)}</td>
        </tr>`);
        if (sub.subtasks && sub.subtasks.length) walk(sub.subtasks, path);
      });
    }
    walk(task.subtasks || [], "");
    return rows.join("");
  }

  function renderTaskTableRow(task, projObj) {
    const health = healthOf(task);
    const hasSubtasks = (task.subtasks || []).length > 0;
    const status = task.status || "Not Started";
    const statusCls = `ptl-status-${status.toLowerCase().replace(/\s+/g, "-")}`;
    const dueStr = task.end ? fmtDate(task.end) : "";
    const projName = projObj ? (projObj.name || projObj.id || "") : "";
    const sectionName = sectionForTask.get(task.id) || "";
    return `<tr class="ptl-task-row" data-task-id="${escapeHtml(task.id)}">
      <td class="ptl-td-expand">
        ${hasSubtasks
          ? `<button type="button" class="ptl-expand-btn" data-expand-task="${escapeHtml(task.id)}" aria-label="Toggle subtasks"><i class="bx bx-chevron-right"></i></button>`
          : ""}
      </td>
      <td class="ptl-td-health"><span class="ptl-health ${healthPtlClass(health)}" title="${escapeHtml(health)}"></span></td>
      <td class="ptl-td-title">
        <button type="button" class="ptl-title-btn" data-open-task="${escapeHtml(task.id)}">${escapeHtml(task.title || "Untitled")}</button>
      </td>
      ${multi ? `<td class="ptl-td-project">${escapeHtml(projName)}</td>` : ""}
      <td class="ptl-td-section">${escapeHtml(sectionName)}</td>
      <td class="ptl-td-assignee">${escapeHtml(task.assignee || "")}</td>
      <td class="ptl-td-status"><span class="ptl-status-pill ${escapeHtml(statusCls)}">${escapeHtml(status)}</span></td>
      <td class="ptl-td-due">${escapeHtml(dueStr)}</td>
    </tr>
    ${hasSubtasks ? renderSubTableRows(task) : ""}`;
  }

  /* Build all <tbody> rows */
  const tbodyRows = [];
  let lastProjId = null;
  groups.forEach((group) => {
    /* Project header row in multi-project mode when proj changes */
    if (multi && group.kind === "project") {
      const taskCount = group.tasks.length;
      tbodyRows.push(`<tr class="ptl-project-row">
        <td class="ptl-project-cell" colspan="${colSpan}">
          <strong>${escapeHtml(group.label)}</strong>
          <span class="ptl-project-id">${escapeHtml(group.projId || "")}</span>
          <span class="ptl-badge">${taskCount} task${taskCount !== 1 ? "s" : ""}</span>
        </td>
      </tr>`);
      lastProjId = group.projId;
    } else if (group.kind === "divider" || group.kind === "ungrouped") {
      /* In multi mode, show project header before first divider for this project if not yet shown */
      if (multi && group.projId && group.projId !== lastProjId) {
        const bundle = bundles.find((b) => b.proj && b.proj.id === group.projId);
        const projName = bundle ? (bundle.proj.name || bundle.proj.id || "") : group.projId;
        const projTaskCount = groups.filter((g) => g.projId === group.projId).reduce((n, g) => n + (g.tasks || []).length, 0);
        tbodyRows.push(`<tr class="ptl-project-row">
          <td class="ptl-project-cell" colspan="${colSpan}">
            <strong>${escapeHtml(projName)}</strong>
            <span class="ptl-project-id">${escapeHtml(group.projId)}</span>
            <span class="ptl-badge">${projTaskCount} task${projTaskCount !== 1 ? "s" : ""}</span>
          </td>
        </tr>`);
        lastProjId = group.projId;
      }
      /* Divider/section group row */
      tbodyRows.push(`<tr class="ptl-group-row">
        <td class="ptl-group-cell" colspan="${colSpan}">
          ${group.health ? `<span class="ptl-health ${healthPtlClass(group.health)}" title="${escapeHtml(group.health)}"></span>` : ""}
          <strong>${escapeHtml(group.label)}</strong>
          ${group.meta ? `<span class="ptl-group-meta">${escapeHtml(group.meta)}</span>` : ""}
        </td>
      </tr>`);
    }
    /* Task rows for this group */
    (group.tasks || []).forEach((task) => {
      const projObj = projectForTask.get(task.id);
      tbodyRows.push(renderTaskTableRow(task, projObj));
    });
  });

  const totalShown = groups.reduce((n, g) => n + (g.tasks || []).length, 0);
  const filtersActive = !!(filters.status || filters.health || filters.assignee || filters.endItem || filters.portfolio || (filters.sortBy && filters.sortBy !== "hierarchy"));
  if (filters.showMore == null) filters.showMore = false;

  mount.innerHTML = `
    <div class="ptl-shell">
      <div class="ptl-toolbar">
        <div class="ptl-seg" role="group" aria-label="Health filter">
          ${[
            ["", "All"],
            ["high", "High"],
            ["off", "Off"],
            ["at", "At risk"],
            ["on", "On track"]
          ].map(([val, label]) => `
            <button type="button" class="ptl-seg-btn ${filters.health === val ? "is-active" : ""}" data-health-filter="${val}">${label}</button>
          `).join("")}
        </div>
        <select class="select-aewttr ptl-sort" id="filter-sort-${stateKey}" aria-label="Sort">
          <option value="hierarchy" ${filters.sortBy === "hierarchy" ? "selected" : ""}>By divider</option>
          <option value="health-desc" ${filters.sortBy === "health-desc" ? "selected" : ""}>Health ↓</option>
          <option value="health-asc" ${filters.sortBy === "health-asc" ? "selected" : ""}>Health ↑</option>
          <option value="due" ${filters.sortBy === "due" ? "selected" : ""}>Due date</option>
          <option value="name" ${filters.sortBy === "name" ? "selected" : ""}>Name</option>
        </select>
        <button type="button" class="btn-aewttr-ghost btn-aewttr-sm ptl-more-btn ${filters.showMore ? "is-open" : ""}" id="ptl-more-${stateKey}">
          More filters${filtersActive && (filters.status || filters.assignee || filters.endItem || filters.portfolio) ? " ·" : ""}
        </button>
        ${filtersActive ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-sm ptl-clear-btn" id="clear-filters-${stateKey}">Clear</button>` : ""}
        <span class="ptl-count">${totalShown} task${totalShown !== 1 ? "s" : ""}</span>
      </div>
      <div class="ptl-more-panel ${filters.showMore ? "is-open" : ""}" id="ptl-more-panel-${stateKey}">
        <select class="select-aewttr" id="filter-status-${stateKey}" aria-label="Status">
          <option value="">Status: any</option>
          ${Array.from(allStatuses).sort().map((s) => `<option value="${escapeHtml(s)}" ${filters.status === s ? "selected" : ""}>${escapeHtml(s)}</option>`).join("")}
        </select>
        <select class="select-aewttr" id="filter-assignee-${stateKey}" aria-label="Assignee">
          <option value="">Assigned: anyone</option>
          ${Array.from(allAssignees).sort().map((m) => `<option value="${escapeHtml(m)}" ${filters.assignee === m ? "selected" : ""}>${escapeHtml(m)}</option>`).join("")}
        </select>
        ${allEndItems.size ? `
          <select class="select-aewttr" id="filter-enditem-${stateKey}" aria-label="End item">
            <option value="">End item: any</option>
            ${Array.from(allEndItems).sort().map((e) => `<option value="${escapeHtml(e)}" ${filters.endItem === e ? "selected" : ""}>${escapeHtml(e)}</option>`).join("")}
          </select>` : ""}
        ${allPortfolios.size ? `
          <select class="select-aewttr" id="filter-portfolio-${stateKey}" aria-label="Portfolio">
            <option value="">Portfolio: any</option>
            ${Array.from(allPortfolios).sort().map((p) => `<option value="${escapeHtml(p)}" ${filters.portfolio === p ? "selected" : ""}>${escapeHtml(p)}</option>`).join("")}
          </select>` : ""}
      </div>
      <div class="ptl-table-wrap">
        ${tbodyRows.length && totalShown ? `
        <table class="ptl-table">
          <thead>
            <tr class="ptl-thead-row">
              <th class="ptl-th ptl-th-expand"></th>
              <th class="ptl-th ptl-th-health">Health</th>
              <th class="ptl-th ptl-th-title">Task</th>
              <th class="ptl-th ptl-th-project${multi ? "" : " ptl-hidden"}">Project</th>
              <th class="ptl-th ptl-th-section">Section</th>
              <th class="ptl-th ptl-th-assignee">Assignee</th>
              <th class="ptl-th ptl-th-status">Status</th>
              <th class="ptl-th ptl-th-due">Due</th>
            </tr>
          </thead>
          <tbody>
            ${tbodyRows.join("")}
          </tbody>
        </table>` : `
        <table class="ptl-table">
          <tbody>
            <tr class="ptl-empty-row"><td colspan="${colSpan}">No tasks match your filters.</td></tr>
          </tbody>
        </table>`}
      </div>
    </div>`;

  function redraw() {
    renderMeetingTasksList(mount, proj, allTasks, visibleTasks, onOpenEditor, onChange, stateKey);
  }
  const bindSelect = (id, key) => {
    const el = mount.querySelector(`#${CSS.escape(id)}`);
    if (!el) return;
    el.addEventListener("change", () => {
      filters[key] = el.value || "";
      redraw();
    });
  };
  mount.querySelectorAll("[data-health-filter]").forEach((btn) => {
    btn.addEventListener("click", () => {
      filters.health = btn.dataset.healthFilter || "";
      redraw();
    });
  });
  bindSelect(`filter-sort-${stateKey}`, "sortBy");
  bindSelect(`filter-status-${stateKey}`, "status");
  bindSelect(`filter-assignee-${stateKey}`, "assignee");
  bindSelect(`filter-enditem-${stateKey}`, "endItem");
  bindSelect(`filter-portfolio-${stateKey}`, "portfolio");

  const moreBtn = mount.querySelector(`#ptl-more-${CSS.escape(stateKey)}`);
  if (moreBtn) {
    moreBtn.addEventListener("click", () => {
      filters.showMore = !filters.showMore;
      redraw();
    });
  }
  const clearBtn = mount.querySelector(`#clear-filters-${CSS.escape(stateKey)}`);
  if (clearBtn) {
    clearBtn.addEventListener("click", () => {
      filters.status = "";
      filters.health = "";
      filters.assignee = "";
      filters.endItem = "";
      filters.portfolio = "";
      filters.sortBy = "hierarchy";
      filters.showMore = false;
      redraw();
    });
  }

  mount.querySelectorAll("[data-open-task]").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.preventDefault();
      if (typeof onOpenEditor === "function") onOpenEditor(btn.dataset.openTask);
    });
  });
  mount.querySelectorAll("[data-open-subtask]").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.preventDefault();
      const raw = btn.dataset.openSubtask || "";
      const colon = raw.indexOf(":");
      const taskId = colon >= 0 ? raw.slice(0, colon) : raw;
      const path = colon >= 0 ? raw.slice(colon + 1) : "";
      if (typeof onOpenEditor === "function") onOpenEditor(taskId, path);
    });
  });
  mount.querySelectorAll("[data-expand-task]").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const taskId = btn.dataset.expandTask;
      const taskRow = btn.closest(".ptl-task-row");
      if (!taskRow) return;
      const expanded = taskRow.classList.toggle("ptl-task-row--expanded");
      mount.querySelectorAll(`.ptl-sub-row[data-subtask^="${CSS.escape(taskId)}:"]`).forEach((sub) => {
        sub.classList.toggle("ptl-sub-hidden", !expanded);
      });
    });
  });
}

/* ---------- entry point used by Projects > Meeting tab ---------- */
function openMeetingConfigModal(proj, onDone) {
  const extra = ensureProjectExtra(proj.id);
  const cfg = extra.meetingConfig || {};
  const FREQ_OPTIONS = ["Weekly", "Bi-weekly", "Monthly", "Bi-monthly", "Quarterly", "Ad hoc"];
  let agendaItems = Array.isArray(cfg.agendaItems) ? cfg.agendaItems.map(a => Object.assign({}, a)) : [];

  function renderAgendaItems(list) {
    if (!list) return;
    if (!agendaItems.length) {
      list.innerHTML = `<div class="mtg-cfg-agenda-empty">No agenda items yet — click "+ Add item".</div>`;
      return;
    }
    list.innerHTML = agendaItems.map((item, idx) => `
      <div class="mtg-cfg-agenda-row" data-aidx="${idx}">
        <i class="bx bx-menu" style="cursor:grab;color:var(--aewttr-muted);flex-shrink:0;"></i>
        <input class="input-aewttr mtg-cfg-agenda-input" data-aidx="${idx}" value="${escapeHtml(item.text || "")}" placeholder="Agenda item…">
        <button type="button" class="mtg-cfg-agenda-delete" data-aidx="${idx}" title="Remove"><i class="bx bx-trash"></i></button>
      </div>`).join("");
    $all(".mtg-cfg-agenda-input", list).forEach(input => input.addEventListener("change", () => {
      agendaItems[Number(input.dataset.aidx)].text = input.value.trim();
    }));
    $all(".mtg-cfg-agenda-delete", list).forEach(btn => btn.addEventListener("click", () => {
      agendaItems.splice(Number(btn.dataset.aidx), 1);
      renderAgendaItems(list);
    }));
  }

  const modal = openModal(`
    <div class="aewttr-modal-head">
      <h3>Meeting Settings</h3>
      <button class="aewttr-modal-close" type="button">&times;</button>
    </div>
    <div class="aewttr-modal-body">
      <div class="form-grid-2">
        <div class="form-row"><label>Meeting name</label><input class="input-aewttr" id="mcfg-name" value="${escapeHtml(cfg.name || "")}" placeholder="e.g., Weekly IPT, Sprint Review…"></div>
        <div class="form-row"><label>Frequency</label><select class="select-aewttr" id="mcfg-freq">${FREQ_OPTIONS.map(f => `<option ${cfg.frequency === f ? "selected" : ""}>${f}</option>`).join("")}</select></div>
      </div>
      <div class="form-row"><label>External stakeholders <small style="font-weight:400;color:var(--aewttr-muted);">(comma-separated)</small></label><input class="input-aewttr" id="mcfg-external" value="${escapeHtml((cfg.externalStakeholders || []).join(", "))}" placeholder="e.g., Range Safety Office, Program Executive…"></div>
      <div class="form-row" style="margin-bottom:4px;"><label>Notes / standing context</label><textarea class="textarea-aewttr" id="mcfg-notes" rows="2" placeholder="Location, dial-in info, recurring context…">${escapeHtml(cfg.notes || "")}</textarea></div>
      <div class="mtg-cfg-agenda-head">
        <span>Standard agenda</span>
        <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="mcfg-add-agenda"><i class="bx bx-plus"></i> Add item</button>
      </div>
      <div id="mcfg-agenda-list" class="mtg-cfg-agenda-list"></div>
    </div>
    <div class="aewttr-modal-foot">
      <button type="button" class="btn-aewttr-ghost" id="mcfg-cancel">Cancel</button>
      <button type="button" class="btn-aewttr" id="mcfg-save">Save Settings</button>
    </div>
  `);

  const agendaList = $("#mcfg-agenda-list", modal);
  renderAgendaItems(agendaList);

  $("#mcfg-add-agenda", modal).addEventListener("click", () => {
    agendaItems.push({ id: uid("agi"), text: "" });
    renderAgendaItems(agendaList);
    const inputs = $all(".mtg-cfg-agenda-input", modal);
    if (inputs.length) inputs[inputs.length - 1].focus();
  });

  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#mcfg-cancel", modal).addEventListener("click", closeModal);
  $("#mcfg-save", modal).addEventListener("click", () => {
    $all(".mtg-cfg-agenda-input", modal).forEach(input => {
      const idx = Number(input.dataset.aidx);
      if (agendaItems[idx]) agendaItems[idx].text = input.value.trim();
    });
    const externalRaw = $("#mcfg-external", modal).value.trim();
    extra.meetingConfig = {
      name: $("#mcfg-name", modal).value.trim(),
      frequency: $("#mcfg-freq", modal).value,
      externalStakeholders: externalRaw ? externalRaw.split(",").map(s => s.trim()).filter(Boolean) : [],
      notes: $("#mcfg-notes", modal).value.trim(),
      agendaItems: agendaItems.filter(a => a.text)
    };
    aewttrSaveStore();
    closeModal();
    toast("Meeting settings saved", "success");
    if (onDone) onDone();
  });
}

window.AEWTTR.renderProjectMeetingApp = function (body, proj) {
  /* Render a lightweight meeting view inside the project page shell.
     Does NOT call renderMeetingApp (which would add its own sidebar and
     call setTopbar, making it feel like you left the project). */
  const scope = { type: "project", project: proj };
  const active = meetingIsActive(scope);
  const canFacilitate = canManageMeetings();
  const mtgData = meetingData(scope);
  const planning = !active && mtgData.currentSession && mtgData.currentSession.sessionStatus === "planned" ? mtgData.currentSession : null;
  const session = active || planning || meetingStatus(scope) === "ended" ? (active ? ensureMeetingSession(scope) : mtgData.currentSession) : null;
  const participants = meetingParticipants(scope);
  const me = currentMeetingParticipant(scope);

  // Live-tab state: "tasks" or "notes" (only used during active meeting)
  if (!window.AEWTTR.state.projMtgTab) window.AEWTTR.state.projMtgTab = {};
  const tabKey = "proj:" + proj.id;
  const liveTab = window.AEWTTR.state.projMtgTab[tabKey] || "tasks";

  // Plan-tab state: "agenda" or "notes" (only used during planning)
  if (!window.AEWTTR.state.projMtgPlanTab) window.AEWTTR.state.projMtgPlanTab = {};
  const planTab = window.AEWTTR.state.projMtgPlanTab[tabKey] || "agenda";

  function attendanceBarHtml() {
    if (!session || !participants.length) return "";
    const hereCount = participants.filter(p => session.attendance[p.id] === "Here").length;
    return `
      <div class="proj-mtg-attendance-bar">
        <span class="proj-mtg-attendance-label"><i class="bx bx-group"></i> ${hereCount}/${participants.length} present</span>
        <div class="proj-mtg-attendance-pills">
          ${participants.map(p => {
            const here = session.attendance[p.id] === "Here";
            const editable = canFacilitate || (me && p.id === me.id);
            return `<button type="button" class="proj-mtg-presence-pill ${here ? "here" : "out"}" data-presence="${escapeHtml(p.id)}" aria-label="${escapeHtml(p.name)}" aria-pressed="${here}" ${editable ? "" : "disabled"}${tip(p.name)}>
              ${userAvatarHtml({ name: p.name, size: 24 })}
              <span class="proj-mtg-pill-name">${escapeHtml(p.name.split(" ")[0])}</span>
              <span class="proj-mtg-pill-status">${here ? "Here" : "Out"}</span>
            </button>`;
          }).join("")}
        </div>
        ${!canFacilitate && me && session.attendance[me.id] !== "Here" ? `<button class="btn-aewttr btn-aewttr-sm" id="proj-mtg-join-self"><i class="bx bx-log-in"></i> Join</button>` : ""}
      </div>`;
  }

  const meetingCfg = (ensureProjectExtra(proj.id).meetingConfig) || {};
  function meetingCfgBannerHtml() {
    if (!meetingCfg.name && !meetingCfg.frequency && !(meetingCfg.agendaItems || []).length) return "";
    const agendaItems = (meetingCfg.agendaItems || []).filter(a => a.text);
    const ext = (meetingCfg.externalStakeholders || []).filter(Boolean);
    return `<div class="proj-mtg-cfg-banner">
      ${meetingCfg.name ? `<span class="proj-mtg-cfg-name"><i class="bx bx-calendar-event"></i> ${escapeHtml(meetingCfg.name)}</span>` : ""}
      ${meetingCfg.frequency ? `<span class="proj-mtg-cfg-chip">${escapeHtml(meetingCfg.frequency)}</span>` : ""}
      ${ext.length ? `<span class="proj-mtg-cfg-chip" title="${escapeHtml(ext.join(", "))}"><i class="bx bx-group"></i> ${ext.length} external</span>` : ""}
      ${agendaItems.length ? `<span class="proj-mtg-cfg-chip"><i class="bx bx-list-ul"></i> ${agendaItems.length} agenda item${agendaItems.length !== 1 ? "s" : ""}</span>` : ""}
      ${meetingCfg.notes ? `<span class="proj-mtg-cfg-notes">${escapeHtml(meetingCfg.notes)}</span>` : ""}
    </div>`;
  }

  const hasPlanningNotes = planning && Array.isArray(planning.docBlocks) && planning.docBlocks.some(b => b.content);

  body.innerHTML = `
    <div class="proj-meeting-shell">
      <div class="proj-meeting-header">
        ${active ? `<span class="kc-badge proj-meeting-status-badge">Meeting active</span>` : ``}
        ${planning ? `<span class="kc-badge proj-meeting-status-badge proj-meeting-status-badge--plan"><i class="bx bx-edit-alt"></i> Pre-meeting</span>` : ``}
        ${!active && !planning && meetingStatus(scope) === "ended" ? `<span class="kc-badge proj-meeting-status-badge proj-meeting-status-badge--ended">Meeting ended</span>` : ``}
        <span class="proj-meeting-scope-label">${escapeHtml(proj.name || "Untitled project")} · meeting</span>
        ${active ? `
          <div class="segmented proj-mtg-live-tabs">
            <button class="segmented-opt ${liveTab === "tasks" ? "active" : ""}" data-proj-mtg-tab="tasks"><i class="bx bx-list-check"></i> Tasks</button>
            <button class="segmented-opt ${liveTab === "notes" ? "active" : ""}" data-proj-mtg-tab="notes"><i class="bx bx-note"></i> Notes</button>
          </div>
        ` : ""}
        ${planning ? `
          <div class="segmented proj-mtg-live-tabs">
            <button class="segmented-opt ${planTab === "agenda" ? "active" : ""}" data-proj-plan-tab="agenda"><i class="bx bx-list-ul"></i> Agenda</button>
            <button class="segmented-opt ${planTab === "notes" ? "active" : ""}" data-proj-plan-tab="notes"><i class="bx bx-note"></i> Notes</button>
          </div>
        ` : ""}
        <div class="proj-meeting-header-actions">
          ${session && active ? `<button class="btn-aewttr-ghost btn-aewttr-sm" id="proj-mtg-attendance-modal"${tip("View full attendance list")}><i class="bx bx-user-check"></i></button>` : ""}
          <button class="btn-aewttr-ghost btn-aewttr-sm" id="proj-mtg-settings"${tip("Configure recurring meeting settings")}><i class="bx bx-cog"></i></button>
          ${!active && !planning ? `<button class="btn-aewttr-outline btn-aewttr-sm" id="proj-mtg-plan"${tip("Set agenda and notes before the meeting starts")}><i class="bx bx-edit-alt"></i> Plan Meeting</button>` : ``}
          ${planning ? `<button class="btn-aewttr-outline btn-aewttr-sm" id="proj-mtg-cancel-plan"${tip("Discard the pre-meeting draft")}><i class="bx bx-x"></i> Discard Draft</button>` : ``}
          ${canFacilitate && active ? `<button class="btn-aewttr-outline btn-aewttr-sm" id="proj-mtg-end"><i class="bx bx-archive"></i> End Meeting</button>` : ``}
          ${canFacilitate && !active ? `<button class="btn-aewttr btn-aewttr-sm" id="proj-mtg-start"><i class="bx bx-play"></i> Start Meeting</button>` : ``}
        </div>
      </div>
      ${active ? attendanceBarHtml() : ""}
      ${!active && !planning ? meetingCfgBannerHtml() : ""}
      <div class="proj-meeting-body" id="proj-mtg-body"></div>
    </div>`;

  function redraw() {
    window.AEWTTR.renderProjectMeetingApp(body, proj);
  }

  const settingsBtn = body.querySelector("#proj-mtg-settings");
  if (settingsBtn) settingsBtn.addEventListener("click", () => openMeetingConfigModal(proj, redraw));

  const endBtn = body.querySelector("#proj-mtg-end");
  if (endBtn) endBtn.addEventListener("click", () => openEndMeetingModal(scope, redraw));

  const startBtn = body.querySelector("#proj-mtg-start");
  if (startBtn) startBtn.addEventListener("click", () => startMeeting(scope, redraw));

  const planBtn = body.querySelector("#proj-mtg-plan");
  if (planBtn) planBtn.addEventListener("click", async () => { await ensurePlanningSession(scope); redraw(); });

  const cancelPlanBtn = body.querySelector("#proj-mtg-cancel-plan");
  if (cancelPlanBtn) cancelPlanBtn.addEventListener("click", async () => {
    const ok = await confirmDialog({ title: "Discard draft?", message: "Remove the pre-meeting notes draft?", confirmLabel: "Discard", danger: true });
    if (!ok) return;
    mtgData.currentSession = null;
    mtgData.meetingStatus = (mtgData.sessions && mtgData.sessions.length) ? "ended" : "idle";
    try { await Repo.save("meetingSession", { _deleted: true }, { projectCode: meetingProjectCode(scope), immediate: true }); } catch (e) { /* best-effort */ }
    redraw();
  });

  const attendanceModalBtn = body.querySelector("#proj-mtg-attendance-modal");
  if (attendanceModalBtn) attendanceModalBtn.addEventListener("click", () => openMeetingAttendanceModal(scope, redraw));

  body.querySelectorAll("[data-proj-mtg-tab]").forEach(btn => {
    btn.addEventListener("click", () => {
      window.AEWTTR.state.projMtgTab[tabKey] = btn.dataset.projMtgTab;
      redraw();
    });
  });

  body.querySelectorAll("[data-proj-plan-tab]").forEach(btn => {
    btn.addEventListener("click", () => {
      window.AEWTTR.state.projMtgPlanTab[tabKey] = btn.dataset.projPlanTab;
      redraw();
    });
  });

  const joinSelfBtn = body.querySelector("#proj-mtg-join-self");
  if (joinSelfBtn && session && me) {
    joinSelfBtn.addEventListener("click", () => toggleMeetingPresence(scope, body, session, me.id, { canFacilitate, me, allParticipants: participants }));
  }

  if (session && active) {
    body.querySelectorAll("[data-presence]").forEach(btn => {
      btn.addEventListener("click", () => {
        toggleMeetingPresence(scope, body, session, btn.dataset.presence, { canFacilitate, me, allParticipants: participants });
        const here = session.attendance[btn.dataset.presence] === "Here";
        btn.classList.toggle("here", here);
        btn.classList.toggle("out", !here);
        btn.querySelector(".proj-mtg-pill-status").textContent = here ? "Here" : "Out";
        const label = body.querySelector(".proj-mtg-attendance-label");
        if (label) {
          const hc = participants.filter(p => session.attendance[p.id] === "Here").length;
          label.innerHTML = `<i class="bx bx-group"></i> ${hc}/${participants.length} present`;
        }
      });
    });
  }

  const taskMount = body.querySelector("#proj-mtg-body");
  if (!taskMount) return;

  if (planning) {
    // Pre-meeting planning: agenda or notes tab
    taskMount.innerHTML = `<div class="proj-mtg-plan-head"><div><div class="side-panel-title">Pre-Meeting Planning</div><p>Build your agenda and notes before the meeting starts. These carry over when you begin.</p></div>${canFacilitate ? `<button class="btn-aewttr" id="proj-mtg-start-from-plan"><i class="bx bx-play"></i> Start Meeting</button>` : ""}</div><div id="proj-mtg-plan-tab-body"></div>`;
    const startFromPlanBtn = taskMount.querySelector("#proj-mtg-start-from-plan");
    if (startFromPlanBtn) startFromPlanBtn.addEventListener("click", () => startMeeting(scope, redraw));
    const planTabBody = taskMount.querySelector("#proj-mtg-plan-tab-body");
    if (planTabBody) {
      if (planTab === "notes") {
        renderMeetingDocEditor(planTabBody, scope, planning);
      } else {
        renderMeetingAgendaTab(planTabBody, scope, planning);
      }
    }
  } else if (active && liveTab === "notes") {
    // Active meeting notes tab
    renderMeetingDocEditor(taskMount, scope, session);
  } else {
    // Tasks view (default for active and idle/ended states)
    const projects = meetingProjects(scope);
    renderMeetingFullTasksList(taskMount, scope, projects, redraw);
  }
};
