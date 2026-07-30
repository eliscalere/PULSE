/* Admins and Finance Admins open Travel to work the queue, not to look at their
   own trips, so the tab lands them on All Travel. This only seeds the initial
   value — switching the view is remembered for the rest of the session, and an
   explicit route ("travel/mine") still forces its own view. */
function defaultTravelListViewForUser() {
  const canApprove = typeof canApproveTravelRequests === "function" && canApproveTravelRequests();
  const canFinance = typeof canAssignTravelCo === "function" && canAssignTravelCo();
  return (canApprove || canFinance) ? "All Travel" : "My Travel";
}

PAGE_RENDERERS.travel = function (parts) {
  if (!parts || !parts.length) {
    /* The Travel tab lands here. Approvers and Finance Admins open Travel to
       work the queue, not to look at their own trips, so they go to All Travel;
       everyone else keeps My Travel. */
    navigate(defaultTravelListViewForUser() === "All Travel" ? "travel/all" : "travel/mine");
    return;
  }
  renderTravelPage(parts);
};

/* My Travel, All Travel, and Approvals used to be three separate sidebar
   pages; they're now one ("list") with a My Travel / All Travel toggle at
   the top — approving/denying/assigning a C/O happens inline in All Travel
   for whoever has that permission, instead of a separate page. "mine",
   "all", and "approve" stay as route keys (see renderTravelPage) purely so
   existing bookmarks, emailed notification links, and the Overview page's
   "Open Travel" button keep landing somewhere correct. */
const TRAVEL_TOOLS = [
  { key: "submit", title: "Submit a Request", navLabel: "New request", sub: "Start a new TDY, conference, or training travel request.", icon: "bx-send" },
  { key: "list", title: "Travel", navLabel: "Travel", sub: "Track requests, review approvals, and browse every trip.", icon: "bx-list-ul" },
  { key: "debrief", title: "Submit a Travel Debrief", navLabel: "Debrief", sub: "File a debrief for a trip that's already been approved.", icon: "bx-receipt" },
  { key: "calendar", title: "Travel Calendar", navLabel: "Calendar", sub: "See approved travel laid out on a monthly calendar.", icon: "bx-calendar" },
  { key: "events", title: "Team Events", navLabel: "Events", sub: "Schedule and manage non-travel team events.", icon: "bx-calendar-event" }
];

const TRAVEL_ACTION_ICONS = {
  view: "bx-show",
  edit: "bx-edit",
  revoke: "bx-undo",
  cancel: "bx-calendar-x",
  concur: "bx-check-shield",
  "assign-co": "bx-receipt",
  export: "bx-file-blank",
  "open-export": "bx-link-external",
  debrief: "bx-receipt",
  "view-debrief": "bx-file-find",
  "set-status": "bx-transfer-alt",
  "nudge-debrief": "bx-bell-ring",
  "delete-request": "bx-trash",
  close: "bx-x"
};

const TRAVEL_APPROVE_FILTER_TIPS = {
  All: "Show all travel requests",
  Upcoming: "Active requests and upcoming trips",
  Submitted: "Submitted and awaiting administrative actions",
  "Awaiting Finance": "Submitted travel still needing a charge object (C/O) number assigned",
  Withdrawn: "Withdrawn by the requester",
  Cancelled: "Cancelled by requester or admin",
  Completed: "Travel completed"
};

function travelActionBtn(kind, idOrOpts, maybeOpts) {
  let id = idOrOpts;
  let options = maybeOpts || {};
  if (idOrOpts && typeof idOrOpts === "object" && !maybeOpts) {
    options = idOrOpts;
    id = options.id;
  }
  const defaultLabels = {
    view: "View",
    edit: "Edit",
    revoke: "Withdraw",
    cancel: "Cancel",
    concur: "Mark Customer Concurrence",
    "assign-co": "Assign C/O",
    export: "PDF",
    "open-export": "Open PDF",
    debrief: "Debrief",
    "view-debrief": "View Debrief",
    "set-status": "Change Status",
    "nudge-debrief": "Nudge for Debrief",
    "delete-request": "Delete",
    close: "Close"
  };
  const label = options.label || defaultLabels[kind] || kind;
  const tone = options.tone || "default";
  const toneClass = tone === "primary" ? " btn-travel-action--primary" : tone === "danger" ? " btn-travel-action--danger" : "";
  const icon = options.icon || TRAVEL_ACTION_ICONS[kind] || "bx-dots-horizontal-rounded";
  let attrs = "";
  if (options.elementId) attrs += ` id="${escapeHtml(options.elementId)}"`;
  else if (id != null && id !== "") attrs += ` data-${kind}="${escapeHtml(id)}"`;
  const tipText = options.tip || label;
  return `<button type="button" class="btn-travel-action${toneClass}"${attrs}${tip(tipText)} aria-label="${escapeHtml(label)}"><i class="bx ${icon}"></i><span>${escapeHtml(label)}</span></button>`;
}

function travelActionLink(kind, href, options = {}) {
  const label = options.label || "Open";
  const icon = options.icon || TRAVEL_ACTION_ICONS[kind] || "bx-link-external";
  const tipText = options.tip || label;
  return `<a class="btn-travel-action" href="${escapeHtml(href)}" target="_blank" rel="noopener"${options.elementId ? ` id="${escapeHtml(options.elementId)}"` : ""}${tip(tipText)} aria-label="${escapeHtml(label)}"><i class="bx ${icon}"></i><span>${escapeHtml(label)}</span></a>`;
}

function travelActionsWrap(buttons) {
  return `<div class="travel-row-actions">${buttons.filter(Boolean).join("")}</div>`;
}

function travelTileAlertCount(key) {
  const alerts = typeof getAppAlertCounts === "function" ? getAppAlertCounts() : {};
  if (key === "list") return (alerts.travelMine || 0) + (alerts.travelApprove || 0);
  return 0;
}

function renderTravelSidebar(activeSub) {
  const tools = (window.PULSE_PORT_CONFIG && window.PULSE_PORT_CONFIG.sidebarTools)
    ? window.PULSE_PORT_CONFIG.sidebarTools
    : TRAVEL_TOOLS;
  const visibleTools = tools.filter((t) => !t.approverOnly || canAccessTravelApprovals());
  const currentHashRoute = String(location.hash || "").replace(/^#\/?/, "");

  return `
    <aside class="project-spo-nav travel-spo-nav">
      <div class="project-spo-nav-head travel-spo-nav-head">
        <div class="project-spo-id">${escapeHtml((window.PULSE_PORT_CONFIG && window.PULSE_PORT_CONFIG.sidebarTitle) || "Travel")}</div>
        <div class="project-spo-name">${escapeHtml((window.PULSE_PORT_CONFIG && window.PULSE_PORT_CONFIG.sidebarSubtitle) || "Requests & calendar")}</div>
      </div>
      <nav class="project-spo-menu travel-spo-menu" aria-label="Travel sections">
        ${visibleTools.map((tool) => {
          const count = travelTileAlertCount(tool.key);
          const isSubActive = tool.routeOverride
            ? currentHashRoute === tool.routeOverride || (activeSub === tool.key && !currentHashRoute.includes("/"))
            : activeSub === tool.key;
          return `
            <button type="button" class="project-spo-link travel-spo-link ${isSubActive ? "active" : ""}" data-travel-sub="${tool.key}" ${tool.routeOverride ? `data-route-override="${escapeHtml(tool.routeOverride)}"` : ""}>
              <i class="bx ${tool.icon}"></i>
              <span>${escapeHtml(tool.navLabel || tool.title)}</span>
              ${count ? `<em class="project-spo-badge">${count}</em>` : ""}
            </button>`;
        }).join("")}
      </nav>
    </aside>`;
}

function wireTravelSidebar(mount) {
  if (!mount) return;
  $all("[data-travel-sub]", mount).forEach((link) => {
    link.addEventListener("click", () => {
      if (link.dataset.routeOverride) {
        navigate(link.dataset.routeOverride);
      } else {
        navigate(`travel/${link.dataset.travelSub}`);
      }
    });
  });
}

function renderTravelShell(activeSub, renderContent) {
  $("#page-content").innerHTML = `
    <div class="project-spo-layout project-spo-layout--scroll travel-spo-layout">
      ${renderTravelSidebar(activeSub)}
      <main class="project-spo-main travel-spo-main" id="travel-body"></main>
    </div>
  `;
  wireTravelSidebar($("#page-content"));
  renderContent($("#travel-body"));
}

function openTravelTypeChooser() {
  navigate("travel/submit");
}

function renderTravelPage(parts) {
  const sub = parts[0];
  const mode = parts[1] || "";

  if (sub === "edit" && parts[1]) {
    return renderTravelShell("submit", (body) => {
      const editing = (window.AEWTTR.db.travelRequests || []).find((r) => r.id === parts[1]);
      const isLeaveEdit = editing && travelCategory(editing) === "Leave";
      setTopbar(
        isLeaveEdit ? "Edit Leave Request" : "Edit Travel Request",
        isLeaveEdit ? "Update your leave request." : "Update your travel request.",
        ""
      );
      drawSubmitRequest(body, "", parts[1]);
    });
  }

  // Back-compat: these three used to be separate pages. Now they're all
  // the unified "list" page with a view forced to match what the old link
  // meant — "approve" also jumps the status filter to Pending, matching
  // the old Approvals queue's default. "finance" is the Finance Admin
  // landing route — All Travel pre-filtered to requests awaiting a C/O.
  if (sub === "mine" || sub === "all" || sub === "approve" || sub === "finance") {
    if (sub === "approve" && !canAccessTravelApprovals()) {
      navigate("travel/mine");
      toast("Travel approvals are only available to Admins and Finance Admins.", "error");
      return;
    }
    if (sub === "finance" && !canAssignTravelCo()) {
      navigate("travel/mine");
      toast("Awaiting Finance is only available to Admins and Finance Admins.", "error");
      return;
    }
    const listTool = TRAVEL_TOOLS.find((t) => t.key === "list");
    return renderTravelShell("list", (body) => {
      setTopbar(listTool.title, listTool.sub, "");
      if (sub === "approve") window.AEWTTR.state.travelApproveFilter = "Submitted";
      if (sub === "finance") window.AEWTTR.state.travelApproveFilter = "Awaiting Finance";
      drawTravelList(body, sub === "mine" ? "My Travel" : "All Travel");
    });
  }

  if (sub === "events") {
    const evSub = parts[1] || "";
    const evId = parts[2] || "";
    if (evSub === "new") {
      return renderTravelShell("events", (body) => {
        setTopbar("New Team Event", "Create a non-travel calendar event visible to the whole team.", "");
        drawTeamEventForm(body, null);
      });
    }
    if (evSub === "edit" && evId) {
      return renderTravelShell("events", (body) => {
        setTopbar("Edit Team Event", "Update this event's details.", "");
        drawTeamEventForm(body, evId);
      });
    }
    return renderTravelShell("events", (body) => {
      setTopbar("Team Events", "Non-travel team events and calendar items.", "");
      drawTeamEventsList(body);
    });
  }

  const tool = TRAVEL_TOOLS.find((t) => t.key === sub);
  if (!tool) {
    /* The Travel tab lands here. It used to redirect everyone to travel/mine,
       which forces the My Travel view — so an approver always started on their
       own trips rather than the queue they came to work. */
    navigate(defaultTravelListViewForUser() === "All Travel" ? "travel/all" : "travel/mine");
    return;
  }

  renderTravelShell(sub, (body) => {
    setTopbar(tool.title, tool.sub, "");
    if (sub === "submit") return drawSubmitRequest(body, mode);
    if (sub === "list") return drawTravelList(body);
    if (sub === "debrief") return drawSubmitDebrief(body);
    if (sub === "calendar") return drawTravelCalendar(body);
  });
}

/* Requests where the current user is either the requester or listed as one
   of the travelers — so a co-traveler added to someone else's request can
   track its status too, not just the person who submitted it. */
function isCurrentUserOnRequest(r) {
  const db = window.AEWTTR.db;
  const myEmail = normalizeTravelerKey(db.user.email);
  const myName = normalizeTravelerKey(db.user.name);
  if (myEmail && normalizeTravelerKey(r.requesterEmail) === myEmail) return true;
  if (!myEmail && normalizeTravelerKey(r.requester) === myName) return true;
  return (r.travelers || []).some((t) => {
    return (myEmail && normalizeTravelerKey(t.email) === myEmail) || (!myEmail && normalizeTravelerKey(t.name) === myName);
  });
}

function travelRequestIsAllDay(request) {
  if (!request) return true;
  if (request.allDay === false) return false;
  if (request.allDay === true) return true;
  return !(request.startTime || request.endTime);
}

function formatTravelTime(timeStr) {
  if (!timeStr) return "";
  const parts = String(timeStr).split(":");
  const hours = Number(parts[0]);
  const mins = parts[1] || "00";
  if (!Number.isFinite(hours)) return String(timeStr);
  const suffix = hours >= 12 ? "PM" : "AM";
  const h12 = hours % 12 || 12;
  return `${h12}:${mins} ${suffix}`;
}

function formatTravelDateRange(request) {
  const start = fmtDate(request.start);
  const end = fmtDate(request.end);
  if (travelRequestIsAllDay(request)) return `${start} – ${end}`;
  const startTime = formatTravelTime(request.startTime);
  const endTime = formatTravelTime(request.endTime);
  return `${start}${startTime ? ` ${startTime}` : ""} – ${end}${endTime ? ` ${endTime}` : ""}`;
}

function isCurrentUserTravelRequester(r) {
  const db = window.AEWTTR.db;
  const myEmail = normalizeTravelerKey(db.user.email);
  const myName = normalizeTravelerKey(db.user.name);
  if (myEmail && normalizeTravelerKey(r.requesterEmail) === myEmail) return true;
  return !myEmail && normalizeTravelerKey(r.requester) === myName;
}

function isTravelRequestPending(r) {
  return !!(r && r.status === "Submitted");
}

function canRequesterManagePendingTravel(r) {
  return isCurrentUserTravelRequester(r) && isTravelRequestPending(r);
}

function canRequesterEditTravel(r) {
  if (typeof canSubmitForms === "function" && !canSubmitForms()) return false;
  if (!isCurrentUserTravelRequester(r)) return false;
  return r && (r.status === "Submitted" || r.status === "Completed");
}

function setTravelRadioValue(body, name, value) {
  if (!value) return;
  $all(`input[name="${name}"]`, body).forEach((input) => {
    input.checked = input.value === value;
  });
}

function hydrateTravelSubmitForm(body, request) {
  const form = request.engineeringForm || {};
  const set = (id, value) => {
    const node = $("#" + id, body);
    if (node) node.value = value == null ? "" : String(value);
  };
  set("tr-title", request.tripTitle || "");
  set("tr-dest", request.destination || "");
  set("tr-start", request.start || "");
  set("tr-end", request.end || "");
  set("tr-start-time", request.startTime || "");
  set("tr-end-time", request.endTime || "");
  const trAllDay = $("#tr-allday", body);
  if (trAllDay) trAllDay.checked = travelRequestIsAllDay(request);
  set("tr-purpose", request.purpose || "");
  set("tr-impact", request.impactIfNotApproved || "");
  set("tr-alternatives", (request.alternatives || []).join("\n"));
  set("tr-cost", request.cost || "");
  set("tr-notes", request.notes || "");
  set("tr-leave-notes", request.notes || "");
  const leaveConcurrenceCheck = $("#tr-leave-concurrence", body);
  if (leaveConcurrenceCheck) leaveConcurrenceCheck.checked = !!request.requiresConcurrence;
  set("tr-cost-eng", request.cost || "");
  set("tr-notes-eng", request.notes || "");
  const typeSelect = $("#tr-type", body);
  if (typeSelect && request.type) {
    const typeVal = String(request.type || "");
    if (typeVal.startsWith("Other:")) {
      typeSelect.value = "Other";
      set("tr-type-other", typeVal.slice(6).trim());
    } else {
      typeSelect.value = typeVal;
    }
  }

  const primaryName = form.travelerName || request.requester || "";
  set("eng-traveler-name", primaryName);
  set("eng-traveler-email", form.travelerEmail || request.requesterEmail || "");
  set("eng-date-of-request", form.dateOfRequest || new Date().toISOString().slice(0, 10));
  set("eng-phone", form.phoneNumber || "");
  set("eng-purpose", form.purposeOfTdy || request.purpose || "");
  set("eng-start", form.tdyStartDate || request.start || "");
  set("eng-end", form.tdyReturnDate || request.end || "");
  set("eng-start-time", request.startTime || "");
  set("eng-end-time", request.endTime || "");
  const engAllDay = $("#eng-allday", body);
  if (engAllDay) engAllDay.checked = travelRequestIsAllDay(request);
  set("eng-location", form.tdyLocation || request.destination || "");
  set("eng-days", form.numberOfDays || calcTravelDaysBetween(form.tdyStartDate || request.start, form.tdyReturnDate || request.end));
  set("eng-seat-preference", form.seatPreference || "");
  set("eng-flight-from", form.flightFrom || "");
  set("eng-flight-to", form.flightTo || "");
  set("eng-flight-return", form.flightReturn || "");
  set("eng-airport-upgrade", (form.airportTransport && form.airportTransport.upgradeJustification) || "");
  set("eng-airport-vehicle", (form.airportTransport && form.airportTransport.vehicleType) || "");
  set("eng-tdy-upgrade", (form.tdyTransport && form.tdyTransport.upgradeJustification) || "");
  set("eng-tdy-vehicle", (form.tdyTransport && form.tdyTransport.vehicleType) || "");
  set("eng-comments", form.comments || request.notes || "");
  set("eng-bfm", form.bfm || "");
  setTravelRadioValue(body, "eng-on-base", form.onBase);
  setTravelRadioValue(body, "eng-flying", form.flying);
  setTravelRadioValue(body, "eng-airport-transport", form.airportTransport && form.airportTransport.mode);
  setTravelRadioValue(body, "eng-tdy-transport", form.tdyTransport && form.tdyTransport.mode);
  (form.lodging || []).forEach((entry, index) => {
    const i = index + 1;
    set(`eng-lodging-location-${i}`, entry.location || "");
    set(`eng-lodging-checkin-${i}`, entry.checkInDate || "");
    set(`eng-lodging-checkout-${i}`, entry.checkOutDate || "");
    set(`eng-lodging-option1-${i}`, entry.option1 || "");
    set(`eng-lodging-option2-${i}`, entry.option2 || "");
  });
  (form.additionalTravelers || []).forEach((traveler, idx) => {
    const i = idx + 1;
    set(`eng-add-traveler-name-${i}`, traveler.name || "");
    set(`eng-add-traveler-notes-${i}`, traveler.transportationNotes || traveler.transportation || "");
  });
}

async function revokeTravelRequest(trId, onDone) {
  const db = window.AEWTTR.db;
  const request = db.travelRequests.find((r) => r.id === trId);
  if (!request || !canRequesterWithdrawTravel(request)) {
    toast("Only the requester can withdraw a submitted request.", "error");
    return;
  }
  const confirmed = await confirmDialog({
    title: `Withdraw ${trId}?`,
    message: "This withdraws your travel request. You can submit a new one later if plans change.",
    confirmLabel: "Withdraw Request"
  });
  if (!confirmed) return;
  request.status = "Withdrawn";
  request._auditAction = "Withdraw";
  request._auditSummary = `Requester withdrew travel request ${trId}`;
  Repo.save("travelRequest", request);
  notifyRequesterTravelStatusUpdate(request);
  refreshTravelNotifications();
  toast(`${trId} withdrawn`, "success");
  if (typeof onDone === "function") onDone();
}

/* Unified My Travel / All Travel page. All Travel folds in what the old
   standalone Approvals page did — status filter pills plus inline
   Approve/Deny/Assign C/O actions for whoever has that permission — so
   there's one place to browse travel instead of three. forcedView lets a
   "travel/mine" or "travel/all" deep link pin the view for this render
   (and remember it) without permanently overriding the user's last choice
   on every future visit to the plain "travel/list" route. */
function drawTravelList(body, forcedView) {
  if (!window.AEWTTR.state.travelListView) {
    window.AEWTTR.state.travelListView = defaultTravelListViewForUser();
  }
  if (forcedView) window.AEWTTR.state.travelListView = forcedView;
  if (!window.AEWTTR.state.travelApproveFilter) window.AEWTTR.state.travelApproveFilter = "Upcoming";
  if (!window.AEWTTR.state.travelListSearch) window.AEWTTR.state.travelListSearch = "";

  const VIEWS = ["My Travel", "All Travel"];
  const STATUS_FILTERS = ["All", "Upcoming", "Submitted", "Awaiting Finance", "Withdrawn", "Cancelled", "Completed"];

  function render() {
    const db = window.AEWTTR.db;
    const view = window.AEWTTR.state.travelListView;
    const isAll = view === "All Travel";
    const requests = db.travelRequests || [];
    const filterCounts = {
      Submitted: isAll && canAccessTravelApprovals()
        ? requests.filter((r) => r.status === "Submitted" && doesTravelNeedCurrentUserAction(r)).length
        : 0,
      "Awaiting Finance": isAll && canAssignTravelCo()
        ? requests.filter((r) => r.status === "Submitted" && r.chargeObjectStatus === "Pending" && travelCategory(r) !== "Leave").length
        : 0
    };

    body.innerHTML = `
      <div class="toolbar-row travel-list-toolbar" style="margin-bottom:14px;flex-wrap:wrap;gap:10px;">
        <div class="filter-pills" id="travel-view-toggle">
          ${VIEWS.map((v) => `<button class="filter-pill ${view === v ? "active" : ""}" data-travel-view="${v}"${tip(v === "My Travel" ? "Requests you submitted or are traveling on" : "Every travel request across the site")}>${v}</button>`).join("")}
        </div>
        ${isAll ? `<div class="search-box travel-list-search"><i class="bx bx-search"></i><input type="search" id="travel-list-search" placeholder="Search by TR ID, destination, or traveler…" value="${escapeHtml(window.AEWTTR.state.travelListSearch)}"></div>` : ""}
      </div>
      <div class="filter-pills" id="travel-status-filters" style="margin-bottom:14px;">
        ${STATUS_FILTERS.map((f) => {
          const count = filterCounts[f] || 0;
          return `<button type="button" class="filter-pill ${window.AEWTTR.state.travelApproveFilter === f ? "active" : ""}" data-f="${f}"${tip(TRAVEL_APPROVE_FILTER_TIPS[f] || f)}>${f}${count && typeof renderAlertIndicator === "function" ? renderAlertIndicator(count, { variant: "pill", label: `${count} request${count === 1 ? "" : "s"} need your action` }) : ""}</button>`;
        }).join("")}
      </div>
      <div class="aewttr-card">
        <table class="aewttr-table">
          <thead><tr><th>TR ID</th><th>${isAll ? "Travelers" : "Role"}</th><th>Destination</th><th>Dates</th><th>Status</th><th>Debrief</th><th></th></tr></thead>
          <tbody id="travel-list-tbody"></tbody>
        </table>
      </div>
    `;
    renderRows();

    $all("[data-travel-view]", $("#travel-view-toggle", body)).forEach((b) => b.addEventListener("click", () => {
      window.AEWTTR.state.travelListView = b.dataset.travelView;
      render();
    }));
    const searchInput = $("#travel-list-search", body);
    if (searchInput) searchInput.addEventListener("input", (e) => {
      window.AEWTTR.state.travelListSearch = e.target.value;
      renderRows();
    });
    $all("[data-f]", $("#travel-status-filters", body)).forEach((b) => b.addEventListener("click", () => {
      window.AEWTTR.state.travelApproveFilter = b.dataset.f;
      render();
    }));
  }

  function renderRows() {
    const db = window.AEWTTR.db;
    const view = window.AEWTTR.state.travelListView;
    const isAll = view === "All Travel";
    const statusFilt = window.AEWTTR.state.travelApproveFilter || "Upcoming";
    const seenMap = getTravelStatusSeenMap();

    const baseRows = isAll
      ? sortTravelRequests(db.travelRequests || [])
      : sortTravelRequests(db.travelRequests.filter(isCurrentUserOnRequest));

    let rows = applyTravelListFilter(baseRows, statusFilt);

    if (isAll) {
      const q = String(window.AEWTTR.state.travelListSearch || "").trim().toLowerCase();
      if (q) {
        rows = rows.filter((r) => {
          const travelers = (r.travelers || []).map((t) => t.name).join(" ");
          const hay = [r.id, r.destination, r.requester, r.tripTitle, travelers].join(" ").toLowerCase();
          return hay.includes(q);
        });
      }
    }

    $("#travel-list-tbody", body).innerHTML = rows.length ? rows.map((r) => {
      const iAmRequester = normalizeTravelerKey(r.requesterEmail || r.requester) === normalizeTravelerKey(db.user.email || db.user.name);
      const isUnreadStatus = !isAll && isCurrentUserTravelRequester(r) && seenMap[r.id] !== r.status;
      const needsAction = isAll && doesTravelNeedCurrentUserAction(r);
      const travelerNames = (r.travelers && r.travelers.length ? r.travelers.map((t) => t.name).filter(Boolean) : [r.requester]).join(", ");
      const isLeave = travelCategory(r) === "Leave";
      const needsConcurrence = r.customerConcurrenceStatus === "Pending" && (isLeave ? r.requiresConcurrence : true);
      return `
        <tr>
          <td>${r.id}${(isUnreadStatus || needsAction) && typeof renderAlertIndicator === "function" ? renderAlertIndicator(1, { variant: "icon", label: needsAction ? "Needs your action" : "Unread status update" }) : ""}<div style="font-size:11px;color:var(--aewttr-muted);margin-top:4px;">${escapeHtml(r.tripTitle || "Travel Request")}</div></td>
          <td>${isAll ? escapeHtml(travelerNames) : (iAmRequester ? "Requester" : "Traveler")}</td>
          <td>${escapeHtml(r.destination || "—")}</td>
          <td>${formatTravelDateRange(r)}</td>
          <td>${travelStatusBadgeGroup(r)}</td>
          <td>${travelDebriefStatusCell(r)}</td>
          <td class="travel-actions-cell">${travelActionsWrap([
            travelActionBtn("view", r.id, { tip: "Open full request details" }),
            !isAll && canRequesterEditTravel(r) ? travelActionBtn("edit", r.id, { tip: r.chargeObjectStatus === "Assigned" ? "Update this request — changes will generate a new document revision" : "Edit this submitted request" }) : "",
            !isAll && iAmRequester && canRequesterWithdrawTravel(r) ? travelActionBtn("revoke", r.id, { tone: "danger", tip: "Withdraw this request" }) : "",
            !isAll && canCancelTravelRequest(r) && !canRequesterWithdrawTravel(r) ? travelActionBtn("cancel", r.id, { tone: "danger", tip: "Cancel this travel request" }) : "",
            !isAll && r.exportFileUrl ? travelActionBtn("open-export", r.id, { label: "Open Document", icon: "bx-desktop", tip: "Open travel document in Word" }) : "",
            !isAll && r.start ? `<a class="btn-travel-action" href="data:text/calendar;charset=utf-8,${encodeURIComponent(buildTravelIcsContent(r))}" download="${escapeHtml(r.id)}.ics" title="Download .ics" aria-label="Download calendar event"><i class="bx bx-calendar-plus"></i><span>Add to Calendar</span></a>` : "",
            isAll && r.status === "Submitted" && needsConcurrence && canRecordConcurrence() ? travelActionBtn("concur", r.id, { tone: "primary", tip: isLeave ? "Record customer concurrence for this leave request" : "Record customer concurrence" }) : "",
            isAll && r.status === "Submitted" && r.chargeObjectStatus === "Pending" && !isLeave && canAssignTravelCo() ? travelActionBtn("assign-co", r.id, { tone: "primary", tip: "Enter charge object number" }) : "",
            isAll && canAdminCancelTravel(r) ? travelActionBtn("cancel", r.id, { tone: "danger", tip: "Cancel this travel request" }) : "",
            travelDebriefRowButtons(r)
          ])}</td>
        </tr>`;
    }).join("") : `<tr><td colspan="7"><div class="empty-state">${statusFilt === "Upcoming" ? "No upcoming travel." : (isAll ? "No travel requests match your filters." : "No travel requests yet — you'll see anything you submit or are a traveler on here.")}</div></td></tr>`;

    $all("[data-view]", $("#travel-list-tbody", body)).forEach((b) => b.addEventListener("click", () => openTravelDetailModal(b.dataset.view, renderRows)));
    $all("[data-edit]", $("#travel-list-tbody", body)).forEach((b) => b.addEventListener("click", () => navigate(`travel/edit/${b.dataset.edit}`)));
    $all("[data-revoke]", $("#travel-list-tbody", body)).forEach((b) => b.addEventListener("click", () => revokeTravelRequest(b.dataset.revoke, renderRows)));
    $all("[data-cancel]", $("#travel-list-tbody", body)).forEach((b) => b.addEventListener("click", () => cancelTravelRequest(b.dataset.cancel, renderRows)));
    $all("[data-concur]", $("#travel-list-tbody", body)).forEach((b) => b.addEventListener("click", async () => {
      await recordCustomerConcurrence(b.dataset.concur, renderRows);
    }));
    $all("[data-assign-co]", $("#travel-list-tbody", body)).forEach((b) => b.addEventListener("click", () => openCoAssignModal(b.dataset.assignCo, renderRows)));
    $all("[data-open-export]", $("#travel-list-tbody", body)).forEach((b) => b.addEventListener("click", () => {
      const req = db.travelRequests.find((x) => x.id === b.dataset.openExport);
      if (req) openTravelDocByPolicy(req);
    }));
    wireTravelDebriefActions($("#travel-list-tbody", body), renderRows);
    if (!isAll) markAllMyTravelStatusSeen();
  }

  render();
  openTravelFromRouteQuery(body, renderRows);
}

function travelCategory(request) {
  if (!request) return "Travel";
  const mode = String(request.formMode || request.requestType || "").toLowerCase();
  if (mode === "leave" || request.requestType === "Personal Leave" || request.category === "Leave") return "Leave";
  return "Travel";
}

function travelStatusSortRank(status) {
  const ranks = { Submitted: 0, Completed: 1, Withdrawn: 2, Cancelled: 3 };
  return ranks[status] != null ? ranks[status] : 99;
}

function sortTravelRequests(list) {
  return (list || []).slice().sort((a, b) => {
    const rankDiff = travelStatusSortRank(a.status) - travelStatusSortRank(b.status);
    if (rankDiff) return rankDiff;
    return String(b.start || b.id || "").localeCompare(String(a.start || a.id || ""));
  });
}

/* Records disagree on the field name depending on where they were created. */
function travelEndDate(request) {
  if (!request) return "";
  return String(request.end || request.endDate || request.tdyReturnDate || "");
}

/* Upcoming means "still going to happen": not finished, and the end date has not
   passed.

   The previous version had this backwards. It returned false for Approved —
   an approved future trip, which is the most upcoming a request can be — and
   true for Completed trips whose end date was still ahead. So the Upcoming tab
   and every count built on it under-reported real travel and included trips that
   were already closed out. */
function isUpcomingOrCurrentTravel(request) {
  if (!request) return false;
  const status = request.status || "";
  if (["Withdrawn", "Cancelled", "Denied", "Completed"].includes(status)) return false;
  const end = travelEndDate(request);
  if (!end) return true; // no dates entered yet, so it is still ahead
  return end >= new Date().toISOString().slice(0, 10);
}

/* A request that was withdrawn or cancelled is finished with, the same as one
   that completed — so Completed is the closed bucket and shows all three. The
   Withdrawn and Cancelled tabs remain for narrowing to one reason. */
const TRAVEL_CLOSED_STATUSES = ["Completed", "Withdrawn", "Cancelled"];

function applyTravelListFilter(list, statusFilt) {
  const filt = statusFilt || "Upcoming";
  return (list || []).filter((r) => {
    if (filt === "Upcoming") return isUpcomingOrCurrentTravel(r);
    if (filt === "Awaiting Finance") return r.status === "Submitted" && r.chargeObjectStatus === "Pending" && travelCategory(r) !== "Leave";
    if (filt === "Completed") return TRAVEL_CLOSED_STATUSES.includes(r.status);
    if (filt !== "All" && r.status !== filt) return false;
    return true;
  });
}

function travelCalendarTypeKey(request) {
  const category = travelCategory(request);
  if (category === "Leave") return "Leave";
  if (request.contractorTravel || String(request.formMode || "").toLowerCase() === "contractor" || request.requestType === "Contractor Travel") return "Contractor";
  const type = String(request.type || "TDY");
  return type.startsWith("Other") ? "Other" : type;
}

const TRAVEL_CALENDAR_COLORS = {
  Leave: "#7B61A8",
  Contractor: "#C62828",
  TDY: "#1565C0",
  Conference: "#2E7D32",
  Training: "#ED6C02",
  Other: "#00897B"
};

function travelCalendarColor(request) {
  return TRAVEL_CALENDAR_COLORS[travelCalendarTypeKey(request)] || TRAVEL_CALENDAR_COLORS.TDY;
}

/* Calendar event titles need every traveler on a multi-traveler request, not
   just the first one — but a full name list gets unreadable fast on a
   month-grid cell. Last names only, comma-separated; if that's still too
   long for a reasonably-sized event, abbreviate each to its first few
   letters + a dot (e.g. "Swan., Rodri., Chen"). */
function travelCalendarTravelerLabel(travelers, fallbackName) {
  const names = (travelers && travelers.length ? travelers : [{ name: fallbackName }])
    .map((t) => String((t && t.name) || "").trim())
    .filter(Boolean);
  if (!names.length) return fallbackName || "Travel";
  const lastNames = names.map((n) => lastNameOf(n) || n);
  const MAX_LEN = 30;
  let joined = lastNames.join(", ");
  if (joined.length <= MAX_LEN || lastNames.length === 1) return joined;
  const abbreviate = (name) => (name.length > 4 ? `${name.slice(0, 4)}.` : name);
  joined = lastNames.map(abbreviate).join(", ");
  if (joined.length <= MAX_LEN) return joined;
  const shown = [];
  let total = 0;
  for (const name of lastNames) {
    const piece = abbreviate(name);
    const nextTotal = total + piece.length + (shown.length ? 2 : 0);
    if (nextTotal > MAX_LEN && shown.length) break;
    shown.push(piece);
    total = nextTotal;
  }
  const remaining = lastNames.length - shown.length;
  return shown.join(", ") + (remaining > 0 ? ` +${remaining} more` : "");
}

function travelDeepLinkActionUrl(route, trId) {
  if (typeof pulseAppRouteUrl === "function") return pulseAppRouteUrl(route, trId ? { tr: trId } : undefined);
  return typeof pulseAppUrl === "function" ? pulseAppUrl() : "";
}

function openTravelFromRouteQuery(body, redrawFn) {
  const query = (typeof currentRoute === "function" ? currentRoute().query : {}) || {};
  const pending = typeof consumePendingRouteAction === "function" ? consumePendingRouteAction() : null;
  const trId = (pending && pending.tr) || query.tr;
  if (!trId) return;
  const request = (window.AEWTTR.db.travelRequests || []).find((r) => r.id === trId);
  if (!request) return;
  if (typeof openTravelDetailModal === "function") openTravelDetailModal(trId, redrawFn);
  else if (typeof redrawFn === "function") redrawFn();
}

function doesTravelNeedCurrentUserAction(request) {
  if (!request || request.status !== "Submitted") return false;
  const isLeave = travelCategory(request) === "Leave";
  if (canRecordConcurrence() && request.customerConcurrenceStatus === "Pending" && (isLeave ? request.requiresConcurrence : true)) return true;
  if (!isLeave && canAssignTravelCo() && request.chargeObjectStatus === "Pending") return true;
  return false;
}

function getTravelMetrics() {
  const requests = (window.AEWTTR.db && window.AEWTTR.db.travelRequests) || [];
  const needsMyAction = requests.filter(doesTravelNeedCurrentUserAction);
  const statusUpdateCount = typeof window.getTravelStatusUpdateCount === "function"
    ? window.getTravelStatusUpdateCount()
    : 0;
  return {
    needsMyActionCount: needsMyAction.length,
    statusUpdateCount,
    totalAlertCount: needsMyAction.length + statusUpdateCount,
    pendingConcurrenceCount: canRecordConcurrence() ? requests.filter((r) => r.status === "Submitted" && r.customerConcurrenceStatus === "Pending").length : 0,
    pendingCoCount: canAssignTravelCo() ? requests.filter((r) => r.status === "Submitted" && r.chargeObjectStatus === "Pending" && travelCategory(r) !== "Leave").length : 0
  };
}

window.getTravelMetrics = getTravelMetrics;

/* Multiple people can be on one travel request, and each files their own
   debrief — so debriefs are keyed by (trId, traveler), not just trId. All
   the debrief-lookup helpers below work off the plural list; the singular
   getDebriefForTravel stays for the one call site that just wants "is
   there any debrief at all" without caring who filed it. */
function getDebriefsForTravel(trId) {
  return (window.AEWTTR.db.debriefs || []).filter((d) => d.trId === trId);
}
function getDebriefForTravel(trId) {
  return getDebriefsForTravel(trId)[0] || null;
}
function travelRequestTravelerRoster(request) {
  const list = (request.travelers && request.travelers.length)
    ? request.travelers
    : [{ name: request.requester, email: request.requesterEmail }];
  return list.filter((t) => t && t.name);
}
function isCurrentUserOnRequest(request) {
  const myKey = normalizeTravelerKey(window.AEWTTR.db.user.email || window.AEWTTR.db.user.name);
  if (!myKey) return false;
  if (normalizeTravelerKey(request.requesterEmail || request.requester) === myKey) return true;
  const travelers = travelRequestTravelerRoster(request);
  return travelers.some((t) => normalizeTravelerKey(t.email || t.name) === myKey);
}
function getMyDebriefForTravel(trId) {
  const db = window.AEWTTR.db;
  const myKey = normalizeTravelerKey(db.user.email || db.user.name);
  if (!myKey) return null;
  return getDebriefsForTravel(trId).find((d) => normalizeTravelerKey(d.travelerEmail || d.travelerName) === myKey) || null;
}

function travelNeedsDebrief(request) {
  if (!request) return false;
  if (travelCategory(request) === "Leave") return false;
  return (request.status === "Submitted" || request.status === "Completed") && request.chargeObjectStatus === "Assigned";
}

function travelDebriefStatusCell(request) {
  if (!travelNeedsDebrief(request)) return `<span style="color:var(--aewttr-muted);">—</span>`;
  const travelers = travelRequestTravelerRoster(request);
  const debriefs = getDebriefsForTravel(request.id);
  if (travelers.length <= 1) {
    return debriefs.length
      ? `<span class="travel-debrief-badge is-filed"><i class="bx bx-check-circle"></i> Filed</span>`
      : `<span class="travel-debrief-badge is-pending"><i class="bx bx-time-five"></i> Needed</span>`;
  }
  const allFiled = debriefs.length >= travelers.length;
  return `<span class="travel-debrief-badge ${allFiled ? "is-filed" : "is-pending"}"><i class="bx ${allFiled ? "bx-check-circle" : "bx-time-five"}"></i> ${debriefs.length}/${travelers.length} filed</span>`;
}

function travelDebriefRowButtons(request) {
  if (!travelNeedsDebrief(request)) return "";
  const debriefs = getDebriefsForTravel(request.id);
  const myDebrief = getMyDebriefForTravel(request.id);
  const buttons = [];
  if (isCurrentUserOnRequest(request) && !myDebrief) {
    buttons.push(travelActionBtn("debrief", request.id, { tone: "primary", tip: "Submit a post-trip debrief for this travel" }));
  }
  if (debriefs.length) {
    buttons.push(travelActionBtn("view-debrief", request.id, {
      label: debriefs.length > 1 ? `Debriefs (${debriefs.length})` : "View Debrief",
      tip: debriefs.length > 1 ? "View and switch between filed debriefs" : `Open debrief ${debriefs[0].id}`
    }));
  }
  return buttons.join("");
}

function wireTravelDebriefActions(scope, redrawFn) {
  $all("[data-debrief]", scope).forEach((btn) => btn.addEventListener("click", () => {
    navigate("travel/debrief", { tr: btn.dataset.debrief });
  }));
  $all("[data-view-debrief]", scope).forEach((btn) => btn.addEventListener("click", () => {
    openDebriefDetailModal(btn.dataset.viewDebrief, redrawFn);
  }));
}

function refreshTravelNotifications() {
  if (typeof renderNav === "function") renderNav();
  if (typeof refreshUserNotifications === "function") refreshUserNotifications();
}

const TRAVEL_STATUS_SEEN_KEY = "aewttr_travel_status_seen";

function getTravelStatusSeenMap() {
  try { return JSON.parse(lsGet(TRAVEL_STATUS_SEEN_KEY) || "{}"); } catch (e) { return {}; }
}

function markTravelRequestStatusSeen(requestId, status) {
  const map = getTravelStatusSeenMap();
  map[requestId] = status;
  lsSet(TRAVEL_STATUS_SEEN_KEY, JSON.stringify(map));
  refreshUserNotifications();
}

function markAllMyTravelStatusSeen() {
  const db = window.AEWTTR.db;
  const map = getTravelStatusSeenMap();
  (db.travelRequests || []).filter(isCurrentUserTravelRequester).forEach((r) => { map[r.id] = r.status; });
  lsSet(TRAVEL_STATUS_SEEN_KEY, JSON.stringify(map));
  refreshUserNotifications();
}

function getTravelStatusUpdateCount() {
  const db = window.AEWTTR.db;
  const map = getTravelStatusSeenMap();
  return (db.travelRequests || []).filter((r) => isCurrentUserTravelRequester(r) && map[r.id] !== r.status).length;
}

window.getTravelStatusUpdateCount = getTravelStatusUpdateCount;

function travelStatusChangePreview(request) {
  if (!request) return "Status changed";
  const isLeave = travelCategory(request) === "Leave";
  switch (request.status) {
    case "Submitted":
      if (isLeave && !request.requiresConcurrence) return "Your leave notice was submitted and added to the calendar.";
      if (isLeave) return "Your leave request is awaiting customer concurrence and has been added to the calendar.";
      return "Your travel request was submitted.";
    case "Withdrawn":
      return "Request withdrawn by the requester.";
    case "Cancelled":
      if (isLeave) return "Leave request was cancelled.";
      return request.cancelledByAdmin
        ? `Travel cancelled by ${request.cancelledBy || "an admin"}.`
        : "Travel request was cancelled.";
    case "Completed":
      return "Travel completed.";
    default:
      return `Status is now ${request.status}`;
  }
}

function travelStatusBadgeGroup(request) {
  if (!request) return "";
  const isLeave = travelCategory(request) === "Leave";
  const parts = [];
  parts.push(statusPill(request.status));
  if (request.status === "Submitted") {
    if (!isLeave) {
      const concStatus = request.customerConcurrenceStatus || "Pending";
      const concColor = concStatus === "Concurred" ? "var(--aewttr-green,#1a8a4a)" : "var(--aewttr-amber,#b45309)";
      parts.push(`<span class="travel-badge" style="color:${concColor};border-color:${concColor};" title="Customer Concurrence: ${concStatus}">Concurrence: ${concStatus}</span>`);
      const coStatus = request.chargeObjectStatus || "Pending";
      const coColor = coStatus === "Assigned" ? "var(--aewttr-green,#1a8a4a)" : "var(--aewttr-muted,#888)";
      const coLabel = coStatus === "Assigned" ? `C/O: ${escapeHtml(request.chargeObject || "Assigned")}` : "C/O: Pending";
      parts.push(`<span class="travel-badge" style="color:${coColor};border-color:${coColor};" title="Charge Object: ${coStatus}">${coLabel}</span>`);
    } else if (request.requiresConcurrence) {
      const concStatus = request.customerConcurrenceStatus || "Pending";
      const concColor = concStatus === "Concurred" ? "var(--aewttr-green,#1a8a4a)" : "var(--aewttr-amber,#b45309)";
      parts.push(`<span class="travel-badge" style="color:${concColor};border-color:${concColor};" title="Customer Concurrence: ${concStatus}">Concurrence: ${concStatus}</span>`);
    }
  }
  return parts.join("<br>");
}

function getTravelNotificationEvents() {
  const db = window.AEWTTR.db;
  const events = [];
  const requests = db.travelRequests || [];
  requests.filter(doesTravelNeedCurrentUserAction).forEach((request) => {
    const isLeave = travelCategory(request) === "Leave";
    const needsConcurrence = request.customerConcurrenceStatus === "Pending" && canRecordConcurrence() && (isLeave ? request.requiresConcurrence : true);
    const needsCo = !isLeave && request.chargeObjectStatus === "Pending" && canAssignTravelCo();
    events.push({
      id: `travel-action-${request.id}`,
      route: "travel/all",
      queryTr: request.id,
      icon: needsConcurrence ? "bx-check-shield" : "bx-receipt",
      tone: needsConcurrence ? "blue" : "purple",
      title: request.id,
      preview: needsConcurrence
        ? (isLeave
          ? `${request.requester} · leave · needs customer concurrence`
          : `${request.requester} → ${request.destination} · needs customer concurrence`)
        : `${request.destination} · assign charge object`,
      time: fmtRelativeTime(request.end),
      category: "Travel",
      sortKey: `2-${request.id}`
    });
  });
  const seenMap = getTravelStatusSeenMap();
  requests
    .filter((request) => isCurrentUserTravelRequester(request) && seenMap[request.id] !== request.status)
    .forEach((request) => {
      events.push({
        id: `travel-status-${request.id}`,
        route: "travel/mine",
        queryTr: request.id,
        icon: request.status === "Submitted" ? "bx-send" : request.status === "Cancelled" ? "bx-x-circle" : "bx-info-circle",
        tone: request.status === "Submitted" ? "green" : request.status === "Cancelled" ? "red" : "amber",
        title: `${request.id} updated`,
        preview: travelStatusChangePreview(request),
        time: fmtRelativeTime(request.end),
        category: "Travel",
        sortKey: `1-${request.id}`
      });
    });
  return events;
}

window.getTravelNotificationEvents = getTravelNotificationEvents;

function canRequesterWithdrawTravel(r) {
  return isCurrentUserTravelRequester(r) && r.status === "Submitted";
}

function canRequesterCancelApprovedTravel(r) {
  return canRequesterWithdrawTravel(r);
}

function canRequesterCancelLeave(r) {
  if (!isCurrentUserTravelRequester(r) || travelCategory(r) !== "Leave") return false;
  return r.status === "Submitted" || r.status === "Completed";
}

/* Statuses an approver may set directly. Deliberately excludes the
   charge-object and concurrence states, which are set by their own flows and
   would be misleading to change by hand. */
const TRAVEL_ADMIN_STATUSES = ["Submitted", "Approved", "Denied", "Completed", "Withdrawn", "Cancelled"];

/* Correcting a record after the fact is an approver action — someone travelled
   but the request was never moved on, or a status was set in error. */
function canAdminSetTravelStatus(r) {
  return !!r && canApproveTravelRequests();
}

/* Deleting removes the record outright, so it is limited to approvers and is
   always confirmed. Withdraw/Cancel remain the normal route; this is for
   duplicates and test entries that should not stay in the list. */
function canAdminDeleteTravel(r) {
  return !!r && canApproveTravelRequests();
}

/* Only worth nudging when the trip is done and at least one traveller still
   owes a debrief. */
function travelDebriefOutstanding(r) {
  if (!r || travelCategory(r) === "Leave") return false;
  if (!["Completed", "Approved"].includes(r.status)) return false;
  const end = r.endDate || r.tdyReturnDate || "";
  if (end && end > new Date().toISOString().slice(0, 10)) return false;
  const filed = typeof getDebriefsForTravel === "function" ? getDebriefsForTravel(r.id) : [];
  const travellers = (r.travelers && r.travelers.length) ? r.travelers.length : 1;
  return filed.length < travellers;
}

async function notifyTravelDebriefNudge(req) {
  const recipients = [];
  if (req.requesterEmail) recipients.push(req.requesterEmail);
  (req.travelers || []).forEach((t) => { if (t && t.email) recipients.push(t.email); });
  const to = Array.from(new Set(recipients.filter(Boolean)));
  if (!isSharePointMode() || !to.length) return false;
  try {
    await notifyUsers({
      to,
      subject: `PULSE Travel: debrief still needed for ${req.id}`,
      area: "Travel",
      kind: "info",
      preview: `A travel debrief is still outstanding for ${req.id}${req.destination ? ` — ${req.destination}` : ""}. Please file it so the trip can be closed out.`,
      facts: travelNotificationFacts(req),
      actionUrl: travelDeepLinkActionUrl("travel/mine", req.id),
      actionTitle: "File debrief"
    });
    return true;
  } catch (e) {
    console.warn("PULSE: debrief nudge failed.", e);
    return false;
  }
}

function canAdminCancelTravel(r) {
  if (!r || !canApproveTravelRequests()) return false;
  return r.status === "Submitted";
}

function canCancelTravelRequest(r) {
  return canRequesterWithdrawTravel(r) || canAdminCancelTravel(r) || canRequesterCancelLeave(r);
}

async function cancelTravelRequest(trId, onDone) {
  const db = window.AEWTTR.db;
  const request = db.travelRequests.find((r) => r.id === trId);
  const asAdmin = canAdminCancelTravel(request);
  const asRequester = canRequesterCancelApprovedTravel(request);
  const asLeaveOwner = canRequesterCancelLeave(request);
  if (!request || (!asAdmin && !asRequester && !asLeaveOwner)) {
    toast("You can't cancel this request.", "error");
    return;
  }
  const isLeave = travelCategory(request) === "Leave";
  const isWithdraw = asRequester && !asAdmin && !isLeave;
  const confirmed = await confirmDialog({
    title: isWithdraw ? `Withdraw ${trId}?` : `Cancel ${trId}?`,
    message: asAdmin && !asRequester && !asLeaveOwner
      ? "This cancels the travel request. The requester will be notified."
      : isLeave
        ? "This cancels your leave request. You can submit a new one if plans change."
        : "This withdraws your travel request.",
    confirmLabel: isWithdraw ? "Withdraw Request" : (isLeave ? "Cancel Leave" : "Cancel Travel")
  });
  if (!confirmed) return;
  const priorStatus = request.status;
  request.status = isWithdraw ? "Withdrawn" : "Cancelled";
  request.cancelledAt = new Date().toISOString();
  request.cancelledBy = (db.user && db.user.name) || "Admin";
  request.cancelledByAdmin = !!asAdmin;
  request._auditAction = asAdmin ? "Cancel (Admin)" : (isWithdraw ? "Withdraw" : "Cancel");
  request._auditSummary = asAdmin
    ? `Admin cancelled travel ${trId} (was ${priorStatus})`
    : isLeave
      ? `Requester cancelled leave ${trId} (was ${priorStatus})`
      : `Requester withdrew travel ${trId}`;
  Repo.save("travelRequest", request);
  notifyTravelCancellation(request, { priorStatus, byAdmin: asAdmin });
  refreshTravelNotifications();
  toast(`${trId} cancelled`, "success");
  if (typeof onDone === "function") onDone();
}

function calcTravelDaysBetween(start, end) {
  if (!start || !end) return "";
  const s = new Date(start + "T00:00:00");
  const e = new Date(end + "T00:00:00");
  if (isNaN(s) || isNaN(e)) return "";
  return String(Math.max(1, Math.round((e - s) / 86400000) + 1));
}

function wireTravelDaysAutofill(body) {
  const update = () => {
    const daysNode = $("#eng-days", body);
    if (!daysNode) return;
    const days = calcTravelDaysBetween(travelWizardReadValue(body, "eng-start"), travelWizardReadValue(body, "eng-end"));
    if (days) daysNode.value = days;
  };
  const startNode = $("#eng-start", body);
  const endNode = $("#eng-end", body);
  if (startNode) startNode.addEventListener("change", update);
  if (endNode) endNode.addEventListener("change", update);
}

function wireTravelTypeOther(body) {
  const typeSelect = $("#tr-type", body);
  const otherRow = $("#tr-type-other-row", body);
  const otherInput = $("#tr-type-other", body);
  if (!typeSelect || !otherRow) return;
  const sync = () => {
    const isOther = typeSelect.value === "Other";
    otherRow.hidden = !isOther;
    if (otherInput) otherInput.required = isOther;
  };
  typeSelect.addEventListener("change", sync);
  sync();
}

/* Shared by both the Standard/Contractor/Leave trip panel (tr-*) and the
   Engineering basics panel (eng-*) — "All day" is checked by default so the
   common case (no times needed) stays a single date field; unchecking
   reveals departure/return time inputs for that same panel. */
function wireAllDayToggle(body, prefix) {
  const checkbox = $(`#${prefix}-allday`, body);
  const timeRow = $(`#${prefix}-time-row`, body);
  if (!checkbox || !timeRow) return;
  const sync = () => { timeRow.hidden = checkbox.checked; };
  checkbox.addEventListener("change", sync);
  sync();
}

/* Leave requests don't need a trip title (auto-generated from the
   requester's name), a destination/travel-type, or a cost estimate — those
   fields belong to actual travel, not time away. Hide them for Leave
   instead of maintaining a whole separate panel, since the dates/
   notes rows are otherwise identical to Standard. Travelers stay hidden
   too (leave is always just the requester). */
function updateTripPanelForMode(body, formMode) {
  const titleRow = $("#tr-title-row", body);
  const destTypeRow = $("#tr-dest-type-row", body);
  const costRow = $("#tr-cost-row", body);
  const budgetTitle = $("#tr-budget-title", body);
  const leaveNotesRow = $("#tr-leave-notes-row", body);
  const leaveRequesterRow = $("#tr-leave-requester-row", body);
  const isLeave = formMode === "Leave";
  if (titleRow) titleRow.hidden = isLeave;
  if (destTypeRow) destTypeRow.hidden = isLeave;
  if (costRow) costRow.hidden = isLeave;
  if (leaveNotesRow) leaveNotesRow.hidden = !isLeave;
  if (leaveRequesterRow) leaveRequesterRow.hidden = !isLeave;
  const leaveConcurrenceRow = $("#tr-leave-concurrence-row", body);
  if (leaveConcurrenceRow) leaveConcurrenceRow.hidden = !isLeave;
  if (budgetTitle) budgetTitle.textContent = isLeave ? "Coverage & notes" : "Estimated costs";
  const aboutTitle = body.querySelector('[data-tw-step="trip"] .travel-wizard-step-card-title');
  if (aboutTitle) aboutTitle.textContent = isLeave ? "Leave details" : "About the trip";
}

function wirePrimaryTravelerPicker(body, travelers, onChange) {
  let selected = currentTravelerRecord();
  const mount = $("#eng-primary-traveler-selected", body);
  const input = $("#eng-primary-traveler-input", body);
  const suggestions = $("#eng-primary-traveler-suggestions", body);
  const hiddenName = $("#eng-traveler-name", body);
  const hiddenEmail = $("#eng-traveler-email", body);
  const directory = getTravelerDirectory();
  if (!mount || !input || !hiddenName) return { refresh() {}, getSelected() { return selected; } };

  function syncHidden() {
    hiddenName.value = selected.name || "";
    if (hiddenEmail) hiddenEmail.value = selected.email || "";
    mount.innerHTML = selected.name
      ? `<span class="traveler-chip"><span>${escapeHtml(selected.name)}${selected.email ? ` <small>${escapeHtml(selected.email)}</small>` : ""}</span><button type="button" id="eng-primary-clear" aria-label="Clear">&times;</button></span>`
      : `<span class="traveler-empty">Select primary traveler</span>`;
    const clearBtn = $("#eng-primary-clear", mount);
    if (clearBtn) clearBtn.addEventListener("click", () => {
      selected = { name: "", email: "", source: "cleared" };
      syncHidden();
      if (typeof onChange === "function") onChange(null);
    });
  }

  function setSelected(entry) {
    if (!entry || !entry.name) return;
    selected = { ...entry };
    syncHidden();
    input.value = "";
    suggestions.innerHTML = "";
    const idx = travelers.findIndex((t) => normalizeTravelerKey(t.email || t.name) === normalizeTravelerKey(entry.email || entry.name));
    if (idx >= 0) travelers.splice(idx, 1);
    travelers.unshift({ ...entry });
    if (typeof onChange === "function") onChange(selected);
  }

  function drawSuggestions() {
    const query = normalizeTravelerKey(input.value);
    if (!query) { suggestions.innerHTML = ""; return; }
    const matches = directory.filter((entry) => {
      return normalizeTravelerKey(entry.name).includes(query) || normalizeTravelerKey(entry.email).includes(query);
    }).slice(0, 6);
    suggestions.innerHTML = matches.length ? matches.map((entry, index) => `
      <button type="button" class="traveler-suggestion" data-primary-suggestion="${index}">
        <strong>${escapeHtml(entry.name)}</strong>
        <span>${escapeHtml(entry.email || "No email on file")}</span>
      </button>`).join("") : `<div class="traveler-suggestion-empty">No matching site users.</div>`;
    $all("[data-primary-suggestion]", suggestions).forEach((button) => {
      button.addEventListener("click", () => setSelected(matches[Number(button.dataset.primarySuggestion)]));
    });
  }

  input.addEventListener("input", drawSuggestions);
  input.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    const query = input.value.trim();
    if (!query) return;
    const exact = directory.find((entry) => normalizeTravelerKey(entry.email) === normalizeTravelerKey(query) || normalizeTravelerKey(entry.name) === normalizeTravelerKey(query));
    if (exact) setSelected(exact);
  });

  const preset = travelers[0] || currentTravelerRecord();
  if (preset && preset.name) setSelected(preset);
  else syncHidden();
  return { refresh: syncHidden, getSelected() { return selected; }, setSelected };
}

/* Travel and Leave requests share one list, so the id prefix itself carries
   the type — "LE-" for Leave, "TR-" for everything else (Standard/
   Contractor/Engineering) — while each prefix keeps its own counter so
   Leave numbering doesn't skip around Travel's. */
function nextTrId(formMode) {
  const db = window.AEWTTR.db;
  const prefix = formMode === "Leave" ? "LE" : "TR";
  const max = db.travelRequests.reduce((m, r) => {
    const parts = String(r.id || "").split("-");
    if (parts[0] !== prefix) return m;
    return Math.max(m, parseInt(parts[1], 10) || 0);
  }, 0);
  return `${prefix}-${String(max + 1).padStart(4, "0")}`;
}

/* SharePoint/AD display names on this tenant are usually "Lastname,
   Firstname MI" — trust the comma over word-splitting so it always wins.
   Names with no comma sometimes carry a trailing branch/status code (e.g.
   "John Smith USA") that a naive last-token split mistakes for the
   surname, so strip a known suffix code before falling back to it. */
const NAME_SUFFIX_CODES = new Set(["USA", "USN", "USAF", "USMC", "USSF", "USCG", "CIV", "CTR", "NG"]);
function lastNameOf(fullName) {
  const raw = String(fullName || "").trim();
  if (!raw) return "";
  const commaIndex = raw.indexOf(",");
  if (commaIndex > 0) return raw.slice(0, commaIndex).trim();
  const parts = raw.split(/\s+/).filter(Boolean);
  if (!parts.length) return "";
  let end = parts.length;
  if (end > 1 && NAME_SUFFIX_CODES.has(parts[end - 1].toUpperCase())) end -= 1;
  return parts[end - 1] || "";
}

/* PDF exports want a normal-order human name ("John A Smith"), not the raw
   SharePoint/AD display-name format this tenant uses ("Smith, John A" —
   see lastNameOf()'s comment above for why). Only flips names that
   actually contain the comma; anything already in normal order (a
   manually-typed guest/contractor name, etc.) passes through untouched. */
function displayNameToActualName(fullName) {
  const raw = String(fullName || "").trim();
  if (!raw) return "";
  const commaIndex = raw.indexOf(",");
  if (commaIndex <= 0) return raw;
  const last = raw.slice(0, commaIndex).trim();
  const rest = raw.slice(commaIndex + 1).trim();
  return rest ? `${rest} ${last}` : last;
}

function currentTravelerRecord() {
  const user = window.AEWTTR.db.user || {};
  return {
    id: user.id || `manual-${normalizeTravelerKey(user.email || user.name || "current-user")}`,
    name: user.name || "Current User",
    email: user.email || "",
    source: "current-user"
  };
}

function normalizeTravelerKey(value) {
  return String(value || "").trim().toLowerCase();
}

function defaultTravelers() {
  const me = currentTravelerRecord();
  return me.name ? [me] : [];
}

function getTravelerDirectory() {
  const directory = typeof getMemberDirectory === "function"
    ? getMemberDirectory()
    : (window.AEWTTR.db.members || []).map((member) => ({
        id: member.id,
        name: member.name,
        email: member.email || ""
      }));
  const siteUsers = directory.map((entry) => ({
    id: entry.id,
    name: entry.name,
    email: entry.email || "",
    source: "site-user"
  })).filter((entry) => entry.name);

  // Include contractor persons (type:"person") from all project rosters
  const db = window.AEWTTR.db;
  const seen = new Set(siteUsers.map((e) => (e.email || e.name).toLowerCase()));
  const contractors = [];
  if (db.projectPeople) {
    Object.values(db.projectPeople).forEach((roster) => {
      (roster || []).forEach((p) => {
        if (p.type !== "person" || !p.label) return;
        const key = (p.email || p.label).toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        contractors.push({ name: p.label, email: p.email || "", company: p.company || "", source: "contractor" });
      });
    });
  }
  return siteUsers.concat(contractors);
}

function renderTravelerChips(mount, travelers, onRemove) {
  if (typeof renderPersonChips === "function") {
    renderPersonChips(mount, travelers, onRemove, { emptyLabel: "No travelers selected." });
    return;
  }
  mount.innerHTML = travelers.length ? travelers.map((traveler, index) => `
    <span class="traveler-chip${traveler && traveler.type === "group" ? " traveler-chip--group" : ""}">
      <span>${escapeHtml(traveler.name)}${traveler.email ? ` <small>${escapeHtml(traveler.email)}</small>` : ""}</span>
      <button type="button" data-remove-traveler="${index}" aria-label="Remove traveler">&times;</button>
    </span>
  `).join("") : `<span class="traveler-empty">No travelers selected.</span>`;
  $all("[data-remove-traveler]", mount).forEach((button) => button.addEventListener("click", () => onRemove(Number(button.dataset.removeTraveler))));
}

function wireTravelerPicker(body, travelers, ids) {
  ids = ids || {
    mount: "tr-travelers-selected",
    input: "tr-travelers-input",
    suggestions: "tr-travelers-suggestions"
  };
  /* Share the same searchable people picker (chips + groups) used app-wide.
     Groups stay as a unit in the traveler list; expand at notify/print time. */
  if (typeof wirePeoplePicker === "function") {
    return wirePeoplePicker(body, travelers, ids, {
      allowManualEmail: true,
      includeGroups: true,
      expandGroups: false,
      emptyLabel: "No travelers selected."
    });
  }
  let onChange = null;
  const mount = $("#" + ids.mount, body);
  const input = $("#" + ids.input, body);
  const suggestions = $("#" + ids.suggestions, body);
  const directory = getTravelerDirectory();
  if (!mount || !input || !suggestions) return { refresh() {} };

  function travelerExists(entry) {
    const emailKey = normalizeTravelerKey(entry.email);
    const nameKey = normalizeTravelerKey(entry.name);
    return travelers.some((traveler) => {
      return (emailKey && normalizeTravelerKey(traveler.email) === emailKey)
        || (!emailKey && nameKey && normalizeTravelerKey(traveler.name) === nameKey);
    });
  }

  function addTraveler(entry) {
    if (!entry || !entry.name || travelerExists(entry)) return;
    travelers.push(entry);
    renderTravelerChips(mount, travelers, removeTravelerAt);
    input.value = "";
    suggestions.innerHTML = "";
    if (typeof onChange === "function") onChange();
  }

  function removeTravelerAt(index) {
    travelers.splice(index, 1);
    renderTravelerChips(mount, travelers, removeTravelerAt);
    if (typeof onChange === "function") onChange();
  }

  function manualTravelerFromText(text) {
    const raw = String(text || "").trim();
    if (!raw) return null;
    const emailMatch = raw.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
    if (!emailMatch) return null;
    const email = emailMatch[0];
    const name = raw.replace(email, "").replace(/[<>]/g, "").trim() || email;
    return {
      id: `manual-${normalizeTravelerKey(email)}`,
      name,
      email,
      source: "manual"
    };
  }

  function drawSuggestions() {
    const query = normalizeTravelerKey(input.value);
    if (!query) {
      suggestions.innerHTML = "";
      return;
    }
    const matches = directory.filter((entry) => {
      if (travelerExists(entry)) return false;
      return normalizeTravelerKey(entry.name).includes(query) || normalizeTravelerKey(entry.email).includes(query);
    }).slice(0, 6);
    suggestions.innerHTML = matches.length ? matches.map((entry, index) => `
      <button type="button" class="traveler-suggestion" data-suggestion-index="${index}">
        <strong>${escapeHtml(entry.name)}</strong>
        <span>${escapeHtml(entry.email || "No email on file")}</span>
      </button>
    `).join("") : `<div class="traveler-suggestion-empty">No match — press Enter to add <strong>${escapeHtml(input.value.trim())}</strong> by name.</div>`;
    $all("[data-suggestion-index]", suggestions).forEach((button) => {
      button.addEventListener("click", () => addTraveler(matches[Number(button.dataset.suggestionIndex)]));
    });
  }

  input.addEventListener("input", drawSuggestions);
  input.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    const query = input.value.trim();
    if (!query) return;
    const exact = directory.find((entry) => {
      return !travelerExists(entry) && (
        normalizeTravelerKey(entry.email) === normalizeTravelerKey(query)
        || normalizeTravelerKey(entry.name) === normalizeTravelerKey(query)
      );
    });
    if (exact) {
      addTraveler(exact);
      return;
    }
    const manual = manualTravelerFromText(query);
    if (manual) {
      addTraveler(manual);
      return;
    }
    // Allow contractor/name-only entries — no email required
    addTraveler({ id: `manual-name-${normalizeTravelerKey(query)}`, name: query, source: "manual" });
  });

  function refresh() {
    renderTravelerChips(mount, travelers, removeTravelerAt);
    drawSuggestions();
  }

  renderTravelerChips(mount, travelers, removeTravelerAt);
  return {
    refresh,
    setOnChange(fn) { onChange = fn; }
  };
}

function travelChoiceGroup(name, options, selectedValue) {
  return options.map((option) => `
    <label class="travel-choice">
      <input type="radio" name="${escapeHtml(name)}" value="${escapeHtml(option)}" ${selectedValue === option ? "checked" : ""}>
      <span>${escapeHtml(option)}</span>
    </label>
  `).join("");
}

function collectEngineeringForm(body, travelers) {
  const value = (id) => {
    const node = $("#" + id, body);
    return node ? node.value.trim() : "";
  };
  const checkedValue = (name) => {
    const node = $(`input[name="${name}"]:checked`, body);
    return node ? node.value : "";
  };
  return {
    dateOfRequest: value("eng-date-of-request") || new Date().toISOString().slice(0, 10),
    travelerName: value("eng-traveler-name"),
    travelerEmail: value("eng-traveler-email"),
    phoneNumber: value("eng-phone"),
    purposeOfTdy: value("eng-purpose"),
    onBase: checkedValue("eng-on-base"),
    tdyLocation: value("eng-location"),
    tdyStartDate: value("eng-start"),
    tdyReturnDate: value("eng-end"),
    numberOfDays: value("eng-days") || calcTravelDaysBetween(value("eng-start"), value("eng-end")),
    additionalTravelers: [1, 2, 3, 4, 5].map((i) => ({
      name: value(`eng-add-traveler-name-${i}`) || "",
      email: "",
      transportationNotes: value(`eng-add-traveler-notes-${i}`) || ""
    })).filter((t) => t.name),
    flying: checkedValue("eng-flying"),
    seatPreference: value("eng-seat-preference"),
    flightFrom: value("eng-flight-from"),
    flightTo: value("eng-flight-to"),
    flightReturn: value("eng-flight-return"),
    airportTransport: {
      mode: checkedValue("eng-airport-transport"),
      upgradeJustification: value("eng-airport-upgrade"),
      vehicleType: value("eng-airport-vehicle")
    },
    tdyTransport: {
      mode: checkedValue("eng-tdy-transport"),
      upgradeJustification: value("eng-tdy-upgrade"),
      vehicleType: value("eng-tdy-vehicle")
    },
    lodging: [1, 2, 3].map((index) => ({
      location: value(`eng-lodging-location-${index}`),
      checkInDate: value(`eng-lodging-checkin-${index}`),
      checkOutDate: value(`eng-lodging-checkout-${index}`),
      option1: value(`eng-lodging-option1-${index}`),
      option2: value(`eng-lodging-option2-${index}`)
    })),
    comments: value("eng-comments"),
    travelCo: "",
    bfm: value("eng-bfm")
  };
}

function buildPrintableTravelHtml(request) {
  const travelers = request.travelers || [];
  const travelerList = travelers.map((traveler) => `${traveler.name}${traveler.email ? ` (${traveler.email})` : ""}`).join(", ");
  if (request.formMode === "Engineering") {
    const form = request.engineeringForm || {};
    const lodgingRows = (form.lodging || []).map((entry, index) => `
      <div class="eng-lodging-block">
        <div class="eng-lodging-row eng-lodging-row--meta">
          <span class="eng-label">Location ${index + 1}:</span>
          <span class="eng-sublabel">City, State</span>
          <span class="eng-line">${escapeHtml(entry.location || "")}</span>
          <span class="eng-sublabel">Check-In Date</span>
          <span class="eng-line eng-line--date">${escapeHtml(entry.checkInDate || "")}</span>
          <span class="eng-sublabel">Check-Out Date</span>
          <span class="eng-line eng-line--date">${escapeHtml(entry.checkOutDate || "")}</span>
        </div>
        <div class="eng-lodging-row eng-lodging-row--option">
          <span class="eng-label">Option 1:</span>
          <span class="eng-line eng-line--wide">${escapeHtml(entry.option1 || "")}</span>
        </div>
        <div class="eng-lodging-row eng-lodging-row--option">
          <span class="eng-label">Option 2:</span>
          <span class="eng-line eng-line--wide">${escapeHtml(entry.option2 || "")}</span>
        </div>
      </div>
    `).join("");
    return `
      <html><head><title>${escapeHtml(request.id)} Engineering Travel Request</title><style>
        body{font-family:"Times New Roman",serif;margin:18px;color:#111}
        .eng-sheet{border:1.5px solid #333;padding:14px;max-width:980px;margin:0 auto}
        .eng-title{font-size:24px;text-align:center;margin:0 0 14px}
        .eng-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}
        .eng-box{border:1px solid #444;padding:10px;min-height:0}
        .eng-line-row{display:grid;grid-template-columns:132px minmax(0,1fr);gap:8px;align-items:center;margin-bottom:8px;font-size:13px}
        .eng-line-row.pair{grid-template-columns:132px minmax(0,1fr) 132px minmax(0,1fr)}
        .eng-line-row.footer-row{grid-template-columns:72px minmax(0,1fr) 48px minmax(0,1fr);max-width:640px}
        .eng-label{font-weight:700;white-space:nowrap}
        .eng-sublabel{font-size:12px;font-weight:700}
        .eng-line{min-height:22px;border:1px solid #444;padding:4px 6px;word-break:break-word}
        .eng-line--date{min-width:88px}
        .eng-line--wide{grid-column:2 / -1}
        .eng-section{border:1px solid #444;padding:10px;margin-bottom:14px;page-break-inside:avoid}
        .eng-section h3{margin:0 0 10px;font-size:16px}
        .eng-columns{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}
        .eng-columns .eng-section{margin-bottom:0}
        .eng-check{display:inline-block;min-width:18px;border:1px solid #444;text-align:center;margin-right:6px}
        .eng-comments{min-height:72px;border:1px solid #444;padding:8px;white-space:pre-wrap}
        .eng-lodging-block{border-top:1px solid #ccc;padding-top:10px;margin-top:10px}
        .eng-lodging-block:first-child{border-top:none;padding-top:0;margin-top:0}
        .eng-lodging-row{display:grid;gap:8px;align-items:center;margin-bottom:8px;font-size:13px}
        .eng-lodging-row--meta{grid-template-columns:78px 72px minmax(0,1.4fr) 84px 92px 92px 92px}
        .eng-lodging-row--option{grid-template-columns:78px minmax(0,1fr)}
        @media print { body{margin:0;padding:0} .eng-sheet{border:none;max-width:none} }
      </style></head><body>
        <div class="eng-sheet">
          <div class="eng-title">TTSD/CL Travel Request Form v2023</div>
          <div class="eng-grid">
            <div class="eng-box">
              <div class="eng-line-row"><span class="eng-label">Date of Request:</span><span class="eng-line">${escapeHtml(form.dateOfRequest || "")}</span></div>
              <div class="eng-line-row"><span class="eng-label">Traveler Name:</span><span class="eng-line">${escapeHtml(form.travelerName || request.requester || "")}</span></div>
              <div class="eng-line-row"><span class="eng-label">Phone Number:</span><span class="eng-line">${escapeHtml(form.phoneNumber || "")}</span></div>
              <div class="eng-line-row"><span class="eng-label">Purpose of TDY:</span><span class="eng-line">${escapeHtml(form.purposeOfTdy || request.purpose || "")}</span></div>
              <div class="eng-line-row pair"><span class="eng-label">TDY Start Date:</span><span class="eng-line">${escapeHtml(form.tdyStartDate || request.start || "")}</span><span class="eng-label">TDY Return Date:</span><span class="eng-line">${escapeHtml(form.tdyReturnDate || request.end || "")}</span></div>
            </div>
            <div class="eng-box">
              <div class="eng-line-row"><span class="eng-label">TDY Location on base:</span><span>${form.onBase === "Yes" ? "[X] Yes [ ] No" : form.onBase === "No" ? "[ ] Yes [X] No" : "[ ] Yes [ ] No"}</span></div>
              <div class="eng-line-row"><span class="eng-label">TDY Location:</span><span class="eng-line">${escapeHtml(form.tdyLocation || request.destination || "")}</span></div>
              <div class="eng-line-row"><span class="eng-label">Number of Days:</span><span class="eng-line short">${escapeHtml(form.numberOfDays || "")}</span></div>
            </div>
          </div>
          <div class="eng-section">
            <h3>Additional Travelers Information</h3>
            ${(form.additionalTravelers || []).length ? form.additionalTravelers.map((traveler, index) => `<div class="eng-line-row"><span class="eng-label">${index + 1}.</span><span class="eng-line wide">${escapeHtml(`${traveler.name}${traveler.email ? ` (${traveler.email})` : ""}`)}</span></div>`).join("") : `<div class="eng-line-row"><span class="eng-line wide">${escapeHtml(travelerList)}</span></div>`}
          </div>
          <div class="eng-columns">
            <div class="eng-section">
              <h3>Flying</h3>
              <div class="eng-line-row"><span>${form.flying === "Yes" ? "[X] Yes [ ] No" : form.flying === "No" ? "[ ] Yes [X] No" : "[ ] Yes [ ] No"}</span></div>
              <div class="eng-line-row"><span class="eng-label">Seat Preference:</span><span class="eng-line">${escapeHtml(form.seatPreference || "")}</span></div>
              <div class="eng-line-row"><span class="eng-label">From:</span><span class="eng-line">${escapeHtml(form.flightFrom || "")}</span></div>
              <div class="eng-line-row"><span class="eng-label">To:</span><span class="eng-line">${escapeHtml(form.flightTo || "")}</span></div>
              <div class="eng-line-row"><span class="eng-label">Return:</span><span class="eng-line">${escapeHtml(form.flightReturn || "")}</span></div>
            </div>
            <div class="eng-section">
              <h3>Transportation To/From Airport</h3>
              <div class="eng-line-row"><span class="eng-label">Mode:</span><span class="eng-line">${escapeHtml((form.airportTransport && form.airportTransport.mode) || "")}</span></div>
              <div class="eng-line-row"><span class="eng-label">Upgrade Justification:</span><span class="eng-line">${escapeHtml((form.airportTransport && form.airportTransport.upgradeJustification) || "")}</span></div>
              <div class="eng-line-row"><span class="eng-label">Type of Vehicle:</span><span class="eng-line">${escapeHtml((form.airportTransport && form.airportTransport.vehicleType) || "")}</span></div>
            </div>
            <div class="eng-section">
              <h3>Transportation at TDY</h3>
              <div class="eng-line-row"><span class="eng-label">Mode:</span><span class="eng-line">${escapeHtml((form.tdyTransport && form.tdyTransport.mode) || "")}</span></div>
              <div class="eng-line-row"><span class="eng-label">Upgrade Justification:</span><span class="eng-line">${escapeHtml((form.tdyTransport && form.tdyTransport.upgradeJustification) || "")}</span></div>
              <div class="eng-line-row"><span class="eng-label">Type of Vehicle:</span><span class="eng-line">${escapeHtml((form.tdyTransport && form.tdyTransport.vehicleType) || "")}</span></div>
            </div>
          </div>
          <div class="eng-section">
            <h3>Lodging</h3>
            ${lodgingRows}
          </div>
          <div class="eng-section">
            <h3>Comments</h3>
            <div class="eng-comments">${escapeHtml(form.comments || request.notes || "")}</div>
          </div>
          <div class="eng-line-row footer-row"><span class="eng-label">Travel C/O:</span><span class="eng-line">${escapeHtml(request.chargeObject || form.travelCo || "")}</span><span class="eng-label">BFM:</span><span class="eng-line">${escapeHtml(form.bfm || "")}</span></div>
        </div>
      </body></html>
    `;
  }

  return `
    <html><head><title>${escapeHtml(request.id)} Standard Travel Request</title><style>
      body{font-family:Arial,sans-serif;margin:28px;color:#111}
      .sheet{max-width:900px;margin:0 auto;border:1px solid #d0d7de;padding:28px}
      h1{margin:0 0 18px;font-size:28px}
      .meta{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;margin-bottom:20px}
      .card{border:1px solid #cfd8e3;padding:12px 14px}
      .label{font-size:11px;text-transform:uppercase;color:#57606a;margin-bottom:6px}
      .value{font-size:15px;line-height:1.5}
      .full{margin-bottom:14px}
      @media print { body{margin:0} .sheet{border:none} }
    </style></head><body>
      <div class="sheet">
        <h1>${escapeHtml(request.tripTitle || "Travel Request")}</h1>
        <div class="meta">
          <div class="card"><div class="label">Request ID</div><div class="value">${escapeHtml(request.id)}</div></div>
          <div class="card"><div class="label">Requester</div><div class="value">${escapeHtml(request.requester)}${request.requesterEmail ? `<br>${escapeHtml(request.requesterEmail)}` : ""}</div></div>
          <div class="card"><div class="label">Destination</div><div class="value">${escapeHtml(request.destination || "")}</div></div>
          <div class="card"><div class="label">Travel Dates</div><div class="value">${escapeHtml(request.start || "")} to ${escapeHtml(request.end || "")}</div></div>
        </div>
        <div class="card full"><div class="label">Travelers</div><div class="value">${escapeHtml(travelerList || request.requester || "")}</div></div>
        <div class="card full"><div class="label">Event Purpose</div><div class="value">${escapeHtml(request.purpose || "")}</div></div>
        <div class="card full"><div class="label">Impact If Not Approved</div><div class="value">${escapeHtml(request.impactIfNotApproved || "")}</div></div>
        <div class="card full"><div class="label">Alternatives / Virtual Substitute</div><div class="value">${escapeHtml((request.alternatives || []).join(", "))}</div></div>
      </div>
    </body></html>
  `;
}

function pdfSafeText(value) {
  return String(value == null ? "" : value)
    .replace(/[^\x20-\x7E\n]/g, "?")
    .replace(/\\/g, "\\\\")
    .replace(/\(/g, "\\(")
    .replace(/\)/g, "\\)");
}

function wrapPdfText(label, value, maxChars) {
  const source = `${label}${value == null || value === "" ? "Not provided" : String(value)}`.replace(/\r/g, "");
  const words = source.split(/\s+/).filter(Boolean);
  const lines = [];
  let current = "";
  words.forEach((word) => {
    const next = current ? `${current} ${word}` : word;
    if (next.length > maxChars && current) {
      lines.push(current);
      current = word;
    } else {
      current = next;
    }
  });
  if (current) lines.push(current);
  return lines.length ? lines : [label];
}

function buildTravelPdfLines(request) {
  const travelers = (request.travelers || []).map((traveler) => `${displayNameToActualName(traveler.name)}${traveler.email ? ` (${traveler.email})` : ""}`).join(", ");
  const lines = [];
  if (request.formMode === "Engineering") {
    const form = request.engineeringForm || {};
    lines.push("TTSD/CL Travel Request Form v2023");
    lines.push("");
    lines.push(...wrapPdfText("Request ID: ", request.id, 92));
    lines.push(...wrapPdfText("Traveler Name: ", displayNameToActualName(form.travelerName || request.requester), 92));
    lines.push(...wrapPdfText("Requester Email: ", request.requesterEmail || "", 92));
    lines.push(...wrapPdfText("Date of Request: ", form.dateOfRequest || "", 92));
    lines.push(...wrapPdfText("Phone Number: ", form.phoneNumber || "", 92));
    lines.push(...wrapPdfText("Purpose of TDY: ", form.purposeOfTdy || request.purpose || "", 92));
    lines.push(...wrapPdfText("TDY Location on Base: ", form.onBase || "", 92));
    lines.push(...wrapPdfText("TDY Location: ", form.tdyLocation || request.destination || "", 92));
    lines.push(...wrapPdfText("TDY Start Date: ", form.tdyStartDate || request.start || "", 92));
    lines.push(...wrapPdfText("TDY Return Date: ", form.tdyReturnDate || request.end || "", 92));
    lines.push(...wrapPdfText("Number of Days: ", form.numberOfDays || "", 92));
    lines.push(...wrapPdfText("All Travelers: ", travelers || displayNameToActualName(request.requester) || "", 92));
    lines.push(...wrapPdfText("Flying: ", form.flying || "", 92));
    lines.push(...wrapPdfText("Seat Preference: ", form.seatPreference || "", 92));
    lines.push(...wrapPdfText("Flight From: ", form.flightFrom || "", 92));
    lines.push(...wrapPdfText("Flight To: ", form.flightTo || "", 92));
    lines.push(...wrapPdfText("Flight Return: ", form.flightReturn || "", 92));
    lines.push(...wrapPdfText("Airport Transport: ", (form.airportTransport && form.airportTransport.mode) || "", 92));
    lines.push(...wrapPdfText("Airport Upgrade Justification: ", (form.airportTransport && form.airportTransport.upgradeJustification) || "", 92));
    lines.push(...wrapPdfText("Airport Vehicle Type: ", (form.airportTransport && form.airportTransport.vehicleType) || "", 92));
    lines.push(...wrapPdfText("TDY Transport: ", (form.tdyTransport && form.tdyTransport.mode) || "", 92));
    lines.push(...wrapPdfText("TDY Upgrade Justification: ", (form.tdyTransport && form.tdyTransport.upgradeJustification) || "", 92));
    lines.push(...wrapPdfText("TDY Vehicle Type: ", (form.tdyTransport && form.tdyTransport.vehicleType) || "", 92));
    (form.lodging || []).forEach((entry, index) => {
      lines.push(...wrapPdfText(`Lodging ${index + 1} Location: `, entry.location || "", 92));
      lines.push(...wrapPdfText(`Lodging ${index + 1} Check-In: `, entry.checkInDate || "", 92));
      lines.push(...wrapPdfText(`Lodging ${index + 1} Check-Out: `, entry.checkOutDate || "", 92));
      lines.push(...wrapPdfText(`Lodging ${index + 1} Option 1: `, entry.option1 || "", 92));
      lines.push(...wrapPdfText(`Lodging ${index + 1} Option 2: `, entry.option2 || "", 92));
    });
    lines.push(...wrapPdfText("Comments: ", form.comments || request.notes || "", 92));
    lines.push(...wrapPdfText("Travel C/O: ", request.chargeObject || form.travelCo || "", 92));
    lines.push(...wrapPdfText("BFM: ", form.bfm || "", 92));
  } else {
    lines.push("PULSE Standard Travel Request");
    lines.push("=".repeat(42));
    lines.push("");
    lines.push(...wrapPdfText("Request ID: ", request.id, 92));
    lines.push(...wrapPdfText("Trip Title: ", request.tripTitle || "", 92));
    lines.push(...wrapPdfText("Requester: ", displayNameToActualName(request.requester) || "", 92));
    lines.push(...wrapPdfText("Requester Email: ", request.requesterEmail || "", 92));
    lines.push(...wrapPdfText("Travelers: ", travelers || displayNameToActualName(request.requester) || "", 92));
    lines.push(...wrapPdfText("Destination: ", request.destination || "", 92));
    lines.push(...wrapPdfText("Departure Date: ", request.start || "", 92));
    lines.push(...wrapPdfText("Return Date: ", request.end || "", 92));
    lines.push(...wrapPdfText("Travel Type: ", request.type || "", 92));
    lines.push(...wrapPdfText("Event Purpose: ", request.purpose || "", 92));
    lines.push(...wrapPdfText("Impact If Not Approved: ", request.impactIfNotApproved || "", 92));
    lines.push(...wrapPdfText("Alternatives / Virtual Substitute: ", (request.alternatives || []).join(", "), 92));
    lines.push(...wrapPdfText("Estimated Cost: ", request.cost ? `$${request.cost}` : "", 92));
    lines.push(...wrapPdfText("Status: ", request.status || "", 92));
    lines.push(...wrapPdfText("Charge Object (C/O): ", request.chargeObject || "", 92));
    lines.push(...wrapPdfText("Notes: ", request.notes || "", 92));
  }
  return lines;
}

function pdfWrappedLinesForWidth(text, maxWidth, fontSize, maxLines) {
  const source = String(text == null || text === "" ? "" : text).replace(/\r/g, "").replace(/\n+/g, " ").trim();
  if (!source) return [""];
  const words = source.split(/\s+/);
  const lines = [];
  let current = "";
  const measure = (value) => pdfSafeText(value).length * fontSize * 0.48;
  words.forEach((word) => {
    const next = current ? `${current} ${word}` : word;
    if (measure(next) > maxWidth && current) {
      lines.push(current);
      current = word;
    } else {
      current = next;
    }
  });
  if (current) lines.push(current);
  if (!maxLines || lines.length <= maxLines) return lines;
  const limited = lines.slice(0, maxLines);
  limited[maxLines - 1] = `${limited[maxLines - 1].slice(0, Math.max(0, limited[maxLines - 1].length - 3))}...`;
  return limited;
}

function pdfMeasureApprox(text, fontSize) {
  return pdfSafeText(text).length * fontSize * 0.48;
}

function pdfNormalizeFieldText(value, maxChars) {
  const clean = String(value == null ? "" : value)
    .replace(/\r/g, " ")
    .replace(/\n+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!maxChars || clean.length <= maxChars) return clean;
  return `${clean.slice(0, Math.max(0, maxChars - 3)).trim()}...`;
}

function pdfClampLine(line, maxWidth, fontSize) {
  const safeLine = String(line || "");
  if (pdfMeasureApprox(safeLine, fontSize) <= maxWidth) return safeLine;
  let trimmed = safeLine;
  while (trimmed.length > 1 && pdfMeasureApprox(`${trimmed}...`, fontSize) > maxWidth) {
    trimmed = trimmed.slice(0, -1).trimEnd();
  }
  return `${trimmed}...`;
}

function pdfFitTextToBox(text, width, height, opts) {
  opts = opts || {};
  const minFontSize = opts.minFontSize || 6.2;
  const maxFontSize = opts.size || 8.5;
  const lineGap = opts.lineGap == null ? 2 : opts.lineGap;
  const explicitMaxLines = opts.maxLines || Infinity;
  const normalized = pdfNormalizeFieldText(text || "", opts.maxChars || 500);
  if (!normalized) {
    return { fontSize: maxFontSize, lines: [""], lineGap };
  }

  let chosen = { fontSize: minFontSize, lines: [pdfClampLine(normalized, width, minFontSize)], lineGap };
  for (let fontSize = maxFontSize; fontSize >= minFontSize; fontSize -= 0.2) {
    const byHeight = Math.max(1, Math.floor((height + lineGap) / (fontSize + lineGap)));
    const allowedLines = Math.max(1, Math.min(explicitMaxLines, byHeight));
    const wrapped = pdfWrappedLinesForWidth(normalized, width, fontSize, allowedLines);
    const neededHeight = wrapped.length * fontSize + Math.max(0, wrapped.length - 1) * lineGap;
    if (neededHeight <= height + 0.1) {
      const clamped = wrapped.map((line, index) => index === wrapped.length - 1 ? pdfClampLine(line, width, fontSize) : line);
      return { fontSize, lines: clamped, lineGap };
    }
    chosen = {
      fontSize,
      lines: wrapped.map((line, index) => index === wrapped.length - 1 ? pdfClampLine(line, width, fontSize) : line),
      lineGap
    };
  }
  return chosen;
}

function createEngineeringPdfLayout(pageHeight) {
  const ops = [];
  const fields = [];
  const pageWidth = 612;
  const topToPdfY = (top) => pageHeight - top;
  const escape = pdfSafeText;

  const rectCmd = (x, top, w, h) => {
    const y = topToPdfY(top + h);
    return `${x.toFixed(2)} ${y.toFixed(2)} ${w.toFixed(2)} ${h.toFixed(2)} re S`;
  };
  const lineCmd = (x1, top1, x2, top2) => {
    return `${x1.toFixed(2)} ${topToPdfY(top1).toFixed(2)} m ${x2.toFixed(2)} ${topToPdfY(top2).toFixed(2)} l S`;
  };
  const overlaps = (a, b, pad) => {
    pad = pad == null ? 1 : pad;
    return a.x < b.x + b.w - pad
      && b.x < a.x + a.w - pad
      && a.top < b.top + b.h - pad
      && b.top < a.top + a.h - pad;
  };
  const findClearTop = (x, w, h, preferredTop, maxBump) => {
    let top = preferredTop;
    const limit = preferredTop + (maxBump == null ? 80 : maxBump);
    while (top <= limit) {
      const candidate = { x, top, w, h };
      const hit = fields.some((existing) => overlaps(existing, candidate));
      if (!hit) return top;
      top += 2;
    }
    return preferredTop;
  };

  const layout = {
    ops,
    fields,
    pageWidth,
    pushRect(x, top, w, h, meta) {
      const safeTop = findClearTop(x, w, h, top);
      fields.push({ x, top: safeTop, w, h, meta: meta || "field" });
      // Only draw the box outline for checkboxes to keep the filled layout clean
      if (meta === "checkbox") {
        ops.push(rectCmd(x, safeTop, w, h));
      }
      return safeTop;
    },
    pushLine(x1, top1, x2, top2) {
      ops.push(lineCmd(x1, top1, x2, top2));
    },
    pushText(text, x, top, size, font) {
      const safe = escape(text);
      if (!safe) return;
      const y = topToPdfY(top) - size;
      ops.push(`BT /${font || "F1"} ${size} Tf 1 0 0 1 ${x.toFixed(2)} ${y.toFixed(2)} Tm (${safe}) Tj ET`);
    },
    pushCenteredText(text, centerX, top, size, font) {
      const safe = escape(text);
      if (!safe) return;
      const width = safe.length * size * 0.48;
      layout.pushText(text, centerX - (width / 2), top, size, font);
    },
    pushTextBox(text, x, top, w, h, opts) {
      opts = opts || {};
      const boxTop = layout.pushRect(x, top, w, h, opts.meta || "textbox");
      const padX = opts.padX == null ? 4 : opts.padX;
      const padTop = opts.padTop == null ? 3 : opts.padTop;
      const innerWidth = Math.max(16, w - padX * 2);
      const innerHeight = Math.max(6, h - padTop * 2);
      const fit = pdfFitTextToBox(text || "", innerWidth, innerHeight, {
        size: opts.size || 8,
        minFontSize: opts.minFontSize || 6,
        maxLines: opts.maxLines || 1,
        maxChars: opts.maxChars || 500,
        lineGap: opts.lineGap == null ? 1.5 : opts.lineGap
      });
      fit.lines.forEach((line, index) => {
        const rowTop = boxTop + padTop + index * (fit.fontSize + fit.lineGap);
        if (rowTop + fit.fontSize > boxTop + h - 0.5) return;
        layout.pushText(line, x + padX, rowTop, fit.fontSize, opts.font || "F1");
      });
      return boxTop;
    },
    pushSection(x, top, w, h) {
      ops.push(rectCmd(x, top, w, h));
      return top;
    },
    pushCheckbox(x, top, size, checked) {
      const boxTop = layout.pushRect(x, top, size, size, "checkbox");
      if (checked) {
        const pad = 2;
        ops.push(lineCmd(x + pad, boxTop + pad, x + size - pad, boxTop + size - pad));
        ops.push(lineCmd(x + size - pad, boxTop + pad, x + pad, boxTop + size - pad));
      }
    },
    labeledField(label, x, top, labelW, fieldW, h, value, opts) {
      opts = opts || {};
      layout.pushText(label, x, top + 1, opts.labelSize || 8.5, "F1");
      layout.pushTextBox(value, x + labelW, top, fieldW, h, {
        size: opts.valueSize || 8.25,
        maxLines: opts.maxLines || 1,
        maxChars: opts.maxChars || 120,
        minFontSize: opts.minFontSize || 6
      });
      return top + h;
    },
    stackedField(label, x, top, fieldW, h, value, opts) {
      opts = opts || {};
      const labelGap = opts.labelGap == null ? 7 : opts.labelGap;
      layout.pushText(label, x, top, opts.labelSize || 7.4, opts.labelFont || "F1");
      layout.pushTextBox(value, x, top + labelGap, fieldW, h, {
        size: opts.valueSize || 7.2,
        maxLines: opts.maxLines || 1,
        maxChars: opts.maxChars || 80,
        minFontSize: opts.minFontSize || 6,
        padTop: opts.padTop == null ? 3 : opts.padTop
      });
      return top + labelGap + h;
    },
    bottom() {
      if (!fields.length) return 0;
      return Math.max(...fields.map((f) => f.top + f.h));
    },
    validate() {
      for (let i = 0; i < fields.length; i++) {
        for (let j = i + 1; j < fields.length; j++) {
          if (overlaps(fields[i], fields[j], 0.5)) {
            throw new Error(`Engineering PDF layout overlap between ${fields[i].meta} and ${fields[j].meta} at (${fields[i].x},${fields[i].top}) and (${fields[j].x},${fields[j].top})`);
          }
        }
      }
    }
  };

  return layout;
}

function createEngineeringTravelPdfBlob(request) {
  const pageWidth = 612;
  const pageHeight = 792;
  const form = request.engineeringForm || {};
  const primaryTraveler = displayNameToActualName(form.travelerName || request.requester || "");
  const additionalTravelers = (form.additionalTravelers || []).slice(0, 5);
  const allTravelers = (request.travelers || []).map((traveler) => displayNameToActualName(traveler.name || "")).filter(Boolean);
  const travelerLines = [];
  for (let index = 0; index < 5; index++) {
    if (additionalTravelers[index]) {
      const traveler = additionalTravelers[index];
      travelerLines.push(`${displayNameToActualName(traveler.name)}${traveler.transportation ? ` — ${traveler.transportation}` : ""}`);
    } else if (!index && allTravelers.length > 1) {
      travelerLines.push(allTravelers.slice(1).join(", "));
    } else {
      travelerLines.push("");
    }
  }

  const formatPdfDate = (value) => {
    if (!value) return "";
    const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return match ? `${match[2]}/${match[3]}/${match[1]}` : String(value);
  };
  const daysValue = form.numberOfDays || (() => {
    if (!(form.tdyStartDate || request.start) || !(form.tdyReturnDate || request.end)) return "";
    const start = new Date((form.tdyStartDate || request.start) + "T00:00:00");
    const end = new Date((form.tdyReturnDate || request.end) + "T00:00:00");
    if (isNaN(start) || isNaN(end)) return "";
    const diff = Math.round((end - start) / 86400000);
    return String(Math.max(1, diff + 1));
  })();

  const layout = createEngineeringPdfLayout(pageHeight);
  const fullWidth = 520;
  const left = 46;
  const outerLeft = 36;
  const outerTop = 78;
  const fieldH = 11;
  const rowGap = 4;

  layout.pushCenteredText("TTSD/CL Travel Request Form v2023", pageWidth / 2, 92, 14, "F3");

  const topFormTop = 106;
  layout.pushSection(left, topFormTop, fullWidth, 140);
  layout.pushLine(266, topFormTop, 266, topFormTop + 140);

  layout.labeledField("Date of Request:", 52, 116, 82, 102, 18, formatPdfDate(form.dateOfRequest || new Date().toISOString().slice(0, 10)), { maxLines: 1, maxChars: 14 });
  layout.labeledField("Traveler Name:", 52, 140, 82, 168, 18, primaryTraveler, { maxLines: 1, maxChars: 42, minFontSize: 6.6 });
  layout.labeledField("Phone Number:", 52, 164, 82, 168, 18, form.phoneNumber || "", { maxLines: 1, maxChars: 24 });
  layout.labeledField("Purpose of TDY:", 52, 188, 82, 168, 18, form.purposeOfTdy || request.purpose || "", { maxLines: 2, maxChars: 52, minFontSize: 6.2 });
  layout.stackedField("TDY Start Date:", 52, 214, 88, 14, formatPdfDate(form.tdyStartDate || request.start || ""), { labelSize: 7.2, valueSize: 7.4, maxChars: 14, padTop: 3, labelGap: 9 });
  layout.stackedField("TDY Return Date:", 154, 214, 88, 14, formatPdfDate(form.tdyReturnDate || request.end || ""), { labelSize: 7.2, valueSize: 7.4, maxChars: 14, padTop: 3, labelGap: 9 });

  layout.pushText("TDY Location on base:", 304, 140, 8.2, "F1");
  layout.pushText("Yes", 442, 140, 8.2, "F1");
  layout.pushCheckbox(462, 134, 18, form.onBase === "Yes");
  layout.pushText("No", 496, 140, 8.2, "F1");
  layout.pushCheckbox(514, 134, 18, form.onBase === "No");
  layout.labeledField("TDY Location:", 304, 164, 74, 170, 18, form.tdyLocation || request.destination || "", { maxLines: 2, maxChars: 38, minFontSize: 6.2 });
  layout.labeledField("Number of Days:", 304, 214, 88, 92, 14, daysValue, { maxLines: 1, maxChars: 6, valueSize: 7.6 });

  const travelerTop = 260;
  layout.pushSection(left, travelerTop, fullWidth, 86);
  layout.pushText("Additional Travelers Information: Name and Transportation", 52, travelerTop + 10, 9, "F2");
  for (let index = 0; index < 5; index++) {
    const rowTop = travelerTop + 24 + index * 12;
    layout.pushText(String(index + 1), 58, rowTop + 2, 8, "F1");
    layout.pushTextBox(travelerLines[index], 70, rowTop, 494, 12, { size: 7.2, maxLines: 1, padTop: 2, padX: 3, minFontSize: 5.8, maxChars: 110, lineGap: 1, meta: `traveler-${index}` });
  }

  const transportTop = 360;
  const transportHeight = 142;
  const colWidth = 168;
  const colGap = 8;
  const col1 = left;
  const col2 = col1 + colWidth + colGap;
  const col3 = col2 + colWidth + colGap;
  layout.pushSection(col1, transportTop, colWidth, transportHeight);
  layout.pushSection(col2, transportTop, colWidth, transportHeight);
  layout.pushSection(col3, transportTop, colWidth, transportHeight);

  layout.pushText("Flying", col1 + 8, transportTop + 10, 10.5, "F2");
  layout.pushText("Yes", col1 + 64, transportTop + 12, 8.5, "F1");
  layout.pushCheckbox(col1 + 88, transportTop + 6, 18, form.flying === "Yes");
  layout.pushText("No", col1 + 114, transportTop + 12, 8.5, "F1");
  layout.pushCheckbox(col1 + 134, transportTop + 6, 18, form.flying === "No");
  layout.labeledField("Seat Preference:", col1 + 8, transportTop + 30, 80, 72, 16, form.seatPreference || "", { maxLines: 1, maxChars: 22, valueSize: 7.4 });
  layout.stackedField("From (Airport Code/Time):", col1 + 8, transportTop + 54, 152, 12, form.flightFrom || "", { labelSize: 6.8, valueSize: 7, minFontSize: 5.6, maxChars: 28, padTop: 3, labelGap: 9 });
  layout.stackedField("To (Airport Code/Time):", col1 + 8, transportTop + 82, 152, 12, form.flightTo || "", { labelSize: 6.8, valueSize: 7, minFontSize: 5.6, maxChars: 28, padTop: 3, labelGap: 9 });
  layout.stackedField("Return from (Airport Code/Time):", col1 + 8, transportTop + 110, 152, 12, form.flightReturn || "", { labelSize: 6.8, valueSize: 7, minFontSize: 5.6, maxChars: 28, padTop: 3, labelGap: 9 });

  layout.pushText("Transportation To /From Airport:", col2 + 8, transportTop + 10, 9.2, "F2");
  ["POV", "Rental", "Passenger"].forEach((option, index) => {
    const rowTop = transportTop + 28 + index * 16;
    layout.pushText(option, col2 + 8, rowTop, 8.5, "F1");
    layout.pushCheckbox(col2 + 138, rowTop - 4, 14, (form.airportTransport && form.airportTransport.mode) === option);
  });
  layout.pushText("If Rental, please fill in the following:", col2 + 8, transportTop + 80, 7, "F1");
  layout.stackedField("Upgrade Required Justification:", col2 + 8, transportTop + 88, 152, 12, (form.airportTransport && form.airportTransport.upgradeJustification) || "", { labelSize: 6.8, valueSize: 7, minFontSize: 5.6, maxChars: 34, padTop: 3, labelGap: 9 });
  layout.stackedField("Type of Vehicle", col2 + 8, transportTop + 114, 152, 12, (form.airportTransport && form.airportTransport.vehicleType) || "", { labelSize: 6.8, valueSize: 7, minFontSize: 5.6, maxChars: 30, padTop: 3, labelGap: 9 });

  layout.pushText("Transportation at TDY:", col3 + 8, transportTop + 10, 9.2, "F2");
  ["Rental", "Passenger", "Public Transportation"].forEach((option, index) => {
    const rowTop = transportTop + 28 + index * 16;
    layout.pushText(option, col3 + 8, rowTop, option === "Public Transportation" ? 7.2 : 8.5, "F1");
    layout.pushCheckbox(col3 + 138, rowTop - 4, 14, (form.tdyTransport && form.tdyTransport.mode) === option);
  });
  layout.pushText("If Rental, please fill in the following:", col3 + 8, transportTop + 80, 7, "F1");
  layout.stackedField("Upgrade Required Justification:", col3 + 8, transportTop + 88, 152, 12, (form.tdyTransport && form.tdyTransport.upgradeJustification) || "", { labelSize: 6.8, valueSize: 7, minFontSize: 5.6, maxChars: 34, padTop: 3, labelGap: 9 });
  layout.stackedField("Type of Vehicle", col3 + 8, transportTop + 114, 152, 12, (form.tdyTransport && form.tdyTransport.vehicleType) || "", { labelSize: 6.8, valueSize: 7, minFontSize: 5.6, maxChars: 30, padTop: 3, labelGap: 9 });

  const rawLodging = (form.lodging || [{}, {}, {}]).slice(0, 3);
  const lodgingHasData = (entry) => [entry.location, entry.checkInDate, entry.checkOutDate, entry.option1, entry.option2]
    .some((value) => String(value || "").trim());
  let lodgingEntries = rawLodging.filter(lodgingHasData);
  while (lodgingEntries.length < 3) lodgingEntries.push(rawLodging[lodgingEntries.length] || {});
  lodgingEntries = lodgingEntries.slice(0, 3);
  
  const lodgingSectionTop = transportTop + transportHeight + 14;
  const lodgingContentTop = lodgingSectionTop + 20;
  const lodgingOpsStart = layout.ops.length;
  let cursorY = lodgingContentTop;
  const cityX = 156;
  const cityW = fullWidth - (cityX - left) - 8;
  const dateCol1 = left + 12;
  const dateCol2 = left + Math.floor(fullWidth / 2) + 4;
  const dateFieldW = Math.floor(fullWidth / 2) - 96;

  lodgingEntries.forEach((entry, index) => {
    if (index > 0) {
      layout.pushLine(left + 8, cursorY, left + fullWidth - 8, cursorY);
      cursorY += 6;
    }
    layout.pushText(`Location ${index + 1}: City, State`, 52, cursorY + 2, 8, "F1");
    layout.pushTextBox(entry.location || "", cityX, cursorY, cityW, fieldH, { size: 7.4, maxLines: 1, maxChars: 40, meta: `lodging-${index}-city` });
    cursorY += fieldH + rowGap + 2;

    layout.pushText("Check-In Date", dateCol1, cursorY + 2, 7.5, "F1");
    layout.pushTextBox(formatPdfDate(entry.checkInDate || ""), dateCol1 + 72, cursorY, dateFieldW, fieldH, { size: 7.2, maxLines: 1, maxChars: 14, meta: `lodging-${index}-in` });
    layout.pushText("Check-Out Date", dateCol2, cursorY + 2, 7.5, "F1");
    layout.pushTextBox(formatPdfDate(entry.checkOutDate || ""), dateCol2 + 76, cursorY, dateFieldW, fieldH, { size: 7.2, maxLines: 1, maxChars: 14, meta: `lodging-${index}-out` });
    cursorY += fieldH + rowGap;

    layout.pushText("Option 1:", 72, cursorY + 2, 8, "F1");
    layout.pushTextBox(entry.option1 || "", cityX, cursorY, cityW, fieldH, { size: 7.4, maxLines: 1, maxChars: 90, meta: `lodging-${index}-opt1` });
    cursorY += fieldH + rowGap;

    layout.pushText("Option 2:", 72, cursorY + 2, 8, "F1");
    layout.pushTextBox(entry.option2 || "", cityX, cursorY, cityW, fieldH, { size: 7.4, maxLines: 1, maxChars: 90, meta: `lodging-${index}-opt2` });
    cursorY += fieldH + rowGap;
  });

  const lodgingHeight = Math.max(72, cursorY - lodgingSectionTop + 6);
  layout.pushText("LODGING", 52, lodgingSectionTop + 8, 14, "F2");
  const lodgingBorderY = pageHeight - (lodgingSectionTop + lodgingHeight);
  layout.ops.splice(lodgingOpsStart, 0, `${left.toFixed(2)} ${lodgingBorderY.toFixed(2)} ${fullWidth.toFixed(2)} ${lodgingHeight.toFixed(2)} re S`);

  const commentsTop = lodgingSectionTop + lodgingHeight + 14;
  const commentsHeight = 50;
  const commentsOpsStart = layout.ops.length;
  layout.pushText("Comments:", 52, commentsTop + 8, 8.5, "F1");
  layout.pushTextBox(form.comments || request.notes || "", 52, commentsTop + 20, fullWidth - 12, 26, { size: 7.6, maxLines: 3, padTop: 3, minFontSize: 6.2, maxChars: 420, lineGap: 1.2, meta: "comments" });
  const commentsBorderY = pageHeight - (commentsTop + commentsHeight);
  layout.ops.splice(commentsOpsStart, 0, `${left.toFixed(2)} ${commentsBorderY.toFixed(2)} ${fullWidth.toFixed(2)} ${commentsHeight.toFixed(2)} re S`);

  const footerTop = commentsTop + commentsHeight + 14;
  layout.pushText("Travel C/O:", 52, footerTop + 4, 8.5, "F1");
  layout.pushTextBox(request.chargeObject || form.travelCo || "", 108, footerTop, 168, 16, { size: 8, maxLines: 1, minFontSize: 6, maxChars: 34, meta: "travel-co" });
  layout.pushText("BFM:", 300, footerTop + 4, 8.5, "F1");
  layout.pushTextBox(form.bfm || "", 334, footerTop, 168, 16, { size: 8, maxLines: 1, minFontSize: 6, maxChars: 34, meta: "bfm" });

  layout.validate();

  const contentBottom = footerTop + 24;
  const outerHeight = Math.max(688, contentBottom - outerTop + 12);
  if (contentBottom > pageHeight - 12) {
    throw new Error(`Engineering travel PDF content exceeds one page (${Math.round(contentBottom)}pt > ${pageHeight - 12}pt).`);
  }

  const ops = ["0 G", "0.75 w"];
  const topToPdfY = (top) => pageHeight - top;
  ops.push(`${outerLeft.toFixed(2)} ${topToPdfY(outerTop + outerHeight).toFixed(2)} 540.00 ${outerHeight.toFixed(2)} re S`);
  ops.push(...layout.ops);

  const joined = ops.join("\n");
  const objects = [];
  objects[1] = "<< /Type /Catalog /Pages 2 0 R >>";
  objects[2] = "<< /Type /Pages /Kids [6 0 R] /Count 1 >>";
  objects[3] = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>";
  objects[4] = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>";
  objects[5] = "<< /Type /Font /Subtype /Type1 /BaseFont /Times-Roman >>";
  objects[6] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Resources << /Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R >> >> /Contents 7 0 R >>`;
  objects[7] = `<< /Length ${joined.length} >>\nstream\n${joined}\nendstream`;

  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  for (let id = 1; id < objects.length; id++) {
    offsets[id] = pdf.length;
    pdf += `${id} 0 obj\n${objects[id]}\nendobj\n`;
  }
  const xrefOffset = pdf.length;
  pdf += `xref\n0 ${objects.length}\n0000000000 65535 f \n`;
  for (let id = 1; id < objects.length; id++) {
    pdf += `${String(offsets[id]).padStart(10, "0")} 00000 n \n`;
  }
  pdf += `trailer\n<< /Size ${objects.length} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;

  const bytes = new Uint8Array(pdf.length);
  for (let index = 0; index < pdf.length; index++) bytes[index] = pdf.charCodeAt(index) & 0xff;
  return new Blob([bytes], { type: "application/pdf" });
}

function createSimpleTravelPdfBlob(request) {
  const lines = buildTravelPdfLines(request);
  const pageWidth = 612;
  const pageHeight = 792;
  const left = 54;
  const top = 740;
  const lineHeight = 16;
  const linesPerPage = 42;
  const chunks = [];
  for (let index = 0; index < lines.length; index += linesPerPage) {
    chunks.push(lines.slice(index, index + linesPerPage));
  }
  const objects = [];
  const pageIds = [];
  const fontObjectId = 3;
  let nextId = 4;

  objects[1] = "<< /Type /Catalog /Pages 2 0 R >>";
  objects[2] = "";
  objects[3] = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>";

  chunks.forEach((pageLines) => {
    const contentId = nextId++;
    const pageId = nextId++;
    const contentStream = [
      "BT",
      "/F1 11 Tf",
      `1 0 0 1 ${left} ${top} Tm`,
      `${lineHeight} TL`
    ];
    pageLines.forEach((line, idx) => {
      const prefix = idx === 0 ? "" : "T* ";
      contentStream.push(`${prefix}(${pdfSafeText(line)}) Tj`);
    });
    contentStream.push("ET");
    const joined = contentStream.join("\n");
    objects[contentId] = `<< /Length ${joined.length} >>\nstream\n${joined}\nendstream`;
    objects[pageId] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Resources << /Font << /F1 ${fontObjectId} 0 R >> >> /Contents ${contentId} 0 R >>`;
    pageIds.push(pageId);
  });

  objects[2] = `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] /Count ${pageIds.length} >>`;

  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  for (let id = 1; id < objects.length; id++) {
    if (!objects[id]) continue;
    offsets[id] = pdf.length;
    pdf += `${id} 0 obj\n${objects[id]}\nendobj\n`;
  }
  const xrefOffset = pdf.length;
  pdf += `xref\n0 ${objects.length}\n`;
  pdf += "0000000000 65535 f \n";
  for (let id = 1; id < objects.length; id++) {
    const offset = offsets[id] || 0;
    pdf += `${String(offset).padStart(10, "0")} 00000 n \n`;
  }
  pdf += `trailer\n<< /Size ${objects.length} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;

  const bytes = new Uint8Array(pdf.length);
  for (let index = 0; index < pdf.length; index++) {
    bytes[index] = pdf.charCodeAt(index) & 0xff;
  }
  return new Blob([bytes], { type: "application/pdf" });
}

function createTravelPdfBlob(request) {
  if (request && request.formMode === "Engineering") {
    if (!request.chargeObject) {
      throw new Error("TTSD/CL PDF is generated after Finance assigns a C/O number.");
    }
    return createEngineeringTravelPdfBlob(request);
  }
  return createSimpleTravelPdfBlob(request);
}

function createTravelHtmlBlob(request) {
  const html = buildPrintableTravelHtml(request);
  return new Blob([html], { type: "text/html;charset=utf-8" });
}

function openTravelDocByPolicy(r) {
  if (!r || !r.exportFileUrl) {
    toast("No document has been generated for this request yet.", "info");
    return;
  }
  const api = window.AEWTTR && window.AEWTTR.OfficeDesktop;
  if (api && typeof api.openSharePointFileByPolicy === "function") {
    api.openSharePointFileByPolicy(r.exportFileUrl, r.exportFileName, r.exportMimeType
      || "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    return;
  }
  window.open(travelExportOpenUrl(r.exportFileUrl, r.exportFileName), "_blank", "noopener");
}

function travelExportOpenUrl(fileUrl, fileName) {
  if (!fileUrl) return "";
  const api = window.AEWTTR && window.AEWTTR.OfficeDesktop;
  if (api && typeof api.sharePointBrowserUrl === "function") {
    return api.sharePointBrowserUrl(fileUrl);
  }
  const ext = String(fileName || "").toLowerCase().split(".").pop();
  if (["doc", "docx", "xls", "xlsx", "ppt", "pptx"].includes(ext)) {
    return `${fileUrl}${fileUrl.includes("?") ? "&" : "?"}web=1`;
  }
  return fileUrl;
}

function reserveTravelExportOpen(fileName, mimeType) {
  const api = window.AEWTTR && window.AEWTTR.OfficeDesktop;
  if (typeof isSharePointMode !== "function" || !isSharePointMode()
    || !api || typeof api.reserveSharePointFileWindow !== "function") {
    return null;
  }
  return api.reserveSharePointFileWindow(fileName, mimeType);
}

function closeTravelExportOpen(popup) {
  const api = window.AEWTTR && window.AEWTTR.OfficeDesktop;
  if (api && typeof api.closeReservedSharePointFileWindow === "function") {
    api.closeReservedSharePointFileWindow(popup);
  } else if (popup && !popup.closed) {
    popup.close();
  }
}

function openTravelExportByPolicy(result, fallbackFileName, popup) {
  const fileUrl = result && result.fileUrl || "";
  if (!fileUrl) {
    closeTravelExportOpen(popup);
    return false;
  }
  const api = window.AEWTTR && window.AEWTTR.OfficeDesktop;
  if (api && typeof api.openSharePointFileByPolicy === "function") {
    return api.openSharePointFileByPolicy(
      fileUrl,
      result.fileName || fallbackFileName,
      result.mimeType || result.contentType || "",
      { popup }
    );
  }
  if (popup && !popup.closed) {
    popup.location.replace(fileUrl);
    return true;
  }
  window.open(fileUrl, "_blank", "noopener");
  return true;
}

/* Hand a generated Blob to the browser as a download. Object URLs are revoked on
   a later tick; revoking synchronously cancels the download in some browsers. */
function downloadBlobAsFile(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 4000);
}

async function saveTravelExportToSharePoint(request, opts) {
  opts = opts || {};
  if (!isSharePointMode()) {
    throw new Error("Travel export requires SharePoint mode so the file can be stored in SharePoint.");
  }
  const modeLabel = String(request.formMode || "Standard").toLowerCase().replace(/\s+/g, "-");
  const baseId = String(request.id || "travel-request").replace(/[^A-Za-z0-9._-]+/g, "-");

  let docxBlob;
  if (typeof window.createTravelDocxBlob === "function") {
    try {
      docxBlob = await window.createTravelDocxBlob(request);
    } catch (e) {
      console.warn("PULSE: DOCX generation failed, falling back to HTML.", e);
    }
  }

  const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  if (docxBlob) {
    const fileName = `${baseId}-${modeLabel}-travel-request.docx`;
    const file = new File([docxBlob], fileName, { type: DOCX_MIME, lastModified: Date.now() });
    const result = await sharePointAdapter.uploadProjectDocument(currentSiteUrl(), "Travel Requests", request.id || "Travel Request", file);
    request.exportFileUrl = result.fileUrl || "";
    request.exportFileName = result.fileName || fileName;
    request.exportMimeType = DOCX_MIME;
    return result;
  }

  // Fallback: HTML
  const htmlBlob = createTravelHtmlBlob(request);
  const fileName = `${baseId}-${modeLabel}-export.html`;
  const file = new File([htmlBlob], fileName, { type: "text/html", lastModified: Date.now() });
  const result = await sharePointAdapter.uploadProjectDocument(currentSiteUrl(), "Travel Requests", request.id || "Travel Request", file);
  request.exportFileUrl = result.fileUrl || "";
  request.exportFileName = result.fileName || fileName;
  request.exportMimeType = "text/html";
  return result;
}

/* ---------- calendar invite helpers ---------- */

function _icsEscape(str) {
  return String(str || "").replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\n/g, "\\n");
}

function buildTravelIcsContent(req) {
  const isAllDay = travelRequestIsAllDay(req);
  const isLeave = travelCategory(req) === "Leave";
  const title = _icsEscape(req.tripTitle || (isLeave ? `Leave — ${req.requester}` : `TDY — ${req.destination || "Travel"}`));
  const uid = `PULSE-${req.id}@aewttr.pulse`;
  const stamp = new Date().toISOString().replace(/[-:.]/g, "").slice(0, 15) + "Z";

  function padDate(iso) { return iso ? iso.replace(/-/g, "") : ""; }
  function padTime(t) { return t ? t.replace(/:/g, "").slice(0, 4).padEnd(6, "0") : "000000"; }

  let dtStart, dtEnd;
  if (isAllDay) {
    dtStart = `DTSTART;VALUE=DATE:${padDate(req.start)}`;
    const endDay = req.end ? new Date(req.end + "T12:00:00Z") : new Date((req.start || req.end) + "T12:00:00Z");
    endDay.setUTCDate(endDay.getUTCDate() + 1);
    dtEnd = `DTEND;VALUE=DATE:${endDay.toISOString().slice(0, 10).replace(/-/g, "")}`;
  } else {
    dtStart = `DTSTART;TZID=America/New_York:${padDate(req.start)}T${padTime(req.startTime)}`;
    dtEnd = `DTEND;TZID=America/New_York:${padDate(req.end)}T${padTime(req.endTime)}`;
  }

  const descLines = [
    `Travel Request: ${req.id}`,
    req.purpose ? `Purpose: ${req.purpose}` : "",
    req.destination ? `Destination: ${req.destination}` : "",
    req.chargeObject ? `Charge Object: ${req.chargeObject}` : ""
  ].filter(Boolean);

  return [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//AEWTTR PULSE//Travel//EN",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
    "BEGIN:VEVENT",
    `UID:${uid}`,
    `DTSTAMP:${stamp}`,
    dtStart,
    dtEnd,
    `SUMMARY:${title}`,
    req.destination ? `LOCATION:${_icsEscape(req.destination)}` : "",
    `DESCRIPTION:${_icsEscape(descLines.join("\n"))}`,
    "STATUS:CONFIRMED",
    "TRANSP:OPAQUE",
    "END:VEVENT",
    "END:VCALENDAR"
  ].filter(Boolean).join("\r\n");
}

function buildTravelOwaLink(req) {
  const isLeave = travelCategory(req) === "Leave";
  const title = req.tripTitle || (isLeave ? `Leave — ${req.requester}` : `TDY — ${req.destination || "Travel"}`);
  const start = req.start && req.startTime ? `${req.start}T${req.startTime}:00` : (req.start || "");
  const end = req.end && req.endTime ? `${req.end}T${req.endTime}:00` : (req.end || "");
  const bodyText = [
    `Travel Request: ${req.id}`,
    req.purpose ? `Purpose: ${req.purpose}` : "",
    req.destination ? `Destination: ${req.destination}` : ""
  ].filter(Boolean).join("\n");
  const params = new URLSearchParams();
  if (start) params.set("startdt", start);
  if (end) params.set("enddt", end);
  params.set("subject", title);
  if (bodyText) params.set("body", bodyText);
  if (req.destination) params.set("location", req.destination);
  if (travelRequestIsAllDay(req)) params.set("allday", "true");
  const attendeeEmails = new Set();
  if (req.requesterEmail) attendeeEmails.add(String(req.requesterEmail).trim().toLowerCase());
  (req.travelers || []).forEach((t) => { if (t.email && t.email.trim()) attendeeEmails.add(t.email.trim().toLowerCase()); });
  if (attendeeEmails.size) params.set("to", Array.from(attendeeEmails).join(";"));
  return `https://outlook.office.com/calendar/action/compose?${params.toString()}`;
}

function downloadTravelIcs(req) {
  const content = buildTravelIcsContent(req);
  const blob = new Blob([content], { type: "text/calendar;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${req.id}.ics`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

function _travelIsoWithOffset(dateStr, timeStr) {
  if (!dateStr) return "";
  const time = timeStr || "08:00";
  const dt = new Date(`${dateStr}T${time}:00`);
  const off = -dt.getTimezoneOffset();
  const sign = off >= 0 ? "+" : "-";
  const h = String(Math.floor(Math.abs(off) / 60)).padStart(2, "0");
  const m = String(Math.abs(off) % 60).padStart(2, "0");
  return `${dateStr}T${time}:00${sign}${h}:${m}`;
}

function openTravelInTeams(req) {
  const isLeave = travelCategory(req) === "Leave";
  const subject = isLeave
    ? `${req.requester || "Leave"} — Leave (${req.id})`
    : `${req.tripTitle || req.destination || "Travel"} — ${req.id}`;
  const startTime = _travelIsoWithOffset(req.start, req.startTime || "08:00");
  const endTime = _travelIsoWithOffset(req.end || req.start, req.endTime || "17:00");
  const content = isLeave
    ? `Leave request ${req.id}${req.purpose ? ` — ${req.purpose}` : ""}`
    : `Travel request ${req.id} — ${req.destination || ""}${req.purpose ? ` | ${req.purpose}` : ""}`;
  const attendeeEmails = new Set();
  if (req.requesterEmail) attendeeEmails.add(String(req.requesterEmail).trim().toLowerCase());
  (req.travelers || []).forEach((t) => { if (t.email && t.email.trim()) attendeeEmails.add(t.email.trim().toLowerCase()); });
  const enc = encodeURIComponent;
  const parts = [`subject=${enc(subject)}`];
  if (startTime) parts.push(`startTime=${enc(startTime)}`);
  if (endTime) parts.push(`endTime=${enc(endTime)}`);
  if (content) parts.push(`content=${enc(content)}`);
  if (attendeeEmails.size) parts.push(`attendees=${Array.from(attendeeEmails).map(enc).join(",")}`);
  window.open(`https://teams.microsoft.com/l/meeting/new?${parts.join("&")}`, "_blank", "noopener,noreferrer");
}

async function notifyRequesterTravelSubmitted(req) {
  if (!isSharePointMode()) return;
  const isLeave = travelCategory(req) === "Leave";
  const recipients = new Set();
  if (req.requesterEmail) recipients.add(String(req.requesterEmail).trim().toLowerCase());
  (req.travelers || []).forEach((t) => {
    if (t.email && t.email.trim()) recipients.add(t.email.trim().toLowerCase());
  });
  const emails = Array.from(recipients).filter(Boolean);
  if (!emails.length) return;

  const subject = isLeave
    ? `PULSE Leave: ${req.id} submitted`
    : `PULSE Travel: ${req.id} submitted — ${req.destination || ""}`;
  const preview = isLeave
    ? `Your leave request has been submitted and is pending review.`
    : `Your travel request to ${req.destination || "your destination"} has been submitted and is pending approval.`;

  try {
    await notifyUsers({
      to: emails,
      subject,
      area: "Travel",
      kind: "success",
      preview,
      facts: travelNotificationFacts(req),
      actionUrl: travelDeepLinkActionUrl("travel/mine", req.id),
      actionTitle: "View request"
    });
  } catch (e) {
    console.warn("PULSE: travel submission confirmation failed.", e);
  }
}

async function notifyFinanceAdminsOfTravelRequest(req) {
  if (!isSharePointMode()) return;
  const db = window.AEWTTR.db;
  const financeEmails = (db.members || []).filter((m) => m.isFinanceAdmin && m.email).map((m) => m.email);
  if (!financeEmails.length) return;
  const subject = `PULSE Travel: ${req.id} needs C/O assignment — ${req.requester}`;
  try {
    await notifyUsers({
      to: financeEmails,
      subject,
      area: "Travel",
      kind: "action",
      preview: `${req.requester} submitted a travel request that needs a charge object (C/O) assigned.`,
      facts: travelNotificationFacts(req),
      actionUrl: travelDeepLinkActionUrl("travel/all", req.id),
      actionTitle: "Assign charge object"
    });
  } catch (e) {
    console.warn("PULSE: finance admin notification failed.", e);
  }
}

async function notifyRequesterTravelStatusUpdate(req) {
  if (!isSharePointMode() || !req.requesterEmail) return;
  const kind = req.status === "Cancelled" ? "cancelled"
    : req.status === "Withdrawn" ? "info"
    : req.status === "Completed" ? "success"
    : "info";
  const subject = `PULSE Travel: ${req.id} — ${req.status}`;
  try {
    await notifyUsers({
      to: [req.requesterEmail],
      subject,
      area: "Travel",
      kind,
      preview: travelStatusChangePreview(req),
      facts: travelNotificationFacts(req),
      actionUrl: travelDeepLinkActionUrl("travel/mine", req.id),
      actionTitle: "Open request"
    });
  } catch (e) {
    console.warn("PULSE: requester status notification failed.", e);
  }
}

async function notifyTravelCancellation(req, opts) {
  opts = opts || {};
  if (!isSharePointMode()) return;
  const db = window.AEWTTR.db;
  const recipients = new Set();
  if (req.requesterEmail) recipients.add(String(req.requesterEmail).trim());
  if (opts.priorStatus === "Submitted" && opts.byAdmin) {
    (db.members || []).filter((m) => m.isFinanceAdmin && m.email).forEach((m) => recipients.add(m.email));
  }
  if (opts.byAdmin && opts.priorStatus === "Submitted") {
    // Finance notified above — requester always included
  }
  const emails = Array.from(recipients).filter(Boolean);
  if (!emails.length) return;
  const byLine = opts.byAdmin
    ? `Cancelled by ${req.cancelledBy || "an admin"}`
    : "Cancelled by the requester";
  try {
    await notifyUsers({
      to: emails,
      subject: `PULSE Travel: ${req.id} cancelled`,
      area: "Travel",
      kind: "cancelled",
      preview: `${byLine}. This trip is no longer active.`,
      facts: travelNotificationFacts(req).concat(
        req.cancelledBy ? [{ title: "Cancelled by", value: req.cancelledBy }] : []
      ),
      actionUrl: travelDeepLinkActionUrl("travel/mine", req.id),
      actionTitle: "View request"
    });
  } catch (e) {
    console.warn("PULSE: travel cancellation notification failed.", e);
  }
}

async function notifyConcurrenceRecorded(req) {
  if (!isSharePointMode()) return;
  const db = window.AEWTTR.db;
  const isLeave = travelCategory(req) === "Leave";
  const recipients = new Set();
  if (req.requesterEmail) recipients.add(String(req.requesterEmail).trim());
  (req.travelers || []).forEach((t) => { if (t.email) recipients.add(String(t.email).trim()); });
  const emails = Array.from(recipients).filter(Boolean);
  if (!emails.length) return;
  const subject = isLeave
    ? `PULSE Leave: ${req.id} — customer concurrence recorded`
    : `PULSE Travel: ${req.id} — customer concurrence recorded`;
  try {
    await notifyUsers({
      to: emails,
      subject,
      area: "Travel",
      kind: "success",
      preview: `Customer concurrence has been recorded by ${req.customerConcurredBy || "an admin"}.`,
      facts: travelNotificationFacts(req),
      actionUrl: travelDeepLinkActionUrl("travel/mine", req.id),
      actionTitle: "Open request"
    });
  } catch (e) {
    console.warn("PULSE: concurrence notification failed.", e);
  }
}

async function notifyCoAssigned(req, isFirstCo) {
  if (!isSharePointMode()) return;
  const recipients = new Set();
  if (req.requesterEmail) recipients.add(String(req.requesterEmail).trim());
  (req.travelers || []).forEach((t) => { if (t.email) recipients.add(String(t.email).trim()); });
  const emails = Array.from(recipients).filter(Boolean);
  if (!emails.length) return;
  const co = req.chargeObject || "";
  const subject = isFirstCo
    ? `PULSE Travel: ${req.id} — charge object assigned${co ? ` (${co})` : ""}`
    : `PULSE Travel: ${req.id} — charge object updated${co ? ` (${co})` : ""}`;
  try {
    await notifyUsers({
      to: emails,
      subject,
      area: "Travel",
      kind: "success",
      preview: isFirstCo
        ? `A charge object has been assigned to your travel request${co ? ` (${co})` : ""}. The travel document is ready.`
        : `The charge object for your travel request has been updated${co ? ` to ${co}` : ""}. A new document revision has been generated.`,
      facts: travelNotificationFacts(req),
      actionUrl: travelDeepLinkActionUrl("travel/mine", req.id),
      actionTitle: "Open request"
    });
  } catch (e) {
    console.warn("PULSE: C/O assigned notification failed.", e);
  }
}

async function notifyAdminsOfTravelRequest(req) {
  if (!isSharePointMode()) return;
  const db = window.AEWTTR.db;
  const adminEmails = (db.members || []).filter((m) => m.isAdmin && m.email).map((m) => m.email);
  if (!adminEmails.length) return;
  const isLeave = travelCategory(req) === "Leave";
  const isEngineering = (req.formMode || req.requestType) === "Engineering";
  const needsConcurrence = isLeave ? !!req.requiresConcurrence : isEngineering;
  const subject = isLeave
    ? (needsConcurrence ? `PULSE Leave: ${req.id} needs customer concurrence — ${req.requester}` : `PULSE Leave: ${req.id} submitted — ${req.requester}`)
    : isEngineering
    ? `PULSE Travel: ${req.id} needs customer concurrence (TTSD/CL) — ${req.requester}`
    : `PULSE Travel: ${req.id} submitted — ${req.requester}`;
  try {
    await notifyUsers({
      to: adminEmails,
      subject,
      area: "Travel",
      kind: "action",
      preview: isLeave && needsConcurrence
        ? `${req.requester} submitted a leave request that requires customer concurrence.`
        : isLeave
        ? `${req.requester} submitted a leave request.`
        : isEngineering
        ? `${req.requester} submitted a TTSD/CL travel request that needs customer concurrence.`
        : `${req.requester} submitted a new ${req.formMode || "Standard"} travel request.`,
      facts: travelNotificationFacts(req),
      actionUrl: travelDeepLinkActionUrl("travel/all", req.id),
      actionTitle: needsConcurrence ? "Record concurrence" : "View request"
    });
  } catch (e) {
    console.warn("PULSE: admin notification email failed.", e);
  }
}

function getTravelWizardSteps(formMode, opts) {
  opts = opts || {};
  const includeType = !opts.existing && !opts.skipType;

  if (formMode === "Leave") {
    // Same wizard chrome as travel — only Details + Review (no Purpose /
    // Budget / Travelers tabs). Type chooser stays when still picking a mode.
    const leaveSteps = [];
    if (includeType) {
      leaveSteps.push({ id: "type", label: "Type", desc: "Pick the request type", icon: "bx-category" });
    }
    leaveSteps.push(
      { id: "trip", label: "Details", desc: "Leave dates, times, and notes", icon: "bx-calendar-x" },
      { id: "review", label: "Review", desc: "Confirm everything looks right", icon: "bx-check-circle" }
    );
    return leaveSteps;
  }

  if (formMode === "Contractor") {
    const contractorSteps = [];
    if (includeType) {
      contractorSteps.push({ id: "type", label: "Type", desc: "Pick the request type", icon: "bx-category" });
    }
    contractorSteps.push(
      { id: "trip", label: "Trip", desc: "Details of your trip", icon: "bx-map" },
      { id: "purpose", label: "Purpose", desc: "Why this travel is needed", icon: "bx-target-lock" },
      { id: "impact", label: "Impact", desc: "Risk and alternatives", icon: "bx-shield" },
      { id: "budget", label: "Budget", desc: "Cost estimate and extra notes", icon: "bx-dollar" },
      { id: "review", label: "Review", desc: "Confirm everything looks right", icon: "bx-check-circle" }
    );
    return contractorSteps;
  }
  
  if (formMode === "Engineering") {
    const engSteps = [];
    if (includeType) {
      engSteps.push({ id: "type", label: "Type", desc: "Pick the request type", icon: "bx-category" });
    }
    engSteps.push(
      { id: "info", label: "Traveler", desc: "Who is traveling", icon: "bx-user" },
      { id: "basics", label: "Trip", desc: "TDY location and purpose", icon: "bx-map" },
      { id: "flights", label: "Flights", desc: "Air travel details", icon: "bx-plane" },
      { id: "transport", label: "Transport", desc: "Ground transportation", icon: "bx-bus" },
      { id: "lodging", label: "Lodging", desc: "Hotel options", icon: "bx-hotel" },
      { id: "extras", label: "Details", desc: "Comments and admin fields", icon: "bx-detail" },
      { id: "review", label: "Review", desc: "Confirm and submit", icon: "bx-check-circle" }
    );
    return engSteps;
  }

  const standardSteps = [];
  if (includeType) {
    standardSteps.push({ id: "type", label: "Type", desc: "Pick the request type", icon: "bx-category" });
  }
  standardSteps.push(
    { id: "trip", label: "Trip", desc: "Destination and dates", icon: "bx-map" },
    { id: "purpose", label: "Purpose", desc: "Why this travel is needed", icon: "bx-target-lock" },
    { id: "impact", label: "Impact", desc: "Risk and alternatives", icon: "bx-shield" },
    { id: "info", label: "Traveler", desc: "Who is traveling and contact info", icon: "bx-user" },
    { id: "flights", label: "Flights", desc: "Air travel details", icon: "bx-plane" },
    { id: "transport", label: "Transport", desc: "Ground transportation", icon: "bx-bus" },
    { id: "lodging", label: "Lodging", desc: "Hotel options", icon: "bx-hotel" },
    { id: "extras", label: "Details", desc: "Comments and additional info", icon: "bx-detail" },
    { id: "budget", label: "Budget", desc: "Cost estimate and notes", icon: "bx-dollar" },
    { id: "review", label: "Review", desc: "Confirm everything looks right", icon: "bx-check-circle" }
  );
  return standardSteps;
}

function travelWizardReadValue(body, id) {
  const node = $("#" + id, body);
  return node ? node.value.trim() : "";
}

function validateTravelWizardStep(stepId, body, formMode, travelers) {
  if (stepId === "type" || stepId === "review") return { ok: true };
  if (formMode === "Leave") {
    if (stepId === "trip") {
      const start = travelWizardReadValue(body, "tr-start");
      const end = travelWizardReadValue(body, "tr-end");
      if (!start || !end) return { ok: false, message: "Departure and return dates are required." };
      if (end < start) return { ok: false, message: "Return date must be on or after the departure date." };
      const allDayNode = $("#tr-allday", body);
      const isAllDay = allDayNode ? allDayNode.checked : true;
      if (!isAllDay) {
        const startTime = travelWizardReadValue(body, "tr-start-time");
        const endTime = travelWizardReadValue(body, "tr-end-time");
        if (!startTime || !endTime) return { ok: false, message: "Enter departure and return times, or choose All day." };
      }
    }
    return { ok: true };
  }
  if (formMode === "Standard" || formMode === "Contractor") {
    if (stepId === "trip") {
      if (!travelWizardReadValue(body, "tr-title")) return { ok: false, message: "Trip title is required." };
      if (!travelers.length) return { ok: false, message: "Select at least one traveler." };
      if (!travelWizardReadValue(body, "tr-dest") || !travelWizardReadValue(body, "tr-start") || !travelWizardReadValue(body, "tr-end")) {
        return { ok: false, message: "Destination and dates are required." };
      }
      if ($("#tr-type", body) && $("#tr-type", body).value === "Other" && !travelWizardReadValue(body, "tr-type-other")) {
        return { ok: false, message: "Describe the travel type when Other is selected." };
      }
      const allDayNode = $("#tr-allday", body);
      const isAllDay = allDayNode ? allDayNode.checked : true;
      if (!isAllDay) {
        const startTime = travelWizardReadValue(body, "tr-start-time");
        const endTime = travelWizardReadValue(body, "tr-end-time");
        if (!startTime || !endTime) return { ok: false, message: "Enter departure and return times, or choose All day." };
      }
    }
    if (stepId === "purpose" && !travelWizardReadValue(body, "tr-purpose")) {
      return { ok: false, message: "Event purpose is required." };
    }
  }
  if (stepId === "info" && (formMode === "Engineering" || formMode === "Standard")) {
    if (!travelWizardReadValue(body, "eng-traveler-name")) {
      return { ok: false, message: "Primary traveler name is required." };
    }
  }
  if (formMode === "Engineering") {
    if (stepId === "basics") {
      if (!travelWizardReadValue(body, "eng-purpose")) return { ok: false, message: "Purpose of TDY is required." };
      if (!travelWizardReadValue(body, "eng-start") || !travelWizardReadValue(body, "eng-end")) {
        return { ok: false, message: "TDY start and return dates are required." };
      }
    }
    return { ok: true };
  }
  return { ok: true };
}

function buildTravelReviewHtml(body, formMode, travelers) {
  const row = (label, value) => `<div class="travel-review-row"><span class="k">${escapeHtml(label)}</span><span class="v">${escapeHtml(value || "—")}</span></div>`;
  const block = (stepId, title, rowsHtml) => `
    <section class="travel-review-block">
      <div class="travel-review-head">
        <h4>${escapeHtml(title)}</h4>
        <button type="button" class="btn-aewttr-outline btn-aewttr-sm" data-tw-goto="${stepId}"${tip(`Jump back to edit ${title}`)}><i class="bx bx-edit"></i> Edit</button>
      </div>
      <div class="travel-review-body">${rowsHtml}</div>
    </section>`;
  const travelerList = travelers.map((t) => `${t.name}${t.email ? ` (${t.email})` : ""}`).join(", ");
  const dateTimeRangeText = (startId, endId, allDayId, startTimeId, endTimeId) => {
    const start = travelWizardReadValue(body, startId);
    const end = travelWizardReadValue(body, endId);
    const allDayNode = $("#" + allDayId, body);
    const isAllDay = allDayNode ? allDayNode.checked : true;
    if (isAllDay) return `${start} → ${end}`;
    const startTime = travelWizardReadValue(body, startTimeId);
    const endTime = travelWizardReadValue(body, endTimeId);
    return `${start}${startTime ? ` ${startTime}` : ""} → ${end}${endTime ? ` ${endTime}` : ""}`;
  };
  const sections = [];
  if (formMode === "Leave") {
    const me = currentTravelerRecord();
    sections.push(block("trip", "Leave details", [
      row("Form type", "Leave"),
      row("Requester", me.name || travelerList),
      row("Dates", dateTimeRangeText("tr-start", "tr-end", "tr-allday", "tr-start-time", "tr-end-time")),
      row("Notes", travelWizardReadValue(body, "tr-leave-notes") || travelWizardReadValue(body, "tr-notes"))
    ].join("")));
  } else if (formMode === "Standard" || formMode === "Contractor") {
    sections.push(block("trip", "Trip overview", [
      row("Form type", formMode),
      row("Trip title", travelWizardReadValue(body, "tr-title")),
      row("Destination", travelWizardReadValue(body, "tr-dest")),
      row("Dates", dateTimeRangeText("tr-start", "tr-end", "tr-allday", "tr-start-time", "tr-end-time")),
      row("Travel type", (() => {
        const t = travelWizardReadValue(body, "tr-type") || "TDY";
        if (t === "Other") return `Other: ${travelWizardReadValue(body, "tr-type-other") || "—"}`;
        return t;
      })()),
      row("Travelers", travelerList)
    ].join("")));
    sections.push(block("purpose", "Purpose", [
      row("Purpose", travelWizardReadValue(body, "tr-purpose"))
    ].join("")));
    sections.push(block("impact", "Impact & alternatives", [
      row("Impact if not approved", travelWizardReadValue(body, "tr-impact") || "—"),
      row("Alternatives", travelWizardReadValue(body, "tr-alternatives") || "—")
    ].join("")));
    if (formMode === "Standard") {
      const engForm = collectEngineeringForm(body, []);
      sections.push(block("flights", "Flights", [
        row("Flying", engForm.flying),
        engForm.flying === "Yes" ? row("Seat preference", engForm.seatPreference) : "",
        engForm.flying === "Yes" ? row("From", engForm.flightFrom) : "",
        engForm.flying === "Yes" ? row("To", engForm.flightTo) : "",
        engForm.flying === "Yes" ? row("Return flight", engForm.flightReturn) : ""
      ].filter(Boolean).join("")));
      sections.push(block("transport", "Ground transport", [
        row("Airport transport", engForm.airportTransport && engForm.airportTransport.mode),
        row("TDY transport", engForm.tdyTransport && engForm.tdyTransport.mode)
      ].join("")));
      const filledLodging = (engForm.lodging || []).filter((l) => l.location || l.option1);
      if (filledLodging.length) {
        sections.push(block("lodging", "Lodging", filledLodging.map((l, li) => [
          row(`Location ${li + 1}`, l.location),
          row("Check-in", l.checkInDate),
          row("Check-out", l.checkOutDate),
          row("Option 1", l.option1),
          row("Option 2", l.option2)
        ].join("")).join("")));
      }
      if (engForm.comments) {
        sections.push(block("extras", "Additional details", [
          row("Comments", engForm.comments)
        ].join("")));
      }
    }
    sections.push(block("budget", "Budget & notes", [
      row("Estimated cost", travelWizardReadValue(body, "tr-cost") ? `$${travelWizardReadValue(body, "tr-cost")}` : ""),
      row("Notes", travelWizardReadValue(body, "tr-notes"))
    ].join("")));
  } else if (formMode === "Engineering") {
    const engForm = collectEngineeringForm(body, []);
    const addlNames = (engForm.additionalTravelers || []).map((t) => t.name).filter(Boolean);
    sections.push(block("info", "Traveler", [
      row("Traveler", engForm.travelerName),
      row("Phone", engForm.phoneNumber),
      addlNames.length ? row("Additional travelers", addlNames.join(", ")) : ""
    ].filter(Boolean).join("")));
    sections.push(block("basics", "TDY details", [
      row("Purpose", engForm.purposeOfTdy),
      row("Dates", `${engForm.tdyStartDate} — ${engForm.tdyReturnDate}`),
      row("Location", engForm.tdyLocation),
      row("Days", engForm.numberOfDays)
    ].filter(Boolean).join("")));
    sections.push(block("flights", "Flights", [
      row("Flying", engForm.flying),
      engForm.flying === "Yes" ? row("Seat preference", engForm.seatPreference) : "",
      engForm.flying === "Yes" ? row("From", engForm.flightFrom) : "",
      engForm.flying === "Yes" ? row("To", engForm.flightTo) : "",
      engForm.flying === "Yes" ? row("Return flight", engForm.flightReturn) : ""
    ].filter(Boolean).join("")));
    sections.push(block("transport", "Ground transport", [
      row("Airport transport", engForm.airportTransport && engForm.airportTransport.mode),
      row("TDY transport", engForm.tdyTransport && engForm.tdyTransport.mode)
    ].join("")));
    const filledLodging = (engForm.lodging || []).filter((l) => l.location || l.option1);
    if (filledLodging.length) {
      sections.push(block("lodging", "Lodging", filledLodging.map((l, li) => [
        row(`Location ${li + 1}`, l.location),
        row("Check-in", l.checkInDate),
        row("Check-out", l.checkOutDate),
        row("Option 1", l.option1),
        row("Option 2", l.option2)
      ].join("")).join("")));
    }
    if (engForm.comments || engForm.bfm) {
      sections.push(block("extras", "Additional details", [
        engForm.comments ? row("Comments", engForm.comments) : "",
        engForm.bfm ? row("BFM", engForm.bfm) : ""
      ].filter(Boolean).join("")));
    }
  }
  return `<div class="travel-review-stack">${sections.join("")}</div>
    <p style="margin:14px 0 0;font-size:12.5px;color:var(--aewttr-muted);">Tap Edit on any section to jump back and change your answers.</p>`;
}

function travelWizardPanelsHtml(db, editing) {
  const requester = editing ? editing.requester : db.user.name;
  const travelersBlock = `
    <div class="form-row">
      <label>Travelers</label>
      <div class="traveler-picker">
        <div id="tr-travelers-selected" class="traveler-chip-list"></div>
        <input class="input-aewttr" id="tr-travelers-input" placeholder="Search people or groups…">
        <div id="tr-travelers-suggestions" class="traveler-suggestions"></div>
      </div>
      <p style="font-size:11.5px;color:var(--aewttr-muted);margin:6px 0 0;">Your account is included by default.</p>
    </div>`;
  return `
    <div class="tw-step-panel" data-tw-step="type" data-tw-mode="both" hidden>
      <div class="travel-type-cards" id="tw-type-cards">
        <button type="button" class="travel-type-card" data-mode="Standard">
          <i class="bx bx-file" style="background:#0078D4;"></i>
          <strong>Standard</strong>
          <span>Quick TDY, conference, or training request.</span>
        </button>
        <button type="button" class="travel-type-card" data-mode="Leave">
          <i class="bx bx-calendar-x" style="background:#7B61A8;"></i>
          <strong>Leave</strong>
          <span>Personal leave — same request window, dates and notes only.</span>
        </button>
        <button type="button" class="travel-type-card" data-mode="Contractor">
          <i class="bx bx-briefcase-alt-2" style="background:#546B2F;"></i>
          <strong>Contractor Travel</strong>
          <span>Travel for contractors or external team members.</span>
        </button>
        <button type="button" class="travel-type-card" data-mode="TeamEvent">
          <i class="bx bx-calendar-event" style="background:${TEAM_EVENT_COLOR};"></i>
          <strong>Team Event</strong>
          <span>Schedule a meeting, training day, or team activity on the calendar.</span>
        </button>
      </div>
    </div>
    <div class="tw-step-panel" data-tw-step="info" data-tw-mode="Engineering" hidden>
      <div class="travel-wizard-step-card">
        <p class="travel-wizard-step-card-title">Traveler information</p>
        <div class="form-grid-2">
          <div class="form-row"><label>Date of request</label><input type="date" class="input-aewttr" id="eng-date-of-request"></div>
          <div class="form-row"><label>Phone number</label><input class="input-aewttr" id="eng-phone" placeholder="(xxx) xxx-xxxx"></div>
        </div>
        <div class="form-row">
          <label>Primary traveler <span class="required-star">*</span></label>
          <div class="traveler-picker">
            <div id="eng-primary-traveler-selected" class="traveler-chip-list"></div>
            <input class="input-aewttr" id="eng-primary-traveler-input" placeholder="Search by name or email…">
            <div id="eng-primary-traveler-suggestions" class="traveler-suggestions"></div>
          </div>
          <input type="hidden" id="eng-traveler-name">
          <input type="hidden" id="eng-traveler-email">
        </div>
        <div class="form-row">
          <label>Additional travelers <small style="font-weight:400;color:var(--aewttr-muted);">(up to 5 — leave blank if not applicable)</small></label>
          <table style="width:100%;border-collapse:collapse;margin-top:6px;">
            <thead><tr>
              <th style="text-align:left;font-size:11px;padding:0 8px 6px 0;color:var(--aewttr-muted);font-weight:600;">Name</th>
              <th style="text-align:left;font-size:11px;padding:0 0 6px 0;color:var(--aewttr-muted);font-weight:600;">Transportation notes</th>
            </tr></thead>
            <tbody>
              ${[1,2,3,4,5].map((i) => `<tr>
                <td style="padding:3px 8px 3px 0;"><input class="input-aewttr" id="eng-add-traveler-name-${i}" placeholder="Traveler ${i}"></td>
                <td style="padding:3px 0;"><input class="input-aewttr" id="eng-add-traveler-notes-${i}" placeholder="e.g. Flight, POV, rental"></td>
              </tr>`).join("")}
            </tbody>
          </table>
        </div>
      </div>
    </div>
    <div class="tw-step-panel" data-tw-step="basics" data-tw-mode="Engineering" hidden>
      <div class="travel-wizard-step-card">
        <p class="travel-wizard-step-card-title">TDY details</p>
        <div class="form-row">
          <label>Purpose of TDY <span class="required-star">*</span></label>
          <textarea class="textarea-aewttr travel-wizard-textarea travel-wizard-textarea--short" id="eng-purpose" placeholder="State the purpose of travel"></textarea>
        </div>
        <div class="form-grid-2">
          <div class="form-row"><label>TDY start date <span class="required-star">*</span></label><input type="date" class="input-aewttr" id="eng-start"></div>
          <div class="form-row"><label>TDY return date <span class="required-star">*</span></label><input type="date" class="input-aewttr" id="eng-end"></div>
        </div>
        <div class="form-grid-2">
          <div class="form-row">
            <label>TDY on base?</label>
            <div class="travel-choice-group">${travelChoiceGroup("eng-on-base", ["Yes", "No"], "")}</div>
          </div>
          <div class="form-row"><label>Number of days</label><input class="input-aewttr" id="eng-days" placeholder="Auto-calculated"></div>
        </div>
        <div class="form-row"><label>TDY location (City, State)</label><input class="input-aewttr" id="eng-location" placeholder="e.g. San Diego, CA"></div>
      </div>
    </div>
    <div class="tw-step-panel" data-tw-step="flights" data-tw-mode="Engineering" hidden>
      <div class="travel-wizard-step-card">
        <p class="travel-wizard-step-card-title">Flight details</p>
        <div class="form-row">
          <label>Flying?</label>
          <div class="travel-choice-group">${travelChoiceGroup("eng-flying", ["Yes", "No"], "")}</div>
        </div>
        <div class="form-row"><label>Seat preference</label><input class="input-aewttr" id="eng-seat-preference" placeholder="e.g. Aisle"></div>
        <div class="form-grid-2">
          <div class="form-row"><label>Outbound — From (airport / date, time)</label><input class="input-aewttr" id="eng-flight-from" placeholder="e.g. DCA / 10 Aug, 0800"></div>
          <div class="form-row"><label>Outbound — To (airport / date, time)</label><input class="input-aewttr" id="eng-flight-to" placeholder="e.g. SAN / 10 Aug, 1100"></div>
        </div>
        <div class="form-row"><label>Return flight — From (airport / date, time)</label><input class="input-aewttr" id="eng-flight-return" placeholder="e.g. SAN / 14 Aug, 1600"></div>
      </div>
    </div>
    <div class="tw-step-panel" data-tw-step="transport" data-tw-mode="Engineering" hidden>
      <div class="travel-wizard-step-card">
        <p class="travel-wizard-step-card-title">Ground transportation</p>
        <div class="form-row">
          <label>Transportation to/from airport</label>
          <div class="travel-choice-group">${travelChoiceGroup("eng-airport-transport", ["POV", "Rental", "Passenger"], "")}</div>
        </div>
        <div class="form-grid-2">
          <div class="form-row"><label>Rental upgrade justification</label><input class="input-aewttr" id="eng-airport-upgrade" placeholder="Reason for non-standard vehicle"></div>
          <div class="form-row"><label>Vehicle type</label><input class="input-aewttr" id="eng-airport-vehicle" placeholder="e.g. Midsize sedan"></div>
        </div>
        <hr style="border:none;border-top:1px solid var(--border-light,#e0e4e8);margin:16px 0;">
        <div class="form-row">
          <label>Transportation at TDY location</label>
          <div class="travel-choice-group">${travelChoiceGroup("eng-tdy-transport", ["Rental", "Passenger", "Public Transportation"], "")}</div>
        </div>
        <div class="form-grid-2">
          <div class="form-row"><label>Rental upgrade justification</label><input class="input-aewttr" id="eng-tdy-upgrade" placeholder="Reason for non-standard vehicle"></div>
          <div class="form-row"><label>Vehicle type</label><input class="input-aewttr" id="eng-tdy-vehicle" placeholder="e.g. Full-size SUV"></div>
        </div>
      </div>
    </div>
    <div class="tw-step-panel" data-tw-step="lodging" data-tw-mode="Engineering" hidden>
      <div class="travel-wizard-step-card">
        <p class="travel-wizard-step-card-title">Lodging options</p>
        ${[1,2,3].map((i) => `
        <div style="${i > 1 ? "margin-top:16px;padding-top:16px;border-top:1px solid var(--border-light,#e0e4e8);" : ""}">
          ${i > 1 ? `<p style="font-size:12px;color:var(--aewttr-muted);margin:0 0 10px;font-weight:600;">Lodging ${i} (optional)</p>` : ""}
          <div class="form-grid-2">
            <div class="form-row"><label>Location ${i > 1 ? i : ""} (City, State)</label><input class="input-aewttr" id="eng-lodging-location-${i}" placeholder="e.g. San Diego, CA"></div>
            <div></div>
          </div>
          <div class="form-grid-2" style="margin-top:8px;">
            <div class="form-row"><label>Check-in date</label><input type="date" class="input-aewttr" id="eng-lodging-checkin-${i}"></div>
            <div class="form-row"><label>Check-out date</label><input type="date" class="input-aewttr" id="eng-lodging-checkout-${i}"></div>
          </div>
          <div class="form-row" style="margin-top:8px;"><label>Option 1</label><input class="input-aewttr" id="eng-lodging-option1-${i}" placeholder="Hotel name or preference"></div>
          <div class="form-row"><label>Option 2</label><input class="input-aewttr" id="eng-lodging-option2-${i}" placeholder="Alternate hotel or preference"></div>
        </div>`).join("")}
      </div>
    </div>
    <div class="tw-step-panel" data-tw-step="extras" data-tw-mode="Engineering" hidden>
      <div class="travel-wizard-step-card">
        <p class="travel-wizard-step-card-title">Additional details</p>
        <div class="form-row">
          <label>Comments <small style="font-weight:400;color:var(--aewttr-muted);">(optional)</small></label>
          <textarea class="textarea-aewttr travel-wizard-textarea" id="eng-comments" placeholder="Additional notes, instructions, or context"></textarea>
        </div>
        <div class="form-row"><label>BFM <small style="font-weight:400;color:var(--aewttr-muted);">(optional)</small></label><input class="input-aewttr" id="eng-bfm" placeholder="e.g. BFM-17"></div>
        <p style="margin:14px 0 0;font-size:12px;color:var(--aewttr-muted);">Travel C/O is assigned by Finance after customer concurrence.</p>
      </div>
    </div>
    <div class="tw-step-panel" data-tw-step="trip" data-tw-mode="Standard" hidden>
      <div class="travel-wizard-step-card">
        <p class="travel-wizard-step-card-title">About the trip</p>
        <div class="form-row" id="tr-leave-requester-row" hidden>
          <label>Requester</label>
          <input class="input-aewttr" id="tr-leave-requester" value="${escapeHtml(requester)}" disabled>
        </div>
        <div class="form-row" id="tr-title-row"><label>Trip title</label><input class="input-aewttr" id="tr-title" placeholder="e.g. Ship check in San Diego"></div>
        <div class="form-grid-2" id="tr-dest-type-row">
          <div class="form-row"><label>Destination</label><input class="input-aewttr" id="tr-dest" placeholder="City, State"></div>
          <div class="form-row"><label>Travel type</label><select class="select-aewttr" id="tr-type"><option>TDY</option><option>Conference</option><option>Training</option><option>Other</option></select></div>
        </div>
        <div class="form-row" id="tr-type-other-row" hidden>
          <label>Describe travel type</label>
          <input class="input-aewttr" id="tr-type-other" placeholder="e.g. Local visit, vendor meeting">
        </div>
        <div class="form-grid-2">
          <div class="form-row"><label>Departure date</label><input type="date" class="input-aewttr" id="tr-start"></div>
          <div class="form-row"><label>Return date</label><input type="date" class="input-aewttr" id="tr-end"></div>
        </div>
        <label class="ssp-link-row"><input type="checkbox" id="tr-allday" checked> <span>All day <small>— uncheck to set specific departure/return times</small></span></label>
        <div class="form-grid-2" id="tr-time-row" hidden>
          <div class="form-row"><label>Departure time</label><input type="time" class="input-aewttr" id="tr-start-time"></div>
          <div class="form-row"><label>Return time</label><input type="time" class="input-aewttr" id="tr-end-time"></div>
        </div>
        <div class="form-row" id="tr-leave-notes-row" hidden>
          <label>Notes <small style="font-weight:400;color:var(--aewttr-muted);">(optional)</small></label>
          <textarea class="textarea-aewttr travel-wizard-textarea travel-wizard-textarea--short" id="tr-leave-notes" placeholder="Coverage, handoff, or anything admins should know"></textarea>
        </div>
        <div id="tr-leave-concurrence-row" hidden>
          <label class="ssp-link-row"><input type="checkbox" id="tr-leave-concurrence"> <span>Request customer concurrence <small style="font-weight:400;color:var(--aewttr-muted);">— check if this leave requires customer sign-off</small></span></label>
        </div>
      </div>
    </div>
    <div class="tw-step-panel" data-tw-step="purpose" data-tw-mode="Standard" hidden>
      <div class="travel-wizard-step-card">
        <p class="travel-wizard-step-card-title">Why this travel matters</p>
        <div class="form-row"><label>Event purpose <span class="required-star">*</span></label><textarea class="textarea-aewttr travel-wizard-textarea" id="tr-purpose" placeholder="What will you accomplish on this trip?"></textarea></div>
      </div>
    </div>
    <div class="tw-step-panel" data-tw-step="impact" data-tw-mode="Standard" hidden>
      <div class="travel-wizard-step-card">
        <p class="travel-wizard-step-card-title">Risk and alternatives</p>
        <div class="form-row"><label>Impact if not approved <small style="font-weight:400;color:var(--aewttr-muted);">(optional)</small></label><textarea class="textarea-aewttr travel-wizard-textarea" id="tr-impact" placeholder="What is at risk if this travel is denied?"></textarea></div>
        <div class="form-row"><label>Alternatives or virtual substitute <small style="font-weight:400;color:var(--aewttr-muted);">(optional)</small></label><textarea class="textarea-aewttr travel-wizard-textarea travel-wizard-textarea--short" id="tr-alternatives" placeholder="One alternative per line"></textarea></div>
      </div>
    </div>
    <div class="tw-step-panel" data-tw-step="budget" data-tw-mode="Standard" hidden>
      <div class="travel-wizard-step-card">
        <p class="travel-wizard-step-card-title" id="tr-budget-title">Estimated costs</p>
        <div class="form-row" id="tr-cost-row"><label>Estimated cost ($)</label><input type="number" class="input-aewttr" id="tr-cost" min="0"></div>
        <div class="form-row"><label>Requester</label><input class="input-aewttr" id="tr-requester" value="${escapeHtml(requester)}" disabled></div>
        <div class="form-row"><label>Additional notes</label><textarea class="textarea-aewttr travel-wizard-textarea travel-wizard-textarea--short" id="tr-notes"></textarea></div>
      </div>
    </div>
    <div id="tw-travelers-shared" class="travel-wizard-step-card" hidden>
      ${travelersBlock}
    </div>
    <div class="tw-step-panel" data-tw-step="review" data-tw-mode="both" hidden>
      <div id="tw-review-content"></div>
    </div>`;
}

async function submitTravelRequest(body, formMode, travelers, db, editing) {
  const costNode = $("#tr-cost", body);
  const notesNode = $("#tr-notes", body);
  const dest = travelWizardReadValue(body, "tr-dest");
  const start = travelWizardReadValue(body, "tr-start");
  const end = travelWizardReadValue(body, "tr-end");
  const req = editing ? editing : {
    id: nextTrId(formMode),
    requester: db.user.name,
    requesterEmail: db.user.email || "",
    status: "Submitted",
    chargeObject: "",
    chargeObjectStatus: "Pending",
    customerConcurrenceStatus: "Pending",
    projectIds: [],
    exportFileUrl: "",
    exportFileName: ""
  };
  const tripTitle = formMode === "Leave"
      ? `Leave - ${lastNameOf(req.requester) || "Request"}`
      : travelWizardReadValue(body, "tr-title");
  req.formMode = formMode;
  req.requestType = formMode === "Leave" ? "Personal Leave" : formMode === "Contractor" ? "Contractor Travel" : formMode;
  req.category = formMode === "Leave" ? "Leave" : "Travel";
  req.contractorTravel = formMode === "Contractor";
  req.tripTitle = tripTitle;
  /* Leave is always just the current user — no traveler picker. */
  req.travelers = formMode === "Leave"
    ? [currentTravelerRecord()]
    : travelers.map((traveler) => ({ ...traveler }));
  req.destination = formMode === "Leave" ? "Leave" : dest;
  req.start = start;
  req.end = end;
  const allDayCheckbox = $("#tr-allday", body);
  const startTimeNode = $("#tr-start-time", body);
  const endTimeNode = $("#tr-end-time", body);
  req.allDay = allDayCheckbox ? !!allDayCheckbox.checked : true;
  req.startTime = req.allDay ? "" : (startTimeNode ? startTimeNode.value : "");
  req.endTime = req.allDay ? "" : (endTimeNode ? endTimeNode.value : "");
  const leaveNotes = travelWizardReadValue(body, "tr-leave-notes");
  const notesValue = formMode === "Leave"
    ? (leaveNotes || (notesNode ? notesNode.value : "") || "")
    : (notesNode ? notesNode.value : "");
  req.purpose = formMode === "Leave" ? (notesValue || tripTitle)
      : travelWizardReadValue(body, "tr-purpose");
  req.impactIfNotApproved = formMode === "Leave" ? "" : travelWizardReadValue(body, "tr-impact");
  req.alternatives = formMode === "Leave" ? [] : travelWizardReadValue(body, "tr-alternatives").split("\n").map((line) => line.trim()).filter(Boolean);
  req.cost = formMode === "Leave" ? 0 : (costNode ? (+costNode.value || 0) : 0);
  let travelType = "";
  if (formMode !== "Leave") {
    travelType = $("#tr-type", body) ? $("#tr-type", body).value : "TDY";
    if (travelType === "Other") {
      const other = travelWizardReadValue(body, "tr-type-other");
      travelType = other ? `Other: ${other}` : "Other";
    }
  }
  req.type = travelType;
  req.notes = notesValue;
  if (formMode === "Engineering") {
    const engForm = collectEngineeringForm(body, []);
    req.engineeringForm = engForm;
    const primaryTraveler = { name: engForm.travelerName || req.requester, email: engForm.travelerEmail || req.requesterEmail || "" };
    const additionals = (engForm.additionalTravelers || []).map((t) => ({ name: t.name, email: t.email || "" }));
    req.travelers = [primaryTraveler, ...additionals].filter((t) => t.name);
    req.destination = engForm.tdyLocation || dest || "TDY";
    req.start = engForm.tdyStartDate || start;
    req.end = engForm.tdyReturnDate || end;
    req.purpose = engForm.purposeOfTdy || req.purpose;
    req.tripTitle = `TTSD/CL Travel — ${engForm.tdyLocation || "TDY"}`;
    req.requestType = "Engineering";
  }
  if (formMode === "Standard") {
    const engForm = collectEngineeringForm(body, []);
    req.engineeringForm = engForm;
  }
  const hasCo = !!(editing && editing.chargeObjectStatus === "Assigned");
  if (!editing) {
    req.status = "Submitted";
    if (formMode === "Leave") {
      const concToggle = $("#tr-leave-concurrence", body);
      req.requiresConcurrence = !!(concToggle && concToggle.checked);
      req.customerConcurrenceStatus = req.requiresConcurrence ? "Pending" : "Concurred";
      req.chargeObjectStatus = null;
    } else {
      req.chargeObject = "";
      req.chargeObjectStatus = "Pending";
      req.customerConcurrenceStatus = "Pending";
    }
  }
  if (!editing) db.travelRequests.unshift(req);
  req._auditAction = editing ? "Update" : "Submit";
  req._auditSummary = `${editing ? "Updated" : "Submitted"} ${formMode === "Leave" ? "leave" : "travel"} ${req.id}${formMode === "Leave" ? "" : ` — ${req.destination}`}`;
  Repo.save("travelRequest", req);
  markTravelRequestStatusSeen(req.id, req.status);
  if (!editing) notifyAdminsOfTravelRequest(req);
  if (!editing && formMode !== "Leave") notifyFinanceAdminsOfTravelRequest(req);
  if (!editing) notifyRequesterTravelSubmitted(req);
  refreshTravelNotifications();

  if (editing && hasCo && req.docReviewId && isSharePointMode()) {
    (async () => {
      try {
        await saveTravelExportToSharePoint(req, { force: true });
        await Repo.save("travelRequest", req);
        const allDocs = Object.values(db.docs || {}).flat();
        const doc = allDocs.find((d) => d.id === req.docReviewId);
        if (doc) {
          const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
          const nextNum = typeof nextDocRevisionNumber === "function"
            ? nextDocRevisionNumber(doc)
            : ((doc.revisions || []).length + 1);
          const newRevision = {
            id: uid("REV"),
            number: nextNum,
            fileUrl: req.exportFileUrl || "",
            serverRelativeUrl: "",
            fileName: req.exportFileName || `${req.id}-travel-request.docx`,
            fileType: (req.exportFileName || "").endsWith(".docx") ? "docx" : "html",
            mimeType: req.exportMimeType || DOCX_MIME,
            uploadedBy: typeof currentUserIdentity === "function" ? (currentUserIdentity().name || "") : "",
            uploadedDate: typeof todayIsoDate === "function" ? todayIsoDate() : new Date().toISOString().slice(0, 10),
            source: "Travel Request Updated"
          };
          if (typeof resetReviewersForNewRevision === "function") resetReviewersForNewRevision(doc);
          doc.revisions = doc.revisions || [];
          doc.revisions.unshift(newRevision);
          if (typeof setDocActiveRevision === "function") setDocActiveRevision(doc, newRevision.id);
          if (req.travelers && req.travelers.length) {
            const existingEmails = new Set((doc.reviewers || []).map((r) => (r.email || "").toLowerCase()));
            req.travelers.forEach((t) => {
              if (t.email && !existingEmails.has(t.email.toLowerCase())) {
                doc.reviewers = doc.reviewers || [];
                doc.reviewers.push({ name: t.name || "", email: t.email, decision: "Pending", reviewedAt: "", note: "", lastNotifiedAt: "" });
              }
            });
          }
          if (typeof saveDocReview === "function") {
            await saveDocReview(doc, { action: "Revision Added", text: `Travel request ${req.id} was updated — new document revision uploaded.` });
          }
          if (typeof notifyPendingReviewers === "function") {
            await notifyPendingReviewers(doc, { kind: "revision", force: true });
          }
        }
      } catch (revErr) {
        console.warn("PULSE: doc review revision after travel update failed.", revErr);
        toast(`${req.id} saved, but the Document Review revision could not be updated.`, "warn");
      }
    })();
  }

  if (editing) {
    toast(`${req.id} updated`, "success");
    navigate("travel/mine");
  } else {
    const isLeave = travelCategory(req) === "Leave";
    body.innerHTML = `<div class="travel-confirm-wrap">
      <div class="travel-confirm-card">
        <div class="travel-confirm-icon"><i class="bx bx-check-circle"></i></div>
        <h2 class="travel-confirm-title">${isLeave ? "Leave request submitted" : "Travel request submitted"}</h2>
        <p class="travel-confirm-id">${escapeHtml(req.id)}</p>
        ${!isLeave && req.destination ? `<p class="travel-confirm-dest">${escapeHtml(req.destination)}</p>` : ""}
        <p class="travel-confirm-msg">Your request has been submitted and is pending review. You'll receive a confirmation email shortly.</p>
        <div class="travel-confirm-actions">
          <a href="data:text/calendar;charset=utf-8,${encodeURIComponent(buildTravelIcsContent(req))}" download="${escapeHtml(req.id)}.ics" class="btn-aewttr btn-aewttr-secondary travel-confirm-teams-btn">
            <i class="bx bx-calendar-plus"></i>&ensp;Add to Calendar
          </a>
          <button type="button" class="btn-aewttr" id="travel-confirm-view-mine">View My Travel</button>
        </div>
      </div>
    </div>`;
    const viewMineBtn = $("#travel-confirm-view-mine", body);
    if (viewMineBtn) viewMineBtn.addEventListener("click", () => navigate("travel/mine"));
  }
}

function drawSubmitRequest(body, initialMode, editId) {
  if (typeof canSubmitForms === "function" && !canSubmitForms()) {
    body.innerHTML = `
      <div class="lock-block" style="padding:48px 24px;text-align:center;">
        <i class="bx bx-lock-alt" style="font-size:48px;color:var(--aewttr-muted);margin-bottom:12px;"></i>
        <h3>Viewer Access Only</h3>
        <p>The Viewer role has read-only access and cannot submit or edit travel requests or forms.</p>
        <button class="btn-aewttr btn-aewttr-sm" onclick="location.hash='#/travel/mine'" style="margin-top:16px;">View Travel Requests</button>
      </div>`;
    return;
  }
  const db = window.AEWTTR.db;
  const editing = editId ? db.travelRequests.find((r) => r.id === editId) : null;
  if (editId && (!editing || !canRequesterEditTravel(editing))) {
    toast("This request can't be edited.", "error");
    navigate("travel/mine");
    return;
  }
  const travelers = editing ? (editing.travelers || []).map((t) => ({ ...t })) : defaultTravelers();
  const initialModeNorm = String(initialMode || "").toLowerCase();
  const skipType = !!editing || !!initialModeNorm;
  let formMode = editing
    ? (travelCategory(editing) === "Leave" ? "Leave"
      : (editing.formMode || editing.requestType || "Standard"))
    : (initialModeNorm === "leave" ? "Leave"
      : initialModeNorm === "contractor" ? "Contractor"
      : initialModeNorm === "engineering" ? "Engineering"
      : "Standard");
  if (formMode === "Personal Leave") formMode = "Leave";
  if (formMode === "Contractor Travel") formMode = "Contractor";
  let steps = getTravelWizardSteps(formMode, { existing: editing, skipType });
  let stepIndex = 0;

  body.innerHTML = `
    <div class="travel-wizard travel-wizard--fit">
      ${editing && editing.chargeObjectStatus === "Assigned" ? `<div class="travel-wizard-banner travel-wizard-banner--info"><i class="bx bx-info-circle"></i> Saving changes to a request with an assigned charge object will generate a new document revision — all Document Review reviewers will be notified.</div>` : ""}
      <div class="travel-wizard-progress" id="tw-progress"></div>
      <div class="travel-wizard-panel aewttr-card aewttr-card-pad travel-wizard-panel--compact">
        <div class="travel-wizard-step-head travel-wizard-step-head--compact">
          <span class="travel-wizard-eyebrow" id="tw-eyebrow"></span>
          <h2 id="tw-step-title"></h2>
          <p id="tw-step-desc"></p>
        </div>
        <div id="tw-form-root">${travelWizardPanelsHtml(db, editing)}</div>
        <div class="travel-wizard-actions">
          <div class="tw-back-group">
            <button type="button" class="btn-aewttr-ghost" id="tw-back" style="visibility:hidden;"${tip("Go to the previous step")}><i class="bx bx-chevron-left"></i> Back</button>
            ${!editing && !skipType ? `<button type="button" class="btn-aewttr-ghost" id="tw-back-to-options" hidden${tip("Return to request type selection")}><i class="bx bx-grid-alt"></i> Back to options</button>` : ""}
          </div>
          <div class="travel-wizard-actions-right">
            ${editing ? `<button type="button" class="btn-aewttr-outline" id="tw-cancel"${tip("Discard changes and return to My Travel")}>Cancel</button>` : ""}
            <button type="button" class="btn-aewttr" id="tw-next"${tip("Continue to the next step")}>Continue</button>
          </div>
        </div>
      </div>
    </div>`;

  const root = $("#tw-form-root", body);
  const travelerPicker = wireTravelerPicker(body, travelers);
  travelerPicker.refresh();
  wirePrimaryTravelerPicker(body, travelers);
  wireTravelDaysAutofill(body);
  if (editing) {
    hydrateTravelSubmitForm(body, editing);
  }
  wireTravelTypeOther(body);
  wireAllDayToggle(body, "tr");
  updateTripPanelForMode(body, formMode);

  function syncTypeCards() {
    $all(".travel-type-card", body).forEach((card) => {
      card.classList.toggle("selected", card.dataset.mode === formMode);
    });
  }

  function refreshStepsAfterModeChange() {
    const prevId = steps[stepIndex] && steps[stepIndex].id;
    steps = getTravelWizardSteps(formMode, { existing: editing, skipType });
    const nextIndex = steps.findIndex((s) => s.id === prevId);
    stepIndex = nextIndex >= 0 ? nextIndex : Math.min(stepIndex, steps.length - 1);
  }

  function showStepPanels() {
    const step = steps[stepIndex];
    if (!step) return;
    $all(".tw-step-panel", root).forEach((panel) => {
      const mode = panel.dataset.twMode || formMode;
      const sharedSteps = {
        trip: ["Standard", "Contractor", "Leave"],
        purpose: ["Standard", "Contractor"],
        impact: ["Standard", "Contractor"],
        budget: ["Standard", "Contractor"]
      };
      const stepId = panel.dataset.twStep;
      const engStdSharedSteps = ["info", "flights", "transport", "lodging", "extras"];
      const modeOk = mode === "both" || mode === formMode
        || ((sharedSteps[stepId] || []).includes(formMode) && (mode === "Standard" || mode === "Leave" || mode === "Contractor"))
        || (mode === "Engineering" && formMode === "Standard" && engStdSharedSteps.includes(stepId));
      panel.hidden = !(stepId === step.id && modeOk);
    });
    const travelersShared = $("#tw-travelers-shared", root);
    if (travelersShared) {
      travelersShared.hidden = formMode === "Leave" || formMode === "Engineering"
        || !(step.id === "trip" || step.id === "info");
    }
    if (step.id === "review") {
      $("#tw-review-content", root).innerHTML = buildTravelReviewHtml(body, formMode, travelers);
      $all("[data-tw-goto]", root).forEach((btn) => btn.addEventListener("click", () => {
        const target = steps.findIndex((s) => s.id === btn.dataset.twGoto);
        if (target >= 0) { stepIndex = target; renderWizard(); }
      }));
    }
  }

  function renderWizard() {
    const step = steps[stepIndex];
    if (!step) return;
    $("#tw-progress", body).innerHTML = steps.map((s, i) => `
      <button type="button" class="travel-wizard-pill ${i === stepIndex ? "active" : ""} ${i < stepIndex ? "done" : ""}" data-tw-step-index="${i}"${tip(`Jump to ${s.label}`)}>
        <i class="bx ${s.icon}"></i>${escapeHtml(s.label)}
      </button>`).join("");
    $all("[data-tw-step-index]", body).forEach((pill) => pill.addEventListener("click", () => {
      const target = Number(pill.dataset.twStepIndex);
      if (target === stepIndex) return;
      stepIndex = target;
      renderWizard();
    }));
    $("#tw-eyebrow", body).textContent = `Step ${stepIndex + 1} of ${steps.length}`;
    $("#tw-step-title", body).textContent = step.id === "review"
      ? (editing ? `Review changes to ${editing.id}` : "Review your request")
      : step.label;
    $("#tw-step-desc", body).textContent = step.desc;
    const backBtn = $("#tw-back", body);
    const isTypeStep = step.id === "type";
    const isFirstRealStep = !isTypeStep && stepIndex === 1 && steps[0] && steps[0].id === "type";
    backBtn.style.visibility = stepIndex > 0 && !isFirstRealStep ? "visible" : "hidden";
    const backToOptionsBtn = $("#tw-back-to-options", body);
    if (backToOptionsBtn) backToOptionsBtn.hidden = !isFirstRealStep;
    const nextBtn = $("#tw-next", body);
    nextBtn.hidden = isTypeStep;
    nextBtn.innerHTML = step.id === "review"
      ? `<i class="bx ${editing ? "bx-save" : "bx-send"}"></i> ${editing ? "Save Changes" : "Submit Request"}`
      : "Continue";
    syncTypeCards();
    updateTripPanelForMode(body, formMode);
    showStepPanels();
  }

  $all(".travel-type-card", body).forEach((card) => card.addEventListener("click", () => {
    if (card.dataset.mode === "TeamEvent") {
      navigate("travel/events/new");
      return;
    }
    formMode = card.dataset.mode;
    syncTypeCards();
    refreshStepsAfterModeChange();
    updateTripPanelForMode(body, formMode);
    stepIndex = 1; // Auto-advance past type selection
    renderWizard();
  }));

  $("#tw-back", body).addEventListener("click", () => {
    if (stepIndex > 0) { stepIndex -= 1; renderWizard(); }
  });

  const backToOptionsEl = $("#tw-back-to-options", body);
  if (backToOptionsEl) {
    backToOptionsEl.addEventListener("click", () => {
      stepIndex = 0;
      renderWizard();
    });
  }

  if (editing) $("#tw-cancel", body).addEventListener("click", () => navigate("travel/mine"));

  $("#tw-next", body).addEventListener("click", async () => {
    const step = steps[stepIndex];
    if (step.id === "type" && !formMode) {
      toast("Select a request type to continue.", "error");
      return;
    }
    if (step.id !== "review") {
      const check = validateTravelWizardStep(step.id, body, formMode, travelers);
      if (!check.ok) { toast(check.message, "error"); return; }
      stepIndex += 1;
      renderWizard();
      return;
    }
    const finalCheck = validateTravelWizardStep(formMode === "Engineering" ? "basics" : "trip", body, formMode, travelers);
    if (!finalCheck.ok) { toast(finalCheck.message, "error"); return; }
    if (formMode === "Standard" && !travelWizardReadValue(body, "tr-purpose")) {
      toast("Event purpose is required.", "error");
      return;
    }
    travelerPicker.refresh();
    const nextBtn = $("#tw-next", body);
    nextBtn.disabled = true;
    try {
      await submitTravelRequest(body, formMode, travelers, db, editing);
    } finally {
      nextBtn.disabled = false;
    }
  });

  if (skipType || editing) syncTypeCards();
  renderWizard();
}


async function createTravelDocReview(request) {
  const db = window.AEWTTR.db;
  if (!db.docs) db.docs = {};
  if (!db.docs["Not Started"]) db.docs["Not Started"] = [];

  const allDocs = Object.values(db.docs).flat();
  const maxNum = allDocs.reduce((m, d) => {
    const parts = String(d.id || "").split("-");
    if (parts[0] !== "DOC") return m;
    return Math.max(m, parseInt(parts[1], 10) || 0);
  }, 0);
  const docId = `DOC-${String(maxNum + 1).padStart(4, "0")}`;

  const travelers = request.travelers || [];
  const reviewers = travelers.map((t) => ({
    name: t.name || "",
    email: t.email || "",
    decision: "Pending",
    reviewedAt: "",
    note: "",
    lastNotifiedAt: ""
  }));
  if (typeof ensureSubmitterReviewer === "function") {
    ensureSubmitterReviewer(reviewers, request.requester, request.requesterEmail || "");
  }

  const revId = uid("REV");
  const modeLabel = request.formMode === "Engineering" ? "TTSD/CL"
    : request.formMode === "Leave" ? "Leave"
    : request.formMode === "Contractor" ? "Contractor"
    : "Travel";
  const doc = {
    id: docId,
    title: `${request.id} — ${request.tripTitle || modeLabel + " Request"}`,
    submitter: request.requester || "",
    submitterEmail: request.requesterEmail || "",
    documentAdmin: request.requester || "",
    documentAdminEmail: request.requesterEmail || "",
    projectCode: "",
    reviewers,
    revisions: [{
      id: revId,
      number: 1,
      fileUrl: request.exportFileUrl || "",
      fileName: request.exportFileName || `${request.id}-travel-export.docx`,
      fileType: (request.exportFileName || "").endsWith(".docx") ? "docx" : "html",
      mimeType: request.exportMimeType || "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      uploadedBy: request.requester || "Travel System",
      uploadedDate: (typeof todayIsoDate === "function" ? todayIsoDate() : new Date().toISOString().slice(0, 10)),
      source: "Travel Request Workflow"
    }],
    activeRevisionId: revId,
    status: "Not Started",
    _column: "Not Started",
    reviewActivity: [],
    travelRequestId: request.id
  };

  if (typeof normalizeDocReview === "function") normalizeDocReview(doc);

  db.docs["Not Started"].unshift(doc);
  await Repo.save("docReviewItem", doc, { column: "Not Started" });

  if (typeof notifyPendingReviewers === "function") {
    try { await notifyPendingReviewers(doc, { kind: "ready" }); } catch (_) {}
  }

  return doc;
}

function openCoAssignModal(trId, onDone) {
  const db = window.AEWTTR.db;
  const r = db.travelRequests.find(x => x.id === trId);
  if (!r || !canAssignTravelCo()) return;
  const isFirstCo = r.chargeObjectStatus !== "Assigned";
  const genLoaderHtml = `
    <div class="co-gen-state" id="co-gen-state" hidden>
      <svg class="co-gen-svg" viewBox="0 0 156 48" width="156" height="48" aria-hidden="true">
        <text class="co-gen-letter" x="4"  y="36">P</text>
        <text class="co-gen-letter" x="36" y="36">U</text>
        <text class="co-gen-letter" x="68" y="36">L</text>
        <text class="co-gen-letter" x="100" y="36">S</text>
        <text class="co-gen-letter" x="132" y="36">E</text>
        <circle class="co-gen-dot co-gen-dot--upper" cx="50" cy="6"  r="5"/>
        <circle class="co-gen-dot co-gen-dot--lower" cx="118" cy="42" r="5"/>
      </svg>
      <span class="co-gen-label" id="co-gen-label">Generating document…</span>
    </div>`;
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>${isFirstCo ? "Assign" : "Update"} C/O — ${escapeHtml(trId)}</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body" id="co-modal-body">
      <p style="margin-top:0;color:var(--aewttr-muted);font-size:13px;">${isFirstCo ? "Enter the charge object number. A travel document will be generated and a Document Review will be created." : "Updating the C/O will generate a new document revision in the existing Document Review."}</p>
      <div class="form-row"><label>Travel C/O</label><input class="input-aewttr" id="travel-co-input" placeholder="e.g. 65-xxxx" value="${escapeHtml(r.chargeObject || "")}"></div>
      ${genLoaderHtml}
    </div>
    <div class="aewttr-modal-foot" id="co-modal-foot">
      <button class="btn-aewttr-ghost" id="co-cancel"${tip("Close without assigning")}>Cancel</button>
      <button class="btn-aewttr" id="co-confirm"${tip(isFirstCo ? "Assign charge object and generate travel document" : "Update charge object and generate new revision")}><i class="bx bx-receipt"></i> ${isFirstCo ? "Assign C/O" : "Update C/O"}</button>
    </div>
  `);
  function showGenerating(labelText) {
    const formRows = modal.querySelectorAll(".form-row, p");
    formRows.forEach(el => { el.hidden = true; });
    const genState = $("#co-gen-state", modal);
    const genLabel = $("#co-gen-label", modal);
    const foot = $("#co-modal-foot", modal);
    if (genState) genState.hidden = false;
    if (genLabel && labelText) genLabel.textContent = labelText;
    if (foot) foot.hidden = true;
    const closeBtn = $(".aewttr-modal-close", modal);
    if (closeBtn) closeBtn.hidden = true;
  }
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#co-cancel", modal).addEventListener("click", closeModal);
  $("#co-confirm", modal).addEventListener("click", async () => {
    const co = ($("#travel-co-input", modal).value || "").trim();
    if (!co) { toast("Charge object (C/O) is required.", "error"); return; }
    const confirmBtn = $("#co-confirm", modal);
    confirmBtn.disabled = true;
    showGenerating(isFirstCo ? "Generating travel document…" : "Generating new document revision…");
    const prevCo = r.chargeObject;
    r.chargeObject = co;
    if (r.engineeringForm) r.engineeringForm.travelCo = co;
    const now = new Date().toISOString();
    if (isFirstCo) {
      r.chargeObjectStatus = "Assigned";
      r.chargeObjectAssignedBy = (db.user && db.user.name) || "";
      r.chargeObjectAssignedAt = now;
      r._auditAction = "Assign C/O";
      r._auditSummary = `C/O ${co} assigned to ${trId} by ${r.chargeObjectAssignedBy}`;
    } else {
      r._auditAction = "Update C/O";
      r._auditSummary = `C/O changed from ${prevCo || "none"} to ${co} on ${trId}`;
    }
    await Repo.save("travelRequest", r);
    if (isSharePointMode()) {
      let docGenFailed = false;
      try {
        await saveTravelExportToSharePoint(r, { force: true });
        await Repo.save("travelRequest", r);
        if (isFirstCo && !r.docReviewId) {
          const docReview = await createTravelDocReview(r);
          r.docReviewId = docReview.id;
          await Repo.save("travelRequest", r);
        } else if (!isFirstCo && r.docReviewId) {
          const allDocs = Object.values(db.docs || {}).flat();
          const doc = allDocs.find(d => d.id === r.docReviewId);
          if (doc) {
            const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            const nextNum = typeof nextDocRevisionNumber === "function" ? nextDocRevisionNumber(doc) : ((doc.revisions || []).length + 1);
            const newRevision = {
              id: uid("REV"),
              number: nextNum,
              fileUrl: r.exportFileUrl || "",
              serverRelativeUrl: "",
              fileName: r.exportFileName || `${r.id}-travel-request.docx`,
              fileType: (r.exportFileName || "").endsWith(".docx") ? "docx" : "html",
              mimeType: r.exportMimeType || DOCX_MIME,
              uploadedBy: typeof currentUserIdentity === "function" ? (currentUserIdentity().name || "") : "",
              uploadedDate: typeof todayIsoDate === "function" ? todayIsoDate() : now.slice(0, 10),
              source: `C/O Updated to ${co}`
            };
            if (typeof resetReviewersForNewRevision === "function") resetReviewersForNewRevision(doc);
            doc.revisions = doc.revisions || [];
            doc.revisions.unshift(newRevision);
            if (typeof setDocActiveRevision === "function") setDocActiveRevision(doc, newRevision.id);
            if (typeof saveDocReview === "function") await saveDocReview(doc, { action: "Revision Added", text: `C/O updated to ${co} — new revision generated.` });
            if (typeof notifyPendingReviewers === "function") await notifyPendingReviewers(doc, { kind: "revision", force: true });
          }
        }
      } catch (e) {
        docGenFailed = true;
        console.warn("PULSE: C/O assign — document generation failed.", e);
        r._docGenFailed = true;
        await Repo.save("travelRequest", r);
        toast(`C/O ${co} saved. Document generation failed — use the "Retry" action on the request to regenerate.`, "warn");
        closeModal();
        refreshTravelNotifications();
        confirmBtn.disabled = false;
        if (onDone) onDone();
        return;
      }
      if (!docGenFailed) {
        await notifyCoAssigned(r, isFirstCo);
      }
    }
    closeModal();
    toast(`C/O ${co} ${isFirstCo ? "assigned" : "updated"} for ${trId}${r.exportFileUrl ? " — document generated" : ""}`, "success");
    refreshTravelNotifications();
    confirmBtn.disabled = false;
    if (onDone) onDone();
  });
}

async function recordCustomerConcurrence(trId, onDone) {
  const db = window.AEWTTR.db;
  const r = db.travelRequests.find(x => x.id === trId);
  if (!r || !canRecordConcurrence()) return;
  if (r.customerConcurrenceStatus === "Concurred") {
    toast("Customer concurrence is already recorded for this request.", "info");
    return;
  }
  const now = new Date().toISOString();
  r.customerConcurrenceStatus = "Concurred";
  r.customerConcurredBy = (db.user && db.user.name) || "";
  r.customerConcurredAt = now;
  r._auditAction = "Customer Concurrence";
  r._auditSummary = `Customer concurrence recorded for ${trId} by ${r.customerConcurredBy}`;
  Repo.save("travelRequest", r);
  await notifyConcurrenceRecorded(r);
  refreshTravelNotifications();
  toast(`${trId} — customer concurrence recorded`, "success");
  if (typeof onDone === "function") onDone();
}

/* onChanged is the caller's own redraw — the list passes renderRows. Actions in
   here must never re-render the whole route: renderTravelPage forces the list
   view from the route ("travel/mine" forces My Travel), so a full re-render
   throws an admin who had switched to All Travel back to My Travel. Redraw the
   rows in place instead. */
function openTravelDetailModal(trId, onChanged) {
  const db = window.AEWTTR.db;
  const r = db.travelRequests.find(x => x.id === trId);
  if (!r) return;
  markTravelRequestStatusSeen(trId, r.status);
  const isEngineering = (r.formMode || r.requestType) === "Engineering";
  const isLeave = travelCategory(r) === "Leave";
  const canEdit = canRequesterEditTravel(r);
  const canWithdraw = canRequesterWithdrawTravel(r);
  const canCancel = canCancelTravelRequest(r);
  const showConcur = r.status === "Submitted" && r.customerConcurrenceStatus === "Pending"
    && canRecordConcurrence()
    && (isLeave ? r.requiresConcurrence : true);
  const showAssignCo = r.status === "Submitted" && r.chargeObjectStatus === "Pending"
    && !isLeave && canAssignTravelCo();
  const showUpdateCo = r.chargeObjectStatus === "Assigned" && canAssignTravelCo();
  const travelers = (r.travelers || []).map((traveler) => traveler.name).filter(Boolean).join(", ");

  function concurrenceRow() {
    if (isLeave && !r.requiresConcurrence) return "";
    const status = r.customerConcurrenceStatus || "Pending";
    const by = r.customerConcurredBy ? ` by ${r.customerConcurredBy}` : "";
    const at = r.customerConcurredAt ? ` on ${r.customerConcurredAt.slice(0, 10)}` : "";
    return `<div class="travel-detail-row"><span class="k">Customer Concurrence</span><span class="v">${escapeHtml(status)}${status === "Concurred" ? escapeHtml(by + at) : ""}</span></div>`;
  }
  function coRow() {
    if (isLeave) return "";
    const status = r.chargeObjectStatus || "Pending";
    const co = r.chargeObject || "";
    const by = r.chargeObjectAssignedBy ? ` by ${r.chargeObjectAssignedBy}` : "";
    return `<div class="travel-detail-row"><span class="k">Charge Object</span><span class="v">${co ? escapeHtml(`${co} (${status}${by})`) : escapeHtml(status)}</span></div>`;
  }

  const modal = openModal(`
    <div class="aewttr-modal-head travel-detail-head">
      <div class="travel-detail-head-copy">
        <div class="travel-detail-kicker">${escapeHtml(r.formMode || r.requestType || "Travel")}</div>
        <h3>${escapeHtml(r.tripTitle || r.destination || "Travel request")}</h3>
        <div class="travel-detail-sub">${escapeHtml(r.id)} · ${escapeHtml(r.requester)}</div>
      </div>
      <div class="travel-detail-head-meta">
        ${travelStatusBadgeGroup(r)}
        <button class="aewttr-modal-close">&times;</button>
      </div>
    </div>
    <div class="aewttr-modal-body travel-detail-body">
      <div class="travel-detail-hero">
        <div class="travel-detail-hero-item">
          <span class="k">When</span>
          <strong>${escapeHtml(formatTravelDateRange(r))}</strong>
        </div>
        ${!isLeave ? `
        <div class="travel-detail-hero-item">
          <span class="k">Where</span>
          <strong>${escapeHtml(r.destination || "—")}</strong>
        </div>
        <div class="travel-detail-hero-item">
          <span class="k">Est. cost</span>
          <strong>$${Number(r.cost || 0).toLocaleString()}</strong>
        </div>` : ""}
      </div>

      <div class="travel-detail-sections">
        <section class="travel-detail-section">
          <h4>${isLeave ? "Requester" : "People"}</h4>
          <div class="travel-detail-rows">
            <div class="travel-detail-row"><span class="k">${isLeave ? "Person" : "Travelers"}</span><span class="v">${escapeHtml(isLeave ? (r.requester || travelers) : (travelers || r.requester))}</span></div>
          </div>
        </section>

        ${!isLeave ? `
        <section class="travel-detail-section">
          <h4>Administrative status</h4>
          <div class="travel-detail-rows">
            ${concurrenceRow()}
            ${coRow()}
            ${r.exportFileUrl ? `<div class="travel-detail-row"><span class="k">Document</span><span class="v"><a href="${escapeHtml(travelExportOpenUrl(r.exportFileUrl, r.exportFileName))}" target="_blank" rel="noopener">Open travel document</a></span></div>` : ""}
            ${r._docGenFailed ? `<div class="travel-detail-row"><span class="k">Document</span><span class="v" style="color:var(--aewttr-danger,#c0392b);">Generation failed — retry below</span></div>` : ""}
          </div>
        </section>
        <section class="travel-detail-section">
          <h4>Trip details</h4>
          <div class="travel-detail-rows">
            <div class="travel-detail-row"><span class="k">Type</span><span class="v">${escapeHtml(r.type || "—")}</span></div>
          </div>
        </section>` : (r.requiresConcurrence ? `
        <section class="travel-detail-section">
          <h4>Administrative status</h4>
          <div class="travel-detail-rows">
            ${concurrenceRow()}
          </div>
        </section>` : "")}

        <section class="travel-detail-section">
          <h4>${isLeave ? "Notes" : "Context"}</h4>
          <div class="travel-detail-rows">
            ${isLeave
              ? `<div class="travel-detail-row travel-detail-row--stack"><span class="k">Notes</span><span class="v">${escapeHtml(r.notes || "—")}</span></div>`
              : `<div class="travel-detail-row travel-detail-row--stack"><span class="k">Purpose</span><span class="v">${escapeHtml(r.purpose || "—")}</span></div>
            ${r.impactIfNotApproved ? `<div class="travel-detail-row travel-detail-row--stack"><span class="k">Impact</span><span class="v">${escapeHtml(r.impactIfNotApproved)}</span></div>` : ""}
            ${r.notes ? `<div class="travel-detail-row travel-detail-row--stack"><span class="k">Notes</span><span class="v">${escapeHtml(r.notes)}</span></div>` : ""}`}
          </div>
        </section>
      </div>
    </div>
    <div class="aewttr-modal-foot travel-detail-foot">
      <div class="travel-detail-foot-actions">
        ${canEdit ? travelActionBtn("edit", { elementId: "td-edit", tip: r.chargeObjectStatus === "Assigned" ? "Update this request — a new document revision will be generated" : "Edit this request" }) : ""}
        ${canWithdraw ? travelActionBtn("revoke", { elementId: "td-revoke", tone: "danger", tip: "Withdraw this request" }) : ""}
        ${showConcur ? travelActionBtn("concur", { elementId: "td-concur", tone: "primary", tip: isLeave ? "Record customer concurrence for this leave request" : "Record customer concurrence" }) : ""}
        ${showAssignCo ? travelActionBtn("assign-co", { elementId: "td-assign-co", tone: "primary", tip: "Assign charge object and generate travel document" }) : ""}
        ${showUpdateCo && !showAssignCo ? travelActionBtn("assign-co", { elementId: "td-assign-co", label: "Update C/O", tip: "Change the charge object and generate a new document revision" }) : ""}
        ${r._docGenFailed ? travelActionBtn("export", { elementId: "td-retry-doc", label: "Retry Document", tone: "primary", tip: "Retry generating the travel document" }) : ""}
        ${canCancel && !canWithdraw ? travelActionBtn("cancel", { elementId: "td-cancel-travel", tip: canAdminCancelTravel(r) ? "Cancel this travel request" : "Cancel this travel request", tone: "danger" }) : ""}
        ${travelDebriefRowButtons(r)}
        ${travelActionBtn("export", { elementId: "td-download-docx", label: "Download DOCX", icon: "bx-download", tip: "Generate and download the travel request document — available at any stage" })}
        ${canAdminSetTravelStatus(r) ? travelActionBtn("set-status", { elementId: "td-set-status", tip: "Correct this request's status after the fact" }) : ""}
        ${canAdminSetTravelStatus(r) && travelDebriefOutstanding(r) ? travelActionBtn("nudge-debrief", { elementId: "td-nudge-debrief", tip: "Remind the travellers that a debrief is still outstanding" }) : ""}
        ${canAdminDeleteTravel(r) ? travelActionBtn("delete-request", { elementId: "td-delete-request", tone: "danger", tip: "Permanently delete this request — use Withdraw or Cancel unless this is a duplicate or test entry" }) : ""}
        ${r.exportFileUrl ? travelActionBtn("open-export", { elementId: "td-open-export", label: "Open Document", icon: "bx-desktop", tip: "Open the filed travel document in Word" }) : ""}
      </div>
      <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="td-close">Close</button>
    </div>
  `, { className: "travel-detail-modal" });
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  const editBtn = $("#td-edit", modal);
  if (editBtn) editBtn.addEventListener("click", () => { closeModal(); navigate(`travel/edit/${trId}`); });
  const revokeBtn = $("#td-revoke", modal);
  if (revokeBtn) revokeBtn.addEventListener("click", async () => { closeModal(); await revokeTravelRequest(trId); });
  const concurBtn = $("#td-concur", modal);
  if (concurBtn) concurBtn.addEventListener("click", async () => {
    closeModal();
    await recordCustomerConcurrence(trId);
    refreshTravelNotifications();
  });
  const assignCoBtn = $("#td-assign-co", modal);
  if (assignCoBtn) assignCoBtn.addEventListener("click", () => { closeModal(); openCoAssignModal(trId); });
  const retryDocBtn = $("#td-retry-doc", modal);
  if (retryDocBtn) retryDocBtn.addEventListener("click", async () => {
    retryDocBtn.disabled = true;
    try {
      await saveTravelExportToSharePoint(r, { force: true });
      await Repo.save("travelRequest", r);
      if (!r.docReviewId) {
        const docReview = await createTravelDocReview(r);
        r.docReviewId = docReview.id;
        await Repo.save("travelRequest", r);
      }
      delete r._docGenFailed;
      await Repo.save("travelRequest", r);
      closeModal();
      toast(`${trId} — document generated successfully.`, "success");
      openTravelDetailModal(trId);
    } catch (e) {
      toast((e && e.message) || "Document generation still failing.", "error");
      retryDocBtn.disabled = false;
    }
  });
  /* A requester needs their own copy of the form long before finance assigns a
     charge object, which is when the filed document is generated. This builds
     the same DOCX from the same template client-side and hands it straight to
     the browser: no SharePoint write, no revision, no effect on the document
     review. The filed copy is still what the review tracks. */
  const downloadDocxBtn = $("#td-download-docx", modal);
  if (downloadDocxBtn) downloadDocxBtn.addEventListener("click", async () => {
    if (typeof window.createTravelDocxBlob !== "function") {
      toast("The travel document generator is not loaded on this page.", "error");
      return;
    }
    const original = downloadDocxBtn.innerHTML;
    downloadDocxBtn.disabled = true;
    try {
      const blob = await window.createTravelDocxBlob(r);
      const baseId = String(r.id || "travel-request").replace(/[^A-Za-z0-9._-]+/g, "-");
      const mode = String(r.formMode || "Standard").toLowerCase().replace(/\s+/g, "-");
      downloadBlobAsFile(blob, `${baseId}-${mode}-travel-request.docx`);
      toast("Travel request document downloaded.", "success");
    } catch (e) {
      console.error("PULSE: travel DOCX download failed.", e);
      toast("Could not generate the travel document. Report an issue if this continues.", "error");
    } finally {
      downloadDocxBtn.disabled = false;
      downloadDocxBtn.innerHTML = original;
    }
  });

  /* Change a status after the fact. An inline select rather than a modal —
     the correction is one field and a modal on top of a modal is worse. */
  const setStatusBtn = $("#td-set-status", modal);
  if (setStatusBtn) setStatusBtn.addEventListener("click", () => {
    if (modal.querySelector(".td-status-editor")) return;
    const editor = document.createElement("div");
    editor.className = "td-status-editor";
    editor.innerHTML =
      '<label>Set status to</label>' +
      '<select class="select-aewttr" id="td-status-pick">' +
      TRAVEL_ADMIN_STATUSES.map((st) => `<option value="${escapeHtml(st)}"${st === r.status ? " selected" : ""}>${escapeHtml(st)}</option>`).join("") +
      '</select>' +
      '<button type="button" class="btn-aewttr btn-aewttr-sm" id="td-status-apply">Apply</button>' +
      '<button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="td-status-cancel">Cancel</button>';
    setStatusBtn.parentElement.insertBefore(editor, setStatusBtn.nextSibling);
    $("#td-status-cancel", modal).addEventListener("click", () => editor.remove());
    $("#td-status-apply", modal).addEventListener("click", async () => {
      const next = $("#td-status-pick", modal).value;
      if (!next || next === r.status) { editor.remove(); return; }
      const previous = r.status;
      r.status = next;
      r.statusChangedBy = (db.user && db.user.name) || "";
      r.statusChangedAt = new Date().toISOString();
      r._auditAction = "Set Status";
      r._auditSummary = `${r.id} status changed from ${previous} to ${next} by ${r.statusChangedBy}`;
      await Repo.save("travelRequest", r);
      markTravelRequestStatusSeen(r.id, r.status);
      notifyRequesterTravelStatusUpdate(r);
      refreshTravelNotifications();
      toast(`${r.id} set to ${next}.`, "success");
      /* Which footer actions apply is computed from the status when the modal
         renders, so patching a label in place would leave the rest stale.
         Reopen the same request instead: the route never changes, so the admin
         stays exactly where they were, and the record shows its new state. */
      if (typeof onChanged === "function") onChanged();
      closeModal();
      openTravelDetailModal(trId, onChanged);
    });
  });

  /* Nudge every traveller on the trip, not just the requester — the debrief is
     owed per traveller. */
  const nudgeBtn = $("#td-nudge-debrief", modal);
  if (nudgeBtn) nudgeBtn.addEventListener("click", async () => {
    nudgeBtn.disabled = true;
    const sent = await notifyTravelDebriefNudge(r);
    nudgeBtn.disabled = false;
    if (sent) {
      r._auditAction = "Nudge Debrief";
      r._auditSummary = `Debrief reminder sent for ${r.id}`;
      await Repo.save("travelRequest", r);
      toast("Debrief reminder sent.", "success");
    } else {
      toast(isSharePointMode() ? "No traveller email addresses on this request." : "Reminders are only sent in SharePoint mode.", "error");
    }
  });

  /* Delete is confirm-then-act on the button itself, matching how ticket
     deletion works elsewhere, so there is no accidental single-click removal. */
  const deleteBtn = $("#td-delete-request", modal);
  if (deleteBtn) deleteBtn.addEventListener("click", async () => {
    if (deleteBtn.dataset.confirmed !== "1") {
      deleteBtn.dataset.confirmed = "1";
      const original = deleteBtn.innerHTML;
      deleteBtn.innerHTML = '<i class="bx bx-trash"></i><span>Confirm delete</span>';
      setTimeout(() => {
        if (deleteBtn.dataset.confirmed === "1") { deleteBtn.dataset.confirmed = ""; deleteBtn.innerHTML = original; }
      }, 4000);
      return;
    }
    deleteBtn.disabled = true;
    const idx = (db.travelRequests || []).findIndex((x) => x.id === r.id);
    if (idx !== -1) db.travelRequests.splice(idx, 1);
    r._auditAction = "Delete";
    r._auditSummary = `${r.id} deleted by ${(db.user && db.user.name) || ""}`;
    try {
      if (typeof Repo !== "undefined" && Repo && typeof Repo.remove === "function") await Repo.remove("travelRequest", r);
    } catch (e) {
      console.warn("PULSE: travel request delete failed.", e);
    }
    refreshTravelNotifications();
    toast(`${r.id} deleted.`, "success");
    closeModal();
    if (typeof onChanged === "function") onChanged();
  });

  const openDocBtn = $("#td-open-export", modal);
  if (openDocBtn) openDocBtn.addEventListener("click", () => openTravelDocByPolicy(r));
  const cancelBtn = $("#td-cancel-travel", modal);
  if (cancelBtn) cancelBtn.addEventListener("click", async () => { closeModal(); await cancelTravelRequest(trId); });
  $("#td-close", modal).addEventListener("click", closeModal);
  wireTravelDebriefActions(modal);
}

/* When a trip has multiple travelers, each one's debrief is a separate
   record — this modal shows a switcher row of pills (one per filed
   debrief) so you can click between them without closing and reopening. */
function openDebriefDetailModal(trId, preferredDebriefId) {
  const debriefs = getDebriefsForTravel(trId);
  const request = (window.AEWTTR.db.travelRequests || []).find((r) => r.id === trId);
  if (!debriefs.length || !request) {
    toast("No debrief found for this trip.", "error");
    return;
  }
  let activeId = (preferredDebriefId && debriefs.some((d) => d.id === preferredDebriefId)) ? preferredDebriefId : debriefs[0].id;

  function bodyHtml() {
    const debrief = debriefs.find((d) => d.id === activeId) || debriefs[0];
    const switcherHtml = debriefs.length > 1 ? `
      <div class="travel-debrief-switcher" role="tablist">
        ${debriefs.map((d) => `
          <button type="button" class="travel-debrief-switch-pill ${d.id === activeId ? "active" : ""}" data-switch-debrief="${escapeHtml(d.id)}" role="tab" aria-selected="${d.id === activeId}">
            ${escapeHtml(d.travelerName || d.id)}
          </button>
        `).join("")}
      </div>
    ` : "";
    return `
      <div class="aewttr-modal-head travel-detail-head">
        <div class="travel-detail-head-copy">
          <div class="travel-detail-kicker">Travel debrief</div>
          <h3>${escapeHtml(debrief.id)}${debrief.travelerName ? ` — ${escapeHtml(debrief.travelerName)}` : ""}</h3>
          <div class="travel-detail-sub">${escapeHtml(trId)} · ${escapeHtml(request.destination || request.tripTitle || "Trip")}</div>
        </div>
        <button class="aewttr-modal-close">&times;</button>
      </div>
      <div class="aewttr-modal-body travel-detail-body">
        ${switcherHtml}
        <div class="travel-detail-hero">
          <div class="travel-detail-hero-item"><span class="k">Trip dates</span><strong>${escapeHtml(debrief.dates || `${fmtDate(request.start)} – ${fmtDate(request.end)}`)}</strong></div>
          <div class="travel-detail-hero-item"><span class="k">Classification</span><strong>${escapeHtml(debrief.classification || "—")}</strong></div>
          <div class="travel-detail-hero-item"><span class="k">Systems visited</span><strong>${escapeHtml(debrief.systems || "—")}</strong></div>
        </div>
        <div class="travel-detail-sections">
          <section class="travel-detail-section">
            <h4>Technical summary</h4>
            <div class="travel-debrief-prose">${escapeHtml(debrief.summary || "—").replace(/\n/g, "<br>")}</div>
          </section>
          <section class="travel-detail-section">
            <h4>Issues / follow-up</h4>
            <div class="travel-debrief-prose">${escapeHtml(debrief.followup || "—").replace(/\n/g, "<br>")}</div>
          </section>
        </div>
      </div>
      <div class="aewttr-modal-foot travel-detail-foot">
        <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="db-view-travel">View travel request</button>
        <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="db-close">Close</button>
      </div>
    `;
  }

  function wire(modal) {
    $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
    $("#db-close", modal).addEventListener("click", closeModal);
    $("#db-view-travel", modal).addEventListener("click", () => { closeModal(); openTravelDetailModal(trId); });
    $all("[data-switch-debrief]", modal).forEach((btn) => btn.addEventListener("click", () => {
      activeId = btn.dataset.switchDebrief;
      modal.innerHTML = bodyHtml();
      wire(modal);
    }));
  }

  const modal = openModal(bodyHtml(), { className: "travel-detail-modal travel-debrief-modal" });
  wire(modal);
}

function drawSubmitDebrief(body) {
  const db = window.AEWTTR.db;
  const routeQuery = (typeof currentRoute === "function" ? currentRoute().query : {}) || {};
  const initialTr = routeQuery.tr || "";
  // A trip only drops off this list once *you* have filed for it — other
  // travelers on the same request filing their own debrief doesn't hide it.
  const approved = db.travelRequests.filter((r) => travelNeedsDebrief(r) && (!getMyDebriefForTravel(r.id) || r.id === initialTr));
  const selected = initialTr ? approved.find((r) => r.id === initialTr) : null;
  body.innerHTML = `
    <div class="travel-debrief-page">
      <div class="travel-debrief-main aewttr-card aewttr-card-pad">
        <div class="travel-debrief-head">
          <div>
            <div class="travel-wizard-eyebrow">Post-trip reporting</div>
            <h2 class="travel-debrief-title">Submit a travel debrief</h2>
            <p class="travel-debrief-lead">Link the debrief to an approved trip, summarize what you saw, and note any follow-up actions.</p>
          </div>
        </div>

        <div class="travel-debrief-form">
          <section class="travel-debrief-section">
            <h3>Trip link</h3>
            <div class="form-row">
              <label>Approved travel request</label>
              <select class="select-aewttr" id="db-tr">
                <option value="">Select an approved trip…</option>
                ${approved.map((r) => `<option value="${r.id}" ${r.id === initialTr ? "selected" : ""}>${r.id} — ${escapeHtml(r.destination || r.tripTitle || "Trip")} (${fmtDate(r.start)})</option>`).join("")}
              </select>
            </div>
            <div class="travel-debrief-trip-card" id="db-trip-card" ${selected ? "" : "hidden"}>
              <div class="travel-debrief-trip-card-row"><span class="k">When</span><strong id="db-dates">${selected ? `${fmtDate(selected.start)} – ${fmtDate(selected.end)}` : ""}</strong></div>
              <div class="travel-debrief-trip-card-row"><span class="k">Where</span><strong id="db-where">${selected ? escapeHtml(selected.destination || "—") : ""}</strong></div>
              <div class="travel-debrief-trip-card-row"><span class="k">Travelers</span><strong id="db-travelers">${selected ? escapeHtml((selected.travelers || []).map((t) => t.name).filter(Boolean).join(", ") || selected.requester) : ""}</strong></div>
              <div class="travel-debrief-trip-card-row"><span class="k">Debriefs filed</span><strong id="db-filed-count">${selected ? `${getDebriefsForTravel(selected.id).length}/${travelRequestTravelerRoster(selected).length}` : ""}</strong></div>
            </div>
          </section>

          <section class="travel-debrief-section">
            <h3>Visit details</h3>
            <div class="form-grid-2">
              <div class="form-row"><label>Systems / locations visited</label><input class="input-aewttr" id="db-systems" placeholder="e.g. Windchill, Range 4 lab"></div>
              <div class="form-row"><label>Trip classification</label>
                <select class="select-aewttr" id="db-class">
                  <option>Non-Contractor-Forward</option>
                  <option>Contractor-Forward</option>
                </select>
              </div>
            </div>
          </section>

          <section class="travel-debrief-section">
            <h3>Summary</h3>
            <div class="form-row"><label>Technical summary</label><textarea class="textarea-aewttr travel-debrief-textarea" id="db-summary" placeholder="What did you accomplish, observe, or verify on this trip?"></textarea></div>
            <div class="form-row"><label>Issues / follow-up actions</label><textarea class="textarea-aewttr travel-debrief-textarea travel-debrief-textarea--short" id="db-followup" placeholder="Open items, risks, or next steps"></textarea></div>
          </section>

          <section class="travel-debrief-section">
            <h3>Attachments</h3>
            <div class="form-row"><label>Supporting files</label><input type="file" class="input-aewttr" id="db-files" multiple></div>
          </section>
        </div>

        <div class="travel-debrief-actions">
          <button type="button" class="btn-aewttr" id="db-submit"${tip("Submit travel debrief for the selected trip")}><i class="bx bx-send"></i> Submit debrief</button>
        </div>
      </div>

      <aside class="travel-debrief-aside aewttr-card aewttr-card-pad" data-help>
        <div class="side-panel-title">What happens next</div>
        <div class="side-step"><div class="num">1</div><div class="stext"><strong>Saved to SharePoint</strong>Your debrief is filed for ESDP handoff and future reference.</div></div>
        <div class="side-step"><div class="num">2</div><div class="stext"><strong>Linked to the trip</strong>A View Debrief button appears on the travel record in My Travel and All Travel.</div></div>
        <div class="side-step"><div class="num">3</div><div class="stext"><strong>Linda notified</strong>She gets a direct link to the filed debrief.</div></div>
      </aside>
    </div>
  `;

  function syncTripCard() {
    const trId = $("#db-tr", body).value;
    const r = approved.find((x) => x.id === trId);
    const card = $("#db-trip-card", body);
    if (!r) { card.hidden = true; return; }
    card.hidden = false;
    $("#db-dates", body).textContent = formatTravelDateRange(r);
    $("#db-where", body).innerHTML = escapeHtml(r.destination || "—");
    $("#db-travelers", body).textContent = (r.travelers || []).map((t) => t.name).filter(Boolean).join(", ") || r.requester;
    $("#db-filed-count", body).textContent = `${getDebriefsForTravel(r.id).length}/${travelRequestTravelerRoster(r).length}`;
  }

  $("#db-tr", body).addEventListener("change", syncTripCard);
  syncTripCard();

  $("#db-submit", body).addEventListener("click", async () => {
    const trId = $("#db-tr", body).value;
    if (!trId) { toast("Select a linked travel request", "error"); return; }
    if (getMyDebriefForTravel(trId)) { toast("You already filed a debrief for this trip.", "error"); return; }
    const summary = $("#db-summary", body).value.trim();
    if (!summary) { toast("Technical summary is required.", "error"); return; }
    const request = db.travelRequests.find((r) => r.id === trId);
    const dbid = "DB-" + String((db.debriefs || []).length + 31).padStart(4, "0");
    const debrief = {
      id: dbid,
      trId,
      travelerName: db.user.name || "",
      travelerEmail: db.user.email || "",
      dates: request ? `${request.start} to ${request.end}` : "",
      systems: $("#db-systems", body).value.trim(),
      classification: $("#db-class", body).value,
      summary,
      followup: $("#db-followup", body).value.trim(),
      files: []
    };
    db.debriefs.unshift(debrief);
    await Repo.save("travelDebrief", debrief);
    if (request) {
      request.debriefId = dbid;
      await Repo.save("travelRequest", request);
    }
    toast(`${dbid} saved — linked to ${trId}`, "success");
    navigate("travel/all");
  });
}

const TEAM_EVENT_COLOR = "#0097A7";

function drawTeamEventsList(body) {
  const db = window.AEWTTR.db;
  if (!db.teamEvents) db.teamEvents = [];
  const events = [...db.teamEvents].sort((a, b) => (a.start || "").localeCompare(b.start || ""));
  const canEdit = typeof canCurrentUserEdit === "function" ? canCurrentUserEdit() : true;

  body.innerHTML = `
    <div class="aewttr-card aewttr-card-pad">
      <div class="toolbar-row" style="margin-bottom:16px;">
        <div>
          <div class="side-panel-title">Team Events</div>
          <p style="margin:3px 0 0;color:var(--aewttr-muted);font-size:13px;">Non-travel calendar events visible to the whole team.</p>
        </div>
        ${canEdit ? `<button type="button" class="btn-aewttr" id="te-new-btn"><i class="bx bx-plus"></i> New Event</button>` : ""}
      </div>
      ${events.length ? `
        <table class="aewttr-table">
          <thead><tr><th>Title</th><th>Date</th><th>Time</th><th>Location</th><th>Description</th><th></th></tr></thead>
          <tbody>
            ${events.map((te) => `
              <tr>
                <td><strong>${escapeHtml(te.title)}</strong></td>
                <td style="white-space:nowrap;">${escapeHtml(te.start)}</td>
                <td style="white-space:nowrap;">${te.allDay === false && te.startTime ? escapeHtml(te.startTime) + (te.endTime ? " – " + escapeHtml(te.endTime) : "") : "All day"}</td>
                <td>${escapeHtml(te.location || "—")}</td>
                <td>${escapeHtml((te.description || "").slice(0, 70))}${(te.description || "").length > 70 ? "…" : ""}</td>
                <td class="travel-actions-cell">
                  ${canEdit ? `<button class="btn-aewttr-outline btn-aewttr-sm" data-te-edit="${escapeHtml(te.id)}">Edit</button>` : ""}
                </td>
              </tr>`).join("")}
          </tbody>
        </table>` : `<div class="empty-state" style="padding:32px 0;">No team events yet — click "New Event" to add one.</div>`}
    </div>
  `;

  const newBtn = $("#te-new-btn", body);
  if (newBtn) newBtn.addEventListener("click", () => navigate("travel/events/new"));
  $all("[data-te-edit]", body).forEach((btn) => {
    btn.addEventListener("click", () => navigate("travel/events/edit/" + btn.dataset.teEdit));
  });
}

function drawTeamEventForm(body, eventId) {
  const db = window.AEWTTR.db;
  if (!db.teamEvents) db.teamEvents = [];
  const existing = eventId ? db.teamEvents.find((e) => e.id === eventId) : null;
  const isEdit = !!existing;
  const ev = existing || {};
  const canEdit = typeof canCurrentUserEdit === "function" ? canCurrentUserEdit() : true;

  body.innerHTML = `
    <div style="max-width:640px;">
      <div style="margin-bottom:16px;">
        <button type="button" class="btn-aewttr-ghost" id="te-back"><i class="bx bx-arrow-back"></i> Back to Events</button>
      </div>
      <div class="aewttr-card aewttr-card-pad">
        <h3 style="margin:0 0 18px;">${isEdit ? "Edit Team Event" : "New Team Event"}</h3>
        <div class="form-row"><label>Title <span class="required-star">*</span></label><input class="input-aewttr" id="te-title" placeholder="Event name" value="${escapeHtml(ev.title || "")}"></div>
        <div class="form-row"><label>Event date <span class="required-star">*</span></label><input class="input-aewttr" id="te-start" type="date" value="${ev.start || ""}"></div>
        <label class="ssp-link-row"><input type="checkbox" id="te-all-day" ${ev.allDay !== false ? "checked" : ""}> <span>All day <small>— uncheck to add start and end times</small></span></label>
        <div class="form-grid-2" id="te-time-row"${ev.allDay !== false ? " hidden" : ""}>
          <div><label>Start time</label><input class="input-aewttr" id="te-start-time" type="time" value="${ev.startTime || ""}"></div>
          <div><label>End time</label><input class="input-aewttr" id="te-end-time" type="time" value="${ev.endTime || ""}"></div>
        </div>
        <div class="form-row"><label>Location</label><input class="input-aewttr" id="te-location" placeholder="Building, room, or virtual link…" value="${escapeHtml(ev.location || "")}"></div>
        <div class="form-row"><label>Description / notes</label><textarea class="textarea-aewttr" id="te-desc" rows="4" placeholder="Agenda, details, or any other context…">${escapeHtml(ev.description || "")}</textarea></div>
        <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:8px;">
          ${isEdit && canEdit ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-danger" id="te-delete" style="margin-right:auto;"><i class="bx bx-trash"></i> Delete</button>` : ""}
          <button type="button" class="btn-aewttr-ghost" id="te-cancel">Cancel</button>
          ${canEdit ? `<button type="button" class="btn-aewttr" id="te-save">${isEdit ? "Save Changes" : "Add Event"}</button>` : ""}
        </div>
      </div>
    </div>
  `;

  $("#te-back", body).addEventListener("click", () => navigate("travel/events"));
  const cancelBtn = $("#te-cancel", body);
  if (cancelBtn) cancelBtn.addEventListener("click", () => navigate("travel/events"));
  const allDayToggle = $("#te-all-day", body);
  const timeRow = $("#te-time-row", body);
  if (allDayToggle && timeRow) allDayToggle.addEventListener("change", () => { timeRow.hidden = allDayToggle.checked; });

  const deleteBtn = $("#te-delete", body);
  if (deleteBtn) deleteBtn.addEventListener("click", async () => {
    if (!confirm("Delete this event?")) return;
    db.teamEvents = db.teamEvents.filter((e) => e.id !== eventId);
    try { await Repo.remove("teamEvent", existing); } catch (_) {
      db.teamEvents.push(existing);
      toast("Could not delete event. Please try again.", "error");
      return;
    }
    toast("Event deleted", "success");
    navigate("travel/events");
  });

  const saveBtn = $("#te-save", body);
  if (saveBtn) saveBtn.addEventListener("click", async () => {
    const title = ($("#te-title", body).value || "").trim();
    const start = ($("#te-start", body).value || "").trim();
    if (!title || !start) { toast("Title and start date are required.", "error"); return; }
    saveBtn.disabled = true;
    const payload = {
      id: existing ? existing.id : uid("te"),
      _spId: existing ? existing._spId : undefined,
      title,
      start,
      end: start,
      allDay: !!allDayToggle?.checked,
      startTime: allDayToggle?.checked ? "" : ($("#te-start-time", body).value || "").trim(),
      endTime: allDayToggle?.checked ? "" : ($("#te-end-time", body).value || "").trim(),
      description: ($("#te-desc", body).value || "").trim(),
      location: ($("#te-location", body).value || "").trim(),
      updatedAt: new Date().toISOString()
    };
    const previous = isEdit ? db.teamEvents.find((e) => e.id === eventId) : null;
    if (isEdit) {
      const idx = db.teamEvents.findIndex((e) => e.id === eventId);
      if (idx !== -1) db.teamEvents[idx] = payload;
    } else {
      db.teamEvents.push(payload);
    }
    try { await Repo.save("teamEvent", payload, { immediate: true }); } catch (_) {
      if (isEdit) {
        const idx = db.teamEvents.findIndex((e) => e.id === eventId);
        if (idx !== -1 && previous) db.teamEvents[idx] = previous;
      } else {
        db.teamEvents = db.teamEvents.filter((e) => e !== payload);
      }
      toast("Could not save event to SharePoint. Please try again.", "error");
      saveBtn.disabled = false;
      return;
    }
    toast(isEdit ? "Event updated" : "Event added", "success");
    navigate("travel/events");
  });
}

function openTeamEventModal(eventId, onDone) {
  const db = window.AEWTTR.db;
  if (!db.teamEvents) db.teamEvents = [];
  const existing = eventId ? db.teamEvents.find(e => e.id === eventId) : null;
  const isEdit = !!existing;
  const ev = existing || {};
  const canEdit = typeof canCurrentUserEdit === "function" ? canCurrentUserEdit() : true;

  const modal = openModal(`
    <div class="aewttr-modal-head">
      <h3>${isEdit ? "Edit Event" : "New Team Event"}</h3>
      <button class="aewttr-modal-close" type="button">&times;</button>
    </div>
    <div class="aewttr-modal-body">
      <div class="form-row"><label>Title <span class="required-star">*</span></label><input class="input-aewttr" id="te-title" placeholder="Event title" value="${escapeHtml(ev.title || "")}"></div>
      <div class="form-row form-row-2col">
        <div><label>Start date <span class="required-star">*</span></label><input class="input-aewttr" id="te-start" type="date" value="${ev.start || ""}"></div>
        <div><label>End date</label><input class="input-aewttr" id="te-end" type="date" value="${ev.end || ""}"></div>
      </div>
      <div class="form-row"><label>Description</label><textarea class="textarea-aewttr" id="te-desc" rows="3" placeholder="Optional details…">${escapeHtml(ev.description || "")}</textarea></div>
      <div class="form-row" style="margin-bottom:0;"><label>Location</label><input class="input-aewttr" id="te-location" placeholder="Building, room, or virtual link…" value="${escapeHtml(ev.location || "")}"></div>
    </div>
    <div class="aewttr-modal-foot">
      ${isEdit && canEdit ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-danger" id="te-delete"><i class="bx bx-trash"></i> Delete</button>` : ""}
      <button type="button" class="btn-aewttr-ghost" id="te-cancel">Cancel</button>
      ${canEdit ? `<button type="button" class="btn-aewttr" id="te-save">${isEdit ? "Save" : "Add Event"}</button>` : ""}
    </div>
  `);

  function close() { modal.closest(".aewttr-modal-backdrop")?.remove(); }

  const cancelBtn = $("#te-cancel", modal);
  if (cancelBtn) cancelBtn.addEventListener("click", close);

  const deleteBtn = $("#te-delete", modal);
  if (deleteBtn) deleteBtn.addEventListener("click", async () => {
    if (!confirm("Delete this event?")) return;
    db.teamEvents = db.teamEvents.filter(e => e.id !== eventId);
    try { if (typeof aewttrSaveStore === "function") aewttrSaveStore(); } catch (err) {}
    close();
    if (typeof onDone === "function") onDone();
  });

  const saveBtn = $("#te-save", modal);
  if (saveBtn) saveBtn.addEventListener("click", async () => {
    const title = ($("#te-title", modal).value || "").trim();
    const start = ($("#te-start", modal).value || "").trim();
    if (!title || !start) {
      toast("Title and start date are required.", "error");
      return;
    }
    saveBtn.disabled = true;
    const payload = {
      id: existing ? existing.id : uid("te"),
      title,
      start,
      end: ($("#te-end", modal).value || "").trim() || start,
      description: ($("#te-desc", modal).value || "").trim(),
      location: ($("#te-location", modal).value || "").trim(),
      updatedAt: new Date().toISOString()
    };
    if (isEdit) {
      const idx = db.teamEvents.findIndex(e => e.id === eventId);
      if (idx !== -1) db.teamEvents[idx] = payload;
    } else {
      db.teamEvents.push(payload);
    }
    try {
      if (typeof aewttrSaveStore === "function") aewttrSaveStore();
    } catch (err) {
      toast("Could not save event.", "error");
      saveBtn.disabled = false;
      return;
    }
    close();
    if (typeof onDone === "function") onDone();
  });
}

function drawTravelCalendar(body) {
  const db = window.AEWTTR.db;
  if (!db.teamEvents) db.teamEvents = [];
  if (window.AEWTTR.state.travelCalendarShowContractor == null) window.AEWTTR.state.travelCalendarShowContractor = true;
  const showContractor = window.AEWTTR.state.travelCalendarShowContractor;
  // Show everything except requests that never happened (Cancelled/Withdrawn).
  // A submitted request is visible on the calendar right away.
  const visible = db.travelRequests.filter(r => !["Cancelled", "Withdrawn"].includes(r.status));
  const legendTypes = ["Leave", "TDY", "Conference", "Training", "Contractor", "Other"];
  const canEdit = typeof canCurrentUserEdit === "function" ? canCurrentUserEdit() : true;

  body.innerHTML = `
    <div class="travel-calendar-page">
      <header class="travel-calendar-head">
        <div class="travel-calendar-copy">
          <div class="travel-calendar-label">Team calendar</div>
          <h2>Plan the month at a glance</h2>
          <p>Submitted and active travel, leave, training, conferences, and team events in one shared schedule.</p>
        </div>
        <label class="ps-toggle-row travel-calendar-toggle"><input type="checkbox" id="cal-show-contractor" ${showContractor ? "checked" : ""}> <span>Show contractor travel</span></label>
      </header>
      <div class="travel-calendar-legend" aria-label="Calendar legend">
        <div class="cal-legend">
          ${legendTypes.map((type) => `<span class="li"><span class="dot" style="background:${TRAVEL_CALENDAR_COLORS[type]}"></span>${escapeHtml(type)}</span>`).join("")}
          <span class="li"><span class="dot" style="background:${TEAM_EVENT_COLOR};"></span>Team Event</span>
          <span class="li"><span class="dot" style="background:#9AA4B2;"></span>Pending C/O assignment</span>
        </div>
      </div>
      <section class="travel-calendar-surface">
        <div id="fc-cal"></div>
      </section>
    </div>
  `;

  const events = [];
  visible.forEach((r) => {
    const isContractor = travelCalendarTypeKey(r) === "Contractor";
    if (isContractor && !showContractor) return;
    const travelers = (r.travelers && r.travelers.length)
      ? r.travelers
      : [{ name: r.requester, email: r.requesterEmail || "" }];
    // Leave is visible immediately as a scheduled absence. Non-leave shows
    // as pending (grey) until a charge object is assigned.
    const isPending = travelCalendarTypeKey(r) !== "Leave" && r.chargeObjectStatus !== "Assigned";
    const color = isPending ? "#9AA4B2" : travelCalendarColor(r);
    const typeLabel = travelCalendarTypeKey(r);
    const namesLabel = travelCalendarTravelerLabel(travelers, r.requester);
    // One event per request (not per traveler) — a 4-person TDY used to
    // produce 4 separate calendar blocks for the same trip.
    const isAllDay = travelRequestIsAllDay(r);
    const event = {
      id: r.id,
      title: `${namesLabel} — ${typeLabel}${(r.destination && r.destination !== typeLabel) ? ` · ${r.destination}` : ""}${isPending ? " (Pending)" : ""}`,
      color,
      classNames: isPending ? ["travel-cal-event-pending"] : [],
      extendedProps: { r, isPending, _kind: "travel" }
    };
    if (isAllDay) {
      event.allDay = true;
      event.start = r.start;
      event.end = nextDay(r.end);
    } else {
      event.allDay = false;
      event.start = `${r.start}T${r.startTime || "00:00"}:00`;
      event.end = `${r.end}T${r.endTime || "23:59"}:00`;
    }
    events.push(event);
  });

  db.teamEvents.forEach((te) => {
    const event = {
      id: te.id,
      title: te.title,
      color: TEAM_EVENT_COLOR,
      classNames: ["team-cal-event"],
      extendedProps: { te, _kind: "teamEvent" }
    };
    if (te.allDay === false) {
      event.allDay = false;
      event.start = `${te.start}T${te.startTime || "00:00"}:00`;
      if (te.endTime) event.end = `${te.start}T${te.endTime}:00`;
    } else {
      event.allDay = true;
      event.start = te.start;
      event.end = nextDay(te.start);
    }
    events.push(event);
  });

  const calMount = $("#fc-cal", body);
  $("#cal-show-contractor", body).addEventListener("change", (e) => {
    window.AEWTTR.state.travelCalendarShowContractor = e.target.checked;
    drawTravelCalendar(body);
  });

  if (window.FullCalendar) {
    const calendar = new FullCalendar.Calendar(calMount, {
      initialView: "dayGridMonth",
      headerToolbar: { left: "prev,next today", center: "title", right: "" },
      height: 640,
      events,
      eventDidMount: (info) => {
        info.el.style.cursor = "pointer";
        info.el.title = "Click for details";
      },
      // Reuse the same rich detail modal used on My Travel and Approvals
      // (status, edit/revoke/cancel/export actions) instead of a bespoke
      // read-only summary — one consistent "click to see details" behavior
      // everywhere in the app, which is the whole point of "more intuitive".
      eventClick: (info) => {
        const props = info.event.extendedProps;
        if (props._kind === "teamEvent") {
          navigate("travel/events/edit/" + props.te.id);
        } else {
          openTravelDetailModal(props.r.id);
        }
      }
    });
    calendar.render();
  } else {
    $("#fc-cal", body).innerHTML = `<div class="empty-state">Calendar library failed to load.</div>`;
  }
}
function nextDay(dateStr) {
  const d = new Date(dateStr + "T00:00:00");
  d.setDate(d.getDate() + 1);
  return d.toISOString().slice(0, 10);
}

function migrateTravelRequests(db) {
  if (!db || !db.travelRequests) return;
  let changed = 0;
  db.travelRequests.forEach((r) => {
    let dirty = false;
    // Normalize old statuses to new lifecycle
    if (r.status === "Approved") {
      r.status = "Submitted";
      // If it already had a C/O, mark it assigned
      if (r.chargeObject && !r.chargeObjectStatus) {
        r.chargeObjectStatus = "Assigned";
      }
      dirty = true;
    } else if (r.status === "Pending Finance") {
      r.status = "Submitted";
      dirty = true;
    } else if (r.status === "Pending") {
      r.status = "Submitted";
      dirty = true;
    } else if (r.status === "Revoked") {
      r.status = "Withdrawn";
      dirty = true;
    } else if (r.status === "Denied") {
      r.status = "Cancelled";
      dirty = true;
    }
    // Backfill chargeObjectStatus
    if (!r.chargeObjectStatus && travelCategory(r) !== "Leave") {
      r.chargeObjectStatus = (r.chargeObject) ? "Assigned" : "Pending";
      dirty = true;
    }
    // Backfill customerConcurrenceStatus
    if (!r.customerConcurrenceStatus) {
      const isLeave = travelCategory(r) === "Leave";
      // If there's already a concurredBy value, mark concurred
      if (r.concurredBy || r.approvedBy) {
        r.customerConcurrenceStatus = "Concurred";
        r.customerConcurredBy = r.customerConcurredBy || r.concurredBy || r.approvedBy || "";
      } else if (isLeave) {
        // Legacy leave without requiresConcurrence flag — treat as no concurrence needed
        r.requiresConcurrence = r.requiresConcurrence || false;
        r.customerConcurrenceStatus = "Concurred";
      } else {
        r.customerConcurrenceStatus = r.chargeObjectStatus === "Assigned" ? "Concurred" : "Pending";
      }
      dirty = true;
    }
    if (dirty) changed++;
  });
  if (changed > 0) {
    console.info(`PULSE: migrated ${changed} travel request(s) to new status model`);
  }
}

// Run migration on load (idempotent)
(function () {
  const checkAndMigrate = function () {
    const db = window.AEWTTR && window.AEWTTR.db;
    if (db && db.travelRequests) {
      migrateTravelRequests(db);
    }
  };
  if (document.readyState === "complete") {
    checkAndMigrate();
  } else {
    window.addEventListener("aewttr-db-ready", checkAndMigrate, { once: true });
  }
}());
