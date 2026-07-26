/**
 * PULSE Calendar — v1.0.0
 * AEWTTR IPT DB46200
 *
 * Developer:  Eli Scalere (Contractor)
 * Primary:    elijah.t.scalere.ctr@us.navy.mil
 * Secondary:  eli.scalere@scaleredesign.com  |  (516) 265-2636
 *
 * Standalone travel-calendar companion to PULSE.
 * Reads from the "PULSE Travel Requests" SharePoint list.
 */

(function () {
  const COLORS = { Leave:"#7B61A8", Contractor:"#C62828", TDY:"#1565C0", Conference:"#2E7D32", Training:"#ED6C02", "Team Event":"#E91E63", Other:"#00897B" };
  const LEGEND = ["Leave", "TDY", "Conference", "Training", "Team Event", "Contractor", "Other"];
  const listTitle = "PULSE Travel Requests";
  let requests = [];
  let calendar = null;

  function escapeHtml(value) { const node = document.createElement("div"); node.textContent = value == null ? "" : String(value); return node.innerHTML; }
  function siteUrl() {
    if (window.AEWTTR && window.AEWTTR.adapter && typeof window.AEWTTR.adapter.detectSiteUrl === "function") {
      return window.AEWTTR.adapter.detectSiteUrl();
    }
    if (window._spPageContextInfo && window._spPageContextInfo.webAbsoluteUrl) return window._spPageContextInfo.webAbsoluteUrl.replace(/\/$/, "");
    return window.location.origin;
  }
  function odata(value) { return String(value).replace(/'/g, "''"); }
  function parseTravelers(value) { try { const parsed = JSON.parse(value || "[]"); return Array.isArray(parsed) ? parsed : []; } catch (_) { return []; } }
  function typeOf(r) {
    if (r.FormMode === "Leave" || r.RequestType === "Personal Leave") return "Leave";
    if (r.FormMode === "Team Event" || r.RequestType === "Team Event" || r.TravelType === "Team Event") return "Team Event";
    if (r.FormMode === "Contractor" || r.RequestType === "Contractor Travel") return "Contractor";
    return String(r.TravelType || "TDY").startsWith("Other") ? "Other" : (r.TravelType || "TDY");
  }
  function isLeave(r) { return typeOf(r) === "Leave"; }
  function nextDay(date) { const value = new Date(`${date}T00:00:00`); value.setDate(value.getDate() + 1); return value.toISOString().slice(0, 10); }
  function dateOnly(value) { return value ? String(value).slice(0, 10) : ""; }
  function travelerLabel(r) {
    const people = parseTravelers(r.TravelersJson).map((person) => person && person.name).filter(Boolean);
    const names = people.length ? people : [r.RequesterName || "Travel"];
    const last = names.map((name) => {
      // SharePoint person labels can include a country marker, e.g.
      // "Elijah Scalere (USA)". Remove that marker before taking the surname.
      const cleaned = String(name).replace(/\s*\([^)]*\)\s*$/, "").trim();
      return cleaned.includes(",")
        ? cleaned.split(",")[0].trim()
        : cleaned.split(/\s+/).pop();
    }).join(", ");
    return last.length > 32 ? `${last.slice(0, 29)}…` : last;
  }
  function requestLink(r) {
    return `${siteUrl()}/SitePages/AEWTTR-PULSE.aspx#/travel/all?tr=${encodeURIComponent(r.RequestCode || r.Id)}`;
  }
  function renderLegend() {
    document.getElementById("calendar-legend").innerHTML = LEGEND.map((type) => `<span><i class="pulse-calendar-dot" style="background:${COLORS[type]}"></i>${type}</span>`).join("") + `<span><i class="pulse-calendar-dot" style="background:#9AA4B2"></i>Pending travel approval</span>`;
  }
  function renderRefiners() {
    document.getElementById("calendar-refiners").innerHTML = `<h2>Refiners</h2>
      <section class="pulse-calendar-refiner-group"><h3>Travel &amp; Leave</h3>
        <label class="pulse-calendar-check"><input type="checkbox" id="type-all" checked> Select all</label>
        ${LEGEND.map((type) => `<label class="pulse-calendar-check"><input type="checkbox" class="type-filter" value="${type}" checked><i style="background:${COLORS[type]}"></i>${type}</label>`).join("")}
      </section>`;
    const all = document.getElementById("type-all");
    const filters = () => Array.from(document.querySelectorAll(".type-filter"));
    all.addEventListener("change", () => { filters().forEach((input) => { input.checked = all.checked; }); renderCalendar(); });
    filters().forEach((input) => input.addEventListener("change", () => { all.checked = filters().every((checkbox) => checkbox.checked); renderCalendar(); }));
  }
  function showDetails(r) {
    const modal = document.getElementById("travel-detail-modal");
    const travelers = parseTravelers(r.TravelersJson).map((person) => person && person.name).filter(Boolean).join(", ") || r.RequesterName || "—";
    const type = typeOf(r);
    const status = isLeave(r) ? "Scheduled leave" : (r.ReqStatus || "Pending");
    modal.innerHTML = `<section class="pulse-calendar-dialog" role="dialog" aria-modal="true" aria-labelledby="travel-detail-title">
      <header class="pulse-calendar-dialog-header"><h2 id="travel-detail-title">${escapeHtml(r.TripTitle || `${type} request`)}</h2><button class="pulse-calendar-dialog-close" aria-label="Close">×</button></header>
      <div class="pulse-calendar-dialog-body"><dl class="pulse-calendar-detail">
        <dt>Type</dt><dd>${escapeHtml(type)}</dd><dt>Travelers</dt><dd>${escapeHtml(travelers)}</dd><dt>Dates</dt><dd>${escapeHtml(dateOnly(r.StartDate))} – ${escapeHtml(dateOnly(r.EndDate))}</dd><dt>Location</dt><dd>${escapeHtml(r.Destination || (isLeave(r) ? "Leave" : "—"))}</dd><dt>Status</dt><dd>${escapeHtml(status)}</dd><dt>Details</dt><dd>${escapeHtml(r.Purpose || r.TravelNotes || "—")}</dd>
      </dl></div><footer class="pulse-calendar-dialog-actions"><a class="pulse-calendar-link" href="${escapeHtml(requestLink(r))}" target="_top">Open in PULSE</a></footer></section>`;
    modal.hidden = false;
    const close = () => { modal.hidden = true; modal.innerHTML = ""; };
    modal.querySelector(".pulse-calendar-dialog-close").addEventListener("click", close);
    modal.addEventListener("click", (event) => { if (event.target === modal) close(); }, { once:true });
  }
  function buildEvents() {
    const selectedTypes = new Set(Array.from(document.querySelectorAll(".type-filter:checked")).map((input) => input.value));
    return requests.filter((r) => !["Denied", "Cancelled", "Revoked"].includes(r.ReqStatus)).filter((r) => selectedTypes.has(typeOf(r))).map((r) => {
      const type = typeOf(r);
      const pending = r.ReqStatus !== "Approved" && type !== "Leave";
      const allDay = r.AllDay !== false && r.AllDay !== 0;
      const event = { id:r.RequestCode || String(r.Id), title:`${travelerLabel(r)} — ${type}${pending ? " (Pending)" : ""}`, color:pending ? "#9AA4B2" : (COLORS[type] || COLORS.TDY), classNames:pending ? ["travel-cal-event-pending"] : [], extendedProps:{ request:r } };
      if (allDay) { event.start = dateOnly(r.StartDate); event.end = nextDay(dateOnly(r.EndDate)); event.allDay = true; }
      else { event.start = `${dateOnly(r.StartDate)}T${r.StartTime || "00:00"}:00`; event.end = `${dateOnly(r.EndDate)}T${r.EndTime || "23:59"}:00`; event.allDay = false; }
      return event;
    }).filter((event) => event.start);
  }
  function renderCalendar() {
    if (!window.FullCalendar) throw new Error("Calendar library did not load.");
    if (calendar) calendar.destroy();
    calendar = new FullCalendar.Calendar(document.getElementById("travel-calendar"), { initialView:"dayGridMonth", headerToolbar:{ left:"prev,next today", center:"title", right:"" }, height:"100%", dayMaxEvents:2, events:buildEvents(), eventDidMount:(info) => { info.el.title = "Click for details"; }, eventClick:(info) => showDetails(info.event.extendedProps.request) });
    calendar.render();
  }
  function dismissLoader() {
    const el = document.getElementById("pulse-boot-loader");
    if (!el) return;
    el.style.transition = "opacity 0.3s ease";
    el.style.opacity = "0";
    setTimeout(() => el.remove(), 320);
  }

  async function load() {
    const status = document.getElementById("calendar-status");
    try {
      const travelSelect = ["Id","RequestCode","RequesterName","RequesterEmail","FormMode","RequestType","TripTitle","TravelersJson","Destination","StartDate","EndDate","AllDay","StartTime","EndTime","Purpose","TravelNotes","ReqStatus","TravelType"].join(",");
      const travelEndpoint = `${siteUrl()}/_api/web/lists/getbytitle('${odata("PULSE Travel Requests")}')/items?$top=500&$select=${travelSelect}`;
      
      const teamEventSelect = ["Id","Title","EventCode","EventDate","AllDay","StartTime","EndTime","Location","Description"].join(",");
      const teamEventEndpoint = `${siteUrl()}/_api/web/lists/getbytitle('${odata("PULSE Team Events")}')/items?$top=500&$select=${teamEventSelect}`;

      const [travelRes, teamEventRes] = await Promise.all([
        fetch(travelEndpoint, { credentials:"same-origin", headers:{ Accept:"application/json;odata=nometadata", "Cache-Control":"no-cache" } }).catch(() => ({ ok: false })),
        fetch(teamEventEndpoint, { credentials:"same-origin", headers:{ Accept:"application/json;odata=nometadata", "Cache-Control":"no-cache" } }).catch(() => ({ ok: false }))
      ]);

      let allRequests = [];
      if (travelRes.ok) {
        const travelPayload = await travelRes.json();
        allRequests.push(...(travelPayload.value || []));
      }
      if (teamEventRes.ok) {
        const teamPayload = await teamEventRes.json();
        const teamEvents = (teamPayload.value || []).map(e => ({ 
           ...e, 
           TravelType: "Team Event", 
           FormMode: "Team Event",
           TripTitle: e.Title || e.EventCode || "Team Event",
           Destination: e.Location || "Local",
           StartDate: String(e.EventDate || "").slice(0, 10),
           EndDate: String(e.EventDate || "").slice(0, 10),
           ReqStatus: "Approved" 
        }));
        allRequests.push(...teamEvents);
      }
      
      if (!travelRes.ok && !teamEventRes.ok) {
         throw new Error(`Could not reach SharePoint lists.`);
      }

      requests = allRequests;
      status.textContent = requests.length ? "" : "No travel or leave is currently scheduled.";
      renderCalendar();
    } catch (error) {
      status.textContent = `Calendar could not load: ${error.message}`;
      requests = [];
      renderCalendar();
    } finally {
      dismissLoader();
    }
  }
  renderRefiners();
  renderLegend();
  
  // Initial load
  load();
  
  // Background refresh every 60 seconds (near real-time)
  setInterval(load, 60000);
}());
