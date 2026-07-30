/* AEWTTR-PULSE — Users directory.
   Two tabs:
     • All users — pulled SharePoint directory (Admin → SharePoint Setup pull);
       admins can assign people into PULSE.
     • AEWTTR — people associated with PULSE (active PULSE App Roles). Default tab.
   Person pickers elsewhere use AEWTTR-associated people only. */

PAGE_RENDERERS.users = function () {
  renderUsersPage();
};

function usersPageTab() {
  if (!window.AEWTTR.state) window.AEWTTR.state = {};
  // Default to AEWTTR; only show All users when explicitly selected.
  return window.AEWTTR.state.usersTab === "all" ? "all" : "aewttr";
}

function setUsersPageTab(tab) {
  if (!window.AEWTTR.state) window.AEWTTR.state = {};
  window.AEWTTR.state.usersTab = tab === "all" ? "all" : "aewttr";
}

function usersSearchQuery() {
  if (!window.AEWTTR.state) window.AEWTTR.state = {};
  return String(window.AEWTTR.state.usersSearch || "").trim().toLowerCase();
}

function setUsersSearchQuery(q) {
  if (!window.AEWTTR.state) window.AEWTTR.state = {};
  window.AEWTTR.state.usersSearch = String(q || "");
}

function usersPageScroller(body) {
  return (body && body.querySelector && body.querySelector(".users-page-scroll"))
    || document.querySelector(".aewttr-content");
}

function refreshUsersTableInPlace(body, directory, roleRecords) {
  const scroller = usersPageScroller(body);
  const scrollTop = scroller ? scroller.scrollTop : 0;
  drawUsersTable(body, directory, roleRecords);
  const again = usersPageScroller(body) || scroller;
  if (again) again.scrollTop = scrollTop;
}

function applyLocalAssign(roleRecords, target, result) {
  const match = roleRecordMatchForUser(roleRecords, target);
  if (match && match.record) {
    const record = match.record;
    record.IsActive = true;
    record.Title = target.displayName || record.Title;
    record.UserEmail = target.email || record.UserEmail || "";
    record.UserDisplayName = target.displayName || record.UserDisplayName || "";
    record.SharePointUserId = target.spUserId || record.SharePointUserId || 0;
    record.LoginName = target.loginName || record.LoginName || "";
    record.PrincipalType = target.principalType == null ? (record.PrincipalType || 1) : target.principalType;
    record.IsSiteAdmin = !!target.isSiteAdmin;
    if (!record.Role) record.Role = APP_CONFIG.defaultUserRole;
    if (result && result.id) record.Id = result.id;
    return;
  }
  roleRecords.push({
    Id: (result && result.id) || Date.now(),
    Title: target.displayName || target.email || "User",
    UserEmail: target.email || "",
    UserDisplayName: target.displayName || "",
    SharePointUserId: target.spUserId || 0,
    LoginName: target.loginName || "",
    PrincipalType: target.principalType == null ? 1 : target.principalType,
    IsSiteAdmin: !!target.isSiteAdmin,
    Role: APP_CONFIG.defaultUserRole,
    IsActive: true,
    HideFromMeetings: false,
    Source: "Manual",
    JobTitle: "",
    SpoAccess: "",
    PowerBiAccess: ""
  });
}

function applyLocalUnassign(roleRecords, target) {
  const match = roleRecordMatchForUser(roleRecords, target);
  if (match && match.record) match.record.IsActive = false;
}

function syncMembersCacheQuietly() {
  const siteUrl = window.AEWTTR && window.AEWTTR.siteUrl;
  if (!siteUrl) return;
  if (typeof refreshMembersFromSharePoint === "function") {
    Promise.resolve(refreshMembersFromSharePoint(siteUrl)).catch(() => {});
  } else if (typeof rebuildMembersView === "function") {
    try { rebuildMembersView(); } catch (e) { /* ignore */ }
  }
}

function renderUsersPage() {
  setTopbar("Users", "Directory and PULSE association.", "");
  $("#page-content").innerHTML = `<div id="users-body"><div class="empty-state">Loading users…</div></div>`;
  const body = $("#users-body");

  if (window.AEWTTR.mode !== "sharepoint") {
    drawUsersLocalFallback(body);
    return;
  }

  const siteUrl = window.AEWTTR.siteUrl;
  Promise.allSettled([
    sharePointAdapter.loadPeopleDirectory(siteUrl),
    sharePointAdapter.getRoleRecords(siteUrl)
  ]).then(([directoryResult, roleRecordsResult]) => {
    const directory = directoryResult.status === "fulfilled" ? directoryResult.value : [];
    const roleRecords = roleRecordsResult.status === "fulfilled" ? roleRecordsResult.value : [];

    if (!directory.length && directoryResult.status === "rejected" && roleRecordsResult.status === "rejected") {
      const error = directoryResult.reason || roleRecordsResult.reason;
      const formatted = sharePointAdapter.formatSpError(null, String(error));
      body.innerHTML = `<div class="empty-state">Couldn't load users. ${escapeHtml((error && error.friendly) || formatted.friendly)}</div>`;
      return;
    }

    drawUsersTable(body, directory, roleRecords);

    if (directoryResult.status === "rejected") {
      toast((directoryResult.reason && directoryResult.reason.friendly) || "Couldn't load the pulled directory. Run Pull SharePoint people from Admin → SharePoint Setup.", "error");
    }
    if (roleRecordsResult.status === "rejected") {
      toast((roleRecordsResult.reason && roleRecordsResult.reason.friendly) || "Couldn't load PULSE App Roles.", "error");
    }
  }).catch((e) => {
    const formatted = sharePointAdapter.formatSpError(null, String(e));
    body.innerHTML = `<div class="empty-state">Couldn't load users. ${escapeHtml((e && e.friendly) || formatted.friendly)}</div>`;
  });
}

function drawUsersLocalFallback(body) {
  const db = window.AEWTTR.db;
  body.innerHTML = `
    <div class="aewttr-card" style="margin-bottom:14px;">
      <table class="aewttr-table">
        <thead><tr><th>Person</th><th>App Role</th></tr></thead>
        <tbody>
          ${db.members.map(m => `
            <tr style="cursor:default;">
              <td>
                <strong>${escapeHtml(m.name)}</strong>
                <div style="font-size:11px;color:var(--aewttr-muted);">${escapeHtml(m.email || "—")}</div>
              </td>
              <td>${m.isAdmin ? `<span class="kc-badge">Admin</span>` : escapeHtml(m.role)}</td>
            </tr>`).join("")}
        </tbody>
      </table>
    </div>
    <p style="font-size:12px;color:var(--aewttr-muted);" data-help>Running outside of SharePoint right now, so this shows the local member list. Directory pull and Assign to PULSE appear once this app is hosted in SharePoint.</p>
  `;
}

function roleRecordMatchForUser(roleRecords, user) {
  return sharePointAdapter.findRoleRecordForUser(roleRecords, {
    email: user.email || "",
    loginName: user.loginName || "",
    displayName: user.displayName || user.name || "",
    spUserId: user.spUserId || 0
  });
}

function isPulseAssociatedRecord(record) {
  return !!(record && sharePointAdapter.isRoleRecordActive(record));
}

function mergeDirectoryWithRoles(directory, roleRecords) {
  const merged = [];
  const seen = new Set();

  function keyFor(user) {
    return [
      Number(user.spUserId) || 0,
      String(user.email || "").trim().toLowerCase(),
      String(user.loginName || "").trim().toLowerCase(),
      String(user.displayName || "").trim().toLowerCase()
    ].join("|");
  }

  function pushUser(user) {
    const normalized = {
      spUserId: user.spUserId || 0,
      displayName: user.displayName || user.name || "",
      email: user.email || "",
      loginName: user.loginName || "",
      isSiteAdmin: !!user.isSiteAdmin,
      principalType: user.principalType == null ? 1 : user.principalType,
      isGroupOrSystem: !!user.isGroupOrSystem
    };
    const key = keyFor(normalized);
    if (seen.has(key)) return;
    seen.add(key);
    merged.push(normalized);
  }

  (directory || []).forEach(pushUser);

  // Ensure people already in App Roles still appear on All users even if a
  // directory pull hasn't run yet.
  (roleRecords || []).forEach((record) => {
    pushUser({
      spUserId: record.SharePointUserId,
      displayName: record.UserDisplayName || record.Title,
      email: record.UserEmail,
      loginName: record.LoginName,
      isSiteAdmin: !!record.IsSiteAdmin,
      principalType: record.PrincipalType
    });
  });

  return merged.sort((a, b) => String(a.displayName || "").localeCompare(String(b.displayName || "")));
}

function drawUsersTable(body, directory, roleRecords) {
  const isAdmin = canCurrentUserAccessAdmin();
  const tab = usersPageTab();
  const q = usersSearchQuery();
  const allUsers = mergeDirectoryWithRoles(directory, roleRecords)
    .filter((u) => typeof isHiddenServicePrincipal !== "function" || !isHiddenServicePrincipal(u));
  const associatedUsers = allUsers.filter((u) => {
    const match = roleRecordMatchForUser(roleRecords, u);
    return isPulseAssociatedRecord(match && match.record);
  });
  const pool = tab === "aewttr" ? associatedUsers : allUsers;
  const displayUsers = pool.filter((u) => {
    if (!q) return true;
    const hay = `${u.displayName || ""} ${u.email || ""} ${u.loginName || ""}`.toLowerCase();
    return hay.includes(q);
  });
  const roleChoices = ["Admin", "PM Admin", "Meeting Admin", "Finance Admin", "Document Admin", "Manager", "Member", "Viewer"];
  const sourceHelp = tab === "aewttr"
    ? `Showing ${displayUsers.length} people associated with PULSE (active App Roles). Service accounts are hidden from this list. Unassign removes app access without deleting the directory entry.`
    : `Showing ${displayUsers.length} people from the pulled SharePoint directory. Service accounts are hidden. Assign someone to PULSE to grant app access. Import more people from Admin → SharePoint Setup.`;
  const colCount = isAdmin ? 3 : 2;

  body.innerHTML = `
    <div class="users-page-toolbar toolbar-row">
      <div class="filter-pills" id="users-tabs">
        <button type="button" class="filter-pill ${tab === "all" ? "active" : ""}" data-users-tab="all"${tip("Pulled SharePoint directory")}>All users (${allUsers.length})</button>
        <button type="button" class="filter-pill ${tab === "aewttr" ? "active" : ""}" data-users-tab="aewttr"${tip("People assigned to PULSE")}>AEWTTR (${associatedUsers.length})</button>
      </div>
      <div class="users-page-toolbar-right">
        ${isAdmin && tab === "aewttr" ? `<button class="btn-aewttr btn-aewttr-sm" id="users-sync"${tip("Refresh identity fields on AEWTTR people from SharePoint Members")}><i class="bx bx-sync"></i> Refresh AEWTTR roster</button>` : ""}
        <div class="search-box users-page-search">
          <i class="bx bx-search"></i>
          <input type="search" id="users-search" placeholder="Search people…" value="${escapeHtml(window.AEWTTR.state.usersSearch || "")}">
        </div>
      </div>
    </div>
    <p class="users-page-help" style="font-size:12px;color:var(--aewttr-muted);margin:0 0 14px;" data-help>${sourceHelp}</p>
    <div class="users-page-scroll">
      <div class="aewttr-card">
        <table class="aewttr-table">
          <thead><tr><th>Person</th><th>App Role</th>${isAdmin ? "<th></th>" : ""}</tr></thead>
          <tbody>
            ${displayUsers.length ? displayUsers.map((u) => {
              const match = roleRecordMatchForUser(roleRecords, u);
              const record = match && match.record ? match.record : null;
              const associated = isPulseAssociatedRecord(record);
              const role = associated ? (record.Role || APP_CONFIG.defaultUserRole) : (u.isGroupOrSystem ? "—" : "—");
              const hidden = associated && sharePointAdapter.isHiddenFromMeetings(record);
              return `
              <tr style="cursor:default;${hidden ? "opacity:.72;" : ""}" data-row-email="${escapeHtml(u.email)}">
                <td>
                  <strong>${escapeHtml(u.displayName)}</strong>
                  ${associated ? ` <span class="kc-badge">PULSE</span>` : ""}
                  ${u.isGroupOrSystem ? ` <span class="kc-badge">Group/System</span>` : ""}
                  ${hidden ? ` <span class="kc-badge">Hidden from meetings</span>` : ""}
                  <div style="font-size:11px;color:var(--aewttr-muted);">${escapeHtml(u.email || "—")}</div>
                </td>
                <td>${associated && role === "Admin" ? `<span class="kc-badge">Admin</span>` : escapeHtml(associated ? role : "Not assigned")}</td>
                ${isAdmin ? `<td style="white-space:nowrap;">${!u.isGroupOrSystem ? `
                  ${!associated ? `
                    <button class="btn-aewttr btn-aewttr-sm" data-assign-pulse data-email="${escapeHtml(u.email)}" data-name="${escapeHtml(u.displayName)}" data-spid="${u.spUserId}" data-login="${escapeHtml(u.loginName)}" data-ptype="${u.principalType}" data-siteadmin="${u.isSiteAdmin}"${tip(`Assign ${u.displayName} to PULSE`)}>Assign to PULSE</button>
                  ` : `
                    <select class="select-aewttr role-select" data-email="${escapeHtml(u.email)}" data-name="${escapeHtml(u.displayName)}" data-spid="${u.spUserId}" data-login="${escapeHtml(u.loginName)}" data-ptype="${u.principalType}" data-siteadmin="${u.isSiteAdmin}" style="max-width:110px;display:inline-block;">
                      ${roleChoices.map(r => `<option ${role === r ? "selected" : ""}>${r}</option>`).join("")}
                    </select>
                    <button class="btn-aewttr-outline btn-aewttr-sm" data-edit-profile="${escapeHtml(u.email)}" data-name="${escapeHtml(u.displayName)}" data-spid="${u.spUserId}" data-login="${escapeHtml(u.loginName)}" data-ptype="${u.principalType}" data-siteadmin="${u.isSiteAdmin}" data-jobtitle="${escapeHtml((record && record.JobTitle) || "")}" data-spo="${escapeHtml((record && record.SpoAccess) || "")}" data-pbi="${escapeHtml((record && record.PowerBiAccess) || "")}"${tip(`Edit profile details for ${u.displayName}`)}>Edit</button>
                    ${record ? `<button class="btn-aewttr-ghost btn-aewttr-sm" data-toggle-meeting-hidden="${record.Id}" data-hidden="${hidden ? "1" : "0"}" data-name="${escapeHtml(u.displayName)}"${tip(hidden ? `Show ${u.displayName} in meetings again` : `Hide ${u.displayName} from automatic meeting attendance`)}>${hidden ? "Unhide" : "Hide"}</button>` : ""}
                    <button class="btn-aewttr-ghost btn-aewttr-sm" data-unassign-pulse data-email="${escapeHtml(u.email)}" data-name="${escapeHtml(u.displayName)}" data-spid="${u.spUserId}" data-login="${escapeHtml(u.loginName)}" data-ptype="${u.principalType}" data-siteadmin="${u.isSiteAdmin}"${tip(`Remove ${u.displayName} from PULSE`)}>Unassign</button>
                  `}
                ` : ""}</td>` : ""}
              </tr>`;
            }).join("") : `<tr><td colspan="${colCount}" style="text-align:center;color:var(--aewttr-muted);padding:28px 12px;">
              ${tab === "all"
                ? "No directory people yet. An admin can import them from Admin → SharePoint Setup → Pull SharePoint people."
                : "No one is assigned to PULSE yet. Open All users and use Assign to PULSE."}
            </td></tr>`}
          </tbody>
        </table>
      </div>
    </div>
  `;

  $all("[data-users-tab]", body).forEach((btn) => btn.addEventListener("click", () => {
    setUsersPageTab(btn.dataset.usersTab);
    refreshUsersTableInPlace(body, directory, roleRecords);
  }));

  const searchInput = $("#users-search", body);
  if (searchInput) {
    searchInput.addEventListener("input", () => {
      setUsersSearchQuery(searchInput.value);
      refreshUsersTableInPlace(body, directory, roleRecords);
      const again = $("#users-search", body);
      if (again) {
        again.focus();
        const len = again.value.length;
        again.setSelectionRange(len, len);
      }
    });
  }

  const syncBtn = $("#users-sync", body);
  if (syncBtn) syncBtn.addEventListener("click", async () => {
    syncBtn.disabled = true;
    syncBtn.textContent = "Refreshing…";
    try {
      const result = await sharePointAdapter.syncSiteUsersToAppRoles(window.AEWTTR.siteUrl, window.AEWTTR.currentSpUser);
      toast(`Refreshed — ${result.updated.length} updated, ${result.skipped.length} skipped`, "success");
      syncMembersCacheQuietly();
      const freshRoles = await sharePointAdapter.getRoleRecords(window.AEWTTR.siteUrl);
      roleRecords.length = 0;
      Array.prototype.push.apply(roleRecords, freshRoles);
      refreshUsersTableInPlace(body, directory, roleRecords);
    } catch (e) {
      toast((e && e.friendly) || e.message || "Refresh failed", "error");
      syncBtn.disabled = false;
      syncBtn.textContent = "Refresh AEWTTR roster";
    }
  });

  $all("[data-assign-pulse]", body).forEach((btn) => btn.addEventListener("click", async () => {
    if (!canCurrentUserAccessAdmin()) {
      toast("You do not have permission to assign people to PULSE.", "error");
      return;
    }
    const target = {
      email: btn.dataset.email, displayName: btn.dataset.name, spUserId: +btn.dataset.spid,
      loginName: btn.dataset.login, principalType: +btn.dataset.ptype, isSiteAdmin: btn.dataset.siteadmin === "true"
    };
    btn.disabled = true;
    try {
      const result = await sharePointAdapter.assignUserToPulse(window.AEWTTR.siteUrl, target, window.AEWTTR.currentSpUser);
      applyLocalAssign(roleRecords, target, result);
      toast(`${target.displayName} is now on PULSE`, "success");
      refreshUsersTableInPlace(body, directory, roleRecords);
      syncMembersCacheQuietly();
    } catch (e) {
      toast(e.message || "Couldn't assign to PULSE", "error");
      btn.disabled = false;
    }
  }));

  $all("[data-unassign-pulse]", body).forEach((btn) => btn.addEventListener("click", async () => {
    if (!canCurrentUserAccessAdmin()) {
      toast("You do not have permission to remove people from PULSE.", "error");
      return;
    }
    const target = {
      email: btn.dataset.email, displayName: btn.dataset.name, spUserId: +btn.dataset.spid,
      loginName: btn.dataset.login, principalType: +btn.dataset.ptype, isSiteAdmin: btn.dataset.siteadmin === "true"
    };
    btn.disabled = true;
    try {
      await sharePointAdapter.unassignUserFromPulse(window.AEWTTR.siteUrl, target, window.AEWTTR.currentSpUser);
      applyLocalUnassign(roleRecords, target);
      toast(`${target.displayName} removed from PULSE`, "success");
      refreshUsersTableInPlace(body, directory, roleRecords);
      syncMembersCacheQuietly();
    } catch (e) {
      toast(e.message || "Couldn't unassign from PULSE", "error");
      btn.disabled = false;
    }
  }));

  $all(".role-select", body).forEach(sel => sel.addEventListener("change", async () => {
    if (!canCurrentUserAccessAdmin()) {
      toast("You do not have permission to manage user roles.", "error");
      return;
    }
    const target = {
      email: sel.dataset.email, displayName: sel.dataset.name, spUserId: +sel.dataset.spid,
      loginName: sel.dataset.login, principalType: +sel.dataset.ptype, isSiteAdmin: sel.dataset.siteadmin === "true"
    };
    const previous = (() => {
      const match = roleRecordMatchForUser(roleRecords, target);
      return match && match.record ? match.record.Role : "";
    })();
    try {
      await sharePointAdapter.assignUserRole(window.AEWTTR.siteUrl, target, sel.value, window.AEWTTR.currentSpUser);
      const match = roleRecordMatchForUser(roleRecords, target);
      if (match && match.record) match.record.Role = sel.value;
      toast(`${target.displayName} is now ${sel.value}`, "success");
      refreshUsersTableInPlace(body, directory, roleRecords);
      syncMembersCacheQuietly();
    } catch (e) {
      toast(e.message || "Couldn't update role", "error");
      if (previous) sel.value = previous;
      else refreshUsersTableInPlace(body, directory, roleRecords);
    }
  }));

  $all("[data-edit-profile]", body).forEach(btn => btn.addEventListener("click", () => {
    if (!canCurrentUserAccessAdmin()) {
      toast("You do not have permission to manage user profiles.", "error");
      return;
    }
    openEditProfileModal({
      email: btn.dataset.editProfile, displayName: btn.dataset.name, spUserId: +btn.dataset.spid,
      loginName: btn.dataset.login, principalType: +btn.dataset.ptype, isSiteAdmin: btn.dataset.siteadmin === "true"
    }, {
      jobTitle: btn.dataset.jobtitle,
      spoAccess: btn.dataset.spo,
      powerBiAccess: btn.dataset.pbi
    }, body, directory, roleRecords);
  }));

  $all("[data-toggle-meeting-hidden]", body).forEach(btn => btn.addEventListener("click", async () => {
    if (!canCurrentUserAccessAdmin()) {
      toast("You do not have permission to manage meeting visibility.", "error");
      return;
    }
    const recordId = +btn.dataset.toggleMeetingHidden;
    const hide = btn.dataset.hidden !== "1";
    const memberId = "u" + recordId;
    try {
      await sharePointAdapter.setUserMeetingHidden(window.AEWTTR.siteUrl, recordId, hide, window.AEWTTR.currentSpUser);
      const record = roleRecords.find((r) => Number(r.Id) === recordId);
      if (record) record.HideFromMeetings = hide;
      if (hide && typeof weeklyMeetingDb === "function") {
        const weekly = weeklyMeetingDb();
        weekly.roster = (weekly.roster || []).filter((id) => id !== memberId);
        if (typeof aewttrSaveStore === "function") aewttrSaveStore();
      }
      if (typeof rebuildMembersView === "function") rebuildMembersView();
      toast(`${btn.dataset.name} ${hide ? "hidden from meetings" : "will appear in meetings again"}`, "success");
      refreshUsersTableInPlace(body, directory, roleRecords);
    } catch (e) {
      toast(e.message || "Couldn't update meeting visibility", "error");
    }
  }));
}

function openEditProfileModal(target, current, body, directory, roleRecords) {
  const ACCESS_CHOICES = ["", "None", "View", "Full"];
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>Edit Profile — ${escapeHtml(target.displayName)}</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body">
      <div class="form-row"><label>Job title (shown as their display role)</label><input class="input-aewttr" id="ep-jobtitle" value="${escapeHtml(current.jobTitle || "")}" placeholder="e.g. Cyber Lead"></div>
      <div class="form-grid-2">
        <div class="form-row"><label>SPO Access</label>
          <select class="select-aewttr" id="ep-spo">${ACCESS_CHOICES.map(c => `<option value="${escapeHtml(c)}" ${current.spoAccess === c ? "selected" : ""}>${c || "—"}</option>`).join("")}</select>
        </div>
        <div class="form-row"><label>Power BI Access</label>
          <select class="select-aewttr" id="ep-pbi">${ACCESS_CHOICES.map(c => `<option value="${escapeHtml(c)}" ${current.powerBiAccess === c ? "selected" : ""}>${c || "—"}</option>`).join("")}</select>
        </div>
      </div>
    </div>
    <div class="aewttr-modal-foot">
      <button class="btn-aewttr-ghost" id="ep-cancel">Cancel</button>
      <button class="btn-aewttr" id="ep-save">Save Profile</button>
    </div>
  `);
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#ep-cancel", modal).addEventListener("click", closeModal);
  $("#ep-save", modal).addEventListener("click", async () => {
    try {
      const jobTitle = $("#ep-jobtitle", modal).value.trim();
      const spoAccess = $("#ep-spo", modal).value;
      const powerBiAccess = $("#ep-pbi", modal).value;
      await sharePointAdapter.updateUserProfile(window.AEWTTR.siteUrl, target, {
        jobTitle,
        spoAccess,
        powerBiAccess
      }, window.AEWTTR.currentSpUser);
      const match = roleRecords ? roleRecordMatchForUser(roleRecords, target) : null;
      if (match && match.record) {
        match.record.JobTitle = jobTitle;
        match.record.SpoAccess = spoAccess;
        match.record.PowerBiAccess = powerBiAccess;
      }
      closeModal();
      toast("Profile updated", "success");
      if (body && directory && roleRecords) {
        refreshUsersTableInPlace(body, directory, roleRecords);
        syncMembersCacheQuietly();
      }
    } catch (e) {
      toast(e.message || "Couldn't save profile", "error");
    }
  });
}
