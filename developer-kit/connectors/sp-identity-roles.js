/* ═══════════════════════════════════════════════════════════════════════════
   UNCLASSIFIED

   sp-identity-roles.js — who is this user, and what are they allowed to do.

   Authentication is free (the session cookie). AUTHORIZATION is yours to build,
   and the trap is thinking SharePoint already answered it:

     IsSiteAdmin is a SharePoint permission. It is NOT your app's "Admin" role.

   Treat site-admin as a bootstrap fallback only — "if this app has no Admin at
   all yet, let a site admin in so first-run setup is reachable" — and never as
   the ongoing authorization check. Hardcoding an email is worse.

   Requires sp-rest-core.js.
═══════════════════════════════════════════════════════════════════════════ */

"use strict";

const SPIdentity = (() => {
  /* ── The current user ─────────────────────────────────────────────────── */

  async function currentUser() {
    const u = await SP.get("/_api/web/currentuser");
    return {
      spUserId: u.Id,
      displayName: u.Title,
      email: u.Email,
      loginName: u.LoginName,
      isSiteAdmin: !!u.IsSiteAdmin,
      principalType: u.PrincipalType,
    };
  }

  /* Site users include groups, system accounts, and stale entries. PrincipalType
     1 is a real user; filtering is the difference between a clean people list and
     one containing "SharePoint App" and "Everyone except external users". */
  async function siteUsers() {
    const all = await SP.getAllPages(
      "/_api/web/siteusers?$select=Id,Title,Email,LoginName,PrincipalType,IsSiteAdmin"
    );
    return all
      .filter((u) => u.PrincipalType === 1 && u.Email)
      .map((u) => ({
        spUserId: u.Id,
        displayName: u.Title,
        email: u.Email,
        loginName: u.LoginName,
        isSiteAdmin: !!u.IsSiteAdmin,
      }));
  }

  /* ── Matching a user to their role row ────────────────────────────────── */

  /* Match on whatever is actually present, in descending reliability. A row added
     by hand through the SharePoint UI is routinely missing the login name or the
     user id, and a lookup that only checks one field will fail to find a person
     who is plainly there — which reads to them as "the app forgot who I am". */
  function findRoleRow(rows, user) {
    const norm = (s) => String(s || "").trim().toLowerCase();
    const byEmail = rows.find((r) => norm(r.UserEmail) && norm(r.UserEmail) === norm(user.email));
    if (byEmail) return byEmail;

    const byLogin = rows.find((r) => norm(r.LoginName) && norm(r.LoginName) === norm(user.loginName));
    if (byLogin) return byLogin;

    const byId = rows.find((r) => r.SharePointUserId != null && Number(r.SharePointUserId) === Number(user.spUserId));
    if (byId) return byId;

    return rows.find((r) => norm(r.UserDisplayName) && norm(r.UserDisplayName) === norm(user.displayName)) || null;
  }

  /* ── Resolving the effective role ─────────────────────────────────────── */

  const DEFAULT_ROLE = "Member";

  /* Deliberately does NOT write a row. If resolution writes on every visit, every
     viewer who ever opens the app creates a record, and your roles list becomes a
     visitor log. Provisioning belongs in an explicit admin-triggered sync. */
  function resolveRole(rows, user, options) {
    const opts = options || {};
    const row = findRoleRow(rows, user);

    if (row && row.IsActive === false) {
      return { role: "Disabled", source: "roles-row", row };
    }
    if (row && row.Role) {
      return { role: row.Role, source: "roles-row", row };
    }

    /* Bootstrap: no Admin exists anywhere, so let a SharePoint site admin in —
       purely so first-run setup and recovery are reachable on a fresh site. */
    const anyAdmin = rows.some((r) => r.Role === "Admin" && r.IsActive !== false);
    if (!anyAdmin && user.isSiteAdmin) {
      return {
        role: "Admin",
        source: "bootstrap-site-admin",
        note: "No app Admin exists yet; granting Admin to a SharePoint site admin so setup is reachable. " +
          "Assign a real Admin and this stops applying.",
      };
    }

    return { role: opts.defaultRole || DEFAULT_ROLE, source: "default" };
  }

  /* ── Admin-triggered provisioning ─────────────────────────────────────── */

  /* Syncing every site user is the slowest thing in a boot sequence and almost
     nobody needs it done before they can see their own data. Run it in the
     background after the first render, or better, only when an admin asks.
     See notes/06 item 7 for what happens when you do it on every boot. */
  async function syncUsers(rows, options) {
    const opts = options || {};
    const users = await siteUsers();
    const toCreate = [];
    const toUpdate = [];

    for (const user of users) {
      const row = findRoleRow(rows, user);
      if (!row) {
        toCreate.push({
          UserEmail: user.email,
          UserDisplayName: user.displayName,
          SharePointUserId: user.spUserId,
          LoginName: user.loginName,
          Role: opts.defaultRole || DEFAULT_ROLE,
          IsActive: true,
        });
        continue;
      }
      /* Backfill the fields a hand-added row is usually missing, without ever
         touching Role — an admin's assignment is not the sync's business. */
      const patch = {};
      if (!row.SharePointUserId) patch.SharePointUserId = user.spUserId;
      if (!row.LoginName) patch.LoginName = user.loginName;
      if (!row.UserEmail) patch.UserEmail = user.email;
      if (!row.UserDisplayName) patch.UserDisplayName = user.displayName;
      if (Object.keys(patch).length) toUpdate.push({ row, patch });
    }

    return { toCreate, toUpdate, scanned: users.length };
  }

  /* ── A permission check worth calling ─────────────────────────────────── */

  /* Keep the capability table in one place. Scattering `if (role === "Admin")`
     through the UI guarantees an inconsistency, and read-only users are the ones
     who find it. Every screen has a real view-only audience. */
  const CAPABILITIES = {
    Admin: ["read", "write", "delete", "manageUsers", "runSetup", "export"],
    Editor: ["read", "write", "export"],
    Member: ["read", "export"],
    Viewer: ["read"],
    Disabled: [],
  };

  const can = (role, capability) => (CAPABILITIES[role] || CAPABILITIES.Viewer).includes(capability);

  return { currentUser, siteUsers, findRoleRow, resolveRole, syncUsers, can, CAPABILITIES, DEFAULT_ROLE };
})();

if (typeof module !== "undefined" && module.exports) module.exports = SPIdentity;

/* Suggested roles list — one item per user:

     UserEmail          Text
     UserDisplayName    Text
     SharePointUserId   Number
     LoginName          Text
     Role               Choice  (Admin | Editor | Member | Viewer)
     IsActive           Boolean

   Create it with SPColumns.ensureColumns() so the internal names come out right.

   Boot usage:

     const user  = await SPIdentity.currentUser();
     const rows  = await SPList.getItems("App Roles", { select: ["Id","UserEmail","UserDisplayName","SharePointUserId","LoginName","Role","IsActive"] });
     const { role, source, note } = SPIdentity.resolveRole(rows, user);
     if (note) console.warn(note);
     if (SPIdentity.can(role, "runSetup")) showSetupButton();
*/

/* UNCLASSIFIED */
