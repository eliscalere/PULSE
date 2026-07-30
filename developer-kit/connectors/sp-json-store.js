/* ═══════════════════════════════════════════════════════════════════════════
   UNCLASSIFIED

   sp-json-store.js — a document library as your database.

   THE HONEST PITCH. PULSE began on SharePoint lists and moved here mid-project,
   because column creation in this tenant kept producing mangled internal names,
   wrong types, and hidden duplicates (see sp-columns.js for how to avoid that —
   the knowledge came too late). Storing one JSON document sidesteps schema
   entirely: no columns to create, no internal names to mangle, no views to
   repair.

   WHAT YOU GIVE UP, and you should decide this deliberately:

     · No server-side query. You load the whole document and filter in memory.
       Fine at thousands of records, wrong at hundreds of thousands.
     · No SharePoint UI. Nobody edits your data in a list view, sorts it, or
       exports it to Excel without your app.
     · No Power Automate triggers on record change.
     · Concurrency is yours to handle. Two users saving at once is a lost update
       unless you use the ETag path below.

   If any of those matter, use lists and follow sp-columns.js exactly.

   Requires sp-rest-core.js.
═══════════════════════════════════════════════════════════════════════════ */

"use strict";

const SPJsonStore = (() => {
  /* ── Setup ────────────────────────────────────────────────────────────── */

  /* Resolve the library's REAL root folder rather than assuming the URL matches
     the title. A library titled "PULSE App Data" lives at .../PULSEAppData or
     .../Shared%20Documents/... depending on how it was created; guessing the path
     is a 404 that looks like a permissions problem. */
  async function resolveLibraryRoot(libraryTitle) {
    const list = `getbytitle('${encodeURIComponent(libraryTitle)}')`;
    const info = await SP.get(`/_api/web/lists/${list}/RootFolder?$select=ServerRelativeUrl`);
    if (!info || !info.ServerRelativeUrl) {
      throw new Error(`Could not resolve root folder for library "${libraryTitle}"`);
    }
    return info.ServerRelativeUrl;
  }

  async function ensureLibrary(libraryTitle, description) {
    try {
      return await resolveLibraryRoot(libraryTitle);
    } catch (err) {
      /* 404 → create it. BaseTemplate 101 is a document library. */
    }
    await SP.post("/_api/web/lists", {
      Title: libraryTitle,
      Description: description || "",
      BaseTemplate: 101,
    });
    /* A freshly created library is not always immediately addressable. Retry the
       resolve a few times rather than failing setup on a race. */
    for (let attempt = 0; attempt < 5; attempt++) {
      try {
        return await resolveLibraryRoot(libraryTitle);
      } catch (err) {
        await new Promise((r) => setTimeout(r, 400 * (attempt + 1)));
      }
    }
    return resolveLibraryRoot(libraryTitle);
  }

  /* Create nested folders one level at a time, parent first. SharePoint will not
     create intermediate folders for you, and the failure is a 404 on the leaf. */
  async function ensureFolder(serverRelativeUrl) {
    const clean = String(serverRelativeUrl).replace(/\/+$/, "");
    try {
      await SP.get(`/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(clean)}')?$select=Exists`);
      return clean;
    } catch (err) {
      /* not there yet */
    }
    const parent = clean.slice(0, clean.lastIndexOf("/"));
    if (parent && parent.split("/").length > 2) await ensureFolder(parent);
    await SP.post("/_api/web/folders", { ServerRelativeUrl: clean });
    return clean;
  }

  /* ── Read ─────────────────────────────────────────────────────────────── */

  /* A missing document is a normal first-run state, not an error. Returning the
     fallback instead of throwing is what keeps a fresh site from booting to a
     stack trace. */
  async function readJson(fileServerRelativeUrl, fallback) {
    try {
      const r = await SP.request(
        `/_api/web/GetFileByServerRelativePath(decodedurl='${encodeURIComponent(fileServerRelativeUrl)}')/$value`,
        { raw: true }
      );
      const text = await r.text();
      return text ? JSON.parse(text) : fallback;
    } catch (err) {
      if (/→ 404/.test(String(err.message))) return fallback;
      throw err;
    }
  }

  async function readETag(fileServerRelativeUrl) {
    try {
      const info = await SP.get(
        `/_api/web/GetFileByServerRelativePath(decodedurl='${encodeURIComponent(fileServerRelativeUrl)}')?$select=ETag`
      );
      return (info && info.ETag) || null;
    } catch (err) {
      return null;
    }
  }

  /* ── Write ────────────────────────────────────────────────────────────── */

  async function writeJson(folderServerRelativeUrl, fileName, data, etag) {
    await ensureFolder(folderServerRelativeUrl);
    const body = JSON.stringify(data, null, 2);
    const headers = { "Content-Type": "application/json" };
    /* IF-MATCH with a real ETag turns a blind overwrite into a checked one:
       a 412 means someone else saved first. Omit it and you silently clobber. */
    if (etag) headers["IF-MATCH"] = etag;

    await SP.request(
      `/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelativeUrl)}')` +
        `/Files/add(url='${encodeURIComponent(fileName)}',overwrite=true)`,
      { method: "POST", body, headers }
    );
  }

  /* ── The safe update ──────────────────────────────────────────────────── */

  /* Read, mutate, write — with the ETag checked and a bounded retry. This is the
     minimum for a store that more than one person uses at a time. `mutate` must
     be pure enough to run again on a conflict: it receives the freshly-read
     document each attempt. */
  async function update(folderServerRelativeUrl, fileName, mutate, options) {
    const opts = options || {};
    const filePath = `${folderServerRelativeUrl}/${fileName}`;
    const attempts = opts.attempts || 4;

    for (let attempt = 1; attempt <= attempts; attempt++) {
      const etag = await readETag(filePath);
      const current = await readJson(filePath, opts.initial != null ? opts.initial : {});
      const next = await mutate(current);
      if (next === undefined) return current; // mutate declined to change anything

      try {
        await writeJson(folderServerRelativeUrl, fileName, next, etag);
        return next;
      } catch (err) {
        const conflict = /→ (412|409)/.test(String(err.message));
        if (!conflict || attempt === attempts) throw err;
        /* Someone else won the race. Back off, re-read, re-apply. */
        await new Promise((r) => setTimeout(r, 150 * attempt * attempt));
      }
    }
    throw new Error(`update(${fileName}): exhausted ${attempts} attempts against concurrent writers`);
  }

  /* ── A tiny store facade ──────────────────────────────────────────────── */

  /* Wrap the above into something an app can hold onto. Note there is no
     per-record write: the unit of storage is the whole document, which is
     exactly the tradeoff described at the top of this file. */
  function open(config) {
    const libraryTitle = config.libraryTitle;
    const folderName = config.folderName;
    const fileName = config.fileName || "db.json";
    let folderPath = null;

    return {
      async init() {
        const root = await ensureLibrary(libraryTitle, config.description);
        folderPath = folderName ? await ensureFolder(`${root}/${folderName}`) : root;
        return folderPath;
      },
      path() {
        if (!folderPath) throw new Error("store.init() has not run yet");
        return `${folderPath}/${fileName}`;
      },
      read(fallback) {
        return readJson(this.path(), fallback === undefined ? {} : fallback);
      },
      update(mutate, opts) {
        return update(folderPath, fileName, mutate, opts);
      },
      /* Overwrite without checking. Only for first-run seeding. */
      overwrite(data) {
        return writeJson(folderPath, fileName, data);
      },
    };
  }

  return { ensureLibrary, resolveLibraryRoot, ensureFolder, readJson, readETag, writeJson, update, open };
})();

if (typeof module !== "undefined" && module.exports) module.exports = SPJsonStore;

/* Example — the shape PULSE actually ships:

     const store = SPJsonStore.open({
       libraryTitle: "PULSE App Data",
       folderName:   "AEWTTR-PULSE",
       fileName:     "db.json",
     });
     await store.init();

     const db = await store.read({ projects: [], meetings: [] });

     await store.update((current) => {
       current.projects.push({ id: crypto.randomUUID(), title: "New project" });
       return current;
     });
*/

/* UNCLASSIFIED */
