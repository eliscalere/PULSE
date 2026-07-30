/* ═══════════════════════════════════════════════════════════════════════════
   UNCLASSIFIED

   sp-list-crud.js — read, create, update, delete list items.

   The easy half of SharePoint. Two things to internalize:

     1. Property names in results are INTERNAL names, not display names. If your
        columns were created correctly (sp-columns.js) they are the same string.
        If they were not, this is where you find out.
     2. Reads page at ~100 items. Code that reads `.value` once works until the
        list grows, then silently shows partial data. Use SP.getAllPages().

   Requires sp-rest-core.js.
═══════════════════════════════════════════════════════════════════════════ */

"use strict";

const SPList = (() => {
  const listPath = (title) => `/_api/web/lists/getbytitle('${encodeURIComponent(title)}')`;

  /* ── Read ─────────────────────────────────────────────────────────────── */

  /* Always $select explicitly. `items` with no select returns every field
     including a pile of SharePoint internals, which on a wide list is a
     multi-megabyte response per page. */
  function query(opts) {
    const o = opts || {};
    const parts = [];
    if (o.select) parts.push(`$select=${o.select.join(",")}`);
    if (o.expand) parts.push(`$expand=${o.expand.join(",")}`);
    if (o.filter) parts.push(`$filter=${encodeURIComponent(o.filter)}`);
    if (o.orderBy) parts.push(`$orderby=${encodeURIComponent(o.orderBy)}`);
    if (o.top) parts.push(`$top=${o.top}`);
    return parts.length ? `?${parts.join("&")}` : "";
  }

  const getItems = (listTitle, opts) => SP.getAllPages(`${listPath(listTitle)}/items${query(opts)}`);

  const getItem = (listTitle, id, opts) => SP.get(`${listPath(listTitle)}/items(${id})${query(opts)}`);

  /* ── Write ────────────────────────────────────────────────────────────── */

  /* Returns the created item, including its server-assigned .Id. Keep that Id:
     it is how you update or delete later, and it is how child records reference
     the parent. See the fire-and-forget postmortem in notes/06 — if you create a
     parent and a child in sequence, you must actually wait for this. */
  const createItem = (listTitle, fields) => SP.post(`${listPath(listTitle)}/items`, fields);

  const updateItem = (listTitle, id, fields, etag) =>
    SP.merge(`${listPath(listTitle)}/items(${id})`, fields, etag);

  async function deleteItem(listTitle, id) {
    try {
      await SP.del(`${listPath(listTitle)}/items(${id})`);
    } catch (err) {
      /* Already gone is the outcome the caller wanted. Anything else is real. */
      if (!/→ 404/.test(String(err.message))) throw err;
    }
  }

  /* ── Values that need special shapes ──────────────────────────────────── */

  /* Most field types take a plain JS value. These do not: */
  const values = {
    /* A Person/Group column is written by SharePoint user id, to a property
       named <InternalName>Id — not the column name itself. */
    user: (internalName, spUserId) => ({ [`${internalName}Id`]: spUserId }),

    /* Multi-user takes a results array, again on the Id property. */
    users: (internalName, spUserIds) => ({ [`${internalName}Id`]: { results: spUserIds } }),

    /* MultiChoice wants a results array of strings. */
    multiChoice: (internalName, choices) => ({ [internalName]: { results: choices } }),

    /* A Lookup is written by the target item's id, on <InternalName>Id. */
    lookup: (internalName, targetItemId) => ({ [`${internalName}Id`]: targetItemId }),

    /* URL columns are an object, not a string. */
    url: (internalName, href, description) => ({
      [internalName]: { Url: href, Description: description || href },
    }),

    /* Dates: send ISO 8601. A Date object stringifies to something SharePoint
       accepts inconsistently across tenants — be explicit. */
    date: (internalName, value) => ({
      [internalName]: value == null ? null : new Date(value).toISOString(),
    }),
  };

  /* ── Debounced, serialized saves ──────────────────────────────────────── */

  /* A UI that can fire many rapid writes — dragging a bar on a timeline,
     resizing a date range — needs two things that are easy to get wrong:

       · debounce per object, so a drag is one write and not forty
       · serialize per object, so a create always finishes (and stamps the new
         Id) before a follow-up update for the same object runs

     And critically: the promise this returns must resolve when the write ACTUALLY
     LANDS, not when the debounce timer is set. A save() that resolves early makes
     `await save(parent)` a lie, and the child record you create next writes a null
     foreign key. That bug shipped here once. See notes/06 item 5. */
  function createSaver(writeRecord, delayMs) {
    const timers = new Map();
    const chains = new WeakMap();
    const settlers = new Map();

    return function save(kind, obj) {
      const pending = timers.get(obj);
      if (pending) clearTimeout(pending);

      let settler = settlers.get(obj);
      if (!settler) {
        let resolve;
        let reject;
        const promise = new Promise((res, rej) => {
          resolve = res;
          reject = rej;
        });
        settler = { promise, resolve, reject };
        settlers.set(obj, settler);
      }

      const timer = setTimeout(() => {
        timers.delete(obj);
        const previous = chains.get(obj) || Promise.resolve();
        const next = previous
          .then(() => writeRecord(kind, obj))
          .then(
            (result) => {
              settlers.delete(obj);
              settler.resolve(result);
              return result;
            },
            (err) => {
              settlers.delete(obj);
              settler.reject(err);
              throw err;
            }
          );
        /* Keep the chain alive after a rejection so one failed write does not
           wedge every later write for this object. */
        chains.set(obj, next.catch(() => {}));
      }, delayMs == null ? 500 : delayMs);

      timers.set(obj, timer);
      return settler.promise;
    };
  }

  return { getItems, getItem, createItem, updateItem, deleteItem, values, createSaver, listPath };
})();

if (typeof module !== "undefined" && module.exports) module.exports = SPList;

/* Example:

     const open = await SPList.getItems("Widget Tracker", {
       select: ["Id", "Title", "WidgetCode", "Status", "DueDate"],
       filter: "Status ne 'Closed'",
       orderBy: "DueDate asc",
     });

     const made = await SPList.createItem("Widget Tracker", {
       Title: "Replace intake manifold",
       WidgetCode: "WDG-014",
       Status: "Open",
       ...SPList.values.date("DueDate", "2026-09-01"),
       ...SPList.values.user("AssignedTo", currentUser.spUserId),
     });

     await SPList.updateItem("Widget Tracker", made.Id, { Status: "Active" });
*/

/* UNCLASSIFIED */
