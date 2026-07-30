/* ═══════════════════════════════════════════════════════════════════════════
   UNCLASSIFIED

   sp-columns.js — creating list columns, which is the part that breaks everyone.

   Read this before you create a single column. Getting it wrong does not produce
   an error; it produces columns named `field_7`, invisible on forms, absent from
   views, and impossible to rename afterwards. This program hit that hard enough
   to abandon lists for a JSON store mid-project (see sp-json-store.js and
   notes/05). It was avoidable, and this file is how.

   Requires sp-rest-core.js.
═══════════════════════════════════════════════════════════════════════════ */

"use strict";

const SPColumns = (() => {
  /* ── The one number that matters: Options = 28 ─────────────────────────── */

  /*  28 = 4 | 8 | 16

        4  addToAllContentTypes      → the column shows up on the add/edit form
        8  addFieldInternalNameHint  → SharePoint honours YOUR Name attribute
       16  addFieldToDefaultView     → the column appears when someone opens the list

     Omit it, or pass 0, and every symptom below follows from that single
     mistake:

       | Symptom                                   | Missing flag |
       |-------------------------------------------|--------------|
       | Internal name came out as `field_7`       | 8            |
       | Column missing from the new/edit form     | 4            |
       | List shows only the Title column          | 16           |
       | Writes "succeed" but the SP form is empty | 8 (writing to
       |                                           | a name that
       |                                           | does not exist)

     The last row is the cruel one: your POST returns 201, your data is nowhere,
     and nothing logs an error — because you wrote a property SharePoint invented
     a different internal name for. */
  const OPTIONS_ADD_FIELD = 28;

  /* ── Field XML ────────────────────────────────────────────────────────── */

  const esc = (s) =>
    String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

  /* Use simple alphanumeric names: `ProjectCode`, never `Project Code`. A space
     becomes `_x0020_` in the internal name and every OData query against it
     turns into an escaping puzzle for the life of the app. */
  function fieldXml(name, type, opts) {
    const o = opts || {};
    if (!/^[A-Za-z][A-Za-z0-9]*$/.test(name)) {
      throw new Error(
        `Column name "${name}" must be alphanumeric and start with a letter. ` +
          `Spaces and punctuation become _x0020_-style escapes in the internal name.`
      );
    }
    const common =
      `DisplayName="${esc(o.displayName || name)}" Name="${esc(name)}" StaticName="${esc(name)}" ` +
      `Required="${o.required ? "TRUE" : "FALSE"}"`;

    switch (type) {
      case "Choice": {
        const choices = (o.choices || []).map((c) => `<CHOICE>${esc(c)}</CHOICE>`).join("");
        return `<Field Type="Choice" ${common} Format="Dropdown"><CHOICES>${choices}</CHOICES></Field>`;
      }
      case "MultiChoice": {
        const choices = (o.choices || []).map((c) => `<CHOICE>${esc(c)}</CHOICE>`).join("");
        return `<Field Type="MultiChoice" ${common}><CHOICES>${choices}</CHOICES></Field>`;
      }
      case "Note":
        return `<Field Type="Note" ${common} NumLines="${o.numLines || 6}" RichText="FALSE" />`;
      case "DateTime":
        return `<Field Type="DateTime" ${common} Format="${o.dateOnly ? "DateOnly" : "DateTime"}" />`;
      case "Boolean":
        return `<Field Type="Boolean" ${common}><Default>${o.default ? "1" : "0"}</Default></Field>`;
      case "Number":
        return `<Field Type="Number" ${common}${o.decimals != null ? ` Decimals="${o.decimals}"` : ""} />`;
      case "Currency":
        return `<Field Type="Currency" ${common} LCID="1033" />`;
      case "User":
        return `<Field Type="User" ${common} UserSelectionMode="PeopleOnly" />`;
      case "URL":
        return `<Field Type="URL" ${common} Format="Hyperlink" />`;
      case "Text":
      default:
        return `<Field Type="Text" ${common} MaxLength="${o.maxLength || 255}" />`;
    }
  }

  /* ── Create ───────────────────────────────────────────────────────────── */

  /* NOTE the odata=verbose. This endpoint is one of the few that rejects
     nometadata — it needs the __metadata type on the parameters object. A 400
     here that looks like a malformed payload is usually this. */
  async function createColumn(listTitle, name, type, opts) {
    const list = `getbytitle('${encodeURIComponent(listTitle)}')`;
    await SP.request(`/_api/web/lists/${list}/fields/createfieldasxml`, {
      method: "POST",
      verbose: true,
      body: {
        parameters: {
          __metadata: { type: "SP.XmlSchemaFieldCreationInformation" },
          SchemaXml: fieldXml(name, type, opts),
          Options: OPTIONS_ADD_FIELD,
        },
      },
    });
  }

  async function listFields(listTitle) {
    const list = `getbytitle('${encodeURIComponent(listTitle)}')`;
    const r = await SP.get(`/_api/web/lists/${list}/fields?$select=InternalName,Title,TypeAsString,Hidden`);
    return r.value || [];
  }

  /* ── Verify, then repair ──────────────────────────────────────────────── */

  /* Create-then-verify, always. A column cannot be renamed through REST — if the
     internal name came out wrong, the only fix is delete and recreate, and you
     want to discover that now rather than three weeks of data later. */
  async function ensureColumns(listTitle, schema) {
    const before = await listFields(listTitle);
    const existing = new Set(before.map((f) => f.InternalName));
    const created = [];
    const problems = [];

    for (const col of schema) {
      if (existing.has(col.name)) continue;
      await createColumn(listTitle, col.name, col.type, col);
      created.push(col.name);
    }

    const after = await listFields(listTitle);
    const present = new Map(after.map((f) => [f.InternalName, f]));
    for (const col of schema) {
      const field = present.get(col.name);
      if (!field) {
        problems.push({
          name: col.name,
          issue:
            "not present after creation — SharePoint almost certainly invented a " +
            "different internal name (field_N). Options:28 missing, or a name collision. " +
            "REST cannot rename it; delete and recreate.",
        });
      } else if (col.type && field.TypeAsString && !typeMatches(col.type, field.TypeAsString)) {
        problems.push({
          name: col.name,
          issue: `type is ${field.TypeAsString}, expected ${col.type} — recreate it`,
        });
      }
    }

    const repaired = await repairDefaultView(listTitle, schema.map((c) => c.name));
    return { created, repaired, problems };
  }

  function typeMatches(want, got) {
    if (want === got) return true;
    /* SharePoint reports some types under a different display name. */
    const aliases = { Note: ["Note", "MultiLine"], URL: ["URL", "Hyperlink"], Currency: ["Currency", "Number"] };
    return (aliases[want] || [want]).includes(got);
  }

  /* A column can exist, be correctly named, and still not appear when someone
     opens the list — because it is not in the default view. That happens to
     anything created by an older build that predates Options:28. Make this
     repair idempotent and run it on every setup pass, not once by hand. */
  async function repairDefaultView(listTitle, names) {
    const list = `getbytitle('${encodeURIComponent(listTitle)}')`;
    const view = await SP.get(`/_api/web/lists/${list}/DefaultView/ViewFields`);
    const inView = new Set((view && (view.Items || view.value)) || []);
    const added = [];

    for (const name of names) {
      if (inView.has(name)) continue;
      try {
        await SP.post(`/_api/web/lists/${list}/DefaultView/ViewFields/addviewfield('${encodeURIComponent(name)}')`);
        added.push(name);
      } catch (err) {
        /* Already present, or not a real field. Not worth failing setup over. */
      }
    }
    return added;
  }

  /* ── Lists ────────────────────────────────────────────────────────────── */

  /* BaseTemplate: 100 = generic custom list, 101 = document library. Check for
     existence first: a duplicate title cannot be created anyway, and the check
     makes an idempotent setup routine cheap to re-run. */
  async function ensureList(title, description, baseTemplate) {
    try {
      const existing = await SP.get(`/_api/web/lists/getbytitle('${encodeURIComponent(title)}')?$select=Id,Title`);
      if (existing && existing.Id) return { created: false, id: existing.Id };
    } catch (err) {
      /* 404 → does not exist yet */
    }
    const made = await SP.post("/_api/web/lists", {
      Title: title,
      Description: description || "",
      BaseTemplate: baseTemplate || 100,
    });
    return { created: true, id: made && made.Id };
  }

  return { OPTIONS_ADD_FIELD, fieldXml, createColumn, listFields, ensureColumns, repairDefaultView, ensureList };
})();

if (typeof module !== "undefined" && module.exports) module.exports = SPColumns;

/* Example — an idempotent setup routine you can call on every boot:

     await SPColumns.ensureList("Widget Tracker", "Widgets and their states");
     const report = await SPColumns.ensureColumns("Widget Tracker", [
       { name: "WidgetCode",  type: "Text",     maxLength: 64 },
       { name: "Status",      type: "Choice",   choices: ["Open", "Active", "Closed"] },
       { name: "DueDate",     type: "DateTime", dateOnly: true },
       { name: "IsCritical",  type: "Boolean",  default: false },
       { name: "Notes",       type: "Note" },
     ]);
     if (report.problems.length) console.error("SCHEMA PROBLEMS", report.problems);

   Surface report.problems somewhere a human will see it. A silent schema problem
   becomes a silent data-loss problem.
*/

/* UNCLASSIFIED */
