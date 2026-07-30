/* ═══════════════════════════════════════════════════════════════════════════
   UNCLASSIFIED

   sp-files.js — folders, binary uploads, downloads, ETags.

   The one mistake that matters here: uploading binary content through a text
   path. A TextEncoder- or string-based upload silently corrupts anything that is
   not UTF-8 — a .docx becomes an unopenable file, a .png becomes noise — and the
   request returns 200. Pass the raw ArrayBuffer or Blob as the body.

   Requires sp-rest-core.js.
═══════════════════════════════════════════════════════════════════════════ */

"use strict";

const SPFiles = (() => {
  const folderApi = (url) => `/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(url)}')`;
  const fileApi = (url) => `/_api/web/GetFileByServerRelativePath(decodedurl='${encodeURIComponent(url)}')`;

  /* ── Folders ──────────────────────────────────────────────────────────── */

  async function folderExists(serverRelativeUrl) {
    try {
      const info = await SP.get(`${folderApi(serverRelativeUrl)}?$select=Exists`);
      return !!(info && info.Exists !== false);
    } catch (err) {
      return false;
    }
  }

  /* Parents first, one level at a time — SharePoint does not create intermediate
     folders, and the failure surfaces as a 404 on the leaf that reads like a
     permissions problem. */
  async function ensureFolderPath(serverRelativeUrl) {
    const clean = String(serverRelativeUrl).replace(/\/+$/, "");
    if (await folderExists(clean)) return clean;

    const parent = clean.slice(0, clean.lastIndexOf("/"));
    /* Stop before walking above the site collection (/sites/Foo). */
    if (parent && parent.split("/").filter(Boolean).length > 2) {
      await ensureFolderPath(parent);
    }
    await SP.post("/_api/web/folders", { ServerRelativeUrl: clean });
    return clean;
  }

  async function listFolder(serverRelativeUrl) {
    const [folders, files] = await Promise.all([
      SP.get(`${folderApi(serverRelativeUrl)}/Folders?$select=Name,ServerRelativeUrl,ItemCount`),
      SP.get(`${folderApi(serverRelativeUrl)}/Files?$select=Name,ServerRelativeUrl,Length,TimeLastModified,ETag`),
    ]);
    return {
      folders: (folders.value || []).map((f) => ({
        name: f.Name,
        path: f.ServerRelativeUrl,
        itemCount: f.ItemCount,
        isFolder: true,
      })),
      files: (files.value || []).map((f) => ({
        name: f.Name,
        path: f.ServerRelativeUrl,
        size: Number(f.Length),
        modified: f.TimeLastModified,
        etag: f.ETag,
        isFolder: false,
      })),
    };
  }

  /* ── Upload ───────────────────────────────────────────────────────────── */

  /* `content` must be an ArrayBuffer, a Blob, or a string for text files.
     Never JSON.stringify a binary, never run it through TextEncoder. */
  async function upload(folderServerRelativeUrl, fileName, content, options) {
    const opts = options || {};
    await ensureFolderPath(folderServerRelativeUrl);

    const isBinary = content instanceof ArrayBuffer || content instanceof Blob || content instanceof Uint8Array;
    const headers = {
      "Content-Type": opts.contentType || (isBinary ? "application/octet-stream" : "text/plain;charset=utf-8"),
    };
    if (opts.etag) headers["IF-MATCH"] = opts.etag;

    const overwrite = opts.overwrite === false ? "false" : "true";
    return SP.request(
      `${folderApi(folderServerRelativeUrl)}/Files/add(url='${encodeURIComponent(fileName)}',overwrite=${overwrite})`,
      { method: "POST", body: content, headers }
    );
  }

  /* Files above ~250 MB need the chunked StartUpload/ContinueUpload/FinishUpload
     dance. If you are shipping a single-file HTML app you will never hit that, so
     it is deliberately not implemented here rather than half-implemented. */

  /* ── Download ─────────────────────────────────────────────────────────── */

  async function readText(fileServerRelativeUrl) {
    const r = await SP.request(`${fileApi(fileServerRelativeUrl)}/$value`, { raw: true });
    return r.text();
  }

  async function readBuffer(fileServerRelativeUrl) {
    const r = await SP.request(`${fileApi(fileServerRelativeUrl)}/$value`, { raw: true });
    return r.arrayBuffer();
  }

  /* ── Optimistic concurrency ───────────────────────────────────────────── */

  const etag = async (fileServerRelativeUrl) => {
    const info = await SP.get(`${fileApi(fileServerRelativeUrl)}?$select=ETag`);
    return (info && info.ETag) || null;
  };

  /* Read → transform → write, with the version checked. A 412 or 409 means
     someone saved between your read and your write; re-read and re-apply rather
     than overwriting their work. */
  async function replaceText(fileServerRelativeUrl, transform, attempts) {
    const max = attempts || 4;
    const slash = fileServerRelativeUrl.lastIndexOf("/");
    const folder = fileServerRelativeUrl.slice(0, slash);
    const name = fileServerRelativeUrl.slice(slash + 1);

    for (let attempt = 1; attempt <= max; attempt++) {
      const version = await etag(fileServerRelativeUrl);
      const current = await readText(fileServerRelativeUrl);
      const next = await transform(current);
      try {
        await upload(folder, name, next, { etag: version, contentType: "text/plain;charset=utf-8" });
        return next;
      } catch (err) {
        if (!/→ (412|409)/.test(String(err.message)) || attempt === max) throw err;
        await new Promise((r) => setTimeout(r, 150 * attempt * attempt));
      }
    }
    throw new Error(`replaceText(${name}): lost ${max} races against concurrent writers`);
  }

  async function remove(fileServerRelativeUrl) {
    try {
      await SP.del(fileApi(fileServerRelativeUrl));
    } catch (err) {
      if (!/→ 404/.test(String(err.message))) throw err;
    }
  }

  /* ── Parsing a pasted SharePoint URL ──────────────────────────────────── */

  /* Users paste what is in their address bar, which is one of several shapes.
     Accepting all of them is a small function and a large usability win — this is
     the pattern PULSE CODE uses to let someone connect a project by paste. */
  function parseSharePointUrl(raw) {
    if (!raw) return null;
    try {
      const url = new URL(String(raw).trim());

      /* Library view: ?id=/sites/Foo/Shared Documents/Bar  (or RootFolder=) */
      const idParam = url.searchParams.get("id") || url.searchParams.get("RootFolder");
      if (idParam) {
        const decoded = decodeURIComponent(idParam);
        const m = decoded.match(/^(\/sites\/[^/]+|\/personal\/[^/]+)/i);
        if (m) {
          return {
            siteUrl: `${url.protocol}//${url.host}${m[1]}`,
            folder: decoded.slice(m[1].length).replace(/^\//, ""),
          };
        }
      }

      /* Direct folder path, possibly ending in a page or a Forms/ segment. */
      const pathname = decodeURIComponent(url.pathname)
        .replace(/\/[^/]+\.aspx$/i, "")
        .replace(/\/Forms$/i, "");
      const m2 = pathname.match(/^(\/sites\/[^/]+|\/personal\/[^/]+)(\/.*)?$/i);
      if (m2) {
        return {
          siteUrl: `${url.protocol}//${url.host}${m2[1]}`,
          folder: (m2[2] || "").replace(/^\//, ""),
        };
      }

      return { siteUrl: `${url.protocol}//${url.host}`, folder: pathname.replace(/^\//, "") };
    } catch (err) {
      return null;
    }
  }

  return {
    folderExists,
    ensureFolderPath,
    listFolder,
    upload,
    readText,
    readBuffer,
    etag,
    replaceText,
    remove,
    parseSharePointUrl,
  };
})();

if (typeof module !== "undefined" && module.exports) module.exports = SPFiles;

/* Example — upload a generated .docx without corrupting it:

     const blob = await buildDocx();                       // a real Blob
     await SPFiles.upload("/sites/Team/Shared Documents/Reports", "weekly.docx", blob, {
       contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
     });
*/

/* UNCLASSIFIED */
