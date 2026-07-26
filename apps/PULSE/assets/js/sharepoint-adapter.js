/* AEWTTR-PULSE — SharePoint REST adapter.
   Every call into SharePoint goes through this module. No Microsoft Graph,
   no Azure app registration, no backend — same-site SharePoint REST only,
   using the browser session the user already has (CAC/PIV/SSO happens
   before this page ever loads). */

const sharePointAdapter = (function () {
  const existingFieldCache = {};

  function pushDebugLog(level, area, message, data) {
    if (!window.AEWTTR) window.AEWTTR = {};
    if (!window.AEWTTR.debugLog) window.AEWTTR.debugLog = [];
    window.AEWTTR.debugLog.push({
      ts: new Date().toISOString(),
      level: level || "info",
      area: area || "sharepoint",
      message: message || "",
      data: data == null ? null : data
    });
    if (window.AEWTTR.debugLog.length > 500) {
      window.AEWTTR.debugLog = window.AEWTTR.debugLog.slice(-500);
    }
  }

  function odataStringValue(value) {
    return String(value == null ? "" : value).replace(/'/g, "''");
  }

  function ensureDiagnosticsState() {
    if (!window.AEWTTR) window.AEWTTR = {};
    if (!window.AEWTTR.spDiagnostics) {
      window.AEWTTR.spDiagnostics = {
        detectedSiteUrl: "",
        currentPageUrl: window.location.href,
        hasPageContext: !!window._spPageContextInfo,
        currentUser: null,
        currentRole: null,
        restCallsSucceeding: false,
        lastRestUrl: "",
        lastRestStatus: null,
        lastRestErrorText: ""
      };
    }
    window.AEWTTR.spDiagnostics.currentPageUrl = window.location.href;
    window.AEWTTR.spDiagnostics.hasPageContext = !!window._spPageContextInfo;
    return window.AEWTTR.spDiagnostics;
  }

  function recordDiagnostics(patch) {
    Object.assign(ensureDiagnosticsState(), patch);
  }

  function getDiagnostics() {
    return ensureDiagnosticsState();
  }

  function listGetByTitlePath(listTitle) {
    return `/_api/web/lists/getbytitle('${odataStringValue(listTitle)}')`;
  }

  /* ---------- environment / site detection ----------
     _spPageContextInfo is injected by SharePoint's own master-page chrome.
     It's a reliable signal when present, but Firepit (and other embeds that
     load this app's index.html directly into a web part rather than through
     SharePoint's normal page render) often never injects it, even though
     the page is same-origin and SharePoint REST works fine from it. So
     detection can't depend on that global alone — it falls back to a live
     probe of /_api/web/currentuser against a handful of candidate site
     URLs, which is the real ground truth. */

  function getSiteUrl() {
    ensureDiagnosticsState();
    if (window._spPageContextInfo && window._spPageContextInfo.webAbsoluteUrl) {
      const siteUrl = window._spPageContextInfo.webAbsoluteUrl.replace(/\/$/, "");
      recordDiagnostics({ detectedSiteUrl: siteUrl, detectedSiteUrlMethod: "webAbsoluteUrl" });
      return siteUrl;
    }
    // manualSharePointSiteUrl is checked BEFORE the webServerRelativeUrl guess
    // below on purpose: that guess assumes window.location.origin IS the
    // SharePoint site's own origin, which is only true if the page is loaded
    // through SharePoint's normal page chrome. Firepit and similar embeds
    // can inject a partial _spPageContextInfo (webServerRelativeUrl present,
    // webAbsoluteUrl missing) while actually serving this page's HTML/JS
    // from a completely different host — in that case the guess below
    // silently builds a URL on the WRONG origin, and every REST call fails
    // with an opaque network/CORS error that looks identical to "SharePoint
    // is unreachable." An explicit manualSharePointSiteUrl is always correct
    // by construction, so it should win over an origin-based guess.
    if (APP_CONFIG.manualSharePointSiteUrl) {
      const siteUrl = APP_CONFIG.manualSharePointSiteUrl.replace(/\/$/, "");
      recordDiagnostics({ detectedSiteUrl: siteUrl, detectedSiteUrlMethod: "manualSharePointSiteUrl" });
      return siteUrl;
    }
    if (window._spPageContextInfo && window._spPageContextInfo.webServerRelativeUrl) {
      const siteUrl = (window.location.origin + window._spPageContextInfo.webServerRelativeUrl).replace(/\/$/, "");
      pushDebugLog("info", "site-detection", `Falling back to origin + webServerRelativeUrl guess: ${siteUrl}. If REST calls fail with a network error, this guess is likely wrong for your embed — set APP_CONFIG.manualSharePointSiteUrl instead.`);
      recordDiagnostics({ detectedSiteUrl: siteUrl, detectedSiteUrlMethod: "origin+webServerRelativeUrl (unverified guess)" });
      return siteUrl;
    }
    recordDiagnostics({ detectedSiteUrl: "", detectedSiteUrlMethod: "none" });
    return "";
  }

  async function detectSharePointSite() {
    const siteUrl = getSiteUrl();
    return siteUrl || null;
  }

  /* ---------- error formatting ---------- */

  function formatSpError(response, text, attemptedUrl) {
    const status = response ? response.status : 0;
    let cause = "unknown";
    let friendly = "Something went wrong talking to SharePoint.";
    if (!response) {
      cause = "network";
      friendly = "Could not reach SharePoint. This usually means the app is running outside of SharePoint/Firepit (no same-site REST endpoint available), or the site URL is wrong."
        + (attemptedUrl ? ` Attempted: ${attemptedUrl}` : "");
    } else if (status === 401) {
      cause = "authentication";
      friendly = "SharePoint did not recognize your session. Try reloading the page inside SharePoint — you may need to sign in again.";
    } else if (status === 403) {
      cause = "permissions";
      friendly = "Your SharePoint account doesn't have permission for this action. Ask a site owner/admin to grant access or run setup.";
    } else if (status === 404) {
      cause = "missing-setup";
      const path = response && response.url ? String(response.url).split("/_api/")[1] : "";
      friendly = path
        ? `SharePoint couldn't find "_api/${path}" — that list, field, folder, or item doesn't exist yet. Use the homepage "Run SharePoint Setup" button to create or repair the SharePoint lists.`
        : "SharePoint couldn't find the list or field this action needs. Use the homepage \"Run SharePoint Setup\" button to create or repair the SharePoint lists.";
    } else if (status === 0) {
      cause = "firepit-sandbox";
      friendly = "The request was blocked before it reached SharePoint. If this is happening inside Firepit, the web part sandbox may be restricting same-site requests.";
    }
    return { status, cause, friendly, raw: text || "" };
  }

  function buildUrl(siteUrl, endpoint) {
    if (/^https?:\/\//i.test(endpoint)) return endpoint;
    return `${siteUrl}${endpoint}`;
  }

  async function sendRequest(siteUrl, endpoint, options) {
    if ((!options || !options.method || options.method === "GET") && !endpoint.includes("_t=")) {
      endpoint += (endpoint.includes("?") ? "&" : "?") + "_t=" + Date.now();
    }
    const url = buildUrl(siteUrl, endpoint);
    const requestOptions = {
      credentials: "same-origin",
      headers: {
        "Accept": "application/json;odata=nometadata"
      },
      ...options
    };
    requestOptions.headers = {
      "Accept": "application/json;odata=nometadata",
      "Cache-Control": "no-cache",
      "Pragma": "no-cache",
      ...(options && options.headers ? options.headers : {})
    };

    recordDiagnostics({ lastRestUrl: url, currentPageUrl: window.location.href, hasPageContext: !!window._spPageContextInfo });
    let response;
    let text = "";
    let contentType = "";

    try {
      response = await fetch(url, requestOptions);
      contentType = response.headers.get("content-type") || "";
      text = await response.text().catch(() => "");
    } catch (networkError) {
      recordDiagnostics({
        lastRestStatus: 0,
        restCallsSucceeding: false,
        lastRestErrorText: String(networkError && networkError.message || networkError || "Network error")
      });
      throw formatSpError(null, String(networkError && networkError.message || networkError || "Network error"), url);
    }

    recordDiagnostics({
      lastRestStatus: response.status,
      restCallsSucceeding: response.ok,
      lastRestErrorText: response.ok ? "" : (text || response.statusText || "Unknown error")
    });

    if (!response.ok) {
      throw formatSpError(response, text, url);
    }

    // Every real SharePoint call funnels through here, so this is the one
    // place that can know "we're talking to SharePoint fine right now."
    // Without this, a single transient failure (a digest race, a momentary
    // 429, one bad poll) got recorded via recordSpError() and stuck in
    // localStorage forever — Admin > SharePoint Setup kept showing
    // "Connection: Error" indefinitely even after hundreds of subsequent
    // successful calls, until someone manually clicked "Clear".
    if (typeof clearLastSpError === "function") clearLastSpError();

    if (!text) return {};
    return contentType.includes("application/json") ? JSON.parse(text) : { text };
  }

  /* ---------- current user ---------- */

  async function getCurrentUserFromRest(siteUrl) {
    const data = await sendRequest(siteUrl, "/_api/web/currentuser", { method: "GET" });
    const currentUser = {
      displayName: data.Title || "",
      email: data.Email || data.UserPrincipalName || "",
      spUserId: data.Id,
      loginName: data.LoginName || "",
      isSiteAdmin: !!data.IsSiteAdmin,
      principalType: data.PrincipalType
    };
    recordDiagnostics({ currentUser: currentUser });
    return currentUser;
  }

  function getCurrentUserFromPageContext() {
    const ctx = window._spPageContextInfo;
    if (!ctx) return null;
    return {
      displayName: ctx.userDisplayName || "",
      email: ctx.userEmail || "",
      spUserId: ctx.userId,
      loginName: ctx.userLoginName || "",
      isSiteAdmin: !!ctx.isSiteAdmin,
      principalType: null
    };
  }

  /* ---------- request digest ---------- */

  async function getRequestDigest(siteUrl) {
    const data = await sendRequest(siteUrl, "/_api/contextinfo", { method: "POST" });
    return data.FormDigestValue;
  }

  /* ---------- raw REST verbs ---------- */

  async function spGet(siteUrl, endpoint) {
    return sendRequest(siteUrl, endpoint, { method: "GET" });
  }

  async function spPost(siteUrl, endpoint, body) {
    const digest = await getRequestDigest(siteUrl);
    return sendRequest(siteUrl, endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json;odata=nometadata",
        "X-RequestDigest": digest
      },
      body: body ? JSON.stringify(body) : undefined
    });
  }

  async function spPostVerbose(siteUrl, endpoint, body) {
    const digest = await getRequestDigest(siteUrl);
    return sendRequest(siteUrl, endpoint, {
      method: "POST",
      headers: {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest
      },
      body: body ? JSON.stringify(body) : undefined
    });
  }

  async function spMerge(siteUrl, endpoint, body) {
    const digest = await getRequestDigest(siteUrl);
    await sendRequest(siteUrl, endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json;odata=nometadata",
        "X-RequestDigest": digest,
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE"
      },
      body: JSON.stringify(body)
    });
    return true;
  }

  async function spDelete(siteUrl, endpoint) {
    const digest = await getRequestDigest(siteUrl);
    await sendRequest(siteUrl, endpoint, {
      method: "POST",
      headers: {
        "X-RequestDigest": digest,
        "IF-MATCH": "*",
        "X-HTTP-Method": "DELETE"
      }
    });
    return true;
  }

  /* ---------- document library / folder / file store ---------- */

  function getParentSiteUrl(siteUrl) {
    return String(siteUrl || "").replace(/\/[^/]+$/, "");
  }

  function getSiteServerRelativeUrl(siteUrl) {
    const value = String(siteUrl || "");
    try {
      const parsed = new URL(value, window.location.origin);
      return parsed.pathname.replace(/\/$/, "") || "/";
    } catch (e) {
      const trimmed = value.replace(/^https?:\/\/[^/]+/i, "");
      return trimmed.replace(/\/$/, "") || "/";
    }
  }

  function combineServerRelativePath() {
    const parts = Array.prototype.slice.call(arguments)
      .filter(Boolean)
      .map((part, index) => {
        const value = String(part);
        if (index === 0) return value.replace(/\/$/, "");
        return value.replace(/^\/+|\/+$/g, "");
      })
      .filter(Boolean);
    return parts.join("/") || "/";
  }

  function getDataLibraryTitle() {
    return APP_CONFIG.dataLibraryTitle || "PULSE App Data";
  }

  function getDataFolderName() {
    return APP_CONFIG.dataFolderName || "AEWTTR-PULSE";
  }

  function getDataFileName(kind) {
    const files = APP_CONFIG.dataFiles || {};
    return files[kind] || `${kind}.json`;
  }

  function getDataLibraryServerRelativeUrl(siteUrl) {
    return combineServerRelativePath(getSiteServerRelativeUrl(siteUrl), getDataLibraryTitle());
  }

  function getDataFolderServerRelativeUrl(siteUrl) {
    return combineServerRelativePath(getDataLibraryServerRelativeUrl(siteUrl), getDataFolderName());
  }

  function getDataFileServerRelativeUrl(siteUrl, kind) {
    return combineServerRelativePath(getDataFolderServerRelativeUrl(siteUrl), getDataFileName(kind));
  }

  async function resolveDataLibraryServerRelativeUrl(siteUrl) {
    const info = await getListInfo(siteUrl, getDataLibraryTitle());
    return info && info.serverRelativeUrl ? info.serverRelativeUrl : getDataLibraryServerRelativeUrl(siteUrl);
  }

  async function resolveDataFolderServerRelativeUrl(siteUrl) {
    const libraryUrl = await resolveDataLibraryServerRelativeUrl(siteUrl);
    return combineServerRelativePath(libraryUrl, getDataFolderName());
  }

  async function resolveDataFileServerRelativeUrl(siteUrl, kind) {
    const folderUrl = await resolveDataFolderServerRelativeUrl(siteUrl);
    return combineServerRelativePath(folderUrl, getDataFileName(kind));
  }

  async function getFolderInfoByServerRelativeUrl(siteUrl, serverRelativeUrl) {
    return spGet(
      siteUrl,
      `/_api/web/GetFolderByServerRelativeUrl('${odataStringValue(serverRelativeUrl)}')?$select=Exists,Name,ServerRelativeUrl`
    );
  }

  async function folderExists(siteUrl, serverRelativeUrl) {
    try {
      const info = await getFolderInfoByServerRelativeUrl(siteUrl, serverRelativeUrl);
      // SharePoint's GetFolderByServerRelativeUrl frequently returns HTTP 200
      // with Exists:false for a path that was never created — it does NOT
      // reliably 404. Ignoring that field (as this used to) meant every
      // never-created project subfolder was reported as "already exists",
      // so ensureFolderPath skipped creating it, and the next real operation
      // (adding a file into that folder) 404'd for real.
      return !!(info && info.Exists !== false);
    } catch (e) {
      if (e && e.status === 404) return false;
      throw e;
    }
  }

  async function ensureDocumentLibrary(siteUrl, libraryTitle, description) {
    const info = await getListInfo(siteUrl, libraryTitle);
    if (info.exists) return { created: false, serverRelativeUrl: info.serverRelativeUrl || getDataLibraryServerRelativeUrl(siteUrl) };
    await spPost(siteUrl, "/_api/web/lists", {
      Title: libraryTitle,
      Description: description || "",
      BaseTemplate: 101
    });
    const createdInfo = await getListInfo(siteUrl, libraryTitle);
    return { created: true, serverRelativeUrl: createdInfo.serverRelativeUrl || getDataLibraryServerRelativeUrl(siteUrl) };
  }

  async function ensureFolderPath(siteUrl, serverRelativeUrl) {
    const normalized = String(serverRelativeUrl || "").replace(/\/$/, "");
    if (!normalized || normalized === "/") return { created: false, serverRelativeUrl: "/" };
    if (await folderExists(siteUrl, normalized)) {
      return { created: false, serverRelativeUrl: normalized };
    }
    const parent = normalized.slice(0, normalized.lastIndexOf("/")) || "/";
    if (parent && parent !== normalized && parent !== "/") {
      await ensureFolderPath(siteUrl, parent);
    }
    await spPostVerbose(siteUrl, "/_api/web/folders", {
      __metadata: { type: "SP.Folder" },
      ServerRelativeUrl: normalized
    });
    if (!(await folderExists(siteUrl, normalized))) {
      throw {
        friendly: `SharePoint did not confirm the folder after creation: ${normalized}`,
        cause: "folder-not-visible-after-create"
      };
    }
    return { created: true, serverRelativeUrl: normalized };
  }

  async function readTextFile(siteUrl, serverRelativeUrl) {
    const digestState = ensureDiagnosticsState();
    const url = buildUrl(siteUrl, `/_api/web/GetFileByServerRelativeUrl('${odataStringValue(serverRelativeUrl)}')/$value`);
    recordDiagnostics({ lastRestUrl: url, currentPageUrl: window.location.href, hasPageContext: !!window._spPageContextInfo });
    let response;
    let text = "";
    try {
      response = await fetch(url, {
        method: "GET",
        credentials: "same-origin",
        headers: {
          "Accept": "application/json;odata=nometadata,text/plain,*/*"
        }
      });
      text = await response.text().catch(() => "");
    } catch (networkError) {
      recordDiagnostics({
        lastRestStatus: 0,
        restCallsSucceeding: false,
        lastRestErrorText: String(networkError && networkError.message || networkError || "Network error")
      });
      throw formatSpError(null, String(networkError && networkError.message || networkError || "Network error"));
    }
    digestState.lastRestStatus = response.status;
    digestState.restCallsSucceeding = response.ok;
    digestState.lastRestErrorText = response.ok ? "" : (text || response.statusText || "Unknown error");
    if (!response.ok) {
      throw formatSpError(response, text);
    }
    return text;
  }

  async function fileExists(siteUrl, serverRelativeUrl) {
    try {
      await spGet(
        siteUrl,
        `/_api/web/GetFileByServerRelativeUrl('${odataStringValue(serverRelativeUrl)}')?$select=Name,ServerRelativeUrl,Exists`
      );
      return true;
    } catch (e) {
      if (e && e.status === 404) return false;
      throw e;
    }
  }

  async function writeTextFile(siteUrl, serverRelativeUrl, text) {
    const digest = await getRequestDigest(siteUrl);
    const folderUrl = serverRelativeUrl.slice(0, serverRelativeUrl.lastIndexOf("/")) || "/";
    const fileName = serverRelativeUrl.slice(serverRelativeUrl.lastIndexOf("/") + 1);
    await ensureFolderPath(siteUrl, folderUrl);
    const url = buildUrl(
      siteUrl,
      `/_api/web/GetFolderByServerRelativeUrl('${odataStringValue(folderUrl)}')/Files/add(url='${odataStringValue(fileName)}',overwrite=true)`
    );
    recordDiagnostics({ lastRestUrl: url, currentPageUrl: window.location.href, hasPageContext: !!window._spPageContextInfo });
    let response;
    let responseText = "";
    try {
      response = await fetch(url, {
        method: "POST",
        credentials: "same-origin",
        headers: {
          "Accept": "application/json;odata=verbose",
          "Content-Type": "application/octet-stream",
          "X-RequestDigest": digest
        },
        body: new TextEncoder().encode(String(text == null ? "" : text))
      });
      responseText = await response.text().catch(() => "");
    } catch (networkError) {
      recordDiagnostics({
        lastRestStatus: 0,
        restCallsSucceeding: false,
        lastRestErrorText: String(networkError && networkError.message || networkError || "Network error")
      });
      throw formatSpError(null, String(networkError && networkError.message || networkError || "Network error"));
    }
    recordDiagnostics({
      lastRestStatus: response.status,
      restCallsSucceeding: response.ok,
      lastRestErrorText: response.ok ? "" : (responseText || response.statusText || "Unknown error")
    });
    if (!response.ok) {
      throw formatSpError(response, responseText);
    }
    return responseText ? JSON.parse(responseText) : {};
  }

  /* Upload arbitrary binary content (docx, pdf, images, ...) — writeTextFile's
     TextEncoder path would corrupt non-UTF8 bytes, so this takes the raw
     ArrayBuffer straight through as the fetch body. */
  async function writeBinaryFile(siteUrl, serverRelativeUrl, arrayBuffer) {
    const digest = await getRequestDigest(siteUrl);
    const folderUrl = serverRelativeUrl.slice(0, serverRelativeUrl.lastIndexOf("/")) || "/";
    const fileName = serverRelativeUrl.slice(serverRelativeUrl.lastIndexOf("/") + 1);
    await ensureFolderPath(siteUrl, folderUrl);
    const url = buildUrl(
      siteUrl,
      `/_api/web/GetFolderByServerRelativeUrl('${odataStringValue(folderUrl)}')/Files/add(url='${odataStringValue(fileName)}',overwrite=true)`
    );
    recordDiagnostics({ lastRestUrl: url, currentPageUrl: window.location.href, hasPageContext: !!window._spPageContextInfo });
    let response;
    let responseText = "";
    try {
      response = await fetch(url, {
        method: "POST",
        credentials: "same-origin",
        headers: {
          "Accept": "application/json;odata=nometadata",
          "Content-Type": "application/octet-stream",
          "X-RequestDigest": digest
        },
        body: arrayBuffer
      });
      responseText = await response.text().catch(() => "");
    } catch (networkError) {
      recordDiagnostics({
        lastRestStatus: 0,
        restCallsSucceeding: false,
        lastRestErrorText: String(networkError && networkError.message || networkError || "Network error")
      });
      throw formatSpError(null, String(networkError && networkError.message || networkError || "Network error"));
    }
    recordDiagnostics({
      lastRestStatus: response.status,
      restCallsSucceeding: response.ok,
      lastRestErrorText: response.ok ? "" : (responseText || response.statusText || "Unknown error")
    });
    if (!response.ok) throw formatSpError(response, responseText);
    return responseText ? JSON.parse(responseText) : {};
  }

  function absoluteUrlForServerRelative(siteUrl, serverRelativeUrl) {
    try {
      return new URL(siteUrl, window.location.origin).origin + serverRelativeUrl;
    } catch (e) {
      return serverRelativeUrl;
    }
  }

  /* Uploads a File/Blob into "PULSE Documents/{projectLabel}/{docTitle}" (or
     "General/{docTitle}" when no project is linked) — every revision of the
     same document review item shares one folder, named after the item
     itself, instead of a flat "Documents" folder shared by every unrelated
     upload. Creates the document library and nested folders on demand.
     Returns the file's absolute URL and name for storing on the record. */
  async function uploadProjectDocument(siteUrl, projectLabel, docTitle, file) {
    const libraryTitle = "PULSE Documents";
    await ensureDocumentLibrary(siteUrl, libraryTitle, "Files attached from AEWTTR-PULSE Document Review.");
    const libraryInfo = await getListInfo(siteUrl, libraryTitle);
    const libraryUrl = (libraryInfo && libraryInfo.serverRelativeUrl) || combineServerRelativePath(getSiteServerRelativeUrl(siteUrl), libraryTitle);
    const safeProjectLabel = sanitizeSharePointSegment(projectLabel || "General");
    const safeDocTitle = sanitizeSharePointSegment(docTitle || "Untitled Document");
    const folderUrl = combineServerRelativePath(libraryUrl, safeProjectLabel, safeDocTitle);
    const safeFileName = sanitizeSharePointSegment(file.name || "document", { keepDots: true });
    const fileUrl = combineServerRelativePath(folderUrl, safeFileName);
    const arrayBuffer = await file.arrayBuffer();
    await writeBinaryFile(siteUrl, fileUrl, arrayBuffer);
    return { fileUrl: absoluteUrlForServerRelative(siteUrl, fileUrl), fileName: file.name || safeFileName, serverRelativeUrl: fileUrl };
  }

  /* Store issue screenshots as real image files instead of squeezing them into
     a multiline list column. Keeping the binary in PULSE App Data preserves
     the source pixel dimensions and leaves the issue item small and reliable. */
  async function uploadIssueScreenshot(siteUrl, issueCode, imageBlob) {
    if (!imageBlob || typeof imageBlob.arrayBuffer !== "function") {
      throw new Error("A screenshot image is required before it can be uploaded.");
    }
    const store = await ensureAppDataStore(siteUrl);
    const folderUrl = store.folders && store.folders.issueScreenshots
      ? store.folders.issueScreenshots.serverRelativeUrl
      : combineServerRelativePath(store.folder.serverRelativeUrl, "Issue Screenshots");
    await ensureFolderPath(siteUrl, folderUrl);
    const mimeType = String(imageBlob.type || "image/png").toLowerCase();
    const extension = mimeType.includes("webp") ? "webp" : (mimeType.includes("jpeg") || mimeType.includes("jpg") ? "jpg" : "png");
    const stamp = new Date().toISOString().replace(/[-:.TZ]/g, "");
    const nonce = Math.random().toString(36).slice(2, 7).toUpperCase();
    const safeIssueCode = sanitizeSharePointSegment(issueCode || "Issue");
    const fileName = sanitizeSharePointSegment(`${safeIssueCode}-${stamp}-${nonce}.${extension}`, { keepDots: true });
    const fileUrl = combineServerRelativePath(folderUrl, fileName);
    await writeBinaryFile(siteUrl, fileUrl, await imageBlob.arrayBuffer());
    return {
      fileUrl: absoluteUrlForServerRelative(siteUrl, fileUrl),
      fileName,
      serverRelativeUrl: fileUrl,
      folderServerRelativeUrl: folderUrl,
      size: Number(imageBlob.size) || 0,
      contentType: mimeType
    };
  }

  async function deleteIssueScreenshot(siteUrl, serverRelativeUrl) {
    if (!serverRelativeUrl) return false;
    try {
      await spDelete(siteUrl, `/_api/web/GetFileByServerRelativeUrl('${odataStringValue(serverRelativeUrl)}')`);
      return true;
    } catch (error) {
      if (error && error.status === 404) return false;
      throw error;
    }
  }

  function sanitizeSharePointSegment(value, options) {
    options = options || {};
    let cleaned = String(value || "").replace(/[~#%&*{}\\:<>?|"]/g, "-");
    if (!options.keepDots) cleaned = cleaned.replace(/[+]/g, "-");
    else cleaned = cleaned.replace(/[+]/g, "-");
    cleaned = cleaned.replace(/\//g, "-").trim();
    return cleaned || "Untitled";
  }

  async function resolvePulseDocumentsLibraryUrl(siteUrl) {
    const libraryTitle = "PULSE Documents";
    await ensureDocumentLibrary(siteUrl, libraryTitle, "Project and document review files for AEWTTR-PULSE.");
    const libraryInfo = await getListInfo(siteUrl, libraryTitle);
    return (libraryInfo && libraryInfo.serverRelativeUrl) || combineServerRelativePath(getSiteServerRelativeUrl(siteUrl), libraryTitle);
  }

  async function ensureProjectDocumentsRoot(siteUrl, projectLabel) {
    const libraryUrl = await resolvePulseDocumentsLibraryUrl(siteUrl);
    const safeProjectLabel = sanitizeSharePointSegment(projectLabel || "General");
    const projectFolderUrl = combineServerRelativePath(libraryUrl, safeProjectLabel);
    await ensureFolderPath(siteUrl, projectFolderUrl);
    return { libraryUrl, projectFolderUrl, projectLabel: safeProjectLabel };
  }

  function normalizeFolderListing(items) {
    const list = items && (items.value || items.results || items) || [];
    return Array.isArray(list) ? list : [];
  }

  async function listPulseDocumentsFolder(siteUrl, folderServerRelativeUrl) {
    const folderPath = odataStringValue(folderServerRelativeUrl);
    const [foldersData, filesData] = await Promise.all([
      spGet(siteUrl, `/_api/web/GetFolderByServerRelativeUrl('${folderPath}')/Folders?$select=Name,ServerRelativeUrl,TimeLastModified,ItemCount`),
      spGet(siteUrl, `/_api/web/GetFolderByServerRelativeUrl('${folderPath}')/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length,UniqueId`)
    ]);
    const folders = normalizeFolderListing(foldersData).map((item) => ({
      name: item.Name,
      serverRelativeUrl: item.ServerRelativeUrl,
      modified: item.TimeLastModified || "",
      itemCount: item.ItemCount || 0,
      type: "folder"
    })).sort((a, b) => String(a.name).localeCompare(String(b.name), undefined, { sensitivity: "base" }));
    const files = normalizeFolderListing(filesData).map((item) => ({
      name: item.Name,
      serverRelativeUrl: item.ServerRelativeUrl,
      modified: item.TimeLastModified || "",
      size: Number(item.Length) || 0,
      type: "file",
      fileUrl: absoluteUrlForServerRelative(siteUrl, item.ServerRelativeUrl)
    })).sort((a, b) => String(a.name).localeCompare(String(b.name), undefined, { sensitivity: "base" }));
    return { folders, files };
  }

  async function createPulseDocumentsSubfolder(siteUrl, parentFolderUrl, folderName) {
    const safeName = sanitizeSharePointSegment(folderName || "New Folder");
    const folderUrl = combineServerRelativePath(parentFolderUrl, safeName);
    await ensureFolderPath(siteUrl, folderUrl);
    return { name: safeName, serverRelativeUrl: folderUrl };
  }

  async function uploadPulseDocumentsFile(siteUrl, folderServerRelativeUrl, file) {
    const safeFileName = sanitizeSharePointSegment(file.name || "document", { keepDots: true });
    const fileUrl = combineServerRelativePath(folderServerRelativeUrl, safeFileName);
    const arrayBuffer = await file.arrayBuffer();
    await writeBinaryFile(siteUrl, fileUrl, arrayBuffer);
    return {
      name: file.name || safeFileName,
      serverRelativeUrl: fileUrl,
      fileUrl: absoluteUrlForServerRelative(siteUrl, fileUrl),
      size: file.size || 0
    };
  }

  async function deletePulseDocumentsFile(siteUrl, serverRelativeUrl) {
    await spDelete(siteUrl, `/_api/web/GetFileByServerRelativeUrl('${odataStringValue(serverRelativeUrl)}')`);
    return true;
  }

  async function deletePulseDocumentsFolder(siteUrl, serverRelativeUrl) {
    await spDelete(siteUrl, `/_api/web/GetFolderByServerRelativeUrl('${odataStringValue(serverRelativeUrl)}')`);
    return true;
  }

  /* Moves a file already in a PULSE Documents folder into a named subfolder
     of that same parent (e.g. "Previous revisions") — used when a new
     Document Review revision is uploaded, so the prior revision's file
     physically relocates instead of just sitting alongside the new one.
     Creates the subfolder on demand. flags=1 on moveto means overwrite if a
     same-named file is already there (matches a re-upload of the same
     revision file name). Safe no-op if the source file doesn't exist. */
  async function moveFileToSubfolder(siteUrl, fileServerRelativeUrl, subfolderName) {
    const source = String(fileServerRelativeUrl || "");
    if (!source) return null;
    if (!(await fileExists(siteUrl, source))) return null;
    const slash = source.lastIndexOf("/");
    const parentFolderUrl = slash > 0 ? source.slice(0, slash) : "";
    const fileName = slash >= 0 ? source.slice(slash + 1) : source;
    // If the file is already inside the target subfolder, do not nest deeper.
    const parentName = parentFolderUrl.slice(parentFolderUrl.lastIndexOf("/") + 1);
    const safeSubfolder = sanitizeSharePointSegment(subfolderName || "Previous revisions");
    if (parentName.toLowerCase() === safeSubfolder.toLowerCase()) {
      return { serverRelativeUrl: source, fileUrl: absoluteUrlForServerRelative(siteUrl, source) };
    }
    const targetFolderUrl = combineServerRelativePath(parentFolderUrl, safeSubfolder);
    await ensureFolderPath(siteUrl, targetFolderUrl);
    const targetUrl = combineServerRelativePath(targetFolderUrl, fileName);
    await spPost(siteUrl, `/_api/web/GetFileByServerRelativeUrl('${odataStringValue(source)}')/moveto(newurl='${odataStringValue(targetUrl)}',flags=1)`);
    return { serverRelativeUrl: targetUrl, fileUrl: absoluteUrlForServerRelative(siteUrl, targetUrl) };
  }

  async function getFileETag(siteUrl, serverRelativeUrl) {
    const info = await spGet(
      siteUrl,
      `/_api/web/GetFileByServerRelativeUrl('${odataStringValue(serverRelativeUrl)}')?$select=ETag`
    );
    return (info && info.ETag) || "";
  }

  /* Overwrite an existing file's content with optimistic concurrency: the write
     only succeeds if the file's ETag still matches (nobody else saved since we
     last read). A 412 response surfaces as err.status === 412 — callers merge
     and retry. Pass etag "*" to force-overwrite. */
  async function writeTextFileIfMatch(siteUrl, serverRelativeUrl, text, etag) {
    const digest = await getRequestDigest(siteUrl);
    const url = buildUrl(siteUrl, `/_api/web/GetFileByServerRelativeUrl('${odataStringValue(serverRelativeUrl)}')/$value`);
    recordDiagnostics({ lastRestUrl: url, currentPageUrl: window.location.href, hasPageContext: !!window._spPageContextInfo });
    let response;
    let responseText = "";
    try {
      response = await fetch(url, {
        method: "POST",
        credentials: "same-origin",
        headers: {
          "Accept": "application/json;odata=nometadata",
          "Content-Type": "application/json; charset=utf-8",
          "X-RequestDigest": digest,
          "X-HTTP-Method": "PUT",
          "If-Match": etag || "*"
        },
        body: String(text == null ? "" : text)
      });
      responseText = await response.text().catch(() => "");
    } catch (networkError) {
      recordDiagnostics({
        lastRestStatus: 0,
        restCallsSucceeding: false,
        lastRestErrorText: String(networkError && networkError.message || networkError || "Network error")
      });
      throw formatSpError(null, String(networkError && networkError.message || networkError || "Network error"));
    }
    recordDiagnostics({
      lastRestStatus: response.status,
      restCallsSucceeding: response.ok,
      lastRestErrorText: response.ok ? "" : (responseText || response.statusText || "Unknown error")
    });
    if (!response.ok) {
      const err = formatSpError(response, responseText);
      if (response.status === 412 || response.status === 409) {
        err.status = 412;
        err.friendly = "Someone else saved changes at the same time. Retrying with a merge…";
      }
      throw err;
    }
    return getFileETag(siteUrl, serverRelativeUrl).catch(() => "");
  }

  async function readJsonFile(siteUrl, serverRelativeUrl, fallbackValue) {
    let text = "";
    try {
      text = await readTextFile(siteUrl, serverRelativeUrl);
    } catch (e) {
      if (e && e.status === 404) {
        pushDebugLog("info", "file-store", `JSON file missing, using fallback for ${serverRelativeUrl}`, {
          serverRelativeUrl: serverRelativeUrl,
          fallbackUsed: true
        });
        return fallbackValue;
      }
      throw e;
    }
    if (!text) return fallbackValue;
    try {
      return JSON.parse(text);
    } catch (e) {
      pushDebugLog("error", "file-store", `Invalid JSON in ${serverRelativeUrl}`, { text });
      return fallbackValue;
    }
  }

  async function writeJsonFile(siteUrl, serverRelativeUrl, value) {
    return writeTextFile(siteUrl, serverRelativeUrl, JSON.stringify(value, null, 2));
  }

  async function ensureJsonFile(siteUrl, serverRelativeUrl, initialValue) {
    const exists = await fileExists(siteUrl, serverRelativeUrl);
    if (exists) return { created: false, serverRelativeUrl: serverRelativeUrl };
    await writeJsonFile(siteUrl, serverRelativeUrl, initialValue);
    return { created: true, serverRelativeUrl: serverRelativeUrl };
  }

  async function ensureRoleStoreFile(siteUrl) {
    await ensureAppDataStore(siteUrl);
    const rolesPath = await resolveDataFileServerRelativeUrl(siteUrl, "roles");
    return ensureJsonFile(siteUrl, rolesPath, []);
  }

  async function ensureDbStoreFile(siteUrl) {
    await ensureAppDataStore(siteUrl);
    const dbPath = await resolveDataFileServerRelativeUrl(siteUrl, "db");
    return ensureJsonFile(siteUrl, dbPath, {});
  }

  async function ensureAppDataStore(siteUrl) {
    const library = await ensureDocumentLibrary(siteUrl, getDataLibraryTitle(), "AEWTTR-PULSE SharePoint file backend.");
    const folderUrl = combineServerRelativePath(library.serverRelativeUrl, getDataFolderName());
    const folder = await ensureFolderPath(siteUrl, folderUrl);
    const issueScreenshotsFolder = await ensureFolderPath(siteUrl, combineServerRelativePath(folder.serverRelativeUrl, "Issue Screenshots"));
    const dbFile = await ensureJsonFile(siteUrl, combineServerRelativePath(folder.serverRelativeUrl, getDataFileName("db")), {});
    const rolesFile = await ensureJsonFile(siteUrl, combineServerRelativePath(folder.serverRelativeUrl, getDataFileName("roles")), []);
    const peopleFile = await ensureJsonFile(siteUrl, combineServerRelativePath(folder.serverRelativeUrl, getDataFileName("peopleDirectory")), []);
    return {
      library,
      folder,
      folders: { issueScreenshots: issueScreenshotsFolder },
      files: {
        db: dbFile,
        roles: rolesFile,
        peopleDirectory: peopleFile
      }
    };
  }

  async function ensurePeopleDirectoryFile(siteUrl) {
    await ensureAppDataStore(siteUrl);
    const path = await resolveDataFileServerRelativeUrl(siteUrl, "peopleDirectory");
    return ensureJsonFile(siteUrl, path, []);
  }

  /* ---------- list / field schema ---------- */

  async function listExists(siteUrl, listTitle) {
    try {
      await spGet(siteUrl, listGetByTitlePath(listTitle));
      return true;
    } catch (e) {
      if (e && e.status === 404) return false;
      throw e;
    }
  }

  async function getListInfo(siteUrl, listTitle) {
    try {
      const data = await spGet(
        siteUrl,
        `${listGetByTitlePath(listTitle)}?$select=Title,Hidden,RootFolder/ServerRelativeUrl&$expand=RootFolder`
      );
      return {
        exists: true,
        title: data.Title || listTitle,
        hidden: !!data.Hidden,
        serverRelativeUrl: (data.RootFolder && data.RootFolder.ServerRelativeUrl) || ""
      };
    } catch (e) {
      if (e && e.status === 404) {
        return { exists: false, title: listTitle, hidden: false, serverRelativeUrl: "" };
      }
      throw e;
    }
  }

  async function renameList(siteUrl, oldTitle, newTitle) {
    return spMerge(siteUrl, `/_api/web/lists/getbytitle('${encodeURIComponent(oldTitle)}')`, {
      Title: newTitle
    });
  }

  async function deleteList(siteUrl, listTitle) {
    try {
      await spDelete(siteUrl, `/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')`);
      return true;
    } catch (e) {
      if (e && e.status === 404) return false;
      throw e;
    }
  }

  async function createList(siteUrl, listTitle, description) {
    return spPost(siteUrl, "/_api/web/lists", {
      Title: listTitle,
      Description: description || "",
      BaseTemplate: 100 // Generic List
    });
  }

  function normalizeFieldAlias(value) {
    return String(value == null ? "" : value)
      .trim()
      .toLowerCase()
      .replace(/_x[0-9a-f]{4}_/gi, "")
      .replace(/[^a-z0-9]/g, "");
  }

  function buildFieldAliasSet(field) {
    const aliases = new Set();
    [
      field && field.InternalName,
      field && field.Title,
      field && field.StaticName,
      field && field.EntityPropertyName
    ].forEach((value) => {
      const raw = String(value == null ? "" : value).trim();
      if (!raw) return;
      aliases.add(raw);
      const normalized = normalizeFieldAlias(raw);
      if (normalized) aliases.add(normalized);
    });
    return Array.from(aliases);
  }

  function valuesEquivalent(left, right) {
    const normalize = (value) => {
      if (value == null) return "";
      if (typeof value === "boolean") return value ? "true" : "false";
      if (typeof value === "number") return String(value);
      return String(value).trim();
    };
    return normalize(left) === normalize(right);
  }

  async function getExistingFieldDefinitions(siteUrl, listTitle) {
    const cacheKey = `${siteUrl}::${listTitle}`;
    if (existingFieldCache[cacheKey]) {
      return existingFieldCache[cacheKey].map((field) => ({ ...field, aliases: field.aliases ? field.aliases.slice() : [] }));
    }
    const data = await spGet(
      siteUrl,
      `${listGetByTitlePath(listTitle)}/fields?$select=InternalName,Title,StaticName,EntityPropertyName,TypeAsString,Hidden,ReadOnlyField,Sealed,Required`
    );
    const items = data.value || data.d && data.d.results || [];
    const fieldDefs = items
      .filter((field) => field && field.InternalName)
      .map((field) => ({
        InternalName: field.InternalName,
        Title: field.Title || field.InternalName,
        StaticName: field.StaticName || field.InternalName,
        EntityPropertyName: field.EntityPropertyName || field.InternalName,
        TypeAsString: field.TypeAsString || "",
        Hidden: !!field.Hidden,
        ReadOnlyField: !!field.ReadOnlyField,
        Sealed: !!field.Sealed,
        Required: !!field.Required,
        aliases: buildFieldAliasSet(field)
      }));
    existingFieldCache[cacheKey] = fieldDefs.map((field) => ({ ...field, aliases: field.aliases.slice() }));
    pushDebugLog("info", "fields", `Loaded field definitions for ${listTitle}`, fieldDefs);
    return fieldDefs.map((field) => ({ ...field, aliases: field.aliases.slice() }));
  }

  function isWritableCustomField(field) {
    const internalName = String(field && field.InternalName || "");
    if (!internalName || internalName === "Title") return true;
    if (field && field.Hidden) return false;
    if (field && field.ReadOnlyField) return false;
    if (field && field.Sealed) return false;
    if (/^_/.test(internalName)) return false;
    return true;
  }

  async function getExistingFieldNames(siteUrl, listTitle) {
    const fieldDefs = await getExistingFieldDefinitions(siteUrl, listTitle);
    return fieldDefs.map((field) => field.InternalName).filter(Boolean);
  }

  function clearFieldCache(siteUrl, listTitle) {
    const cacheKey = `${siteUrl}::${listTitle}`;
    delete existingFieldCache[cacheKey];
  }

  function updateFieldCache(siteUrl, listTitle, fieldName) {
    const cacheKey = `${siteUrl}::${listTitle}`;
    if (!existingFieldCache[cacheKey]) existingFieldCache[cacheKey] = [];
    const alreadyThere = existingFieldCache[cacheKey].some((field) => field && field.InternalName === fieldName);
    if (!alreadyThere) {
      existingFieldCache[cacheKey].push({
        InternalName: fieldName,
        Title: fieldName,
        StaticName: fieldName,
        EntityPropertyName: fieldName,
        TypeAsString: "",
        Hidden: false,
        ReadOnlyField: false,
        Sealed: false,
        aliases: buildFieldAliasSet({ InternalName: fieldName, Title: fieldName, StaticName: fieldName, EntityPropertyName: fieldName })
      });
    }
  }

  async function filterItemPayloadForExistingFields(siteUrl, listTitle, payload) {
    if (!payload || typeof payload !== "object") return payload;
    const existingFields = await getExistingFieldDefinitions(siteUrl, listTitle);
    const writableFields = existingFields.filter(isWritableCustomField);
    const internalNames = writableFields.map((field) => field.InternalName);
    const allowed = new Set(internalNames.concat(["Title"]));
    const aliasToInternal = new Map();
    writableFields.forEach((field) => {
      (field.aliases || []).forEach((alias) => {
        if (!aliasToInternal.has(alias)) aliasToInternal.set(alias, field.InternalName);
      });
    });
    const filtered = {};
    const dropped = [];
    const remapped = [];
    Object.keys(payload).forEach((key) => {
      if (allowed.has(key)) {
        filtered[key] = payload[key];
        return;
      }
      const normalized = normalizeFieldAlias(key);
      const mappedInternalName = aliasToInternal.get(key) || aliasToInternal.get(normalized);
      if (mappedInternalName && !filtered.hasOwnProperty(mappedInternalName)) {
        filtered[mappedInternalName] = payload[key];
        remapped.push({ from: key, to: mappedInternalName });
        return;
      }
      dropped.push(key);
    });
    pushDebugLog("info", "payload", `Filtered payload for ${listTitle}`, {
      requestedKeys: Object.keys(payload),
      existingFields: internalNames,
      writableFieldMap: writableFields.map((field) => ({
        InternalName: field.InternalName,
        Title: field.Title,
        EntityPropertyName: field.EntityPropertyName,
        Hidden: field.Hidden,
        ReadOnlyField: field.ReadOnlyField,
        Sealed: field.Sealed
      })),
      keptKeys: Object.keys(filtered),
      droppedKeys: dropped,
      remappedKeys: remapped
    });
    return filtered;
  }

  async function fieldExists(siteUrl, listTitle, internalName) {
    const wanted = normalizeFieldAlias(internalName);
    const fields = await getExistingFieldDefinitions(siteUrl, listTitle);
    return fields.some((field) => {
      // Exact internal-name match always counts, even for a sealed/system
      // field — that would mean we created it ourselves under that name.
      if (String(field.InternalName || "") === internalName) return true;
      // SharePoint's own built-in/computed columns (Created, Modified,
      // BaseName, FileLeafRef, and similar) are Sealed and/or ReadOnlyField.
      // Their display name can coincidentally normalize to the same alias as
      // a field we want to create (e.g. built-in "Created_x0020_Date" and
      // our own "CreatedDate" both normalize to "createddate"). Without this
      // exclusion, fieldExists() wrongly reports our field as already
      // present, and ensureField() never creates our real column at all.
      if (field.Sealed || field.ReadOnlyField) return false;
      if (normalizeFieldAlias(field.InternalName) === wanted) return true;
      if (normalizeFieldAlias(field.Title) === wanted) return true;
      if (normalizeFieldAlias(field.StaticName) === wanted) return true;
      if (normalizeFieldAlias(field.EntityPropertyName) === wanted) return true;
      return false;
    });
  }

  function escapeXmlAttr(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  /* field: { name, type: "Text"|"Note"|"Number"|"DateTime"|"Boolean"|"Choice", choices?: string[], required?: bool, numLines?: number } */
  function buildFieldXml(field) {
    const name = escapeXmlAttr(field.name);
    const required = field.required ? "TRUE" : "FALSE";
    if (field.type === "Choice") {
      const choices = (field.choices || []).map(c => `<CHOICE>${escapeXmlAttr(c)}</CHOICE>`).join("");
      return `<Field Type="Choice" DisplayName="${name}" Name="${name}" StaticName="${name}" Required="${required}" Format="Dropdown"><CHOICES>${choices}</CHOICES></Field>`;
    }
    if (field.type === "Note") {
      return `<Field Type="Note" DisplayName="${name}" Name="${name}" StaticName="${name}" Required="${required}" NumLines="${field.numLines || 6}" RichText="FALSE" />`;
    }
    if (field.type === "DateTime") {
      return `<Field Type="DateTime" DisplayName="${name}" Name="${name}" StaticName="${name}" Required="${required}" Format="DateTime" />`;
    }
    if (field.type === "Boolean") {
      return `<Field Type="Boolean" DisplayName="${name}" Name="${name}" StaticName="${name}" Required="${required}" />`;
    }
    if (field.type === "Number") {
      return `<Field Type="Number" DisplayName="${name}" Name="${name}" StaticName="${name}" Required="${required}" />`;
    }
    // default: Text
    return `<Field Type="Text" DisplayName="${name}" Name="${name}" StaticName="${name}" Required="${required}" />`;
  }

  function buildFieldRestPayload(field) {
    const base = {
      __metadata: { type: "SP.Field" },
      Title: field.name,
      StaticName: field.name,
      InternalName: field.name,
      Required: !!field.required,
      EnforceUniqueValues: false
    };

    if (field.type === "Choice") {
      return {
        ...base,
        FieldTypeKind: 6,
        EditFormat: 0,
        FillInChoice: false,
        Choices: { results: field.choices || [] },
        DefaultValue: field.choices && field.choices.length ? field.choices[0] : ""
      };
    }
    if (field.type === "Note") {
      return {
        ...base,
        FieldTypeKind: 3,
        NumberOfLines: field.numLines || 6,
        RichText: false,
        AppendOnly: false,
        RestrictedMode: false,
        AllowHyperlink: false
      };
    }
    if (field.type === "DateTime") {
      return {
        ...base,
        FieldTypeKind: 4,
        DisplayFormat: 0,
        DateTimeCalendarType: 0,
        FriendlyDisplayFormat: 0
      };
    }
    if (field.type === "Boolean") {
      return {
        ...base,
        FieldTypeKind: 8,
        DefaultValue: "0"
      };
    }
    if (field.type === "Number") {
      return {
        ...base,
        FieldTypeKind: 9
      };
    }
    return {
      ...base,
      FieldTypeKind: 2,
      MaxLength: 255
    };
  }

  async function createFieldAsXml(siteUrl, listTitle, fieldXml) {
    return spPostVerbose(
      siteUrl,
      `${listGetByTitlePath(listTitle)}/fields/createfieldasxml`,
      {
        parameters: {
          __metadata: { type: "SP.XmlSchemaFieldCreationInformation" },
          SchemaXml: fieldXml,
          // 8  = addFieldInternalNameHint: honor the Name attribute in the XML
          //      as the field's InternalName. WITHOUT this flag SharePoint
          //      auto-generates internal names like "field_7", which is exactly
          //      the bug that broke earlier list-based setups on this tenant.
          // 4  = addToAllContentTypes: make the column visible on the list's
          //      default form/content type instead of only the raw list.
          // 16 = addFieldToDefaultView: show the column when the list is opened
          //      (otherwise the default view keeps showing only Title).
          Options: 28
        }
      }
    );
  }

  async function createFieldByType(siteUrl, listTitle, field) {
    return spPostVerbose(
      siteUrl,
      `${listGetByTitlePath(listTitle)}/fields`,
      buildFieldRestPayload(field)
    );
  }

  async function ensureField(siteUrl, listTitle, field) {
    const exists = await fieldExists(siteUrl, listTitle, field.name);
    if (exists) return { name: field.name, created: false, verified: true };
    let creationMethod = "xml";
    try {
      await createFieldAsXml(siteUrl, listTitle, buildFieldXml(field));
      pushDebugLog("success", "fields", `Created field ${field.name} on ${listTitle} using XML`, field);
    } catch (e) {
      creationMethod = "typed-fallback";
      try {
        await createFieldByType(siteUrl, listTitle, field);
        pushDebugLog("success", "fields", `Created field ${field.name} on ${listTitle} using typed fallback`, field);
      } catch (fallbackError) {
        const detail = fallbackError && fallbackError.friendly ? fallbackError : { friendly: String(fallbackError && fallbackError.message || fallbackError), raw: "" };
        detail.friendly = `Failed creating field "${field.name}" on list "${listTitle}". XML method error: ${(e && e.friendly) || String(e && e.message || e)}. Typed fallback error: ${detail.friendly}`;
        pushDebugLog("error", "fields", `Failed creating field ${field.name} on ${listTitle}`, {
          field,
          xmlError: e,
          typedError: fallbackError
        });
        throw detail;
      }
    }
    clearFieldCache(siteUrl, listTitle);
    const fieldsAfter = await getExistingFieldDefinitions(siteUrl, listTitle);
    const exact = fieldsAfter.find((f) => f.InternalName === field.name);
    if (!exact) {
      // The field may exist under a mangled internal name (alias match only).
      // That is the exact failure mode Options:12 prevents — surface it loudly
      // instead of letting writes silently target a column the form can't show.
      const aliasHit = fieldsAfter.find((f) => (f.aliases || []).includes(normalizeFieldAlias(field.name)));
      const detail = {
        friendly: aliasHit
          ? `Field "${field.name}" on "${listTitle}" was created with the wrong internal name ("${aliasHit.InternalName}"). Delete that column in SharePoint list settings and run setup again.`
          : `SharePoint reported success creating field "${field.name}" on list "${listTitle}", but the field still was not returned by the SharePoint fields endpoint after creation.`,
        cause: aliasHit ? "wrong-internal-name-after-create" : "field-not-visible-after-create"
      };
      pushDebugLog("error", "fields", `Field ${field.name} on ${listTitle} did not verify after ${creationMethod}`, { field, detail, aliasHit });
      throw detail;
    }
    updateFieldCache(siteUrl, listTitle, field.name);
    return { name: field.name, created: true, verified: true, method: creationMethod };
  }

  /* Relaxes (or re-applies) the Required flag on a single existing field.
     Used to repair lists whose columns predate the app's current schema —
     a field the app no longer writes to but that SharePoint still requires
     a value for will block every create/update on that list. We never
     delete the column (that's real tenant data an admin may still want);
     we just stop SharePoint from demanding a value the app will never send. */
  async function setFieldRequired(siteUrl, listTitle, internalName, required) {
    await spMerge(
      siteUrl,
      `${listGetByTitlePath(listTitle)}/fields/getbyinternalnameortitle('${odataStringValue(internalName)}')`,
      { __metadata: { type: "SP.Field" }, Required: !!required }
    );
    clearFieldCache(siteUrl, listTitle);
  }

  /* Make sure every schema column is visible in the list's default view.
     Fields created in earlier runs (before the addFieldToDefaultView flag)
     exist on the list but never made it into the view — this repairs them. */
  async function ensureDefaultViewFields(siteUrl, listTitle, fieldNames) {
    const added = [];
    let current = [];
    try {
      const data = await spGet(siteUrl, `${listGetByTitlePath(listTitle)}/DefaultView/ViewFields`);
      current = (data.Items && (data.Items.results || data.Items)) || [];
    } catch (e) {
      pushDebugLog("error", "views", `Could not read default view for ${listTitle}`, e);
      return added;
    }
    const present = new Set(current.map((name) => String(name)));
    for (const fieldName of fieldNames) {
      if (present.has(fieldName)) continue;
      try {
        await spPost(siteUrl, `${listGetByTitlePath(listTitle)}/DefaultView/ViewFields/addviewfield('${odataStringValue(fieldName)}')`, {});
        added.push(fieldName);
      } catch (e) {
        pushDebugLog("error", "views", `Could not add ${fieldName} to default view of ${listTitle}`, e);
      }
    }
    if (added.length) pushDebugLog("success", "views", `Added ${added.length} columns to default view of ${listTitle}`, added);
    return added;
  }

  /* schema: { listTitle: { description, fields: [field, ...] }, ... } */
  async function ensureListSchema(siteUrl, schema) {
    const results = [];
    for (const listTitle of Object.keys(schema)) {
      const def = schema[listTitle];
      const listResult = { listTitle, listCreated: false, fields: [], errors: [] };
      const hadList = await listExists(siteUrl, listTitle);
      if (!hadList) {
        await createList(siteUrl, listTitle, def.description || "");
        listResult.listCreated = true;
        clearFieldCache(siteUrl, listTitle);
      }
      for (const field of def.fields || []) {
        try {
          const fieldResult = await ensureField(siteUrl, listTitle, field);
          listResult.fields.push(fieldResult);
        } catch (error) {
          const normalizedError = error && error.friendly ? error : { friendly: String(error && error.message || error) };
          listResult.fields.push({ name: field.name, created: false, verified: false, error: normalizedError });
          listResult.errors.push({ name: field.name, error: normalizedError });
        }
      }
      // Repair pass: surface every schema column in the default view so the
      // list is actually readable when opened in SharePoint.
      try {
        listResult.viewFieldsAdded = await ensureDefaultViewFields(
          siteUrl, listTitle, (def.fields || []).map((field) => field.name)
        );
      } catch (e) {
        listResult.viewFieldsAdded = [];
      }
      results.push(listResult);
    }
    return results;
  }

  /* ---------- generic item CRUD ---------- */

  async function getItems(siteUrl, listTitle, options) {
    options = options || {};
    const params = [];
    if (options.select) params.push(`$select=${options.select}`);
    if (options.filter) params.push(`$filter=${options.filter}`);
    if (options.orderby) params.push(`$orderby=${options.orderby}`);
    if (options.top) params.push(`$top=${options.top}`);
    const qs = params.length ? `?${params.join("&")}` : "";
    let endpoint = `${listGetByTitlePath(listTitle)}/items${qs}`;
    const all = [];
    let guard = 0;
    while (endpoint && guard < 50) {
      const data = await sendRequest(siteUrl, endpoint, { method: "GET" });
      const items = data.value || [];
      all.push(...items);
      endpoint = data["odata.nextLink"] || null;
      guard++;
    }
    return all;
  }

  async function createItem(siteUrl, listTitle, item) {
    const filtered = await filterItemPayloadForExistingFields(siteUrl, listTitle, item);
    return spPost(siteUrl, `${listGetByTitlePath(listTitle)}/items`, filtered);
  }

  async function updateItem(siteUrl, listTitle, itemId, updates) {
    const filtered = await filterItemPayloadForExistingFields(siteUrl, listTitle, updates);
    if (!filtered || !Object.keys(filtered).length) return true;
    return spMerge(siteUrl, `${listGetByTitlePath(listTitle)}/items(${itemId})`, filtered);
  }

  async function deleteItem(siteUrl, listTitle, itemId) {
    return spDelete(siteUrl, `${listGetByTitlePath(listTitle)}/items(${itemId})`);
  }

  async function getItemById(siteUrl, listTitle, itemId, select) {
    const qs = select ? `?$select=${select}` : "";
    return spGet(siteUrl, `${listGetByTitlePath(listTitle)}/items(${itemId})${qs}`);
  }

  async function getCurrentUser(siteUrl) {
    return getCurrentUserFromRest(siteUrl);
  }

  async function getListItems(siteUrl, listTitle, options) {
    return getItems(siteUrl, listTitle, options);
  }

  async function createListItem(siteUrl, listTitle, item) {
    return createItem(siteUrl, listTitle, item);
  }

  async function updateListItem(siteUrl, listTitle, itemId, updates) {
    return updateItem(siteUrl, listTitle, itemId, updates);
  }

  async function deleteListItem(siteUrl, listTitle, itemId) {
    return deleteItem(siteUrl, listTitle, itemId);
  }

  /* ---------- site users ---------- */

  function resolveEmailFromLogin(loginName) {
    const login = String(loginName || "");
    const claimsMatch = login.match(/[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}/);
    if (claimsMatch) return claimsMatch[0];
    return "";
  }

  /* A real person (PrincipalType 1) can have no Email on their SharePoint
     user record — common for CAC/PIV-provisioned accounts that haven't hit
     every service yet, or whose login claim isn't email-shaped. That's a
     missing field, not evidence of being a group/system account — treating
     "no email" as "isGroupOrSystem" silently dropped real people (and every
     downstream consumer of this flag) out of site-user sync. */
  function mapSharePointPrincipal(u) {
    const email = u.Email || u.UserPrincipalName || resolveEmailFromLogin(u.LoginName);
    return {
      spUserId: u.Id,
      displayName: u.Title || "",
      email: email || "",
      loginName: u.LoginName || "",
      isSiteAdmin: !!u.IsSiteAdmin,
      principalType: u.PrincipalType,
      isGroupOrSystem: u.PrincipalType !== 1
    };
  }

  /* Key on identity (email, then login, then SP user id) — NOT displayName.
     The same real person can legitimately come back with two different
     displayName renderings (e.g. one entity resolves to their full name,
     another only resolves as far as their bare email address); keying on
     displayName too meant those didn't dedupe and the same person showed up
     twice. When two entries share an identity key, keep whichever has the
     more useful (non-email-shaped) display name. */
  function dedupeSharePointUsers(users) {
    const byKey = new Map();
    const order = [];
    users.forEach((user) => {
      const key = normalizeComparable(user.email) || normalizeComparable(user.loginName) || (Number(user.spUserId) || normalizeComparable(user.displayName));
      if (!byKey.has(key)) {
        byKey.set(key, user);
        order.push(key);
        return;
      }
      const existing = byKey.get(key);
      const existingIsEmailish = normalizeComparable(existing.displayName) === normalizeComparable(existing.email);
      const candidateIsEmailish = normalizeComparable(user.displayName) === normalizeComparable(user.email);
      if (existingIsEmailish && !candidateIsEmailish) byKey.set(key, user);
    });
    return order.map((key) => byKey.get(key));
  }

  function normalizeGroupTitle(value) {
    return String(value == null ? "" : value).trim().toLowerCase();
  }

  async function fetchAssociatedMemberGroup(siteUrl) {
    try {
      return await spGet(siteUrl, "/_api/web/AssociatedMemberGroup?$select=Id,Title,LoginName");
    } catch (e) {
      return null;
    }
  }

  async function fetchAssociatedOwnerGroup(siteUrl) {
    try {
      return await spGet(siteUrl, "/_api/web/AssociatedOwnerGroup?$select=Id,Title,LoginName");
    } catch (e) {
      return null;
    }
  }

  async function fetchSiteGroupByName(siteUrl, groupName) {
    if (!groupName) return null;
    try {
      return await spGet(siteUrl, `/_api/web/sitegroups/getbyname('${odataStringValue(groupName)}')?$select=Id,Title,LoginName`);
    } catch (e) {
      return null;
    }
  }

  function cleanGuid(value) {
    const str = String(value == null ? "" : value).trim();
    if (!str) return "";
    const match = str.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i);
    return match ? match[0].toLowerCase() : "";
  }

  function extractAadGroupId(value) {
    return cleanGuid(value);
  }

  function normalizeRestCollection(data) {
    if (!data) return [];
    if (Array.isArray(data.value)) return data.value;
    if (data.d && Array.isArray(data.d.results)) return data.d.results;
    return [];
  }

  function mapGraphMemberToSharePointUser(member) {
    const email = member.mail || member.userPrincipalName || "";
    return {
      spUserId: 0,
      displayName: member.displayName || email || "Unknown",
      email: email || "",
      loginName: member.userPrincipalName || email || "",
      isSiteAdmin: false,
      principalType: 1,
      // Graph already scopes to microsoft.graph.user — missing mail is not a group.
      isGroupOrSystem: false
    };
  }

  function finalizeSiteUsers(users) {
    // Keep users even without email (CAC/PIV accounts); only drop non-user principals.
    return dedupeSharePointUsers(users).filter((user) => !user.isGroupOrSystem);
  }

  function getPageContextInfo() {
    if (window._spPageContextInfo) return window._spPageContextInfo;
    try {
      if (window.parent && window.parent !== window && window.parent._spPageContextInfo) return window.parent._spPageContextInfo;
    } catch (e) { /* cross-origin parent */ }
    return null;
  }

  async function resolveUnifiedGroupId(siteUrl) {
    const ctx = getPageContextInfo();
    const contextIds = [
      ctx && ctx.groupId,
      ctx && ctx.siteGroupId,
      ctx && ctx.aadGroupId,
      ctx && ctx.GroupId
    ];
    for (const candidate of contextIds) {
      const cleaned = extractAadGroupId(candidate);
      if (cleaned) return cleaned;
    }

    try {
      const props = await spGet(siteUrl, "/_api/web/allproperties");
      for (const key of ["GroupId", "groupId", "vti_groupid", "SiteGroupId"]) {
        const cleaned = extractAadGroupId(props[key]);
        if (cleaned) return cleaned;
      }
    } catch (e) { /* no property bag access */ }

    const memberGroup = await fetchAssociatedMemberGroup(siteUrl);
    if (memberGroup && memberGroup.Id) {
      const principals = await fetchGroupPrincipalBatch(siteUrl, memberGroup.Id, 0);
      for (const principal of principals) {
        if (principal && principal.PrincipalType !== 1) {
          const nestedId = extractAadGroupId(principal.LoginName) || extractAadGroupId(principal.Title);
          if (nestedId) return nestedId;
        }
      }
    }
    return "";
  }

  async function acquireGraphAccessToken(siteUrl) {
    // The resource must be the Graph endpoint for YOUR tenant's actual
    // cloud (commercial/GCC/GCC High/DoD) — see APP_CONFIG.graphApiBase.
    // Requesting a token for the wrong cloud's Graph either fails outright
    // or (worse) silently returns a token that Graph itself then rejects.
    const resources = [APP_CONFIG.graphApiBase, "00000003-0000-0000-c000-000000000000"];
    let lastError = null;
    for (const resource of resources) {
      try {
        const data = await spPostVerbose(siteUrl, "/_api/SP.OAuth.Token/Acquire", { resource });
        const token = data.access_token
          || (data.d && data.d.Acquire && data.d.Acquire.access_token)
          || (data.d && data.d.access_token);
        if (token) return token;
      } catch (e) {
        lastError = e;
      }
    }
    throw lastError || new Error("Graph token unavailable");
  }

  async function fetchGraphGroupMembers(groupId, accessToken) {
    let url = `${APP_CONFIG.graphApiBase}/v1.0/groups/${groupId}/members/microsoft.graph.user?$select=id,displayName,mail,userPrincipalName&$top=999`;
    const all = [];
    let guard = 0;
    while (url && guard < 25) {
      const response = await fetch(url, {
        credentials: "omit",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json"
        }
      });
      const raw = await response.text().catch(() => "");
      if (!response.ok) {
        throw formatSpError(response, raw || response.statusText);
      }
      const data = raw ? JSON.parse(raw) : {};
      (data.value || []).forEach((member) => {
        all.push(mapGraphMemberToSharePointUser(member));
      });
      url = data["@odata.nextLink"] || null;
      guard += 1;
    }
    return all;
  }

  async function enrichSharePointUserIds(siteUrl, users) {
    const enriched = [];
    for (const user of users) {
      if (user.spUserId) {
        enriched.push(user);
        continue;
      }
      const login = user.loginName || (user.email ? `i:0#.f|membership|${user.email}` : "");
      if (!login) {
        enriched.push(user);
        continue;
      }
      try {
        const ensured = await spPostVerbose(siteUrl, "/_api/web/ensureuser", { logonName: login });
        const spUser = (ensured && ensured.d) || ensured || {};
        const spId = Number(spUser.Id) || 0;
        enriched.push(Object.assign({}, user, {
          spUserId: spId || user.spUserId,
          loginName: spUser.LoginName || login,
          displayName: user.displayName || spUser.Title || user.email,
          email: user.email || spUser.Email || resolveEmailFromLogin(spUser.LoginName)
        }));
      } catch (e) {
        enriched.push(user);
      }
    }
    return enriched;
  }

  async function fetchGroupPrincipalBatch(siteUrl, groupId, skip) {
    const paths = [
      `/_api/web/sitegroups(${groupId})/users?$select=Id,Title,Email,LoginName,IsSiteAdmin,PrincipalType&$top=500&$skip=${skip}`,
      `/_api/web/sitegroups(${groupId})/Users?$select=Id,Title,Email,LoginName,IsSiteAdmin,PrincipalType&$top=500&$skip=${skip}`
    ];
    for (const path of paths) {
      try {
        const data = await spGet(siteUrl, path);
        const batch = normalizeRestCollection(data);
        if (batch.length) return batch;
      } catch (e) { /* try alternate casing */ }
    }
    return [];
  }

  async function expandSiteGroupMembers(siteUrl, groupId, options) {
    options = options || {};
    const visited = options.visited || new Set();
    const graphToken = options.graphToken || null;
    const unifiedGroupId = options.unifiedGroupId || "";
    if (!groupId || visited.has(groupId)) return [];
    visited.add(groupId);

    const users = [];
    let skip = 0;
    for (let page = 0; page < 50; page++) {
      const batch = await fetchGroupPrincipalBatch(siteUrl, groupId, skip);
      if (!batch.length) break;

      for (const principal of batch) {
        const principalType = principal.PrincipalType;
        if (principalType === 1) {
          users.push(mapSharePointPrincipal(principal));
          continue;
        }
        const nestedAadId = extractAadGroupId(principal.LoginName) || extractAadGroupId(principal.Title);
        if (graphToken && nestedAadId) {
          try {
            const graphUsers = await fetchGraphGroupMembers(nestedAadId, graphToken);
            users.push(...graphUsers);
          } catch (e) {
            pushDebugLog("warn", "site-users", "Nested Graph group expansion failed", { nestedAadId, error: String(e && e.message || e) });
          }
          continue;
        }
        if (principalType === 8 && principal.Id) {
          users.push(...await expandSiteGroupMembers(siteUrl, principal.Id, { visited, graphToken, unifiedGroupId }));
        }
      }

      skip += batch.length;
      if (batch.length < 500) break;
    }

    if (!users.length && graphToken && unifiedGroupId) {
      try {
        users.push(...await fetchGraphGroupMembers(unifiedGroupId, graphToken));
      } catch (e) { /* handled upstream */ }
    }
    return users;
  }

  async function fetchSiteGroupUsersPaginated(siteUrl, groupId) {
    return expandSiteGroupMembers(siteUrl, groupId, { visited: new Set() });
  }

  async function usersFromSiteGroup(siteUrl, group) {
    if (!group || !group.Id) return [];
    const users = await fetchSiteGroupUsersPaginated(siteUrl, group.Id);
    return users;
  }

  /* Both search paths below can return "entities" that aren't actually
     resolved people — e.g. ClientPeoplePickerSearchUser's Description falling
     back to raw, low-confidence claim/crawl metadata when a candidate can't
     be fully resolved (a date string, a "/_layouts/15/.../DispForm.aspx?..."
     item path, etc). Naively using that text as the display name is what
     produced two junk-looking suggestions per real search. Guard against it:
     a genuine person needs either a real email or a proper claims-encoded
     login ("i:0#.f|membership|...", "i:05.t|...", etc — always starts with
     "i:0"), and the resolved name must not itself look like a path or a bare
     date. */
  function isValidEmailAddress(value) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || "").trim());
  }
  function isClaimsUserLogin(value) {
    return /^i:0[#.]/.test(String(value || "").trim());
  }
  function looksLikeJunkPersonName(value) {
    const text = String(value || "").trim();
    if (!text) return true;
    if (/\.aspx(\?|$)/i.test(text) || text.startsWith("/") || text.includes("DispForm")) return true;
    if (/^\d{1,4}[/\-.]\d{1,2}[/\-.]\d{1,4}/.test(text)) return true; // bare date, e.g. 7/9/2026
    if (!/[a-zA-Z]/.test(text)) return true; // no letters at all
    return false;
  }
  function isPlausiblePersonEntry(user) {
    // Requiring a specific claims-encoded login format ("i:0#.f|membership|
    // ...") was too strict for tenants whose AccountName/login comes back in
    // a different shape (e.g. a bare UPN, or Windows-claims federated auth
    // in a hybrid on-prem/DoD setup) — real search hits were being thrown
    // away entirely because of this, even though the Search API is already
    // scoped to the "Local People Results" source (sourceid in the query),
    // so anything it returns is already a real person, not arbitrary site
    // content. Just require SOME account identifier (email or login,
    // whatever shape it's in) plus a name that isn't obviously junk.
    const hasRealIdentity = isValidEmailAddress(user.email) || isClaimsUserLogin(user.loginName) || !!String(user.loginName || "").trim();
    return hasRealIdentity && !looksLikeJunkPersonName(user.displayName);
  }

  function parsePeoplePickerResults(payload) {
    let raw = payload;
    if (payload && typeof payload === "object") {
      raw = payload.value != null ? payload.value
        : (payload.d && payload.d.ClientPeoplePickerSearchUser != null ? payload.d.ClientPeoplePickerSearchUser : raw);
    }
    if (typeof raw === "string") {
      try { raw = JSON.parse(raw); } catch (e) { raw = []; }
    }
    if (!Array.isArray(raw)) return [];
    return raw.map((entry) => {
      const entity = entry.EntityData || entry.entityData || {};
      const description = entry.Description || entry.description || "";
      const email = entity.Email || entity.email || resolveEmailFromLogin(description) || resolveEmailFromLogin(entry.Key || entry.key || "");
      // Only fall back to the raw Description when nothing better is
      // available AND it doesn't look like junk metadata — never use it
      // blindly the way this used to.
      const displayText = entry.DisplayText || entry.displayText || "";
      const name = displayText || email || (looksLikeJunkPersonName(description) ? "" : description);
      return {
        spUserId: 0,
        displayName: name,
        email: email || "",
        loginName: entry.Key || entry.key || "",
        isSiteAdmin: false,
        principalType: 1,
        isGroupOrSystem: false
      };
    }).filter(isPlausiblePersonEntry);
  }

  /* Unwraps a REST collection that may come back three different shapes
     depending on odata mode/endpoint: a plain array, an OData-verbose
     "{ results: [...] }" wrapper, or (only here) the Search REST API's own
     "Table.Rows" being the wrapper object one level before that. Getting
     this wrong meant `list` ended up as a non-array object — silently
     returning zero parsed rows regardless of RowCount, which is exactly
     how "12 raw hits, 0 usable" happened: the raw hits were real and
     counted correctly (RowCount is a separate, correctly-read field), but
     the actual row list was never unwrapped so nothing was ever parsed. */
  function unwrapRestArray(value) {
    if (Array.isArray(value)) return value;
    if (value && Array.isArray(value.results)) return value.results;
    return [];
  }

  function parseSearchPeopleResults(data) {
    const table = (((data || {}).PrimaryQueryResult || {}).RelevantResults || {}).Table || {};
    const list = unwrapRestArray(table.Rows);
    if (!list.length) return [];
    return list.map((row) => {
      const cells = {};
      const cellList = unwrapRestArray(row.Cells);
      cellList.forEach((cell) => { cells[cell.Key] = cell.Value; });
      // Different tenants populate different managed properties for the
      // People search result set — don't assume every tenant has WorkEmail/
      // PreferredName populated the same way. Try several candidates for
      // each field rather than gating Title behind requiring an email
      // first (that gate meant a tenant with an empty WorkEmail/PreferredName
      // schema — real hits, just under different property names — got
      // silently dropped as nameless).
      const email = cells.WorkEmail || cells.Mail || cells.EmailAddress || cells["SPS-EmailAddresses"]
        || resolveEmailFromLogin(cells.AccountName || "") || resolveEmailFromLogin(cells.Title || "") || resolveEmailFromLogin(cells.PreferredName || "");
      const name = cells.PreferredName || cells.Title || cells.AccountName || email || "";
      return {
        spUserId: 0,
        displayName: name || "",
        email: email || "",
        loginName: cells.AccountName || "",
        isSiteAdmin: false,
        principalType: 1,
        isGroupOrSystem: false
      };
    }).filter(isPlausiblePersonEntry);
  }

  /* The same "type a name, see clean matching people" search Teams/Outlook
     use for "start new chat" — Microsoft Graph's /users endpoint with a
     $filter across name/email fields. This is the primary path: it returns
     real directory users directly, with none of the claim-provider
     resolution quirks that made the SharePoint People Picker/Search API
     paths below return junk (dates, DispForm.aspx links, etc) for
     low-confidence matches. Requires ConsistencyLevel: eventual for the
     $filter/$count combination Graph needs for "advanced query" support. */
  async function searchTenantPeopleViaGraph(siteUrl, query, maxResults) {
    const token = await acquireGraphAccessToken(siteUrl);
    const q = String(query || "").trim().replace(/"/g, '\\"');
    if (!q) return [];
    // $search (not $filter+startswith) is what Graph's own docs recommend
    // for "type a partial/full name, get matching people" — it's tokenized
    // and case-INsensitive, so a full display name like "John Smith" still
    // matches regardless of the case the user typed it in. $filter with
    // startswith() is case-sensitive and matches only from the very start
    // of the whole field, which is why typing a real, correctly-spelled
    // full name could still come back with zero matches.
    const clauses = ["displayName", "mail", "userPrincipalName", "givenName", "surname"]
      .map((field) => `"${field}:${q}"`).join(" OR ");
    const url = `${APP_CONFIG.graphApiBase}/v1.0/users?$search=${encodeURIComponent(clauses)}&$select=id,displayName,mail,userPrincipalName&$top=${maxResults}`;
    const response = await fetch(url, {
      credentials: "omit",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        ConsistencyLevel: "eventual"
      }
    });
    const raw = await response.text().catch(() => "");
    if (!response.ok) throw formatSpError(response, raw || response.statusText);
    const data = raw ? JSON.parse(raw) : {};
    return (data.value || []).map(mapGraphMemberToSharePointUser).filter((u) => u.email);
  }

  /* This module's own header says "No Microsoft Graph, no Azure app
     registration" — so Graph is tried only as a last-ditch optional
     extra, never as the thing tenant search depends on. The two calls
     that should actually work with zero extra setup, using only the
     SharePoint session the browser already has, are the People Picker
     and the Search API — both native SharePoint REST, no Graph/Azure
     app registration required. Every attempt's outcome (ok/error/count)
     is recorded, in order, on `attempts` and attached to the returned
     array as a non-enumerable property so a caller (see
     getLastTenantSearchAttempts) can show the real per-source cause
     instead of just "no matches" when a tenant genuinely returns nothing
     from every path. */
  async function searchTenantPeople(siteUrl, query, options) {
    options = options || {};
    const q = String(query || "").trim();
    if (q.length < (options.minLength || 2)) return [];
    const maxResults = options.maxResults || 20;
    const attempts = [];

    try {
      // ClientPeoplePickerSearchUser's `queryParams` argument is typed as a
      // plain string on the server (PeoplePickerQueryParameters gets JSON-
      // deserialized from it internally) — sending the params as a nested
      // JSON object here (instead of a JSON-encoded STRING) makes the model
      // binder reject/ignore it, so this call would silently never match
      // anyone.
      const payload = await spPostVerbose(
        siteUrl,
        "/_api/SP.UI.ApplicationPages.PeoplePickerWebServiceInterface/ClientPeoplePickerSearchUser",
        {
          queryParams: JSON.stringify({
            QueryString: q,
            MaximumEntitySuggestions: maxResults,
            AllowEmailAddresses: true,
            AllowOnlyEmailAddresses: false,
            AllowMultipleEntities: false,
            AllUrlZones: false,
            ForceClaims: false,
            SharePointGroupID: 0,
            UrlZone: 0,
            UrlZoneSpecified: false,
            PrincipalType: 1,
            PrincipalSource: 15,
            WebUrl: siteUrl,
            EnabledClaimProviders: "",
            Required: false
          })
        }
      );
      const rawCount = Array.isArray(payload) ? payload.length
        : (payload && payload.value != null ? (Array.isArray(payload.value) ? payload.value.length : 1)
        : (payload && payload.d && payload.d.ClientPeoplePickerSearchUser != null ? 1 : 0));
      const people = dedupeSharePointUsers(parsePeoplePickerResults(payload));
      attempts.push({ source: "People Picker", ok: true, rawCount, matchCount: people.length });
      pushDebugLog("info", "tenant-search", "People Picker search completed", { query: q, rawCount, matchCount: people.length });
      if (people.length) { finishTenantSearch(attempts); return people; }
    } catch (e) {
      const detail = String((e && e.friendly) || (e && e.message) || e);
      attempts.push({ source: "People Picker", ok: false, error: detail });
      pushDebugLog("warn", "tenant-search", "People picker search failed; trying Search API", { query: q, error: detail });
    }

    try {
      const safe = odataStringValue(q);
      const searchData = await spGet(
        siteUrl,
        `/_api/search/query?querytext='${safe}*'&sourceid='b09a7990-05ea-4af9-8bce-5ea644c81364'&rowlimit=${maxResults}&selectproperties='PreferredName,WorkEmail,AccountName,JobTitle,Title,EmailAddress,SPS-EmailAddresses'`
      );
      const people = dedupeSharePointUsers(parseSearchPeopleResults(searchData));
      const rawRows = (((searchData || {}).PrimaryQueryResult || {}).RelevantResults || {}).RowCount;
      attempts.push({ source: "Search API", ok: true, rawCount: rawRows == null ? people.length : rawRows, matchCount: people.length });
      pushDebugLog("info", "tenant-search", "Search API completed", { query: q, rawRows, matchCount: people.length });
      if (people.length) { finishTenantSearch(attempts); return people; }
    } catch (e) {
      const detail = String((e && e.friendly) || (e && e.message) || e);
      attempts.push({ source: "Search API", ok: false, error: detail });
      pushDebugLog("warn", "tenant-search", "Search API people lookup failed", { query: q, error: detail });
    }

    try {
      const graphResults = dedupeSharePointUsers(await searchTenantPeopleViaGraph(siteUrl, q, maxResults));
      attempts.push({ source: "Graph (optional)", ok: true, matchCount: graphResults.length });
      if (graphResults.length) { finishTenantSearch(attempts); return graphResults; }
    } catch (e) {
      const detail = String((e && e.friendly) || (e && e.message) || e);
      attempts.push({ source: "Graph (optional)", ok: false, error: detail });
      pushDebugLog("warn", "tenant-search", "Graph /users search failed (expected if no Azure app registration is set up)", { query: q, error: detail });
    }

    finishTenantSearch(attempts);
    return [];
  }

  let lastTenantSearchAttempts = [];
  function finishTenantSearch(attempts) {
    lastTenantSearchAttempts = attempts;
  }
  function getLastTenantSearchAttempts() {
    return lastTenantSearchAttempts;
  }

  /* "Everyone who has access" means every principal with an actual
     permission grant on the site — Owners + Members + Visitors + any custom
     group + anyone granted access directly (not through a group) — not just
     the Members group. /_api/web/roleassignments returns exactly that in
     one call: one entry per principal that has a role binding on the web.
     Group entries are then expanded to their member users via the existing
     expandSiteGroupMembers helper; direct user grants are included as-is. */
  async function fetchWebRoleAssignmentPrincipals(siteUrl) {
    const data = await spGet(siteUrl, "/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings");
    return normalizeRestCollection(data)
      .map((assignment) => assignment && assignment.Member)
      .filter(Boolean);
  }

  async function acquireExpansionContext(siteUrl) {
    let graphToken = null;
    let unifiedGroupId = "";
    try {
      graphToken = await acquireGraphAccessToken(siteUrl);
    } catch (e) {
      pushDebugLog("warn", "site-users", "Graph token unavailable for nested group expansion", {
        error: String(e && e.message || e)
      });
    }
    try {
      unifiedGroupId = await resolveUnifiedGroupId(siteUrl);
    } catch (e) { /* optional */ }
    return { graphToken, unifiedGroupId };
  }

  /* Source of truth for the People roster / App Roles sync:
     SharePoint AssociatedMemberGroup (site Members) ∪ AssociatedOwnerGroup,
     with nested SP/AAD group expansion + pagination. Prefer this over
     /siteusers (historical User Information List) or roleassignments
     (anyone with any permission, often 80+ stale/visitor principals). */
  async function resolveAssociatedMembershipGroups(siteUrl) {
    const groups = [];
    const seenIds = new Set();
    function pushGroup(group) {
      if (!group || !group.Id || seenIds.has(group.Id)) return;
      seenIds.add(group.Id);
      groups.push(group);
    }

    pushGroup(await fetchAssociatedMemberGroup(siteUrl));
    pushGroup(await fetchAssociatedOwnerGroup(siteUrl));

    if (!groups.length || !groups.some((g) => normalizeGroupTitle(g.Title).includes("member"))) {
      const configuredNames = (typeof APP_CONFIG !== "undefined" && Array.isArray(APP_CONFIG.siteMemberGroupNames))
        ? APP_CONFIG.siteMemberGroupNames
        : [];
      for (const name of configuredNames) {
        pushGroup(await fetchSiteGroupByName(siteUrl, name));
      }
    }
    return groups;
  }

  async function getAssociatedMembershipUsers(siteUrl) {
    const groups = await resolveAssociatedMembershipGroups(siteUrl);
    if (!groups.length) {
      pushDebugLog("warn", "site-users", "No AssociatedMember/Owner group resolved");
      return [];
    }

    const { graphToken, unifiedGroupId } = await acquireExpansionContext(siteUrl);
    const collected = [];
    for (const group of groups) {
      try {
        const users = await expandSiteGroupMembers(siteUrl, group.Id, {
          visited: new Set(),
          graphToken,
          unifiedGroupId
        });
        collected.push(...users);
        pushDebugLog("info", "site-users", "Expanded associated site group", {
          groupId: group.Id,
          title: group.Title,
          count: users.length
        });
      } catch (e) {
        pushDebugLog("warn", "site-users", "Associated group expansion failed", {
          groupId: group.Id,
          title: group.Title,
          error: String(e && e.message || e)
        });
      }
    }

    let users = finalizeSiteUsers(collected);
    // Graph-sourced members often lack SharePoint user Ids — resolve when possible.
    const needsEnrich = users.some((u) => !(Number(u.spUserId) > 0) && (u.loginName || u.email));
    if (needsEnrich) {
      try {
        users = finalizeSiteUsers(await enrichSharePointUserIds(siteUrl, users));
      } catch (e) {
        pushDebugLog("warn", "site-users", "ensureUser enrichment failed", { error: String(e && e.message || e) });
      }
    }

    pushDebugLog("info", "site-users", "Associated membership users", {
      groupCount: groups.length,
      groupTitles: groups.map((g) => g.Title),
      count: users.length
    });
    if (typeof window !== "undefined" && window.AEWTTR) {
      window.AEWTTR.associatedSiteMemberCount = users.length;
      window.AEWTTR.associatedSiteMemberKeys = new Set(users.flatMap((u) => {
        const keys = [];
        if (u.email) keys.push(normalizeComparable(u.email));
        if (u.loginName) keys.push(normalizeComparable(u.loginName));
        if (Number(u.spUserId) > 0) keys.push("spid:" + Number(u.spUserId));
        return keys;
      }));
    }
    return users;
  }

  /* Broader roster: anyone with a current permission grant (role assignments
     + site group membership) plus paginated /siteusers. Used by the Admin
     "Pull SharePoint people" directory import — not for auto-associating
     anyone into PULSE. */
  async function getSharePointPermissionUsers(siteUrl, options) {
    options = options || {};
    const onProgress = typeof options.onProgress === "function" ? options.onProgress : null;
    const principals = [];
    const { graphToken, unifiedGroupId } = await acquireExpansionContext(siteUrl);

    try {
      const roleMembers = await fetchWebRoleAssignmentPrincipals(siteUrl);
      for (const member of roleMembers) {
        if (!member) continue;
        if (member.PrincipalType === 1) {
          principals.push(mapSharePointPrincipal(member));
          continue;
        }
        if (member.Id) {
          try {
            principals.push(...await expandSiteGroupMembers(siteUrl, member.Id, {
              visited: new Set(),
              graphToken,
              unifiedGroupId
            }));
          } catch (e) {
            const nestedUsers = normalizeRestCollection(member.Users);
            nestedUsers.forEach((groupUser) => {
              if (groupUser && groupUser.PrincipalType === 1) principals.push(mapSharePointPrincipal(groupUser));
            });
          }
        }
      }
      pushDebugLog("info", "site-users", "roleassignments expansion ok", { rawCount: principals.length });
      if (onProgress) onProgress({ stage: "permissions", count: finalizeSiteUsers(principals).length });
    } catch (e) {
      pushDebugLog("warn", "site-users", "roleassignments lookup failed", { error: String(e && e.message || e) });
    }

    try {
      const groupsData = await spGet(
        siteUrl,
        "/_api/web/sitegroups?$select=Id,Title"
      );
      for (const group of normalizeRestCollection(groupsData)) {
        if (!group || !group.Id) continue;
        try {
          principals.push(...await expandSiteGroupMembers(siteUrl, group.Id, {
            visited: new Set(),
            graphToken,
            unifiedGroupId
          }));
        } catch (e) { /* continue other groups */ }
      }
      pushDebugLog("info", "site-users", "sitegroups expansion ok");
      if (onProgress) onProgress({ stage: "sitegroups", count: finalizeSiteUsers(principals).length });
    } catch (e) {
      pushDebugLog("warn", "site-users", "sitegroups lookup failed", { error: String(e && e.message || e) });
    }

    try {
      const siteUsers = await getPaginatedSiteUsers(siteUrl);
      principals.push(...siteUsers);
      pushDebugLog("info", "site-users", "siteusers paginated ok", { count: siteUsers.length });
      if (onProgress) onProgress({ stage: "siteusers", count: finalizeSiteUsers(principals).length });
    } catch (e) {
      pushDebugLog("warn", "site-users", "siteusers lookup failed", { error: String(e && e.message || e) });
    }

    const users = finalizeSiteUsers(principals);
    pushDebugLog("info", "site-users", "Site permission users", { count: users.length });
    return users;
  }

  /* Paginate /_api/web/siteusers — follow odata.nextLink until exhausted. */
  async function getPaginatedSiteUsers(siteUrl) {
    let endpoint = "/_api/web/siteusers?$select=Id,Title,Email,LoginName,IsSiteAdmin,PrincipalType&$top=500";
    const all = [];
    let guard = 0;
    while (endpoint && guard < 40) {
      const data = await sendRequest(siteUrl, endpoint, { method: "GET" });
      all.push(...(data.value || []).map(mapSharePointPrincipal));
      const next = data["odata.nextLink"] || data["@odata.nextLink"] || null;
      if (!next) break;
      try {
        const abs = String(next);
        if (/^https?:\/\//i.test(abs)) {
          const u = new URL(abs);
          endpoint = u.pathname + u.search;
          const apiIdx = endpoint.indexOf("/_api/");
          if (apiIdx >= 0) endpoint = endpoint.slice(apiIdx);
        } else {
          endpoint = abs;
        }
      } catch (e) {
        endpoint = null;
      }
      guard++;
    }
    return all;
  }

  /* options.mode:
       "members" — Associated Members + Owners
       "all" (default for bulk pull) — permission grants + siteusers
     options.hideUnaffiliated — legacy alias; prefer mode */
  async function getSharePointSiteUsers(siteUrl, options) {
    options = options || {};
    const hideUnaffiliated = options.hideUnaffiliated != null
      ? !!options.hideUnaffiliated
      : options.mode === "members";
    const mode = options.mode || (hideUnaffiliated ? "members" : "all");

    if (mode === "members") {
      const associated = await getAssociatedMembershipUsers(siteUrl);
      if (associated.length) return associated;
      pushDebugLog("warn", "site-users", "Associated membership empty; falling back to permission users");
    }
    return getSharePointPermissionUsers(siteUrl, options);
  }

  function tenantPeopleToDirectory(users) {
    const byKey = new Map();
    (users || []).forEach((user) => {
      const name = String(user && (user.displayName || user.email || "") || "").trim();
      const email = String(user && user.email || "").trim().toLowerCase();
      if (!name) return;
      if (/dispform\.aspx/i.test(name) || /\.aspx(\?|$)/i.test(name) || name.startsWith("/") || /^\d{1,4}[\/\-.]\d{1,2}[\/\-.]\d{1,4}/.test(name)) return;
      const key = email || normalizeComparable(user && (user.loginName || user.displayName || ""));
      if (!key) return;
      if (!byKey.has(key)) {
        byKey.set(key, {
          id: String(user.spUserId || user.email || user.loginName || user.displayName),
          name: name,
          email: email,
          spUserId: Number(user.spUserId) || 0,
          loginName: user.loginName || "",
          isSiteAdmin: !!user.isSiteAdmin,
          principalType: user.principalType == null ? 1 : user.principalType
        });
      }
    });
    return Array.from(byKey.values()).sort((a, b) => a.name.localeCompare(b.name));
  }

  function normalizeDirectoryPerson(entry) {
    if (!entry) return null;
    const name = String(entry.displayName || entry.name || "").trim();
    const email = String(entry.email || "").trim().toLowerCase();
    const loginName = String(entry.loginName || "").trim();
    const spUserId = Number(entry.spUserId) || 0;
    if (!name && !email && !loginName && !(spUserId > 0)) return null;
    return {
      spUserId,
      displayName: name || email || loginName || "Unknown",
      email,
      loginName,
      isSiteAdmin: !!entry.isSiteAdmin,
      principalType: entry.principalType == null ? 1 : Number(entry.principalType) || 1,
      isGroupOrSystem: false,
      pulledAt: entry.pulledAt || ""
    };
  }

  function directoryPersonKey(person) {
    const email = normalizeComparable(person && person.email);
    if (email) return "e:" + email;
    const login = normalizeComparable(person && person.loginName);
    if (login) return "l:" + login;
    const spid = Number(person && person.spUserId) || 0;
    if (spid > 0) return "id:" + spid;
    return "n:" + normalizeComparable(person && person.displayName);
  }

  function mergePeopleDirectory(existing, incoming) {
    const byKey = new Map();
    (existing || []).forEach((entry) => {
      const n = normalizeDirectoryPerson(entry);
      if (!n) return;
      byKey.set(directoryPersonKey(n), n);
    });
    let added = 0;
    (incoming || []).forEach((entry) => {
      const n = normalizeDirectoryPerson(entry);
      if (!n) return;
      const key = directoryPersonKey(n);
      if (!byKey.has(key)) {
        added += 1;
        byKey.set(key, n);
      } else {
        const prev = byKey.get(key);
        byKey.set(key, {
          ...prev,
          ...n,
          displayName: n.displayName || prev.displayName,
          email: n.email || prev.email,
          loginName: n.loginName || prev.loginName,
          spUserId: n.spUserId || prev.spUserId,
          pulledAt: n.pulledAt || prev.pulledAt
        });
      }
    });
    const people = Array.from(byKey.values()).sort((a, b) =>
      String(a.displayName || "").localeCompare(String(b.displayName || "")));
    return { people, added, total: people.length };
  }

  async function loadPeopleDirectory(siteUrl) {
    try {
      await ensurePeopleDirectoryFile(siteUrl);
      const path = await resolveDataFileServerRelativeUrl(siteUrl, "peopleDirectory");
      const raw = await readJsonFile(siteUrl, path, []);
      const people = (Array.isArray(raw) ? raw : []).map(normalizeDirectoryPerson).filter(Boolean);
      if (typeof window !== "undefined" && window.AEWTTR) {
        window.AEWTTR.peopleDirectory = people;
        window.AEWTTR.peopleDirectoryCount = people.length;
      }
      return people;
    } catch (e) {
      pushDebugLog("warn", "people-directory", "Load failed", { error: String(e && e.message || e) });
      if (typeof window !== "undefined" && window.AEWTTR && Array.isArray(window.AEWTTR.peopleDirectory)) {
        return window.AEWTTR.peopleDirectory;
      }
      return [];
    }
  }

  async function savePeopleDirectory(siteUrl, people) {
    await ensurePeopleDirectoryFile(siteUrl);
    const path = await resolveDataFileServerRelativeUrl(siteUrl, "peopleDirectory");
    const normalized = (people || []).map(normalizeDirectoryPerson).filter(Boolean);
    await writeJsonFile(siteUrl, path, normalized);
    if (typeof window !== "undefined" && window.AEWTTR) {
      window.AEWTTR.peopleDirectory = normalized;
      window.AEWTTR.peopleDirectoryCount = normalized.length;
    }
    return normalized;
  }

  /* Bulk pull: Associated Members + permission users + paginated siteusers.
     Persists into people-directory.json. Does NOT assign anyone to PULSE. */
  async function pullSharePointPeopleDirectory(siteUrl, options) {
    options = options || {};
    const onProgress = typeof options.onProgress === "function" ? options.onProgress : null;
    const nowIso = new Date().toISOString();
    const collected = [];

    if (onProgress) onProgress({ stage: "start", message: "Starting SharePoint people pull…" });

    try {
      const associated = await getAssociatedMembershipUsers(siteUrl);
      associated.forEach((u) => collected.push({ ...u, pulledAt: nowIso }));
      if (onProgress) onProgress({ stage: "members", count: collected.length, message: `Members/Owners: ${associated.length}` });
    } catch (e) {
      pushDebugLog("warn", "people-directory", "Associated members pull failed", { error: String(e && e.message || e) });
    }

    try {
      const permissionUsers = await getSharePointPermissionUsers(siteUrl, {
        onProgress: (p) => {
          if (onProgress) onProgress({
            stage: p.stage,
            count: collected.length + (p.count || 0),
            message: `Scanning ${p.stage}…`
          });
        }
      });
      permissionUsers.forEach((u) => collected.push({ ...u, pulledAt: nowIso }));
    } catch (e) {
      pushDebugLog("warn", "people-directory", "Permission users pull failed", { error: String(e && e.message || e) });
    }

    const incoming = finalizeSiteUsers(collected).map((u) => ({ ...u, pulledAt: nowIso }));
    const existing = await loadPeopleDirectory(siteUrl);
    const merged = mergePeopleDirectory(existing, incoming);
    await savePeopleDirectory(siteUrl, merged.people);
    pushDebugLog("success", "people-directory", "Pull complete", {
      pulled: incoming.length,
      added: merged.added,
      total: merged.total
    });
    if (onProgress) onProgress({
      stage: "done",
      count: merged.total,
      added: merged.added,
      message: `Saved ${merged.total} people (${merged.added} new)`
    });
    return { pulled: incoming.length, added: merged.added, total: merged.total, people: merged.people };
  }

  async function searchTenantPeopleDirectory(siteUrl, query, options) {
    const users = await searchTenantPeople(siteUrl, query, options);
    return tenantPeopleToDirectory(users);
  }

  /* Warm picker directory from AEWTTR-associated people (active App Roles). */
  async function loadSiteMemberDirectory(siteUrl, options) {
    options = options || {};
    let directory = [];
    try {
      const records = await getActiveRoleRecords(siteUrl);
      directory = tenantPeopleToDirectory(records.map((r) => ({
        spUserId: r.SharePointUserId,
        displayName: r.UserDisplayName || r.Title,
        email: r.UserEmail,
        loginName: r.LoginName,
        isSiteAdmin: !!r.IsSiteAdmin,
        principalType: r.PrincipalType
      })));
    } catch (e) {
      pushDebugLog("warn", "site-users", "Active roles directory failed; falling back", { error: String(e && e.message || e) });
      const users = await getSharePointSiteUsers(siteUrl, { mode: "members" });
      directory = tenantPeopleToDirectory(users);
    }
    if (typeof window !== "undefined" && window.AEWTTR) window.AEWTTR.siteMemberDirectory = directory;
    return directory;
  }

  async function uploadProjectCoverImage(siteUrl, projectCode, file) {
    const { projectFolderUrl } = await ensureProjectDocumentsRoot(siteUrl, projectCode);
    const coverFolder = await createPulseDocumentsSubfolder(siteUrl, projectFolderUrl, "_cover");
    const extMatch = String(file.name || "").match(/\.([a-z0-9]+)$/i);
    const ext = extMatch ? extMatch[1].toLowerCase() : "jpg";
    const coverName = `cover.${ext}`;
    const uploadFile = (file.name && file.name.toLowerCase() === coverName.toLowerCase())
      ? file
      : new File([await file.arrayBuffer()], coverName, { type: file.type || "image/jpeg" });
    return uploadPulseDocumentsFile(siteUrl, coverFolder.serverRelativeUrl, uploadFile);
  }

  /* ---------- app roles ---------- */

  const ROLES_LIST = "PULSE App Roles";
  const FILE_ROLE_SOURCE = "SharePoint File Store";

  function normalizeComparable(value) {
    return String(value == null ? "" : value).trim().toLowerCase();
  }

  function cloneJson(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function normalizeRoleRecord(record, index) {
    const normalized = Object.assign({
      Id: index + 1,
      Title: "",
      UserEmail: "",
      UserDisplayName: "",
      SharePointUserId: 0,
      LoginName: "",
      PrincipalType: 1,
      IsSiteAdmin: false,
      Role: APP_CONFIG.defaultUserRole,
      ThemeMode: "",
      HideFromMeetings: false,
      IsActive: true,
      Source: FILE_ROLE_SOURCE,
      CreatedByEmail: "",
      CreatedDate: "",
      UpdatedByEmail: "",
      UpdatedDate: ""
    }, record || {});
    normalized.Id = Number(normalized.Id) || (index + 1);
    normalized.IsSiteAdmin = !!normalized.IsSiteAdmin;
    normalized.IsActive = normalized.IsActive !== false;
    normalized.HideFromMeetings = normalized.HideFromMeetings === true || normalized.HideFromMeetings === 1
      || String(normalized.HideFromMeetings || "").trim().toLowerCase() === "true";
    normalized.Role = normalized.Role || APP_CONFIG.defaultUserRole;
    normalized.Source = normalized.Source || FILE_ROLE_SOURCE;
    normalized.Title = normalized.Title || normalized.UserDisplayName || normalized.UserEmail || `User ${normalized.Id}`;
    return normalized;
  }

  async function loadRoleStore(siteUrl) {
    await ensureAppDataStore(siteUrl);
    const rolesPath = await resolveDataFileServerRelativeUrl(siteUrl, "roles");
    const records = await readJsonFile(siteUrl, rolesPath, []);
    return (records || []).map((record, index) => normalizeRoleRecord(record, index));
  }

  async function saveRoleStore(siteUrl, records) {
    const normalized = (records || []).map((record, index) => normalizeRoleRecord(record, index));
    const rolesPath = await resolveDataFileServerRelativeUrl(siteUrl, "roles");
    await writeJsonFile(siteUrl, rolesPath, normalized);
    return normalized;
  }

  async function loadAppDataFile(siteUrl) {
    await ensureAppDataStore(siteUrl);
    const dbPath = await resolveDataFileServerRelativeUrl(siteUrl, "db");
    return readJsonFile(siteUrl, dbPath, {});
  }

  /* Read db.json together with its ETag so callers can write back with
     optimistic concurrency. */
  async function loadAppDataFileWithMeta(siteUrl) {
    await ensureAppDataStore(siteUrl);
    const dbPath = await resolveDataFileServerRelativeUrl(siteUrl, "db");
    const data = await readJsonFile(siteUrl, dbPath, {});
    const etag = await getFileETag(siteUrl, dbPath).catch(() => "");
    return { data, etag };
  }

  /* Save db.json. When an etag is supplied the write is conditional (fails
     with status 412 if someone else saved first); without one it falls back
     to a plain overwrite. Returns the file's new ETag. */
  async function saveAppDataFile(siteUrl, data, etag) {
    await ensureAppDataStore(siteUrl);
    const dbPath = await resolveDataFileServerRelativeUrl(siteUrl, "db");
    const body = JSON.stringify(data || {}, null, 2);
    if (etag) {
      return writeTextFileIfMatch(siteUrl, dbPath, body, etag);
    }
    await writeTextFile(siteUrl, dbPath, body);
    return getFileETag(siteUrl, dbPath).catch(() => "");
  }

  function roleRecordDisplayName(record) {
    return record ? (record.UserDisplayName || record.Title || record.UserEmail || "Unknown") : "Unknown";
  }

  function buildRoleRecordLookup(records) {
    const byEmail = new Map();
    const byLogin = new Map();
    const bySpUserId = new Map();
    const byTitle = new Map();
    const byDisplayName = new Map();
    records.forEach((record) => {
      const email = normalizeComparable(record && record.UserEmail);
      const login = normalizeComparable(record && record.LoginName);
      const spUserId = Number(record && record.SharePointUserId) || 0;
      const title = normalizeComparable(record && record.Title);
      const displayName = normalizeComparable(record && record.UserDisplayName);
      if (email && !byEmail.has(email)) byEmail.set(email, record);
      if (login && !byLogin.has(login)) byLogin.set(login, record);
      if (spUserId > 0 && !bySpUserId.has(spUserId)) bySpUserId.set(spUserId, record);
      if (title && !byTitle.has(title)) byTitle.set(title, record);
      if (displayName && !byDisplayName.has(displayName)) byDisplayName.set(displayName, record);
    });
    return { byEmail, byLogin, bySpUserId, byTitle, byDisplayName };
  }

  function findExistingRoleRecord(lookup, user) {
    const email = normalizeComparable(user && user.email);
    const login = normalizeComparable(user && user.loginName);
    const spUserId = Number(user && user.spUserId) || 0;
    const displayName = normalizeComparable(user && user.displayName);
    return (email && lookup.byEmail.get(email))
      || (login && lookup.byLogin.get(login))
      || (spUserId > 0 && lookup.bySpUserId.get(spUserId))
      || (displayName && lookup.byDisplayName.get(displayName))
      || (displayName && lookup.byTitle.get(displayName))
      || null;
  }

  function recordTitleMatchesUser(record, normalizedEmail, normalizedLogin, normalizedName) {
    const title = normalizeComparable(record && record.Title);
    if (!title) return false;
    if (normalizedEmail && title === normalizedEmail) return true;
    if (normalizedLogin && (title === normalizedLogin || normalizedLogin.endsWith(title) || title.endsWith(normalizedLogin))) return true;
    if (normalizedName && title === normalizedName) return true;
    return false;
  }

  function findRoleRecordForUser(records, currentUser) {
    const normalizedEmail = normalizeComparable(currentUser && currentUser.email);
    const normalizedLogin = normalizeComparable(currentUser && currentUser.loginName);
    const normalizedName = normalizeComparable(currentUser && currentUser.displayName);
    const spUserId = Number(currentUser && currentUser.spUserId) || 0;

    const matchers = [
      {
        source: "UserEmail",
        predicate(record) {
          return normalizedEmail && normalizeComparable(record.UserEmail) === normalizedEmail;
        }
      },
      {
        source: "LoginName",
        predicate(record) {
          const recordLogin = normalizeComparable(record.LoginName);
          if (!normalizedLogin || !recordLogin) return false;
          return recordLogin === normalizedLogin
            || recordLogin.endsWith(normalizedEmail)
            || normalizedLogin.endsWith(recordLogin);
        }
      },
      {
        source: "SharePointUserId",
        predicate(record) {
          return spUserId > 0 && Number(record.SharePointUserId) === spUserId;
        }
      },
      {
        source: "UserDisplayName",
        predicate(record) {
          return normalizedName && normalizeComparable(roleRecordDisplayName(record)) === normalizedName;
        }
      },
      {
        source: "Title",
        predicate(record) {
          return recordTitleMatchesUser(record, normalizedEmail, normalizedLogin, normalizedName);
        }
      }
    ];

    for (const matcher of matchers) {
      const record = records.find(matcher.predicate);
      if (record) {
        return { record, source: matcher.source };
      }
    }
    return { record: null, source: "Default" };
  }

  function isRoleRecordActive(record) {
    const value = record ? record.IsActive : null;
    if (value === false || value === 0) return false;
    if (typeof value === "string") {
      const normalized = value.trim().toLowerCase();
      if (normalized === "false" || normalized === "0" || normalized === "no" || normalized === "inactive") return false;
    }
    return true;
  }

  /* Serialize a role record into a list-item payload. Only schema columns are
     sent; DateTime columns must be ISO strings or null (SharePoint rejects ""). */
  function roleRecordToItemPayload(record) {
    const dateOrNull = (value) => (value ? String(value) : null);
    const payload = {};
    if ("Title" in record) payload.Title = record.Title || record.UserDisplayName || record.UserEmail || "Unknown";
    ["UserEmail", "UserDisplayName", "LoginName", "Role", "Source", "CreatedByEmail", "UpdatedByEmail", "JobTitle", "SpoAccess", "PowerBiAccess", "ThemeMode", "NotificationPrefsJson"].forEach((key) => {
      if (key in record) payload[key] = record[key] == null ? "" : String(record[key]);
    });
    if ("SharePointUserId" in record) payload.SharePointUserId = Number(record.SharePointUserId) || 0;
    if ("PrincipalType" in record) payload.PrincipalType = Number(record.PrincipalType) || 1;
    if ("IsSiteAdmin" in record) payload.IsSiteAdmin = !!record.IsSiteAdmin;
    if ("IsActive" in record) payload.IsActive = record.IsActive !== false;
    if ("HideFromMeetings" in record) payload.HideFromMeetings = !!record.HideFromMeetings;
    if ("CreatedDate" in record) payload.CreatedDate = dateOrNull(record.CreatedDate);
    if ("UpdatedDate" in record) payload.UpdatedDate = dateOrNull(record.UpdatedDate);
    return payload;
  }

  async function createRoleItem(siteUrl, record) {
    const payload = roleRecordToItemPayload(record);
    const created = await createItem(siteUrl, ROLES_LIST, payload);
    return (created && created.Id) || 0;
  }

  async function updateRoleItem(siteUrl, itemId, updates) {
    return updateItem(siteUrl, ROLES_LIST, itemId, roleRecordToItemPayload(updates));
  }

  /* Roles live in the "PULSE App Roles" SharePoint list (one item per user).
     A missing list just means setup hasn't run — treat as no roles. */
  async function getRoleRecords(siteUrl) {
    try {
      const items = await getItems(siteUrl, ROLES_LIST);
      return (items || []).map((item, index) => normalizeRoleRecord(item, index));
    } catch (e) {
      if (e && e.status === 404) return [];
      throw e;
    }
  }

  async function getActiveRoleRecords(siteUrl) {
    const all = await getRoleRecords(siteUrl);
    return all.filter(isRoleRecordActive);
  }

  async function getCurrentUserRole(siteUrl, currentUser) {
    const records = await getActiveRoleRecords(siteUrl);
    const match = findRoleRecordForUser(records, currentUser);
    if (match.record) {
      const role = match.record.Role || APP_CONFIG.defaultUserRole;
      const roleInfo = {
        role,
        isAdmin: role === "Admin",
        isMeetingAdmin: role === "Admin" || role === "Meeting Admin",
        isFinanceAdmin: role === "Admin" || role === "Finance Admin",
        isDocAdmin: role === "Admin" || role === "Document Admin",
        roleRecordId: match.record.Id,
        associated: true,
        themeMode: match.record.ThemeMode || "",
        notificationPrefs: safeJsonParse(match.record.NotificationPrefsJson, null),
        source: match.record.Source || match.source,
        matchedBy: match.source,
        matchedRecordName: roleRecordDisplayName(match.record)
      };
      recordDiagnostics({ currentRole: roleInfo });
      return roleInfo;
    }
    const defaultRoleInfo = {
      role: APP_CONFIG.defaultUserRole,
      isAdmin: false,
      isMeetingAdmin: false,
      isFinanceAdmin: false,
      isDocAdmin: false,
      roleRecordId: null,
      associated: false,
      themeMode: "",
      notificationPrefs: null,
      source: "Default",
      matchedBy: "Default",
      matchedRecordName: ""
    };
    recordDiagnostics({ currentRole: defaultRoleInfo });
    return defaultRoleInfo;
  }

  async function isCurrentUserAdmin(siteUrl, currentUser) {
    const role = await getCurrentUserRole(siteUrl, currentUser);
    return role.isAdmin;
  }

  async function syncSiteUsersToAppRoles(siteUrl, currentUser) {
    /* Refresh identity fields on existing AEWTTR (active App Roles) people.
       Does NOT auto-create associations — that is Assign to PULSE / setup bootstrap. */
    const [rawSiteUsers, existing] = await Promise.all([
      getSharePointSiteUsers(siteUrl, { mode: "members" }).catch(() => []),
      getRoleRecords(siteUrl)
    ]);
    const siteUsers = rawSiteUsers.slice();
    const lookup = buildRoleRecordLookup(existing);
    const created = [];
    const updated = [];
    const deactivated = [];
    const skipped = [];
    const nowIso = new Date().toISOString();
    pushDebugLog("info", "roles-sync", "Refreshing existing PULSE App Roles from site Members", {
      rawSiteUserCount: rawSiteUsers.length,
      existingRoleRecordCount: existing.length
    });
    for (const user of siteUsers) {
      if (!user.email && !user.loginName && !(Number(user.spUserId) > 0)) {
        skipped.push({ displayName: user.displayName, reason: "no usable identity" });
        continue;
      }
      const existingRecord = findExistingRoleRecord(lookup, user);
      if (!existingRecord || !isRoleRecordActive(existingRecord)) continue;
      const updates = {
        Title: user.displayName,
        UserEmail: user.email,
        UserDisplayName: user.displayName,
        SharePointUserId: user.spUserId,
        LoginName: user.loginName,
        PrincipalType: user.principalType,
        UpdatedByEmail: (currentUser && currentUser.email) || "",
        UpdatedDate: nowIso
      };
      await updateRoleItem(siteUrl, existingRecord.Id, updates);
      updated.push(user.displayName);
    }
    pushDebugLog("success", "roles-sync", "Finished PULSE App Roles refresh", {
      created, updated, deactivated, skipped
    });
    return { created, updated, deactivated, skipped };
  }

  /* Assign a directory person into PULSE (active App Roles = associated). */
  async function assignUserToPulse(siteUrl, targetUser, currentUser, role) {
    const isAdmin = await isCurrentUserAdmin(siteUrl, currentUser);
    if (!isAdmin) throw new Error("You do not have permission to assign people to PULSE.");
    const all = await getRoleRecords(siteUrl);
    const nowIso = new Date().toISOString();
    const match = findRoleRecordForUser(all, targetUser);
    const desiredRole = role || APP_CONFIG.defaultUserRole;
    if (match.record) {
      await updateRoleItem(siteUrl, match.record.Id, {
        Title: targetUser.displayName || match.record.Title,
        UserEmail: targetUser.email || match.record.UserEmail || "",
        UserDisplayName: targetUser.displayName || match.record.UserDisplayName || "",
        SharePointUserId: targetUser.spUserId || match.record.SharePointUserId || 0,
        LoginName: targetUser.loginName || match.record.LoginName || "",
        PrincipalType: targetUser.principalType || match.record.PrincipalType || 1,
        IsSiteAdmin: !!targetUser.isSiteAdmin,
        Role: match.record.Role || desiredRole,
        IsActive: true,
        UpdatedByEmail: currentUser.email || "",
        UpdatedDate: nowIso
      });
      return { id: match.record.Id, created: false, reactivated: !isRoleRecordActive(match.record) };
    }
    const newId = await createRoleItem(siteUrl, {
      Title: targetUser.displayName || targetUser.email || "User",
      UserEmail: targetUser.email || "",
      UserDisplayName: targetUser.displayName || "",
      SharePointUserId: targetUser.spUserId || 0,
      LoginName: targetUser.loginName || "",
      PrincipalType: targetUser.principalType || 1,
      IsSiteAdmin: !!targetUser.isSiteAdmin,
      Role: desiredRole,
      IsActive: true,
      Source: "Manual",
      CreatedByEmail: currentUser.email || "",
      CreatedDate: nowIso,
      UpdatedByEmail: currentUser.email || "",
      UpdatedDate: nowIso
    });
    return { id: newId, created: true, reactivated: false };
  }

  async function unassignUserFromPulse(siteUrl, targetUser, currentUser) {
    const isAdmin = await isCurrentUserAdmin(siteUrl, currentUser);
    if (!isAdmin) throw new Error("You do not have permission to remove people from PULSE.");
    const all = await getRoleRecords(siteUrl);
    const match = findRoleRecordForUser(all, targetUser);
    if (!match.record || !isRoleRecordActive(match.record)) {
      return { id: match.record ? match.record.Id : null, removed: false };
    }
    if (match.record.Role === "Admin") {
      const activeAdmins = all.filter((r) => isRoleRecordActive(r) && r.Role === "Admin");
      if (activeAdmins.length <= 1) throw new Error("Can't remove the last active Admin. Promote someone else first.");
    }
    await updateRoleItem(siteUrl, match.record.Id, {
      IsActive: false,
      UpdatedByEmail: currentUser.email || "",
      UpdatedDate: new Date().toISOString()
    });
    return { id: match.record.Id, removed: true };
  }

  function isUserPulseAssociated(roleInfo) {
    if (!roleInfo) return false;
    if (roleInfo.isAdmin) return true;
    if (roleInfo.roleRecordId) return true;
    if (roleInfo.source && roleInfo.source !== "Default" && roleInfo.matchedBy && roleInfo.matchedBy !== "Default") return true;
    return false;
  }

  async function probeRoleRecordWrite(siteUrl, currentUser) {
    const nowIso = new Date().toISOString();
    const allRecords = await getRoleRecords(siteUrl);
    const match = findRoleRecordForUser(allRecords, currentUser);
    const payload = {
      Title: currentUser.displayName || currentUser.email || "Current User",
      UserEmail: currentUser.email || "",
      UserDisplayName: currentUser.displayName || "",
      SharePointUserId: currentUser.spUserId || 0,
      LoginName: currentUser.loginName || "",
      PrincipalType: currentUser.principalType || 1,
      IsActive: true,
      Source: match.record && match.record.Source ? match.record.Source : "SharePoint Site Users",
      Role: match.record && match.record.Role ? match.record.Role : APP_CONFIG.defaultUserRole,
      UpdatedByEmail: currentUser.email || "",
      UpdatedDate: nowIso
    };
    const relevantFieldDefinitions = Object.keys(payload).map((key) => ({
      InternalName: key,
      Title: key,
      StaticName: key,
      EntityPropertyName: key,
      TypeAsString: typeof payload[key]
    }));
    const buildComparison = (verified) => {
      const keys = Object.keys(payload);
      const comparison = {};
      keys.forEach((key) => {
        comparison[key] = {
          sent: payload[key],
          stored: verified ? verified[key] : undefined,
          matches: valuesEquivalent(payload[key], verified ? verified[key] : undefined)
        };
      });
      return comparison;
    };

    if (match.record) {
      await updateRoleItem(siteUrl, match.record.Id, payload);
      const verified = await getItemById(siteUrl, ROLES_LIST, match.record.Id).catch(() => null);
      const result = {
        action: "updated",
        itemId: match.record.Id,
        matchedBy: match.source,
        payload,
        relevantFieldDefinitions,
        verified,
        comparison: buildComparison(verified)
      };
      pushDebugLog("success", "roles-probe", "Completed role-record write probe", result);
      return result;
    }

    const createPayload = {
      ...payload,
      CreatedByEmail: currentUser.email || "",
      CreatedDate: nowIso
    };
    const newId = await createRoleItem(siteUrl, createPayload);
    const verified = newId ? await getItemById(siteUrl, ROLES_LIST, newId).catch(() => null) : null;
    const result = {
      action: "created",
      itemId: newId,
      matchedBy: "None",
      payload: createPayload,
      relevantFieldDefinitions,
      verified,
      comparison: buildComparison(verified)
    };
    pushDebugLog("success", "roles-probe", "Completed role-record create probe", result);
    return result;
  }

  async function getWritableRoleFieldMap(siteUrl) {
    try {
      const fields = await getExistingFieldDefinitions(siteUrl, ROLES_LIST);
      const writable = fields.filter(isWritableCustomField);
      if (writable.length) return writable;
    } catch (e) { /* list missing — fall through to the static schema map */ }
    return [
      { InternalName: "Title", Title: "Title", StaticName: "Title", EntityPropertyName: "Title", TypeAsString: "Text", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["Title", "title"] },
      { InternalName: "UserEmail", Title: "UserEmail", StaticName: "UserEmail", EntityPropertyName: "UserEmail", TypeAsString: "Text", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["UserEmail", "useremail"] },
      { InternalName: "UserDisplayName", Title: "UserDisplayName", StaticName: "UserDisplayName", EntityPropertyName: "UserDisplayName", TypeAsString: "Text", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["UserDisplayName", "userdisplayname"] },
      { InternalName: "SharePointUserId", Title: "SharePointUserId", StaticName: "SharePointUserId", EntityPropertyName: "SharePointUserId", TypeAsString: "Number", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["SharePointUserId", "sharepointuserid"] },
      { InternalName: "LoginName", Title: "LoginName", StaticName: "LoginName", EntityPropertyName: "LoginName", TypeAsString: "Text", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["LoginName", "loginname"] },
      { InternalName: "PrincipalType", Title: "PrincipalType", StaticName: "PrincipalType", EntityPropertyName: "PrincipalType", TypeAsString: "Number", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["PrincipalType", "principaltype"] },
      { InternalName: "IsSiteAdmin", Title: "IsSiteAdmin", StaticName: "IsSiteAdmin", EntityPropertyName: "IsSiteAdmin", TypeAsString: "Boolean", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["IsSiteAdmin", "issiteadmin"] },
      { InternalName: "Role", Title: "Role", StaticName: "Role", EntityPropertyName: "Role", TypeAsString: "Choice", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["Role", "role"] },
      { InternalName: "ThemeMode", Title: "ThemeMode", StaticName: "ThemeMode", EntityPropertyName: "ThemeMode", TypeAsString: "Choice", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["ThemeMode", "thememode"] },
      { InternalName: "NotificationPrefsJson", Title: "NotificationPrefsJson", StaticName: "NotificationPrefsJson", EntityPropertyName: "NotificationPrefsJson", TypeAsString: "Note", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["NotificationPrefsJson", "notificationprefsjson"] },
      { InternalName: "HideFromMeetings", Title: "HideFromMeetings", StaticName: "HideFromMeetings", EntityPropertyName: "HideFromMeetings", TypeAsString: "Boolean", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["HideFromMeetings", "hidefrommeetings"] },
      { InternalName: "IsActive", Title: "IsActive", StaticName: "IsActive", EntityPropertyName: "IsActive", TypeAsString: "Boolean", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["IsActive", "isactive"] },
      { InternalName: "Source", Title: "Source", StaticName: "Source", EntityPropertyName: "Source", TypeAsString: "Text", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["Source", "source"] },
      { InternalName: "CreatedByEmail", Title: "CreatedByEmail", StaticName: "CreatedByEmail", EntityPropertyName: "CreatedByEmail", TypeAsString: "Text", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["CreatedByEmail", "createdbyemail"] },
      { InternalName: "CreatedDate", Title: "CreatedDate", StaticName: "CreatedDate", EntityPropertyName: "CreatedDate", TypeAsString: "DateTime", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["CreatedDate", "createddate"] },
      { InternalName: "UpdatedByEmail", Title: "UpdatedByEmail", StaticName: "UpdatedByEmail", EntityPropertyName: "UpdatedByEmail", TypeAsString: "Text", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["UpdatedByEmail", "updatedbyemail"] },
      { InternalName: "UpdatedDate", Title: "UpdatedDate", StaticName: "UpdatedDate", EntityPropertyName: "UpdatedDate", TypeAsString: "DateTime", Hidden: false, ReadOnlyField: false, Sealed: false, aliases: ["UpdatedDate", "updateddate"] }
    ];
  }

  function isHiddenFromMeetings(record) {
    if (!record) return false;
    const value = record.HideFromMeetings;
    return value === true || value === 1 || String(value || "").trim().toLowerCase() === "true";
  }

  async function setUserMeetingHidden(siteUrl, recordId, hidden, currentUser) {
    const isAdmin = await isCurrentUserAdmin(siteUrl, currentUser);
    if (!isAdmin) throw new Error("You do not have permission to manage user visibility.");
    const nowIso = new Date().toISOString();
    await updateRoleItem(siteUrl, recordId, {
      HideFromMeetings: !!hidden,
      UpdatedByEmail: currentUser.email,
      UpdatedDate: nowIso
    });
    return { id: recordId, hidden: !!hidden };
  }

  async function assignUserRole(siteUrl, targetUser, role, currentUser) {
    const isAdmin = await isCurrentUserAdmin(siteUrl, currentUser);
    if (!isAdmin) throw new Error("You do not have permission to manage user roles.");
    const existing = await getActiveRoleRecords(siteUrl);
    const nowIso = new Date().toISOString();
    const mine = existing.find(r => r.UserEmail && targetUser.email && r.UserEmail.toLowerCase() === targetUser.email.toLowerCase());

    if (role !== "Admin" && mine && mine.Role === "Admin") {
      const activeAdmins = existing.filter(r => r.Role === "Admin");
      if (activeAdmins.length <= 1) throw new Error("Can't remove the last active Admin. Promote someone else first.");
    }

    if (mine) {
      await updateRoleItem(siteUrl, mine.Id, { Role: role, UpdatedByEmail: currentUser.email, UpdatedDate: nowIso });
      return { id: mine.Id, created: false };
    }
    const newId = await createRoleItem(siteUrl, {
      Title: targetUser.displayName,
      UserEmail: targetUser.email,
      UserDisplayName: targetUser.displayName,
      SharePointUserId: targetUser.spUserId || 0,
      LoginName: targetUser.loginName || "",
      PrincipalType: targetUser.principalType || 1,
      Role: role,
      IsActive: true,
      Source: "Manual",
      CreatedByEmail: currentUser.email,
      CreatedDate: nowIso,
      UpdatedByEmail: currentUser.email,
      UpdatedDate: nowIso
    });
    return { id: newId, created: true };
  }

  /* Sets JobTitle/SpoAccess/PowerBiAccess on a user's PULSE App
     Roles record — the fields the Users page displays but, until this
     function existed, had no write path anywhere in SharePoint mode. Mirrors
     assignUserRole's find-or-create logic. profile keys are all optional;
     only the ones present get written. */
  async function updateUserProfile(siteUrl, targetUser, profile, currentUser) {
    const isAdmin = await isCurrentUserAdmin(siteUrl, currentUser);
    if (!isAdmin) throw new Error("You do not have permission to manage user profiles.");
    const existing = await getActiveRoleRecords(siteUrl);
    const nowIso = new Date().toISOString();
    const mine = existing.find(r => r.UserEmail && targetUser.email && r.UserEmail.toLowerCase() === targetUser.email.toLowerCase());
    const fields = {};
    if ("jobTitle" in profile) fields.JobTitle = profile.jobTitle || "";
    if ("spoAccess" in profile) fields.SpoAccess = profile.spoAccess || "";
    if ("powerBiAccess" in profile) fields.PowerBiAccess = profile.powerBiAccess || "";
    if ("themeMode" in profile) fields.ThemeMode = profile.themeMode || "";

    if (mine) {
      await updateRoleItem(siteUrl, mine.Id, { ...fields, UpdatedByEmail: currentUser.email, UpdatedDate: nowIso });
      return { id: mine.Id, created: false };
    }
    const newId = await createRoleItem(siteUrl, {
      Title: targetUser.displayName,
      UserEmail: targetUser.email,
      UserDisplayName: targetUser.displayName,
      SharePointUserId: targetUser.spUserId || 0,
      LoginName: targetUser.loginName || "",
      PrincipalType: targetUser.principalType || 1,
      Role: APP_CONFIG.defaultUserRole,
      ...fields,
      IsActive: true,
      Source: "Manual",
      CreatedByEmail: currentUser.email,
      CreatedDate: nowIso,
      UpdatedByEmail: currentUser.email,
      UpdatedDate: nowIso
    });
    return { id: newId, created: true };
  }

  async function saveCurrentUserThemeMode(siteUrl, currentUser, themeMode) {
    const normalizedThemeMode = String(themeMode || "").trim().toLowerCase() === "dark" ? "Dark" : "Light";
    const existing = await getRoleRecords(siteUrl);
    const nowIso = new Date().toISOString();
    const mine = findRoleRecordForUser(existing, currentUser).record;

    if (mine) {
      await updateRoleItem(siteUrl, mine.Id, {
        ThemeMode: normalizedThemeMode,
        UpdatedByEmail: currentUser.email || "",
        UpdatedDate: nowIso
      });
      return { id: mine.Id, created: false, themeMode: normalizedThemeMode };
    }

    const newId = await createRoleItem(siteUrl, {
      Title: currentUser.displayName || currentUser.email || "Current User",
      UserEmail: currentUser.email || "",
      UserDisplayName: currentUser.displayName || "",
      SharePointUserId: currentUser.spUserId || 0,
      LoginName: currentUser.loginName || "",
      PrincipalType: currentUser.principalType || 1,
      Role: APP_CONFIG.defaultUserRole,
      ThemeMode: normalizedThemeMode,
      IsActive: true,
      Source: "Manual",
      CreatedByEmail: currentUser.email || "",
      CreatedDate: nowIso,
      UpdatedByEmail: currentUser.email || "",
      UpdatedDate: nowIso
    });
    return { id: newId, created: true, themeMode: normalizedThemeMode };
  }

  async function saveCurrentUserNotificationPrefs(siteUrl, currentUser, prefs) {
    const prefsJson = JSON.stringify(prefs || {});
    const existing = await getRoleRecords(siteUrl);
    const nowIso = new Date().toISOString();
    const mine = findRoleRecordForUser(existing, currentUser).record;

    if (mine) {
      await updateRoleItem(siteUrl, mine.Id, {
        NotificationPrefsJson: prefsJson,
        UpdatedByEmail: currentUser.email || "",
        UpdatedDate: nowIso
      });
      return { id: mine.Id, created: false, prefs };
    }

    const newId = await createRoleItem(siteUrl, {
      Title: currentUser.displayName || currentUser.email || "Current User",
      UserEmail: currentUser.email || "",
      UserDisplayName: currentUser.displayName || "",
      SharePointUserId: currentUser.spUserId || 0,
      LoginName: currentUser.loginName || "",
      PrincipalType: currentUser.principalType || 1,
      Role: APP_CONFIG.defaultUserRole,
      NotificationPrefsJson: prefsJson,
      IsActive: true,
      Source: "Manual",
      CreatedByEmail: currentUser.email || "",
      CreatedDate: nowIso,
      UpdatedByEmail: currentUser.email || "",
      UpdatedDate: nowIso
    });
    return { id: newId, created: true, prefs };
  }

  async function deactivateUserRole(siteUrl, roleItemId, currentUser) {
    const isAdmin = await isCurrentUserAdmin(siteUrl, currentUser);
    if (!isAdmin) throw new Error("You do not have permission to manage user roles.");
    const existing = await getActiveRoleRecords(siteUrl);
    const target = existing.find(r => r.Id === roleItemId);
    if (target && target.Role === "Admin") {
      const activeAdmins = existing.filter(r => r.Role === "Admin");
      if (activeAdmins.length <= 1) throw new Error("Can't deactivate the last active Admin. Promote someone else first.");
    }
    await updateRoleItem(siteUrl, roleItemId, {
      IsActive: false,
      UpdatedByEmail: currentUser.email,
      UpdatedDate: new Date().toISOString()
    });
  }

  async function queueTeamsNotification(siteUrl, { to, subject, bodyHtml, emailBodyHtml, teamsText, area }) {
    const recipients = (Array.isArray(to) ? to : [to]).filter(Boolean);
    if (!recipients.length) return { queued: false, reason: "no-recipients" };
    await createItem(siteUrl, "PULSE Notifications", {
      Title: String(subject || "PULSE notification").slice(0, 255),
      ToEmails: recipients.join(";"),
      Subject: subject || "",
      BodyHtml: bodyHtml || "",
      EmailBodyHtml: emailBodyHtml || "",
      TeamsText: teamsText || "",
      NotificationArea: area || "General",
      DeliveryStatus: "Pending"
    });
    return { queued: true, count: recipients.length };
  }

  /* ---------- email notifications ----------
     Uses SharePoint's built-in SP.Utilities.Utility.SendEmail REST endpoint —
     no Graph, no Power Automate, no separate mail infrastructure. Some
     tenants disable outbound mail via this endpoint; failures are caught by
     callers (a missed notification shouldn't block the underlying save). */
  async function sendEmail(siteUrl, { to, subject, body }) {
    const recipients = (Array.isArray(to) ? to : [to]).filter(Boolean);
    if (!recipients.length) return { sent: false, reason: "no-recipients" };
    await spPostVerbose(siteUrl, "/_api/SP.Utilities.Utility.SendEmail", {
      properties: {
        __metadata: { type: "SP.Utilities.EmailProperties" },
        To: { results: recipients },
        Subject: subject || "",
        Body: body || ""
      }
    });
    return { sent: true };
  }

  /* ---------- bootstrap admin (setup-time only) ---------- */

  async function bootstrapAdminIfNeeded(siteUrl, currentUser) {
    if (!currentUser) return { bootstrapped: false, skipped: true };
    const existing = await getRoleRecords(siteUrl);
    const activeAdmins = existing.filter((r) => isRoleRecordActive(r) && r.Role === "Admin");
    if (activeAdmins.length) return { bootstrapped: false, skipped: true, reason: "admin-exists" };
    const match = findRoleRecordForUser(existing, currentUser);
    const nowIso = new Date().toISOString();
    if (match.record) {
      await updateRoleItem(siteUrl, match.record.Id, {
        Role: "Admin",
        IsActive: true,
        Source: match.record.Source || "Bootstrap Admin",
        UpdatedByEmail: currentUser.email || "",
        UpdatedDate: nowIso
      });
      return { bootstrapped: true, created: false, id: match.record.Id };
    }
    const newId = await createRoleItem(siteUrl, {
      Title: currentUser.displayName || currentUser.email || "Admin",
      UserEmail: currentUser.email || "",
      UserDisplayName: currentUser.displayName || "",
      SharePointUserId: currentUser.spUserId || 0,
      LoginName: currentUser.loginName || "",
      PrincipalType: currentUser.principalType || 1,
      IsSiteAdmin: !!currentUser.isSiteAdmin,
      Role: "Admin",
      IsActive: true,
      Source: "Bootstrap Admin",
      CreatedByEmail: currentUser.email || "",
      CreatedDate: nowIso,
      UpdatedByEmail: currentUser.email || "",
      UpdatedDate: nowIso
    });
    return { bootstrapped: true, created: true, id: newId };
  }

  return {
    detectSharePointSite,
    getSiteUrl,
    getDiagnostics,
    formatSpError,
    getCurrentUser,
    getCurrentUserFromRest,
    getCurrentUserFromPageContext,
    getRequestDigest,
    spGet, spPost, spPostVerbose, spMerge, spDelete,
    getSiteServerRelativeUrl,
    getDataLibraryTitle,
    getDataFolderName,
    getDataLibraryServerRelativeUrl,
    getDataFolderServerRelativeUrl,
    getDataFileServerRelativeUrl,
    ensureDocumentLibrary,
    ensureFolderPath,
    folderExists,
    fileExists,
    readTextFile,
    writeTextFile,
    writeBinaryFile,
    uploadProjectDocument,
    uploadIssueScreenshot,
    deleteIssueScreenshot,
    sanitizeSharePointSegment,
    resolvePulseDocumentsLibraryUrl,
    ensureProjectDocumentsRoot,
    listPulseDocumentsFolder,
    createPulseDocumentsSubfolder,
    uploadPulseDocumentsFile,
    deletePulseDocumentsFile,
    deletePulseDocumentsFolder,
    moveFileToSubfolder,
    readJsonFile,
    writeJsonFile,
    ensureJsonFile,
    ensureRoleStoreFile,
    ensureDbStoreFile,
    ensureAppDataStore,
    loadAppDataFile,
    loadAppDataFileWithMeta,
    saveAppDataFile,
    listExists, getListInfo, createList, renameList, deleteList, fieldExists, getExistingFieldDefinitions, getExistingFieldNames, createFieldAsXml, createFieldByType, ensureField, ensureListSchema, ensureDefaultViewFields, setFieldRequired,
    buildFieldXml, buildFieldRestPayload,
    getListItems,
    createListItem,
    updateListItem,
    deleteListItem,
    getItems, createItem, updateItem, deleteItem, getItemById,
    getSharePointSiteUsers,
    getAssociatedMembershipUsers,
    getSharePointPermissionUsers,
    getPaginatedSiteUsers,
    loadPeopleDirectory,
    savePeopleDirectory,
    pullSharePointPeopleDirectory,
    loadSiteMemberDirectory,
    searchTenantPeople,
    searchTenantPeopleDirectory,
    getLastTenantSearchAttempts,
    uploadProjectCoverImage,
    sendEmail,
    queueTeamsNotification,
    getRoleRecords,
    normalizeRoleRecord,
    getWritableRoleFieldMap,
    getActiveRoleRecords,
    findRoleRecordForUser,
    getCurrentUserRole, isCurrentUserAdmin, syncSiteUsersToAppRoles, assignUserRole, assignUserToPulse, unassignUserFromPulse, updateUserProfile, setUserMeetingHidden, deactivateUserRole, saveCurrentUserThemeMode, saveCurrentUserNotificationPrefs,
    probeRoleRecordWrite,
    bootstrapAdminIfNeeded,
    isRoleRecordActive,
    isUserPulseAssociated,
    isHiddenFromMeetings,
    ensurePeopleDirectoryFile,
    ROLES_LIST
  };
})();
