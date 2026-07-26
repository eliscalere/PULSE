/* Shared file-opening policy for files stored in SharePoint.
   Word documents open in the installed app. Every other file opens its
   SharePoint browser URL. */
(function () {
  window.AEWTTR = window.AEWTTR || {};

  function cleanSharePointFileUrl(fileUrl) {
    if (!fileUrl) return "";
    try {
      const url = new URL(fileUrl, window.location.origin);
      url.search = "";
      url.hash = "";
      return url.toString();
    } catch (e) {
      return String(fileUrl).split("#")[0].split("?")[0];
    }
  }

  function officeSchemeForFileType(fileType) {
    const kind = String(fileType || "").toLowerCase();
    if (kind === "docx" || kind === "doc") return "ms-word";
    if (kind === "pptx" || kind === "ppt") return "ms-powerpoint";
    if (kind === "xlsx" || kind === "xls") return "ms-excel";
    return "";
  }

  function officeSchemeFromName(fileName, mimeType) {
    const name = String(fileName || "").toLowerCase();
    const type = String(mimeType || "").toLowerCase();
    const ext = (name.split(".").pop() || "").replace(/[^a-z0-9]/g, "");
    if (["doc", "docx"].includes(ext) || type.includes("wordprocessingml") || type.includes("msword") || (type.includes("word") && !type.includes("powerpoint"))) {
      return "ms-word";
    }
    if (["ppt", "pptx"].includes(ext) || type.includes("presentationml") || type.includes("powerpoint")) {
      return "ms-powerpoint";
    }
    if (["xls", "xlsx"].includes(ext) || type.includes("spreadsheetml") || type.includes("ms-excel") || type === "application/vnd.ms-excel") {
      return "ms-excel";
    }
    return "";
  }

  function fileNameFromUrl(fileUrl) {
    try {
      const url = new URL(fileUrl, window.location.origin);
      return decodeURIComponent((url.pathname || "").split("/").pop() || "");
    } catch (e) {
      return String(fileUrl || "").split("#")[0].split("?")[0].split("/").pop() || "";
    }
  }

  function isWordDocument(fileName, mimeType, fileUrl) {
    return officeSchemeFromName(fileName || fileNameFromUrl(fileUrl), mimeType) === "ms-word";
  }

  function sharePointBrowserUrl(fileUrl) {
    const clean = cleanSharePointFileUrl(fileUrl);
    if (!clean) return "";
    return `${clean}${clean.includes("?") ? "&" : "?"}web=1`;
  }

  /* Build the ms-word: protocol URI.
     Use decodeURI so paths with %20 become literal spaces — this matches how
     SharePoint's own "Open in app" button constructs the ms-word: URI.
     Word on Windows resolves the decoded URL through its own HTTP client and
     handles spaces correctly; the percent-encoded form causes "can't connect"
     errors on some DoD / Flank Speed configurations. */
  function officeDesktopUri(fileUrl, fileTypeOrScheme) {
    const clean = cleanSharePointFileUrl(fileUrl);
    if (!clean || !/^https?:\/\//i.test(clean)) return "";
    let scheme = String(fileTypeOrScheme || "").toLowerCase();
    scheme = officeSchemeForFileType(scheme) || scheme;
    if (!["ms-word", "ms-powerpoint", "ms-excel"].includes(scheme)) return "";
    try {
      return `${scheme}:ofe|u|${decodeURI(clean)}`;
    } catch (_) {
      return `${scheme}:ofe|u|${clean}`;
    }
  }

  /* Open the SharePoint file in the installed Office app using the ms-word:
     protocol — the same mechanism SharePoint's own "Open in app" button uses.
     This opens the live SharePoint-connected copy so changes save back to SP.
     Uses an anchor-click (not iframe) because browsers block protocol-handler
     activation from iframe src navigation in many environments. Falls back to
     the browser viewer if the protocol handler isn't registered. */
  function openSharePointFileInDesktopApp(fileUrl, fileTypeOrScheme) {
    const uri = officeDesktopUri(fileUrl, fileTypeOrScheme);
    if (!uri) return false;
    try {
      const a = document.createElement("a");
      a.href = uri;
      a.style.cssText = "display:none;position:absolute;left:-9999px;";
      document.body.appendChild(a);
      a.click();
      setTimeout(function () { try { a.remove(); } catch (_) {} }, 500);
    } catch (e) {
      const clean = cleanSharePointFileUrl(fileUrl);
      window.open(clean + (clean.includes("?") ? "&" : "?") + "web=1", "_blank");
    }
    return true;
  }

  /* After an async SPO upload, give the write a moment to be visible, then
     open. Returns a Promise that resolves when the open attempt runs. */
  function openSharePointFileInDesktopAppSoon(fileUrl, fileTypeOrScheme, delayMs) {
    const wait = typeof delayMs === "number" ? delayMs : 450;
    return new Promise((resolve) => {
      setTimeout(() => {
        resolve(openSharePointFileInDesktopApp(fileUrl, fileTypeOrScheme));
      }, wait);
    });
  }

  function reserveSharePointFileWindow(fileName, mimeType) {
    if (isWordDocument(fileName, mimeType, "")) return null;
    const popup = window.open("", "_blank");
    if (!popup) return null;
    try {
      popup.opener = null;
      popup.document.title = "Opening file in SharePoint";
      popup.document.body.textContent = "Opening file in SharePoint…";
    } catch (e) {
      /* The reserved window is still usable if its document is inaccessible. */
    }
    return popup;
  }

  function closeReservedSharePointFileWindow(popup) {
    if (!popup || popup.closed) return;
    try { popup.close(); } catch (e) { /* no-op */ }
  }

  function openSharePointFileByPolicy(fileUrl, fileName, mimeType, options) {
    const opts = options || {};
    const scheme = officeSchemeFromName(fileName || fileNameFromUrl(fileUrl), mimeType);
    if (scheme) {
      closeReservedSharePointFileWindow(opts.popup);
      return openSharePointFileInDesktopApp(fileUrl, scheme);
    }
    const href = sharePointBrowserUrl(fileUrl);
    if (!href) {
      closeReservedSharePointFileWindow(opts.popup);
      return false;
    }
    if (opts.popup && !opts.popup.closed) {
      opts.popup.location.replace(href);
      return true;
    }
    const opened = window.open(href, "_blank");
    if (opened) {
      try { opened.opener = null; } catch (e) { /* no-op */ }
    }
    return true;
  }

  window.AEWTTR.OfficeDesktop = {
    cleanSharePointFileUrl,
    officeSchemeForFileType,
    officeSchemeFromName,
    isWordDocument,
    sharePointBrowserUrl,
    officeDesktopUri,
    openSharePointFileInDesktopApp,
    openSharePointFileInDesktopAppSoon,
    reserveSharePointFileWindow,
    closeReservedSharePointFileWindow,
    openSharePointFileByPolicy
  };

  /* Flat globals for pages that already call these by name (doc review). */
  window.cleanSharePointFileUrl = cleanSharePointFileUrl;
  window.officeDesktopUri = officeDesktopUri;
  window.openSharePointFileInDesktopApp = openSharePointFileInDesktopApp;
  window.openSharePointFileInDesktopAppSoon = openSharePointFileInDesktopAppSoon;
  window.openSharePointFileByPolicy = openSharePointFileByPolicy;
})();
