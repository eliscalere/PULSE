/* PULSE issue reporting — captures useful repro context without asking the
   reporter to copy browser diagnostics by hand. The report itself remains a
   normal SharePoint list item, so it works in both SharePoint and local mode. */

const ISSUE_CLIENT_LOG_LIMIT = 160;
const _issueClientLogs = [];

function issueTrim(value, maxLength) {
  const text = String(value == null ? "" : value);
  const limit = maxLength || 700;
  return text.length > limit ? `${text.slice(0, Math.max(0, limit - 1))}…` : text;
}

function issueLogValue(value, seen) {
  if (value instanceof Error) return `${value.name || "Error"}: ${value.message || String(value)}`;
  if (typeof value === "string") return issueTrim(value);
  if (value == null || typeof value === "number" || typeof value === "boolean") return issueTrim(value);
  const visited = seen || new WeakSet();
  try {
    return issueTrim(JSON.stringify(value, (key, child) => {
      if (typeof child === "object" && child !== null) {
        if (visited.has(child)) return "[Circular]";
        visited.add(child);
      }
      return child instanceof Error ? `${child.name}: ${child.message}` : child;
    }));
  } catch (e) {
    return issueTrim(String(value));
  }
}

function recordIssueClientLog(level, args) {
  _issueClientLogs.push({
    ts: new Date().toISOString(),
    level: level || "info",
    message: issueTrim((args || []).map((value) => issueLogValue(value)).join(" "), 1100)
  });
  if (_issueClientLogs.length > ISSUE_CLIENT_LOG_LIMIT) _issueClientLogs.splice(0, _issueClientLogs.length - ISSUE_CLIENT_LOG_LIMIT);
}

(function captureClientConsole() {
  ["log", "info", "warn", "error"].forEach((level) => {
    const original = window.console && window.console[level];
    if (typeof original !== "function") return;
    window.console[level] = function (...args) {
      recordIssueClientLog(level, args);
      return original.apply(window.console, args);
    };
  });
  window.addEventListener("error", (event) => {
    recordIssueClientLog("error", [`${event.message || "Unhandled error"}${event.filename ? ` @ ${event.filename}:${event.lineno || 0}` : ""}`]);
  });
  window.addEventListener("unhandledrejection", (event) => {
    recordIssueClientLog("error", ["Unhandled promise rejection", event.reason]);
  });
}());

function appSettingsFallback() {
  return typeof appSettingsDefaults === "function" ? appSettingsDefaults() : { cuiMarkingEnabled: false };
}

function getAppSettings() {
  const db = window.AEWTTR && window.AEWTTR.db;
  if (!db) return appSettingsFallback();
  if (!db.appSettings || typeof db.appSettings !== "object") db.appSettings = appSettingsFallback();
  db.appSettings.cuiMarkingEnabled = !!db.appSettings.cuiMarkingEnabled;
  return db.appSettings;
}

function cuiMarkingBarHtml(position) {
  if (!getAppSettings().cuiMarkingEnabled) return "";
  return `<div class="cui-bar" data-cui-marking="${position || "top"}" role="note" aria-label="CUI marking active">CUI</div>`;
}

function applyCuiMarking() {
  const shell = document.querySelector(".aewttr-shell");
  if (!shell) return;
  const enabled = !!getAppSettings().cuiMarkingEnabled;
  const existing = Array.from(shell.querySelectorAll("[data-cui-marking]"));
  if (!enabled) {
    existing.forEach((bar) => bar.remove());
    return;
  }
  if (!shell.querySelector('[data-cui-marking="top"]')) shell.insertAdjacentHTML("afterbegin", cuiMarkingBarHtml("top"));
  if (!shell.querySelector('[data-cui-marking="bottom"]')) shell.insertAdjacentHTML("beforeend", cuiMarkingBarHtml("bottom"));
}

function issueCurrentRoute() {
  try {
    const route = typeof currentRoute === "function" ? currentRoute() : null;
    if (route && route.app) return [route.app].concat(route.parts || []).filter(Boolean).join("/");
  } catch (e) { /* retain URL fallback */ }
  return "";
}

function issueCurrentPage() {
  const title = document.querySelector("#page-title");
  return {
    title: issueTrim(title && title.textContent || document.title || "PULSE", 180),
    route: issueCurrentRoute(),
    url: window.location.href
  };
}

function issueErrorCodes(logs, lastSpError) {
  const candidates = [];
  if (lastSpError) candidates.push(String(lastSpError.status || ""), String(lastSpError.code || ""), String(lastSpError.friendly || lastSpError.message || ""));
  (logs || []).filter((entry) => entry.level === "error" || entry.level === "warn").forEach((entry) => candidates.push(entry.message || ""));
  const codes = new Set();
  candidates.join("\n").replace(/(?:http\s*|status(?:\s+code)?\s*|error\s+code\s*|\b)([45]\d{2})\b/gi, (whole, code) => {
    codes.add(`HTTP ${code}`);
    return whole;
  });
  return Array.from(codes).slice(0, 12);
}

function issueReportDiagnostics() {
  const app = window.AEWTTR || {};
  const logs = _issueClientLogs.slice(-80);
  const debugLogs = Array.isArray(app.debugLog) ? app.debugLog.slice(-30).map((entry) => ({
    ts: entry.ts || "",
    level: entry.level || "info",
    area: issueTrim(entry.area || "", 100),
    message: issueTrim(entry.message || "", 700)
  })) : [];
  const lastSpError = typeof getLastSpError === "function" ? getLastSpError() : null;
  const page = issueCurrentPage();
  return {
    page,
    logs,
    errorCodes: issueErrorCodes(logs, lastSpError),
    diagnostics: {
      capturedAt: new Date().toISOString(),
      mode: app.mode || "local",
      pageTitle: page.title,
      route: page.route,
      url: page.url,
      viewport: { width: window.innerWidth, height: window.innerHeight, devicePixelRatio: window.devicePixelRatio || 1 },
      browser: issueTrim(navigator.userAgent || "", 500),
      lastSharePointError: lastSpError ? {
        friendly: issueTrim(lastSpError.friendly || lastSpError.message || "", 1400),
        status: lastSpError.status || "",
        raw: issueTrim(lastSpError.raw || "", 1400),
        recordedAt: lastSpError.recordedAt || ""
      } : null,
      bootMessages: Array.isArray(app.bootMessages) ? app.bootMessages.slice(-40).map((entry) => issueTrim(entry, 500)) : [],
      sharePointDebug: debugLogs
    }
  };
}

function issueCanvasToPng(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new Error("The browser could not encode the screenshot."));
    }, "image/png");
  });
}

function issueBlobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("The captured screenshot could not be previewed."));
    reader.readAsDataURL(blob);
  });
}

function issueCaptureErrorMessage(error) {
  const name = String(error && error.name || "");
  const message = issueTrim(error && error.message || error, 300);
  if (name === "NotAllowedError") return "Capture was cancelled or this app host does not permit screen capture.";
  if (name === "NotFoundError") return "No shareable screen or tab was available to capture.";
  if (name === "AbortError") return "Capture was cancelled before a photo could be taken.";
  return message || "A native tab capture could not be created.";
}

async function captureIssueTabPhoto() {
  if (!navigator.mediaDevices || typeof navigator.mediaDevices.getDisplayMedia !== "function") {
    throw new Error("This browser or embedded app host does not support native tab capture.");
  }
  let stream;
  let video;
  try {
    // getDisplayMedia is the browser-provided capture path. It requires an explicit
    // click and lets the reporter choose This Tab; nothing is reconstructed from DOM.
    stream = await navigator.mediaDevices.getDisplayMedia({
      // Do not request a smaller target size: the browser can then provide the
      // selected tab at its native capture dimensions (including 4K displays).
      video: { displaySurface: "browser" },
      audio: false,
      preferCurrentTab: true,
      selfBrowserSurface: "include",
      surfaceSwitching: "exclude"
    });
    const track = stream.getVideoTracks()[0];
    if (!track) throw new Error("The selected tab did not provide a video stream.");
    const settings = typeof track.getSettings === "function" ? track.getSettings() : {};
    video = document.createElement("video");
    video.autoplay = true;
    video.muted = true;
    video.playsInline = true;
    video.srcObject = stream;
    await new Promise((resolve, reject) => {
      video.onloadedmetadata = resolve;
      video.onerror = () => reject(new Error("The selected tab could not be read."));
    });
    await video.play();
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const sourceWidth = video.videoWidth || settings.width || window.innerWidth;
    const sourceHeight = video.videoHeight || settings.height || window.innerHeight;
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(sourceWidth));
    canvas.height = Math.max(1, Math.round(sourceHeight));
    const context = canvas.getContext("2d");
    context.drawImage(video, 0, 0, canvas.width, canvas.height);
    const blob = await issueCanvasToPng(canvas);
    const dataUrl = await issueBlobToDataUrl(blob);
    return {
      dataUrl,
      blob,
      meta: {
        capturedAt: new Date().toISOString(),
        source: "Native browser tab capture · full-resolution PNG",
        width: canvas.width,
        height: canvas.height,
        fileSize: blob.size,
        contentType: blob.type || "image/png"
      }
    };
  } finally {
    if (video) video.srcObject = null;
    if (stream) stream.getTracks().forEach((track) => track.stop());
  }
}

function safeIssueScreenshotDataUrl(value) {
  const screenshot = String(value || "");
  return /^data:image\/(?:jpeg|png|webp);base64,/i.test(screenshot) && screenshot.length <= 64 * 1024 * 1024 ? screenshot : "";
}

function safeIssueScreenshotFileUrl(value) {
  const fileUrl = String(value || "").trim();
  if (!/^(?:https?:\/\/|\/)/i.test(fileUrl)) return "";
  try {
    const resolved = new URL(fileUrl, window.location.href);
    if (!/^https?:$/.test(resolved.protocol)) return "";
    if (window.location.origin && window.location.origin !== "null" && resolved.origin !== window.location.origin) return "";
    return resolved.href;
  } catch (error) {
    return "";
  }
}

function safeIssueScreenshotSource(issue) {
  if (!issue) return "";
  return safeIssueScreenshotFileUrl(issue.screenshotFileUrl) || safeIssueScreenshotDataUrl(issue.screenshotDataUrl);
}

function formatIssueReportDate(value) {
  if (!value) return "—";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return date.toLocaleString([], { year: "numeric", month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
}

function issueCode() {
  return `ISS-${new Date().toISOString().slice(0, 10).replace(/-/g, "")}-${Math.random().toString(36).slice(2, 7).toUpperCase()}`;
}

function issueSaveMessage(error) {
  const detail = issueTrim(error && (error.friendly || error.message) || error || "Unable to save the report.", 300);
  if (/PULSE Issues|does not exist|not found/i.test(detail)) return "The Issues list is not ready yet. Ask an admin to run Admin → SharePoint Setup, then try again.";
  return detail;
}

function openIssueReportModal() {
  const page = issueCurrentPage();
  const modal = openModal(`
    <form class="issue-report-form" id="issue-report-form">
      <div class="aewttr-modal-head issue-report-head">
        <div>
          <p class="issue-report-kicker"><i class="bx bx-message-error"></i> PULSE Support</p>
          <h3>Report an issue</h3>
          <p>Describe the problem and we'll attach your page context automatically.</p>
        </div>
        <button class="aewttr-modal-close" type="button" id="issue-report-close" aria-label="Close">&times;</button>
      </div>
      <div class="aewttr-modal-body issue-report-body">

        <div class="issue-report-context-strip">
          <span><i class="bx bx-file-blank"></i> <strong>Page:</strong> <span id="issue-report-page">${escapeHtml(page.title)}</span></span>
          <span class="aewttr-mono" style="font-size:11px;color:var(--aewttr-muted);" id="issue-report-route">${escapeHtml(page.route || "dashboard")}</span>
        </div>

        <div class="form-row">
          <label for="issue-type">Issue type</label>
          <select class="select-aewttr" id="issue-type" required>
            <option value="Bug / error">Bug or error</option>
            <option value="Data / saving">Data or saving</option>
            <option value="Access / permissions">Access or permissions</option>
            <option value="Display / usability">Display or usability</option>
            <option value="Feature Request">Feature request</option>
            <option value="Other">Other</option>
          </select>
        </div>

        <div class="form-row">
          <label for="issue-description">What happened? <span style="color:var(--aewttr-red)">*</span></label>
          <textarea class="input-aewttr" id="issue-description" rows="4" required aria-required="true" placeholder="Tell us what you were trying to do and what went wrong."></textarea>
        </div>

        <div class="issue-report-form-grid">
          <div class="form-row">
            <label for="issue-expected">Expected result <span class="field-optional">Optional</span></label>
            <textarea class="input-aewttr" id="issue-expected" rows="2" placeholder="What should have happened…"></textarea>
          </div>
          <div class="form-row">
            <label for="issue-context">Additional context <span class="field-optional">Optional</span></label>
            <textarea class="input-aewttr" id="issue-context" rows="2" placeholder="Steps, timing, or impact…"></textarea>
          </div>
        </div>

        <div class="issue-screenshot-section">
          <div class="issue-screenshot-toggle-row">
            <label class="issue-screenshot-toggle-label">
              <input type="checkbox" id="attach-screenshot-check">
              <span>Attach a screenshot</span>
            </label>
            <span class="issue-screenshot-state" id="issue-screenshot-state"></span>
          </div>
          <div id="issue-wizard-capture-container" hidden>
            <div class="issue-wizard-capture-copy">
              <i class="bx bx-screenshot"></i>
              <div>
                <strong>Capture this tab</strong>
                <span>The form briefly hides so the screenshot shows the page behind it. Choose <b>This Tab</b> when prompted.</span>
              </div>
            </div>
            <button class="btn-aewttr" type="button" id="issue-capture-tab" style="margin-top:10px;">
              <i class="bx bx-camera"></i> Capture screenshot
            </button>
          </div>
          <div class="issue-screenshot-preview issue-wizard-capture-preview" id="issue-screenshot-preview" hidden></div>
        </div>

        <div class="issue-report-diagnostics">
          <i class="bx bx-shield-quarter"></i>
          <span id="issue-diagnostic-summary">Collecting recent logs and error codes…</span>
        </div>

        <div class="issue-report-save-status" id="issue-report-save-status" aria-live="polite"></div>
      </div>
      <div class="aewttr-modal-foot issue-wizard-footer">
        <button class="btn-aewttr-ghost" type="button" id="issue-report-cancel">Cancel</button>
        <div class="issue-wizard-footer-actions">
          <button class="btn-aewttr" type="button" id="issue-report-submit"><i class="bx bx-send"></i> Submit report</button>
        </div>
      </div>
    </form>
  `, { wide: true, className: "issue-report-modal" });


  let snapshot = { dataUrl: "", meta: { status: "not-captured" } };
  let pendingIssue = null;
  const snapshotState = document.querySelector("#issue-screenshot-state", modal);
  const snapshotPreview = document.querySelector("#issue-screenshot-preview", modal);
  const diagnosticSummary = document.querySelector("#issue-diagnostic-summary", modal);
  const saveStatus = document.querySelector("#issue-report-save-status", modal);
  const submit = document.querySelector("#issue-report-submit", modal);
  const initialDiagnostics = issueReportDiagnostics();
  if (diagnosticSummary) diagnosticSummary.textContent = `${initialDiagnostics.errorCodes.length} recent error code${initialDiagnostics.errorCodes.length === 1 ? "" : "s"} and ${initialDiagnostics.logs.length} browser log entries will be attached.`;

  const dismiss = () => closeModal();
  document.querySelector("#issue-report-close", modal).addEventListener("click", dismiss);
  document.querySelector("#issue-report-cancel", modal).addEventListener("click", dismiss);

  const capture = document.querySelector("#issue-capture-tab", modal);
  const captureBackdrop = modal.closest(".aewttr-modal-backdrop");
  let capturing = false;
  capture.addEventListener("click", async () => {
    if (capture.disabled) return;
    capturing = true;
    capture.disabled = true;
    capture.innerHTML = `<i class="bx bx-loader-alt bx-spin"></i> Capturing…`;
    if (snapshotState) snapshotState.innerHTML = `<i class="bx bx-loader-alt bx-spin"></i> Waiting for browser capture…`;
    try {
      // Hide the form before the browser supplies the first stream frame so
      // the attachment documents the reported page, not the report dialog.
      if (captureBackdrop) {
        captureBackdrop.classList.add("issue-capture-pending");
        void captureBackdrop.offsetWidth;
      }
      snapshot = await captureIssueTabPhoto();
      if (snapshotState) snapshotState.innerHTML = `<i class="bx bx-check-circle"></i> Full-resolution tab photo attached`;
      if (snapshotPreview) {
        snapshotPreview.hidden = false;
        snapshotPreview.innerHTML = `<img src="${escapeHtml(snapshot.dataUrl)}" alt="Photo captured from the reported browser tab"><span><i class="bx bx-check"></i> Screenshot attached</span>`;
      }
    } catch (error) {
      snapshot = { dataUrl: "", meta: { status: "unavailable", reason: issueCaptureErrorMessage(error) } };
      if (snapshotState) snapshotState.innerHTML = `<i class="bx bx-info-circle"></i> ${escapeHtml(snapshot.meta.reason)}`;
    } finally {
      capturing = false;
      if (captureBackdrop) captureBackdrop.classList.remove("issue-capture-pending");
      capture.disabled = false;
      capture.innerHTML = `<i class="bx bx-camera"></i> ${snapshot.dataUrl ? "Retake screenshot" : "Capture screenshot"}`;
    }
  });

  const captureContainer = document.querySelector("#issue-wizard-capture-container", modal);
  const attachCheck = document.querySelector("#attach-screenshot-check", modal);
  if (attachCheck) {
    attachCheck.addEventListener("change", () => {
      if (captureContainer) captureContainer.hidden = !attachCheck.checked;
      if (!attachCheck.checked) {
        // Clear snapshot if user unchecks
        snapshot = { dataUrl: "", meta: { status: "not-captured" } };
        if (snapshotState) snapshotState.innerHTML = "";
        if (snapshotPreview) { snapshotPreview.hidden = true; snapshotPreview.innerHTML = ""; }
      }
    });
  }
  let submitting = false;
  const submitIssueReport = async (event) => {
    if (event) event.preventDefault();
    if (submitting) return;
    if (capturing) {
      if (saveStatus) saveStatus.innerHTML = `<i class="bx bx-info-circle"></i> Finish or cancel the tab capture before submitting.`;
      return;
    }
    const description = document.querySelector("#issue-description", modal).value.trim();
    if (!description) {
      const input = document.querySelector("#issue-description", modal);
      if (saveStatus) saveStatus.innerHTML = `<i class="bx bx-error-circle"></i> Tell us what happened before submitting.`;
      if (input) input.focus();
      toast("Add a short description before submitting the issue.", "error");
      return;
    }
    submitting = true;
    submit.disabled = true;
    submit.innerHTML = `<i class="bx bx-loader-alt bx-spin"></i> Submitting…`;
    if (saveStatus) saveStatus.textContent = "Saving your report…";
    const diagnostics = issueReportDiagnostics();
    const db = window.AEWTTR && window.AEWTTR.db;
    const wasNew = !pendingIssue;
    pendingIssue = Object.assign(pendingIssue || {}, {
        id: pendingIssue && pendingIssue.id || issueCode(),
        title: `${document.querySelector("#issue-type", modal).value}: ${diagnostics.page.title}`,
        type: document.querySelector("#issue-type", modal).value,
        status: "New",
        reportedBy: db && db.user && db.user.name || "",
        reporterEmail: db && db.user && db.user.email || "",
        pageTitle: diagnostics.page.title,
        route: diagnostics.page.route,
        pageUrl: diagnostics.page.url,
        createdAt: pendingIssue && pendingIssue.createdAt || new Date().toISOString(),
        description,
        expectedBehavior: document.querySelector("#issue-expected", modal).value.trim(),
        additionalContext: document.querySelector("#issue-context", modal).value.trim(),
        errorCodes: diagnostics.errorCodes,
        logs: diagnostics.logs,
        diagnostics: diagnostics.diagnostics,
        screenshotDataUrl: snapshot.upload ? "" : snapshot.dataUrl,
        screenshotFileUrl: snapshot.upload && snapshot.upload.fileUrl || "",
        screenshotServerRelativeUrl: snapshot.upload && snapshot.upload.serverRelativeUrl || "",
        screenshotFileName: snapshot.upload && snapshot.upload.fileName || "",
        screenshotMeta: Object.assign({}, snapshot.meta, snapshot.upload ? {
          storage: "PULSE App Data / AEWTTR-PULSE / Issue Screenshots",
          fileSize: snapshot.upload.size || snapshot.meta.fileSize || 0,
          contentType: snapshot.upload.contentType || snapshot.meta.contentType || "image/png"
        } : {}),
        resolutionNote: pendingIssue && pendingIssue.resolutionNote || ""
      });
    if (wasNew && db) {
      if (!Array.isArray(db.issues)) db.issues = [];
      db.issues.unshift(pendingIssue);
    }
    try {
      const app = window.AEWTTR || {};
      if (app.mode === "sharepoint" && snapshot.blob && !snapshot.upload) {
        if (!app.siteUrl || typeof sharePointAdapter === "undefined" || typeof sharePointAdapter.uploadIssueScreenshot !== "function") {
          throw new Error("Full-resolution screenshot storage is unavailable. Reload the current package and try again.");
        }
        if (saveStatus) saveStatus.textContent = "Uploading the full-resolution screenshot to PULSE App Data…";
        snapshot.upload = await sharePointAdapter.uploadIssueScreenshot(app.siteUrl, pendingIssue.id, snapshot.blob);
        pendingIssue.screenshotDataUrl = "";
        pendingIssue.screenshotFileUrl = snapshot.upload.fileUrl;
        pendingIssue.screenshotServerRelativeUrl = snapshot.upload.serverRelativeUrl;
        pendingIssue.screenshotFileName = snapshot.upload.fileName;
        pendingIssue.screenshotMeta = Object.assign({}, snapshot.meta, {
          storage: "PULSE App Data / AEWTTR-PULSE / Issue Screenshots",
          fileSize: snapshot.upload.size || snapshot.meta.fileSize || 0,
          contentType: snapshot.upload.contentType || snapshot.meta.contentType || "image/png"
        });
        if (saveStatus) saveStatus.textContent = "Screenshot uploaded. Saving your report…";
      }
      await Repo.save("issue", pendingIssue, { immediate: true });
      if (typeof logUserAction === "function") logUserAction({ action: "Submit", area: "System", summary: `Submitted issue report ${pendingIssue.id}`, recordId: pendingIssue.id });
      closeModal();
      toast("Issue reported. Thank you — the support details are attached.", "success");
    } catch (error) {
      const message = issueSaveMessage(error);
      if (wasNew && db) {
        const pendingIndex = db.issues.indexOf(pendingIssue);
        if (pendingIndex >= 0) db.issues.splice(pendingIndex, 1);
        pendingIssue = null;
      }
      if (saveStatus) saveStatus.innerHTML = `<i class="bx bx-error-circle"></i> ${escapeHtml(message)}`;
      toast(`Issue not submitted: ${message}`, "error");
      submit.disabled = false;
      submit.innerHTML = `<i class="bx bx-send"></i> Submit issue`;
      submitting = false;
    }
  };
  // Use a direct click handler: some Forge/Firepit hosts suppress native form
  // submission while the same button click is still delivered to the iframe.
  submit.addEventListener("click", submitIssueReport);
  document.querySelector("#issue-report-form", modal).addEventListener("submit", (event) => {
    event.preventDefault();
    submitIssueReport(event);
  });
}

function issueStatusClass(status) {
  return `issue-status--${String(status || "New").toLowerCase().replace(/[^a-z0-9]+/g, "-")}`;
}
