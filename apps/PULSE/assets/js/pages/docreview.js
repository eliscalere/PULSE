async function uploadDocReviewFile(projectCode, docTitle, file) {
  if (!file) return { fileUrl: "", fileName: "" };
  if (!isSharePointMode()) {
    toast("File upload requires SharePoint mode — recorded the file name only.", "error");
    return { fileUrl: "", fileName: file.name };
  }
  const db = window.AEWTTR.db;
  const project = projectCode ? db.projects.find((p) => p.id === projectCode) : null;
  const label = project
    ? (typeof projectDocumentsLabel === "function" ? projectDocumentsLabel(project) : `${project.id} - ${project.name}`)
    : "General";
  const result = await sharePointAdapter.uploadProjectDocument(currentSiteUrl(), label, docTitle || "Untitled Document", file);
  return { fileUrl: result.fileUrl, fileName: result.fileName, serverRelativeUrl: result.serverRelativeUrl || "" };
}

/* Eight statuses, replacing the old 3-bucket "Not Started / In Review /
   Final / Signed" board. "Review Complete" is deliberately its own state,
   distinct from "Signed" — a document that's finished review but still
   needs signatures (or was never signed at all) is never called "Final".
   Grouped into 4 board views so the toolbar doesn't have to show 8 kanban
   columns at once; the underlying status is still one of these 8 exact
   values everywhere else (cards, filters, notifications). */
const DOC_REVIEW_STATUSES = [
  { key: "Not Started", hint: "No reviewer has looked at this yet." },
  { key: "In Review", hint: "At least one reviewer has responded; none have requested changes." },
  { key: "Changes Requested", hint: "A reviewer asked for changes on the current revision." },
  { key: "Review Complete", hint: "Every reviewer approved. No signature is required for this document." },
  { key: "Awaiting Final Pack", hint: "Every reviewer approved. Upload the final packed PDF to start sequential signing." },
  { key: "Signing in Progress", hint: "Final pack uploaded; signatures are being collected one signer at a time." },
  { key: "Signed", hint: "Every required signer has signed." },
  { key: "Archived", hint: "Archived — kept for record, not part of active review." }
];
const DOC_REVIEW_BOARD_VIEWS = {
  "Active": ["Not Started", "In Review", "Changes Requested"],
  "Complete": ["Review Complete", "Signed"],
  "Archived": ["Archived"],
  "All": ["Not Started", "In Review", "Changes Requested", "Review Complete", "Signed", "Archived"]
};
const DOC_REVIEW_SIGN_BOARD_VIEWS = {
  "Needs Signature": ["Awaiting Final Pack", "Signing in Progress"],
  "Queued": ["Not Started", "In Review", "Changes Requested"],
  "Signed": ["Signed"],
  "All": ["Not Started", "In Review", "Changes Requested", "Awaiting Final Pack", "Signing in Progress", "Signed"]
};
const DOC_REVIEW_TOOL_VIEWS = ["Review", "Sign"];
const INLINE_PREVIEW_EXTENSIONS = ["txt", "md", "markdown", "csv", "json", "xml", "html", "htm", "js", "ts", "css", "log"];
const DOC_REVIEW_REMINDER_INTERVAL_MS = 24 * 60 * 60 * 1000;
let _docReviewArchiveSweepStarted = false;
let _docReviewReminderSweepRunning = false;

function docFileExtension(fileName) {
  const match = String(fileName || "").toLowerCase().match(/\.([a-z0-9]+)$/);
  return match ? match[1] : "";
}

function docFileType(fileName, mimeType) {
  const ext = docFileExtension(fileName);
  const type = String(mimeType || "").toLowerCase();
  if (ext === "pdf" || type.includes("pdf")) return "pdf";
  if (["doc", "docx"].includes(ext) || type.includes("wordprocessingml") || type.includes("msword") || (type.includes("word") && !type.includes("powerpoint"))) return "docx";
  if (["ppt", "pptx"].includes(ext) || type.includes("presentationml") || type.includes("ms-powerpoint") || type.includes("powerpoint")) return "pptx";
  if (INLINE_PREVIEW_EXTENSIONS.includes(ext) || type.startsWith("text/")) return "text";
  return ext || "file";
}

/* Real SPO subfolder (next to the live review file) where prior revision
   binaries are moved on each new upload — not a UI metaphor in the modal. */
const DOC_REVIEW_PREVIOUS_REVISIONS_FOLDER = "Previous revisions";
const DOC_REVIEW_UPLOAD_ACCEPT = ".pdf,.doc,.docx,.ppt,.pptx,.txt,.md,.csv,.json,.xml,.html";

function todayIsoDate() {
  return new Date().toISOString().slice(0, 10);
}

function isoDaysBetween(startIso, endIso) {
  if (!startIso || !endIso) return 0;
  const start = new Date(`${startIso}T00:00:00Z`);
  const end = new Date(`${endIso}T00:00:00Z`);
  return Math.floor((end - start) / 86400000);
}

function currentUserIdentity() {
  const db = window.AEWTTR.db || {};
  const user = db.user || {};
  return {
    name: String(user.name || "").trim(),
    email: String(user.email || "").trim().toLowerCase(),
    isAdmin: !!user.isAdmin || currentAppRole() === "Admin"
  };
}

/* "Document Admins" and "upper-level users" (Manager) from the spec.
   "Document Admin" is a real assignable role (Users admin page, PULSE App
   Roles Role choice, isDocAdmin threaded through boot) — checked here by
   role string, same as isAdmin/Manager. */
function canManageDocReview() {
  const me = currentUserIdentity();
  if (me.isAdmin) return true;
  const role = typeof currentAppRole === "function" ? currentAppRole() : "";
  return role === "Document Admin" || role === "Manager";
}

function normalizeReviewer(record) {
  const legacyDone = !!record.done;
  const decision = record.decision || (legacyDone ? "Approved" : "Pending");
  return {
    name: record.name || "Reviewer",
    email: record.email || "",
    decision: decision,
    reviewedAt: record.reviewedAt || record.doneDate || "",
    note: record.note || "",
    lastNotifiedAt: record.lastNotifiedAt || "",
    isSigner: !!record.isSigner,
    signedAt: record.signedAt || "",
    signOrder: Number(record.signOrder || record.order) || 0,
    signNotifiedAt: record.signNotifiedAt || record.notifiedAt || ""
  };
}

function normalizeSigner(record) {
  const signedAt = record && record.signedAt || "";
  const order = Number(record && (record.order != null ? record.order : record.signOrder)) || 0;
  return {
    name: (record && record.name) || "Signer",
    email: (record && record.email) || "",
    order: order,
    signedAt: signedAt,
    status: signedAt ? "Signed" : "Pending",
    notifiedAt: (record && (record.notifiedAt || record.signNotifiedAt)) || ""
  };
}

function ensureSubmitterReviewer(reviewers, submitterName, submitterEmail) {
  const submitter = { name: submitterName || "Submitter", email: submitterEmail || "" };
  if (!(reviewers || []).some((reviewer) => samePersonByNameOrEmail(reviewer, submitter))) {
    reviewers.unshift({
      name: submitter.name,
      email: submitter.email,
      decision: "Pending",
      reviewedAt: "",
      note: ""
    });
  }
}

function appendDocActivity(doc, action, note) {
  if (!Array.isArray(doc.reviewActivity)) doc.reviewActivity = [];
  doc.reviewActivity.unshift({
    id: uid("DRA"),
    action: action,
    note: note || "",
    author: currentUserIdentity().name || "User",
    date: new Date().toISOString()
  });
  doc.reviewActivity = doc.reviewActivity.slice(0, 100);
}

function docSignatureRequired(doc) {
  return (doc.reviewers || []).some((reviewer) => reviewer.isSigner)
    || (Array.isArray(doc.signers) && doc.signers.length > 0);
}

function docSigners(doc) {
  return (doc.reviewers || []).filter((reviewer) => reviewer.isSigner);
}

function docSignersOrdered(doc) {
  return docSigners(doc).slice().sort((a, b) => {
    const orderA = Number(a.signOrder) || 0;
    const orderB = Number(b.signOrder) || 0;
    if (orderA && orderB && orderA !== orderB) return orderA - orderB;
    if (orderA && !orderB) return -1;
    if (!orderA && orderB) return 1;
    return String(a.name || "").localeCompare(String(b.name || ""));
  });
}

function docPendingSigners(doc) {
  return docSignersOrdered(doc).filter((reviewer) => !reviewer.signedAt);
}

function docCurrentSigner(doc) {
  if (!docHasFinalPack(doc)) return null;
  return docPendingSigners(doc)[0] || null;
}

function docIsFullyReviewed(doc) {
  const decisions = (doc.reviewers || []).map((reviewer) => reviewer.decision || "Pending");
  return decisions.length > 0 && decisions.every((decision) => decision === "Approved");
}

function docHasFinalPack(doc) {
  return !!(doc && doc.finalPackUrl);
}

function docSigningStageFlags(doc) {
  const signatureRequired = docSignatureRequired(doc);
  const fullyReviewed = docIsFullyReviewed(doc);
  const hasPack = docHasFinalPack(doc);
  const pending = docPendingSigners(doc);
  const signers = docSigners(doc);
  const awaitingFinalPack = !!(signatureRequired && fullyReviewed && !hasPack && !doc.isArchived);
  const finalPacked = !!(signatureRequired && fullyReviewed && hasPack);
  const signingInProgress = !!(finalPacked && pending.length > 0);
  const fullySigned = !!(signatureRequired && fullyReviewed && hasPack && signers.length > 0 && !pending.length);
  return { awaitingFinalPack, finalPacked, signingInProgress, fullySigned, signatureRequired, fullyReviewed };
}

function isPdfFile(file) {
  if (!file) return false;
  const name = String(file.name || "").toLowerCase();
  const type = String(file.type || "").toLowerCase();
  return name.endsWith(".pdf") || type === "application/pdf" || type.includes("pdf");
}

function shuffleArrayInPlace(items) {
  const list = items || [];
  for (let i = list.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    const tmp = list[i];
    list[i] = list[j];
    list[j] = tmp;
  }
  return list;
}

/* Prefer live reviewer.isSigner flags; if SignersJson has people that
   reviewers don't yet flag, fold them in so older rows and fresh uploads
   both land on one roster. Always rewrite doc.signers as the persisted
   snapshot for SignersJson. */
function mergeSignersIntoReviewers(doc) {
  const fromJson = Array.isArray(doc.signers) ? doc.signers.map(normalizeSigner) : [];
  fromJson.forEach((signer) => {
    let reviewer = (doc.reviewers || []).find((entry) => samePersonByNameOrEmail(entry, signer));
    if (!reviewer) {
      reviewer = normalizeReviewer({
        name: signer.name,
        email: signer.email,
        isSigner: true,
        signedAt: signer.signedAt || "",
        signOrder: signer.order || 0,
        signNotifiedAt: signer.notifiedAt || ""
      });
      doc.reviewers = (doc.reviewers || []).concat([reviewer]);
    } else {
      reviewer.isSigner = true;
      if (signer.signedAt && !reviewer.signedAt) reviewer.signedAt = signer.signedAt;
      if (signer.order && !reviewer.signOrder) reviewer.signOrder = signer.order;
      if (signer.notifiedAt && !reviewer.signNotifiedAt) reviewer.signNotifiedAt = signer.notifiedAt;
    }
  });
}

function syncDocSignersSnapshot(doc) {
  doc.signers = docSignersOrdered(doc).map((reviewer, index) => normalizeSigner({
    name: reviewer.name,
    email: reviewer.email,
    order: Number(reviewer.signOrder) || (index + 1),
    signedAt: reviewer.signedAt || "",
    notifiedAt: reviewer.signNotifiedAt || ""
  }));
  doc.signers.forEach((signer, index) => {
    const reviewer = (doc.reviewers || []).find((entry) => samePersonByNameOrEmail(entry, signer));
    if (reviewer) {
      if (!reviewer.signOrder) reviewer.signOrder = signer.order || (index + 1);
      signer.order = reviewer.signOrder;
    }
  });
  return doc;
}

/* Replace the required-signer set with `people`. Existing signatures are
   kept when the same person remains a signer; removed people lose isSigner
   and signedAt. People not already on the review team are added as
   reviewers so status derivation stays one roster. Order resets unless
   signing is already locked after final pack. */
function applySignersToDoc(doc, people) {
  const desired = (people || []).filter((person) => person && person.name).map(normalizeSigner);
  const locked = !!doc.signingSequenceLocked && docHasFinalPack(doc);
  (doc.reviewers || []).forEach((reviewer) => {
    const keep = desired.some((signer) => samePersonByNameOrEmail(signer, reviewer));
    if (!keep && reviewer.isSigner) {
      reviewer.isSigner = false;
      reviewer.signedAt = "";
      reviewer.signOrder = 0;
      reviewer.signNotifiedAt = "";
    }
  });
  desired.forEach((signer, index) => {
    let reviewer = (doc.reviewers || []).find((entry) => samePersonByNameOrEmail(entry, signer));
    if (!reviewer) {
      reviewer = normalizeReviewer({
        name: signer.name,
        email: signer.email,
        isSigner: true,
        signedAt: locked ? (signer.signedAt || "") : "",
        signOrder: locked ? (signer.order || 0) : 0,
        signNotifiedAt: locked ? (signer.notifiedAt || "") : ""
      });
      doc.reviewers = (doc.reviewers || []).concat([reviewer]);
    } else {
      reviewer.isSigner = true;
      if (locked && signer.signedAt) reviewer.signedAt = signer.signedAt;
      if (!locked) {
        reviewer.signedAt = signer.signedAt || reviewer.signedAt || "";
      }
    }
    if (!locked) {
      if (!reviewer.signOrder) reviewer.signOrder = 0;
    } else if (signer.order) {
      reviewer.signOrder = signer.order;
    }
  });
  if (!locked) {
    // Provisional roster order for display before shuffle-on-pack.
    desired.forEach((signer, index) => {
      const reviewer = (doc.reviewers || []).find((entry) => samePersonByNameOrEmail(entry, signer));
      if (reviewer && !reviewer.signOrder) reviewer.signOrder = index + 1;
    });
  }
  syncDocSignersSnapshot(doc);
  return doc;
}

function shuffleAndLockSigningOrder(doc) {
  const signers = docSigners(doc).slice();
  shuffleArrayInPlace(signers);
  signers.forEach((reviewer, index) => {
    reviewer.signOrder = index + 1;
    reviewer.signNotifiedAt = "";
    if (!reviewer.signedAt) reviewer.signedAt = "";
  });
  doc.signingSequenceLocked = true;
  doc.currentSignerNotifiedAt = "";
  syncDocSignersSnapshot(doc);
  return doc;
}

function clearFinalPackAndSigning(doc) {
  doc.finalPackUrl = "";
  doc.finalPackFileName = "";
  doc.finalPackUploadedAt = "";
  doc.finalPackUploadedBy = "";
  doc.signingSequenceLocked = false;
  doc.awaitingFinalPackNotifiedAt = "";
  doc.currentSignerNotifiedAt = "";
  (doc.reviewers || []).forEach((reviewer) => {
    if (reviewer.isSigner) {
      reviewer.signedAt = "";
      reviewer.signOrder = 0;
      reviewer.signNotifiedAt = "";
    }
  });
  syncDocSignersSnapshot(doc);
  return doc;
}

/* Central status derivation. Runs on every load/save so the persisted
   ReviewColumn value (still just a legacy Choice field until schema.js is
   free to extend) never has to be trusted on its own — this always
   recomputes the real status from ground truth: reviewer decisions,
   signer assignments/signatures, final pack presence, and the archive flag. */
function syncDocReviewState(doc) {
  doc.reviewers = (doc.reviewers || []).map(normalizeReviewer);
  mergeSignersIntoReviewers(doc);
  ensureSubmitterReviewer(doc.reviewers, doc.submitter, doc.submitterEmail || "");
  syncDocSignersSnapshot(doc);
  const decisions = doc.reviewers.map((reviewer) => reviewer.decision || "Pending");
  const allApproved = decisions.length > 0 && decisions.every((decision) => decision === "Approved");
  const anyRequested = decisions.some((decision) => decision === "Requested Changes");
  const anyReviewed = decisions.some((decision) => decision !== "Pending");
  const stages = docSigningStageFlags(doc);

  if (doc.isArchived) {
    doc._column = "Archived";
    if (!doc.archivedDate) doc.archivedDate = todayIsoDate();
    return doc;
  }

  if (!allApproved) {
    doc.fullyReviewedDate = "";
    doc.signedDate = "";
    doc._column = anyRequested ? "Changes Requested" : (anyReviewed ? "In Review" : "Not Started");
    return doc;
  }

  if (!doc.fullyReviewedDate) doc.fullyReviewedDate = todayIsoDate();
  if (!docSignatureRequired(doc)) {
    doc._column = "Review Complete";
    doc.signedDate = "";
    return doc;
  }

  if (stages.awaitingFinalPack) {
    doc.signedDate = "";
    doc._column = "Awaiting Final Pack";
    return doc;
  }

  const pendingSigners = docPendingSigners(doc);
  if (!pendingSigners.length) {
    doc._column = "Signed";
    if (!doc.signedDate) doc.signedDate = todayIsoDate();
  } else {
    doc.signedDate = "";
    doc._column = "Signing in Progress";
  }
  return doc;
}

function normalizeDocReview(doc) {
  const safeDoc = doc || {};
  const revisions = Array.isArray(safeDoc.revisions) ? safeDoc.revisions : [];
  if (!revisions.length && (safeDoc.fileUrl || safeDoc.fileName)) {
    revisions.push({
      id: safeDoc.activeRevisionId || uid("REV"),
      number: safeDoc.currentRevisionNumber || 1,
      fileUrl: safeDoc.fileUrl || "",
      fileName: safeDoc.fileName || "",
      fileType: safeDoc.fileType || docFileType(safeDoc.fileName || "", ""),
      uploadedBy: safeDoc.submitter || "",
      uploadedDate: safeDoc.date || "",
      source: "Initial Upload"
    });
  }
  revisions.sort((a, b) => (Number(b && b.number) || 0) - (Number(a && a.number) || 0));
  const activeRevision = revisions.find((revision) => revision.id === safeDoc.activeRevisionId) || revisions[0] || null;
  safeDoc.revisions = revisions;
  safeDoc.activeRevisionId = activeRevision ? activeRevision.id : "";
  safeDoc.currentRevisionNumber = Number(safeDoc.currentRevisionNumber) || (activeRevision ? activeRevision.number : 0);
  safeDoc.fileUrl = activeRevision ? activeRevision.fileUrl || "" : (safeDoc.fileUrl || "");
  safeDoc.fileName = activeRevision ? activeRevision.fileName || "" : (safeDoc.fileName || "");
  safeDoc.fileType = activeRevision ? activeRevision.fileType || docFileType(activeRevision.fileName || "", "") : (safeDoc.fileType || "");
  safeDoc.reviewActivity = Array.isArray(safeDoc.reviewActivity) ? safeDoc.reviewActivity : [];
  safeDoc.isArchived = !!safeDoc.isArchived;
  safeDoc.archivedDate = safeDoc.archivedDate || "";
  safeDoc.fullyReviewedDate = safeDoc.fullyReviewedDate || "";
  safeDoc.submitterEmail = safeDoc.submitterEmail || ((safeDoc.reviewers || []).find((reviewer) => String(reviewer.name || "").trim() === String(safeDoc.submitter || "").trim()) || {}).email || "";
  // Government-vs-contractor identification, contractor name, associated
  // portfolio/config end item, and applied reviewer-group label — dedicated
  // columns on "PULSE Document Review" (see sharepoint-schema.js/repo.js),
  // so these persist and resync like every other field on the record.
  safeDoc.docKind = safeDoc.docKind || "";
  safeDoc.contractorName = safeDoc.contractorName || "";
  safeDoc.portfolios = Array.isArray(safeDoc.portfolios) ? safeDoc.portfolios.map(String).filter(Boolean) : [];
  safeDoc.configEndItems = Array.isArray(safeDoc.configEndItems) ? safeDoc.configEndItems.map(String).filter(Boolean) : [];
  // Keep portfolioOrConfigItem as a legacy computed display string for search/backward compat.
  safeDoc.portfolioOrConfigItem = [...safeDoc.portfolios, ...safeDoc.configEndItems].join("; ") || safeDoc.portfolioOrConfigItem || "";
  safeDoc.reviewerGroupName = safeDoc.reviewerGroupName || "";
  safeDoc.signers = Array.isArray(safeDoc.signers) ? safeDoc.signers.map(normalizeSigner) : [];
  safeDoc.finalPackUrl = safeDoc.finalPackUrl || "";
  safeDoc.finalPackFileName = safeDoc.finalPackFileName || "";
  safeDoc.finalPackUploadedAt = safeDoc.finalPackUploadedAt || "";
  safeDoc.finalPackUploadedBy = safeDoc.finalPackUploadedBy || "";
  safeDoc.signingSequenceLocked = !!safeDoc.signingSequenceLocked;
  safeDoc.awaitingFinalPackNotifiedAt = safeDoc.awaitingFinalPackNotifiedAt || "";
  safeDoc.currentSignerNotifiedAt = safeDoc.currentSignerNotifiedAt || "";
  // Migrate legacy "Awaiting Signature" rows (pre-final-pack workflow).
  if (safeDoc._column === "Awaiting Signature") {
    safeDoc._column = safeDoc.finalPackUrl ? "Signing in Progress" : "Awaiting Final Pack";
  }
  return syncDocReviewState(safeDoc);
}

function getAllDocReviewRecords() {
  const db = window.AEWTTR.db;
  const seen = new Set();
  const docs = [];
  Object.keys(db.docs || {}).forEach((column) => {
    (db.docs[column] || []).forEach((doc) => {
      if (seen.has(doc.id)) return;
      seen.add(doc.id);
      docs.push(normalizeDocReview(doc));
    });
  });
  return docs;
}

function getDocById(id) {
  return getAllDocReviewRecords().find((doc) => doc.id === id) || null;
}

function getDocRevision(doc, revisionId) {
  const safeDoc = normalizeDocReview(doc);
  return safeDoc.revisions.find((revision) => revision.id === (revisionId || safeDoc.activeRevisionId)) || safeDoc.revisions[0] || null;
}

function setDocActiveRevision(doc, revisionId) {
  const revision = getDocRevision(doc, revisionId);
  if (!revision) return doc;
  doc.activeRevisionId = revision.id;
  doc.currentRevisionNumber = revision.number;
  doc.fileUrl = revision.fileUrl || "";
  doc.fileName = revision.fileName || "";
  doc.fileType = revision.fileType || docFileType(revision.fileName || "", "");
  return doc;
}

function nextDocRevisionNumber(doc) {
  return Math.max(0, ...normalizeDocReview(doc).revisions.map((revision) => Number(revision && revision.number) || 0)) + 1;
}

function revisionFileLabel(revision) {
  return revision ? `Rev ${revision.number} · ${revision.fileName || "Untitled file"}` : "No revision";
}

function reviewerStatusBadge(decision) {
  const map = {
    "Pending": "status-Pending",
    "Approved": "status-Approved",
    "Requested Changes": "status-Denied"
  };
  return map[decision] || "status-Pending";
}

function docColumnStatusClass(column) {
  const key = String(column || "Not Started")
    .replace(/\s*\/\s*/g, "-")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-");
  return `status-${key}`;
}

function reviewerPillClass(decision) {
  if (decision === "Approved") return "done";
  if (decision === "Requested Changes") return "requested";
  return "pending";
}

/* Admins can approve/deny without being on the roster; when they act,
   seat them as a reviewer so the decision actually updates board state. */
function ensureActorReviewerSlot(doc) {
  const me = currentUserIdentity();
  let reviewer = (doc.reviewers || []).find((entry) => samePersonByNameOrEmail(entry, me));
  if (reviewer) return reviewer;
  reviewer = normalizeReviewer({ name: me.name || "Reviewer", email: me.email || "" });
  doc.reviewers = (doc.reviewers || []).concat([reviewer]);
  return reviewer;
}

function resolveReviewerEmail(reviewer) {
  const direct = String(reviewer && reviewer.email || "").trim();
  if (direct) return direct;
  const db = window.AEWTTR.db;
  const name = String(reviewer && reviewer.name || "").trim().toLowerCase();
  const match = (db.members || []).find((member) => {
    const memberName = String(member.name || "").trim().toLowerCase();
    const memberEmail = String(member.email || "").trim();
    return memberEmail && name && memberName === name;
  });
  return match ? String(match.email).trim() : "";
}

function getPendingReviewers(doc) {
  if (!doc || doc.isArchived) return [];
  normalizeDocReview(doc);
  return (doc.reviewers || []).filter((reviewer) => (reviewer.decision || "Pending") === "Pending");
}

function docNeedsAnyReview(doc) {
  if (!doc || doc.isArchived) return false;
  if (getPendingReviewers(doc).length > 0) return true;
  return docPendingSigners(doc).length > 0;
}

function latestDocRevisionLabel(doc) {
  const latestRevision = (doc.revisions || []).slice().sort((a, b) => (Number(b && b.number) || 0) - (Number(a && a.number) || 0))[0];
  return latestRevision && latestRevision.number ? `Rev ${latestRevision.number}` : "Current revision";
}

function docReviewNotificationFacts(doc, extraFacts) {
  const facts = [
    { title: "Document", value: (doc && doc.title) || "—" },
    { title: "Project", value: (doc && doc.projectCode) || "—" },
    { title: "Revision", value: latestDocRevisionLabel(doc) },
    { title: "Submitter", value: (doc && doc.submitter) || "—" }
  ];
  if (doc && doc.contractorName) facts.push({ title: "Contractor", value: doc.contractorName });
  if (doc && doc.deadline) facts.push({ title: "Due", value: fmtDate(doc.deadline) });
  return facts.concat(extraFacts || []);
}

function reviewerEligibleForReminder(reviewer, nowMs) {
  if ((reviewer.decision || "Pending") !== "Pending") return false;
  if (!resolveReviewerEmail(reviewer)) return false;
  const last = reviewer.lastNotifiedAt ? new Date(reviewer.lastNotifiedAt).getTime() : 0;
  if (!last || Number.isNaN(last)) return true;
  return nowMs - last >= DOC_REVIEW_REMINDER_INTERVAL_MS;
}

function collectDocStakeholderEmails(doc, excludeEmail) {
  const exclude = String(excludeEmail || "").trim().toLowerCase();
  const emails = new Set();
  const submitterEmail = resolveReviewerEmail({ name: doc.submitter, email: doc.submitterEmail || "" });
  if (submitterEmail && submitterEmail.toLowerCase() !== exclude) emails.add(submitterEmail);
  (doc.reviewers || []).forEach((reviewer) => {
    const email = resolveReviewerEmail(reviewer);
    if (email && email.toLowerCase() !== exclude) emails.add(email);
  });
  return Array.from(emails);
}

/* Deep links land on Document Review and open review vs signing popup.
   `mode`: "review" | "sign" (optional — omitted = review). */
function docReviewAppRoute(doc, mode) {
  if (typeof pulseAppRouteUrl !== "function") return typeof pulseAppUrl === "function" ? pulseAppUrl() : "";
  const query = { doc: doc && doc.id };
  if (mode === "sign") query.mode = "sign";
  else if (mode === "review") query.mode = "review";
  return pulseAppRouteUrl("docreview", query);
}

async function notifyDocReviewStakeholders(doc, { emails, subject, preview, body, facts, kind, mode }) {
  if (!isSharePointMode() || !doc) return;
  const recipients = (Array.isArray(emails) ? emails : [emails]).map((email) => String(email || "").trim()).filter(Boolean);
  if (!recipients.length) return;
  const linkMode = mode === "sign" ? "sign" : "review";
  try {
    await notifyUsers({
      to: recipients,
      subject,
      area: "Documents",
      kind: kind || "info",
      preview: preview || subject,
      facts: facts || docReviewNotificationFacts(doc),
      actionUrl: docReviewAppRoute(doc, linkMode),
      actionTitle: linkMode === "sign" ? "Open signing" : "Open Document Review",
      body
    });
  } catch (e) {
    console.warn("PULSE: document review stakeholder notification failed.", e);
  }
}

async function notifyPendingReviewers(doc, options) {
  options = options || {};
  if (!isSharePointMode() || !doc) return;
  const nowMs = Date.now();
  const nowIso = new Date(nowMs).toISOString();
  const kind = options.kind || "ready";
  const isReminder = kind === "reminder";
  const force = options.force !== false && !isReminder;
  const onlyEmails = options.onlyEmails
    ? new Set(options.onlyEmails.map((email) => String(email || "").trim().toLowerCase()).filter(Boolean))
    : null;
  const excludePeople = options.excludePeople || [];
  const previewByKind = {
    ready: `Your review is needed on ${doc.title}`,
    revision: `A new revision was uploaded — please review ${doc.title} again`,
    reminder: `Reminder: your review is still needed on ${doc.title}`,
    added: `You were added as a reviewer on ${doc.title}`
  };
  const targets = [];
  (doc.reviewers || []).forEach((reviewer) => {
    if ((reviewer.decision || "Pending") !== "Pending") return;
    if (excludePeople.some((person) => samePersonByNameOrEmail(reviewer, person))) return;
    if (options.excludeSubmitter && samePersonByNameOrEmail(reviewer, { name: doc.submitter, email: doc.submitterEmail || "" })) return;
    const email = resolveReviewerEmail(reviewer);
    if (!email) return;
    if (onlyEmails && !onlyEmails.has(email.toLowerCase())) return;
    if (isReminder) {
      if (!reviewerEligibleForReminder(reviewer, nowMs)) return;
    } else if (!force) {
      const last = reviewer.lastNotifiedAt ? new Date(reviewer.lastNotifiedAt).getTime() : 0;
      if (last && nowMs - last < 60000) return;
    }
    targets.push({ reviewer, email });
  });
  if (!targets.length) return;

  const preview = options.preview || previewByKind[kind] || previewByKind.ready;
  const subject = options.subject || `PULSE Document Review — ${doc.title}`;
  try {
    await notifyUsers({
      to: targets.map((target) => target.email),
      subject,
      area: "Documents",
      kind: "action",
      preview,
      facts: options.facts || docReviewNotificationFacts(doc),
      actionUrl: docReviewAppRoute(doc, "review"),
      actionTitle: "Open Document Review",
      body: options.body || undefined
    });
    targets.forEach(({ reviewer }) => {
      reviewer.lastNotifiedAt = nowIso;
    });
    if (options.persist !== false) {
      await Repo.save("docReviewItem", doc, { column: doc._column || "Not Started" });
    }
  } catch (e) {
    console.warn("PULSE: pending reviewer notification failed.", e);
  }
}

async function runDocReviewReminderSweep() {
  if (!isSharePointMode() || _docReviewReminderSweepRunning) return;
  _docReviewReminderSweepRunning = true;
  try {
    const docs = getAllDocReviewRecords().filter(docNeedsAnyReview);
    for (const doc of docs) {
      try {
        await notifyPendingReviewers(doc, { kind: "reminder", force: false });
      } catch (e) {}
    }
  } finally {
    _docReviewReminderSweepRunning = false;
  }
}

window.runDocReviewReminderSweep = runDocReviewReminderSweep;

function canCurrentUserReviewDoc(doc) {
  const me = currentUserIdentity();
  if (me.isAdmin) return true;
  return (doc.reviewers || []).some((reviewer) => samePersonByNameOrEmail(reviewer, me));
}

function isDocOwnerOrCreator(doc) {
  const me = currentUserIdentity();
  return samePersonByNameOrEmail({ name: doc.submitter, email: doc.submitterEmail || "" }, me)
    || samePersonByNameOrEmail({ name: doc.owner || doc.ownerName || "", email: "" }, me);
}

/* Document admins, app admins, and the person who owns/created the review
   may select/unselect signers. Everyone else can view the roster only. */
function canEditDocSigners(doc) {
  if (!doc || doc.isArchived) return false;
  if (doc.signingSequenceLocked && docHasFinalPack(doc)) return false;
  const me = currentUserIdentity();
  if (me.isAdmin) return true;
  if (canManageDocReview()) return true;
  return isDocOwnerOrCreator(doc);
}

function canCurrentUserSignDoc(doc) {
  if (!doc || doc.isArchived) return false;
  if (!docIsFullyReviewed(doc) || !docHasFinalPack(doc)) return false;
  const current = docCurrentSigner(doc);
  if (!current) return false;
  return samePersonByNameOrEmail(current, currentUserIdentity());
}

function canUploadFinalPack(doc) {
  if (!doc || doc.isArchived) return false;
  if (!docSignatureRequired(doc) || docHasFinalPack(doc)) return false;
  const me = currentUserIdentity();
  if (me.isAdmin) return true;
  if (canManageDocReview()) return true;
  return isDocOwnerOrCreator(doc);
}

function canCurrentUserPublishDocRevision(doc) {
  const me = currentUserIdentity();
  if (me.isAdmin) return true;
  return samePersonByNameOrEmail({ name: doc.submitter, email: doc.submitterEmail || "" }, me);
}

function canCurrentUserManageDoc(doc) {
  const me = currentUserIdentity();
  if (me.isAdmin) return true;
  return samePersonByNameOrEmail({ name: doc.submitter, email: doc.submitterEmail || "" }, me)
    || (doc.reviewers || []).some((reviewer) => samePersonByNameOrEmail(reviewer, me));
}

function canCurrentUserDeleteDoc(doc) {
  const me = currentUserIdentity();
  if (me.isAdmin) return true;
  return samePersonByNameOrEmail({ name: doc.submitter, email: doc.submitterEmail || "" }, me);
}

/* Archive/restore is Document Admins + leadership (Manager/system admins)
   only — deliberately NOT the submitter, unlike delete. A document owner
   shouldn't be able to pull their own doc out of the active review record. */
function canCurrentUserArchiveDoc() {
  return canManageDocReview();
}

function canCurrentUserAddReviewers(doc) {
  return canCurrentUserDeleteDoc(doc) || canManageDocReview();
}

/* Download: Document Admins, Manager-tier, Admins, and the current signer
   (so they can pull the packed PDF to sign offline). */
function canDownloadDoc(doc) {
  const me = currentUserIdentity();
  if (me.isAdmin) return true;
  const role = typeof currentAppRole === "function" ? currentAppRole() : "";
  if (role === "Document Admin" || role === "Manager") return true;
  const current = docCurrentSigner(doc);
  return !!(current && samePersonByNameOrEmail(current, me));
}

function collectDocAdminEmails() {
  const db = window.AEWTTR.db || {};
  const emails = new Set();
  (db.members || []).forEach((member) => {
    const email = String(member.email || "").trim();
    if (!email) return;
    if (member.isAdmin || member.isDocAdmin) emails.add(email);
  });
  return Array.from(emails);
}

function collectDocOwnerAndAdminEmails(doc) {
  const emails = new Set(collectDocAdminEmails());
  const submitterEmail = resolveReviewerEmail({ name: doc.submitter, email: doc.submitterEmail || "" });
  if (submitterEmail) emails.add(submitterEmail);
  return Array.from(emails);
}

async function notifyFinalPackNeeded(doc) {
  if (!doc || doc.awaitingFinalPackNotifiedAt) return false;
  const emails = collectDocOwnerAndAdminEmails(doc);
  if (!emails.length) return false;
  await notifyDocReviewStakeholders(doc, {
    emails,
    subject: `PULSE Document Review — Final Pack Needed — ${doc.title}`,
    preview: `${doc.title} is fully reviewed — upload the final packed PDF to start signing`,
    kind: "action",
    mode: "review",
    facts: docReviewNotificationFacts(doc, [{ title: "Next step", value: "Upload final packed PDF (PDF only)" }]),
    body: `<p><strong>${escapeHtml(doc.title)}</strong> is fully reviewed and needs a <strong>final packed PDF</strong> before sequential signing can begin.</p><p>This is not a revision upload — open Document Review and upload the final package.</p>`
  });
  doc.awaitingFinalPackNotifiedAt = new Date().toISOString();
  return true;
}

async function notifyCurrentSignerTurn(doc) {
  const current = docCurrentSigner(doc);
  if (!current) return false;
  const email = resolveReviewerEmail(current);
  if (!email) return false;
  const order = Number(current.signOrder) || 0;
  const total = docSigners(doc).length;
  await notifyDocReviewStakeholders(doc, {
    emails: [email],
    subject: `PULSE Document Review — Your Turn to Sign — ${doc.title}`,
    preview: `You are next to sign ${doc.title} (${order} of ${total})`,
    kind: "action",
    mode: "sign",
    facts: docReviewNotificationFacts(doc, [
      { title: "Signing order", value: `${order} of ${total}` },
      { title: "Action", value: "Download packed PDF, sign offline, upload signed PDF" }
    ]),
    body: `<p>You are the current signer for <strong>${escapeHtml(doc.title)}</strong> (${order} of ${total}).</p><p>Download the packed PDF, affix your signature offline, then upload the signed PDF in PULSE Document Review → Sign.</p>`
  });
  current.signNotifiedAt = new Date().toISOString();
  doc.currentSignerNotifiedAt = current.signNotifiedAt;
  syncDocSignersSnapshot(doc);
  return true;
}

async function notifyFullySigned(doc) {
  const emails = collectDocOwnerAndAdminEmails(doc);
  if (!emails.length) return false;
  await notifyDocReviewStakeholders(doc, {
    emails,
    subject: `PULSE Document Review — Fully Signed — ${doc.title}`,
    preview: `${doc.title} has been fully signed by all required signers`,
    kind: "success",
    mode: "sign",
    facts: docReviewNotificationFacts(doc, [{ title: "Status", value: "Fully signed" }]),
    body: `<p><strong>${escapeHtml(doc.title)}</strong> is now <strong>fully signed</strong>. All required signers have completed the signing sequence.</p>`
  });
  return true;
}

async function maybeNotifyFinalPackAfterReview(doc) {
  syncDocReviewState(doc);
  if (!docSigningStageFlags(doc).awaitingFinalPack) return false;
  const notified = await notifyFinalPackNeeded(doc);
  if (notified) await Repo.save("docReviewItem", doc, { column: doc._column || "Awaiting Final Pack" });
  return notified;
}

function renameFileForFinalPack(file, docTitle) {
  const base = String(docTitle || "document").trim().replace(/[^A-Za-z0-9._-]+/g, "-").replace(/^-+|-+$/g, "") || "document";
  const newName = `${base}-final-pack.pdf`;
  return new File([file], newName, { type: "application/pdf", lastModified: file.lastModified || Date.now() });
}

function renameFileForSignedPack(file, docTitle, signerName, order) {
  const base = String(docTitle || "document").trim().replace(/[^A-Za-z0-9._-]+/g, "-").replace(/^-+|-+$/g, "") || "document";
  const who = String(signerName || "signer").trim().replace(/[^A-Za-z0-9._-]+/g, "-").replace(/^-+|-+$/g, "") || "signer";
  const newName = `${base}-signed-${order || 0}-${who}.pdf`;
  return new File([file], newName, { type: "application/pdf", lastModified: file.lastModified || Date.now() });
}

async function uploadFinalPackedCopy(doc, file, actorName) {
  if (!isPdfFile(file)) throw new Error("Final packed copy must be a PDF.");
  const renamed = renameFileForFinalPack(file, doc.title);
  const fileInfo = await uploadDocReviewFile(doc.projectCode || "", doc.title, renamed);
  doc.finalPackUrl = fileInfo.fileUrl || "";
  doc.finalPackFileName = fileInfo.fileName || renamed.name;
  doc.finalPackUploadedAt = new Date().toISOString();
  doc.finalPackUploadedBy = actorName || currentUserIdentity().name || "";
  shuffleAndLockSigningOrder(doc);
  syncDocReviewState(doc);
  await saveDocReview(doc, {
    action: "Final Pack Uploaded",
    text: `${actorName || "A user"} uploaded the final packed PDF. Signing order was randomized.`
  });
  moveDocBetweenBuckets(doc);
  await notifyCurrentSignerTurn(doc);
  await saveDocReview(doc, null);
  return doc;
}

async function uploadSignedPackedCopy(doc, file, actorName) {
  if (!isPdfFile(file)) throw new Error("Signed copy must be a PDF.");
  if (!canCurrentUserSignDoc(doc)) throw new Error("It is not your turn to sign, or the document is not ready.");
  const me = currentUserIdentity();
  const current = docCurrentSigner(doc);
  if (!current || !samePersonByNameOrEmail(current, me)) throw new Error("You are not the current signer.");
  const renamed = renameFileForSignedPack(file, doc.title, current.name, current.signOrder);
  const fileInfo = await uploadDocReviewFile(doc.projectCode || "", doc.title, renamed);
  // Accumulate signatures on the current packed artifact (not a review revision).
  doc.finalPackUrl = fileInfo.fileUrl || doc.finalPackUrl;
  doc.finalPackFileName = fileInfo.fileName || renamed.name;
  current.signedAt = todayIsoDate();
  syncDocReviewState(doc);
  await saveDocReview(doc, {
    action: "Signed",
    text: `${actorName || me.name} uploaded a signed PDF (signer ${current.signOrder} of ${docSigners(doc).length}).`
  });
  moveDocBetweenBuckets(doc);
  if (doc._column === "Signed") {
    await notifyFullySigned(doc);
  } else {
    await notifyCurrentSignerTurn(doc);
    await saveDocReview(doc, null);
  }
  return doc;
}

function logDocDownload(doc, revision) {
  if (typeof logUserAction !== "function") return;
  logUserAction({
    action: "Download",
    area: "Documents",
    summary: `Downloaded "${doc.title}" — ${revision ? revisionFileLabel(revision) : (doc.finalPackFileName || "packed copy")}`,
    recordId: doc.id,
    detail: { docId: doc.id, revisionId: revision && revision.id, revisionNumber: revision && revision.number, finalPack: !revision }
  });
}

async function appendReviewersToDoc(doc, people, actorName) {
  const additions = [];
  (people || []).forEach((person) => {
    if (!person || !person.name) return;
    if ((doc.reviewers || []).some((reviewer) => samePersonByNameOrEmail(reviewer, person))) return;
    additions.push(normalizeReviewer({ name: person.name, email: person.email || "", isSigner: !!person.isSigner }));
  });
  if (!additions.length) return 0;
  doc.reviewers = (doc.reviewers || []).concat(additions);
  syncDocReviewState(doc);
  const names = additions.map((reviewer) => reviewer.name).join(", ");
  await saveDocReview(doc, { action: "Reviewers Added", text: names });
  moveDocBetweenBuckets(doc);
  const addedEmails = additions.map((reviewer) => resolveReviewerEmail(reviewer)).filter(Boolean);
  await notifyPendingReviewers(
    doc,
    {
      kind: "added",
      subject: `PULSE Document Review — Reviewers Added — ${doc.title}`,
      preview: `${actorName || "A user"} added you as a reviewer on ${doc.title}`,
      body: `<p><strong>${escapeHtml(actorName || "A user")}</strong> added reviewer(s) to <strong>${escapeHtml(doc.title)}</strong>: ${escapeHtml(names)}.</p>`,
      onlyEmails: addedEmails,
      force: true
    }
  );
  return additions.length;
}

async function removeReviewerFromDoc(doc, index, actorName) {
  const reviewer = (doc.reviewers || [])[index];
  if (!reviewer) return false;
  doc.reviewers.splice(index, 1);
  // If this person was a signer, remove them from the signers snapshot too — otherwise
  // syncDocReviewState → mergeSignersIntoReviewers immediately re-adds them.
  if (reviewer.isSigner && Array.isArray(doc.signers)) {
    doc.signers = doc.signers.filter((s) => !samePersonByNameOrEmail(s, reviewer));
  }
  syncDocReviewState(doc);
  await saveDocReview(doc, { action: "Reviewer Removed", text: `${actorName || "A user"} removed ${reviewer.name} from the review team.` });
  moveDocBetweenBuckets(doc);
  return true;
}

/* Flip isSigner on one reviewer and keep SignersJson in sync. Snapshot
   MUST be rewritten before syncDocReviewState → mergeSignersIntoReviewers,
   or an unselect is immediately undone by the stale signers list. */
function applyReviewerSignerToggle(doc, index) {
  if (!canEditDocSigners(doc)) return null;
  const reviewer = (doc.reviewers || [])[index];
  if (!reviewer) return null;
  reviewer.isSigner = !reviewer.isSigner;
  if (!reviewer.isSigner) {
    reviewer.signedAt = "";
    reviewer.signOrder = 0;
    reviewer.signNotifiedAt = "";
  }
  syncDocSignersSnapshot(doc);
  syncDocReviewState(doc);
  return reviewer;
}

async function toggleReviewerSigner(doc, index, actorName) {
  const reviewer = applyReviewerSignerToggle(doc, index);
  if (!reviewer) return false;
  await saveDocReview(doc, { action: reviewer.isSigner ? "Signer Assigned" : "Signer Removed", text: `${actorName || "A user"} ${reviewer.isSigner ? "assigned" : "removed"} ${reviewer.name} ${reviewer.isSigner ? "as a required signer" : "from required signers"}.` });
  moveDocBetweenBuckets(doc);
  await maybeNotifyFinalPackAfterReview(doc);
  return true;
}

async function saveDocSigners(doc, people, actorName) {
  if (!canEditDocSigners(doc)) return 0;
  const beforeKeys = docSigners(doc).map((signer) => normalizePersonKey(signer.email || signer.name)).sort().join("|");
  applySignersToDoc(doc, people);
  syncDocReviewState(doc);
  const afterList = docSigners(doc);
  const afterKeys = afterList.map((signer) => normalizePersonKey(signer.email || signer.name)).sort().join("|");
  if (beforeKeys === afterKeys) return 0;
  const names = afterList.map((signer) => signer.name);
  await saveDocReview(doc, {
    action: "Signers Updated",
    text: names.length
      ? `${actorName || "A user"} set required signers: ${names.join(", ")}.`
      : `${actorName || "A user"} cleared required signers.`
  });
  moveDocBetweenBuckets(doc);
  await maybeNotifyFinalPackAfterReview(doc);
  return names.length || 1;
}

function isCurrentUserOnDoc(doc) {
  const me = currentUserIdentity();
  return samePersonByNameOrEmail({ name: doc.submitter, email: doc.submitterEmail || "" }, me)
    || (doc.reviewers || []).some((reviewer) => samePersonByNameOrEmail(reviewer, me));
}

function doesDocNeedCurrentUserReview(doc) {
  if (!doc || doc.isArchived) return false;
  const me = currentUserIdentity();
  if (samePersonByNameOrEmail({ name: doc.submitter, email: doc.submitterEmail || "" }, me)) {
    const selfReviewer = (doc.reviewers || []).find((reviewer) => samePersonByNameOrEmail(reviewer, me));
    if (selfReviewer && selfReviewer.decision === "Pending") return true;
  }
  if ((doc.reviewers || []).some((reviewer) => samePersonByNameOrEmail(reviewer, me) && reviewer.decision === "Pending")) return true;
  if (canCurrentUserSignDoc(doc)) return true;
  if (canUploadFinalPack(doc)) return true;
  return false;
}

function getDocReviewMetrics() {
  const docs = getAllDocReviewRecords();
  const myDocs = docs.filter(isCurrentUserOnDoc);
  const needsMyReview = docs.filter(doesDocNeedCurrentUserReview);
  return {
    allCount: docs.length,
    myCount: myDocs.length,
    needsMyReviewCount: needsMyReview.length
  };
}

window.getDocReviewMetrics = getDocReviewMetrics;

function getDocReviewNotificationEvents() {
  const docs = getAllDocReviewRecords().filter(doesDocNeedCurrentUserReview);
  return docs.map((doc) => {
    const latestRevision = (doc.revisions || []).slice().sort((a, b) => (b.number || 0) - (a.number || 0))[0];
    const revisionLabel = latestRevision && latestRevision.number ? `Rev ${latestRevision.number}` : "Current revision";
    const needsSignature = canCurrentUserSignDoc(doc);
    const needsPack = canUploadFinalPack(doc);
    const mode = needsSignature && !needsPack ? "sign" : "review";
    return {
      id: `doc-review-${doc.id}`,
      route: "docreview",
      queryDoc: doc.id,
      queryMode: mode,
      icon: needsPack ? "bx-package" : (needsSignature ? "bx-pen" : "bx-file-blank"),
      tone: "teal",
      title: doc.title || doc.id,
      preview: `${doc.projectCode || "No project"} · ${revisionLabel} · ${needsPack ? "final pack PDF needed" : (needsSignature ? "your signature is needed" : "your review is needed")}`,
      time: fmtRelativeTime((latestRevision && latestRevision.uploadedAt) || doc.submittedAt || ""),
      category: "Document Review",
      area: "Documents",
      sortKey: `0-${doc.title || doc.id}`
    };
  });
}

window.getDocReviewNotificationEvents = getDocReviewNotificationEvents;

function cleanSharePointFileUrl(fileUrl) {
  if (window.AEWTTR && window.AEWTTR.OfficeDesktop && typeof window.AEWTTR.OfficeDesktop.cleanSharePointFileUrl === "function") {
    return window.AEWTTR.OfficeDesktop.cleanSharePointFileUrl(fileUrl);
  }
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

function previewUrlForRevision(revision) {
  if (!revision || !revision.fileUrl) return "";
  const fileType = docFileType(revision.fileName, revision.mimeType);
  if (fileType === "docx" || fileType === "pptx") {
    return `${revision.fileUrl}${revision.fileUrl.includes("?") ? "&" : "?"}web=1&action=embedview`;
  }
  return revision.fileUrl;
}

function sharePointOpenUrlForRevision(revision) {
  if (!revision || !revision.fileUrl) return "";
  const fileType = docFileType(revision.fileName, revision.mimeType);
  if (fileType === "docx" || fileType === "pptx") {
    return `${revision.fileUrl}${revision.fileUrl.includes("?") ? "&" : "?"}web=1`;
  }
  return revision.fileUrl;
}

function sharePointOpenLabelForRevision(revision) {
  const fileType = docFileType(revision && revision.fileName, revision && revision.mimeType);
  if (fileType === "docx") return "Open in Word app";
  if (fileType === "pptx") return "Open in SharePoint";
  return "Open in SharePoint";
}

function revisionSupportsDesktopAppOpen(revision) {
  const fileType = docFileType(revision && revision.fileName, revision && revision.mimeType);
  return fileType === "docx";
}

/* Office URI scheme against the shared SPO file URL — opens the live copy in
   the installed desktop app (not a download). Same helper as PPTX export. */
function officeDesktopUriForRevision(revision) {
  if (!revisionSupportsDesktopAppOpen(revision)) return "";
  const fileType = docFileType(revision.fileName, revision.mimeType);
  const api = window.AEWTTR && window.AEWTTR.OfficeDesktop;
  if (api && typeof api.officeDesktopUri === "function") {
    return api.officeDesktopUri(revision && revision.fileUrl, fileType);
  }
  const fileUrl = cleanSharePointFileUrl(revision && revision.fileUrl);
  if (!fileUrl || !/^https?:\/\//i.test(fileUrl)) return "";
  return `ms-word:ofe|u|${encodeURI(fileUrl)}`;
}

function openRevisionInDesktopApp(revision) {
  const fileType = docFileType(revision && revision.fileName, revision && revision.mimeType);
  const api = window.AEWTTR && window.AEWTTR.OfficeDesktop;
  if (api && typeof api.openSharePointFileInDesktopApp === "function") {
    if (!api.openSharePointFileInDesktopApp(revision && revision.fileUrl, fileType)) {
      toast("Open in app is available for Word (.docx) files.", "error");
    }
    return;
  }
  const uri = officeDesktopUriForRevision(revision);
  if (!uri) {
    toast("Open in app is available for Word (.docx) files.", "error");
    return;
  }
  const anchor = document.createElement("a");
  anchor.setAttribute("href", uri);
  anchor.style.display = "none";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
}

function sharePointFolderUrlForRevision(revision) {
  if (!revision || !revision.fileUrl) return "";
  try {
    const url = new URL(revision.fileUrl, window.location.origin);
    url.search = "";
    url.hash = "";
    const path = decodeURIComponent(url.pathname || "");
    const slash = path.lastIndexOf("/");
    if (slash <= 0) return "";
    const folderPath = path.slice(0, slash);
    const parts = folderPath.split("/").filter(Boolean);
    let libraryEnd = -1;
    parts.forEach((part, index) => {
      if (/documents$/i.test(part) || /^shared documents$/i.test(part)) libraryEnd = index;
    });
    if (libraryEnd < 0) {
      return `${url.origin}/${parts.map((segment) => encodeURIComponent(segment)).join("/")}`;
    }
    const libraryPath = "/" + parts.slice(0, libraryEnd + 1).join("/");
    return `${url.origin}${libraryPath}/Forms/AllItems.aspx?id=${encodeURIComponent(folderPath)}`;
  } catch (e) {
    return "";
  }
}

function openRevisionInSharePoint(revision) {
  if (!revision || !revision.fileUrl) return;
  const api = window.AEWTTR && window.AEWTTR.OfficeDesktop;
  if (api && typeof api.openSharePointFileByPolicy === "function") {
    api.openSharePointFileByPolicy(revision.fileUrl, revision.fileName, revision.mimeType);
    return;
  }
  window.open(sharePointOpenUrlForRevision(revision), "_blank", "noopener");
}

function openRevisionFolderInSharePoint(revision) {
  const folderUrl = sharePointFolderUrlForRevision(revision);
  if (!folderUrl) {
    toast("Couldn't find a SharePoint folder for this file.", "error");
    return;
  }
  window.open(folderUrl, "_blank", "noopener");
}

function revisionHasOfficeOnlineOpen(revision) {
  const fileType = docFileType(revision && revision.fileName, revision && revision.mimeType);
  return fileType === "docx" || fileType === "pptx";
}

function revisionHasWordOnlineOpen(revision) {
  return revisionHasOfficeOnlineOpen(revision);
}

function revisionServerPathLooksArchived(serverRelativeUrl) {
  const path = String(serverRelativeUrl || "");
  return /\/previous revisions\//i.test(path);
}

/* After a new current file is uploaded, relocate every prior revision file
   that still sits in the document folder root into the real SPO
   "Previous revisions" subfolder. Updates revision URLs in place. Failures
   toast + log but do not unwind the new revision. */
async function movePriorRevisionFilesToSharePointArchive(doc, keepRevisionId) {
  if (!isSharePointMode() || !doc || typeof sharePointAdapter === "undefined" || !sharePointAdapter.moveFileToSubfolder) {
    return { moved: 0, failed: 0 };
  }
  const siteUrl = currentSiteUrl();
  let moved = 0;
  let failed = 0;
  const priors = (doc.revisions || []).filter((rev) => rev && rev.id !== keepRevisionId && rev.serverRelativeUrl && !revisionServerPathLooksArchived(rev.serverRelativeUrl));
  for (const prior of priors) {
    try {
      const result = await sharePointAdapter.moveFileToSubfolder(siteUrl, prior.serverRelativeUrl, DOC_REVIEW_PREVIOUS_REVISIONS_FOLDER);
      if (result && result.serverRelativeUrl) {
        prior.serverRelativeUrl = result.serverRelativeUrl;
        prior.fileUrl = result.fileUrl || cleanSharePointFileUrl(prior.fileUrl) || prior.fileUrl;
        moved += 1;
      }
    } catch (e) {
      failed += 1;
      console.warn("PULSE: could not move previous revision into Previous revisions folder.", e);
    }
  }
  if (failed) {
    toast(`New revision uploaded, but ${failed} previous file${failed === 1 ? "" : "s"} could not be moved into “${DOC_REVIEW_PREVIOUS_REVISIONS_FOLDER}”. Check SharePoint permissions.`, "error");
  } else if (moved) {
    console.info(`PULSE: moved ${moved} previous revision file(s) into “${DOC_REVIEW_PREVIOUS_REVISIONS_FOLDER}”.`);
  }
  return { moved, failed };
}

function renderRevisionViewer(doc, revision, large) {
  if (!revision) return `<div class="docreview-viewer-empty">No uploaded file yet.</div>`;
  const fileType = docFileType(revision.fileName, revision.mimeType);
  if (fileType === "text") {
    return `<pre class="docreview-text-preview">${escapeHtml(revision.inlineText || "Open this file to view its contents.")}</pre>`;
  }
  if (revision.fileUrl) {
    return `<iframe class="docreview-iframe-viewer" src="${escapeHtml(previewUrlForRevision(revision))}" title="${escapeHtml(revision.fileName || "Document preview")}"></iframe>`;
  }
  return `<div class="docreview-viewer-empty">No preview is available for this revision yet.</div>`;
}

function openDocViewerModal(doc, revision) {
  const canDownload = canDownloadDoc(doc);
  const fileType = docFileType(revision && revision.fileName, revision && revision.mimeType);
  const desktopUri = officeDesktopUriForRevision(revision);
  const modal = openModal(`
    <div class="aewttr-modal-head">
      <h3>${escapeHtml(doc.title)} — ${escapeHtml(revisionFileLabel(revision))}</h3>
      <button class="aewttr-modal-close">&times;</button>
    </div>
    <div class="aewttr-modal-body">
      <div class="docreview-viewer-toolbar">
        ${revision && revision.fileUrl ? `<button type="button" class="btn-aewttr btn-aewttr-sm" id="dv-open-file"${tip(desktopUri ? "Open this file in Word" : sharePointOpenLabelForRevision(revision))}><i class="bx ${desktopUri ? "bx-desktop" : "bx-link-external"}"></i> ${desktopUri ? "Open in Word" : "Open"}</button>` : ""}
      </div>
      <div class="docreview-viewer-large">
        ${renderRevisionViewer(doc, revision, true)}
      </div>
    </div>
  `, { xwide: true });
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  const openFileBtn = $("#dv-open-file", modal);
  if (openFileBtn) openFileBtn.addEventListener("click", () => {
    if (revisionSupportsDesktopAppOpen(revision) && officeDesktopUriForRevision(revision)) {
      openRevisionInDesktopApp(revision);
    } else {
      openRevisionInSharePoint(revision);
    }
    logDocDownload(doc, revision);
  });
}

function renameFileForDocRevision(file, nextNumber, docTitle) {
  const ext = docFileExtension(file && file.name || "");
  const base = String(docTitle || "document").trim().replace(/[^A-Za-z0-9._-]+/g, "-").replace(/^-+|-+$/g, "") || "document";
  const cleanExt = ext ? `.${ext}` : "";
  const newName = `${base}-rev-${nextNumber}${cleanExt}`;
  return new File([file], newName, { type: file.type || "application/octet-stream", lastModified: file.lastModified || Date.now() });
}

async function createRevisionFromUpload(doc, file, sourceLabel) {
  const nextNumber = nextDocRevisionNumber(doc);
  const renamedFile = renameFileForDocRevision(file, nextNumber, doc.title);
  const fileInfo = await uploadDocReviewFile(doc.projectCode || "", doc.title, renamedFile);
  return {
    id: uid("REV"),
    number: nextNumber,
    fileUrl: fileInfo.fileUrl || "",
    serverRelativeUrl: fileInfo.serverRelativeUrl || "",
    fileName: fileInfo.fileName || renamedFile.name,
    fileType: docFileType(renamedFile.name, renamedFile.type),
    mimeType: renamedFile.type || "",
    uploadedBy: currentUserIdentity().name,
    uploadedDate: todayIsoDate(),
    source: sourceLabel || "Manual Upload"
  };
}

/* New revision = restart the review: every reviewer's decision resets to
   Pending, signer assignments are kept (who must sign doesn't change) but
   their signatures reset since they're signing a different revision.
   Final pack + signing sequence also reset — a new revision restarts
   packaging after the next full approval. */
function resetReviewersForNewRevision(doc) {
  doc.reviewers = (doc.reviewers || []).map((reviewer) => ({
    name: reviewer.name,
    email: reviewer.email || "",
    decision: "Pending",
    reviewedAt: "",
    note: "",
    lastNotifiedAt: "",
    isSigner: !!reviewer.isSigner,
    signedAt: "",
    signOrder: 0,
    signNotifiedAt: ""
  }));
  doc.isArchived = false;
  doc.archivedDate = "";
  doc.fullyReviewedDate = "";
  doc.signedDate = "";
  clearFinalPackAndSigning(doc);
  syncDocReviewState(doc);
}

async function saveDocReview(doc, note) {
  syncDocReviewState(doc);
  if (note) appendDocActivity(doc, note.action, note.text || "");
  // If a background refresh replaced window.AEWTTR.db while a modal held
  // this doc by reference, re-seat it so redraws see the mutation.
  const live = getDocById(doc.id);
  if (live !== doc) moveDocBetweenBuckets(doc);
  await Repo.save("docReviewItem", doc, { column: doc._column || "Not Started" });
}

function moveDocBetweenBuckets(doc) {
  const db = window.AEWTTR.db;
  Object.keys(db.docs || {}).forEach((column) => {
    db.docs[column] = (db.docs[column] || []).filter((entry) => entry.id !== doc.id);
  });
  const bucket = doc._column || "Not Started";
  if (!db.docs[bucket]) db.docs[bucket] = [];
  db.docs[bucket].unshift(doc);
}

async function maybeAutoArchiveDoc(doc) {
  normalizeDocReview(doc);
  if (doc.isArchived || !doc.fullyReviewedDate) return false;
  // Don't auto-archive a document that's still waiting on signatures —
  // "Review Complete" (no signature needed) can archive after a week, but
  // "Awaiting Final Pack" / "Signing in Progress" must reach "Signed" first.
  if (docSignatureRequired(doc) && doc._column !== "Signed") return false;
  if (isoDaysBetween(doc.fullyReviewedDate, todayIsoDate()) < 7) return false;
  doc.isArchived = true;
  doc.archivedDate = todayIsoDate();
  syncDocReviewState(doc);
  await saveDocReview(doc, { action: "Auto Archived", text: "Moved to Archive one week after full review." });
  moveDocBetweenBuckets(doc);
  await notifyDocReviewStakeholders(doc, {
    emails: collectDocStakeholderEmails(doc),
    subject: `PULSE Document Review Archived — ${doc.title}`,
    preview: `${doc.title} was automatically archived after one week in a fully reviewed state.`,
    kind: "info",
    body: `<p><strong>${escapeHtml(doc.title)}</strong> was automatically archived after one week in a fully reviewed state.</p>`
  });
  return true;
}

function docStatusSummary(doc) {
  const total = (doc.reviewers || []).length;
  const approved = (doc.reviewers || []).filter((reviewer) => reviewer.decision === "Approved").length;
  const requested = (doc.reviewers || []).filter((reviewer) => reviewer.decision === "Requested Changes").length;
  const pending = total - approved - requested;
  const signers = docSignersOrdered(doc);
  const signed = signers.filter((reviewer) => reviewer.signedAt).length;
  const stages = docSigningStageFlags(doc);
  const current = docCurrentSigner(doc);
  return {
    total, approved, requested, pending,
    signerTotal: signers.length,
    signed,
    stages,
    currentSigner: current,
    signingPosition: current ? (Number(current.signOrder) || (signed + 1)) : (stages.fullySigned ? signers.length : 0)
  };
}

function docWorkflowStageIndex(doc) {
  const stages = docSigningStageFlags(doc);
  if (doc.isArchived || stages.fullySigned || doc._column === "Signed" || doc._column === "Review Complete") return 3;
  if (stages.signingInProgress || stages.finalPacked) return 2;
  if (stages.awaitingFinalPack || stages.fullyReviewed) return stages.signatureRequired ? 1 : 3;
  return 0;
}

function renderDocWorkflowStages(doc) {
  const summary = docStatusSummary(doc);
  const active = docWorkflowStageIndex(doc);
  const signingLabel = summary.signerTotal
    ? (summary.stages.signingInProgress
      ? `Signing (${summary.signed} of ${summary.signerTotal})`
      : (summary.stages.fullySigned ? "Complete" : "Signing"))
    : "Signing";
  const labels = [
    { key: "review", label: "In review" },
    { key: "pack", label: "Awaiting final pack" },
    { key: "sign", label: signingLabel },
    { key: "done", label: "Complete" }
  ];
  if (!summary.stages.signatureRequired) {
    labels[1].label = "Review complete";
    labels[2].label = "—";
  }
  return `
    <nav class="docreview-stage-band" aria-label="Document workflow stage">
      <ol class="docreview-stage-track">
        ${labels.map((stage, index) => `
          <li class="docreview-stage-step ${index < active ? "is-done" : ""} ${index === active ? "is-active" : ""} ${!summary.stages.signatureRequired && index === 2 ? "is-skipped" : ""}">
            <span class="docreview-stage-index">${index < active ? "<i class='bx bx-check'></i>" : (index + 1)}</span>
            <span class="docreview-stage-label">${escapeHtml(stage.label)}</span>
          </li>`).join("")}
      </ol>
    </nav>`;
}

/* ---------- people groups (global — managed in Admin) ----------
   Document Review uses the shared db.groups store. Selecting a group in a
   reviewer picker expands members into individuals on this document. */

function loadReviewerGroups() {
  return typeof loadPulseGroups === "function" ? loadPulseGroups() : (window.AEWTTR.db.groups || window.AEWTTR.db.docReviewerGroups || []);
}

/* ---------- search & filters ---------- */

function docReviewSearchState() {
  if (!window.AEWTTR.state.docSearch) {
    window.AEWTTR.state.docSearch = { q: "", project: "", contractor: "", reviewer: "", signer: "", status: "", dueBefore: "" };
  }
  return window.AEWTTR.state.docSearch;
}

function docReviewDistinctValues(docs) {
  const projects = new Map();
  const contractors = new Set();
  const people = new Map();
  const signers = new Map();
  docs.forEach((doc) => {
    if (doc.projectCode) {
      const proj = window.AEWTTR.db.projects.find((p) => p.id === doc.projectCode);
      projects.set(doc.projectCode, proj ? proj.name : doc.projectCode);
    }
    if (doc.contractorName) contractors.add(doc.contractorName);
    (doc.reviewers || []).forEach((reviewer) => {
      if (!reviewer.name) return;
      people.set(normalizePersonKey(reviewer.email || reviewer.name), reviewer.name);
      if (reviewer.isSigner) signers.set(normalizePersonKey(reviewer.email || reviewer.name), reviewer.name);
    });
  });
  const byName = (a, b) => a.name.localeCompare(b.name);
  return {
    projects: Array.from(projects.entries()).map(([id, name]) => ({ id, name })),
    contractors: Array.from(contractors).sort(),
    people: Array.from(people.entries()).map(([key, name]) => ({ key, name })).sort(byName),
    signers: Array.from(signers.entries()).map(([key, name]) => ({ key, name })).sort(byName)
  };
}

function docMatchesSearch(doc, state) {
  const q = String(state.q || "").trim().toLowerCase();
  if (q) {
    const proj = doc.projectCode ? window.AEWTTR.db.projects.find((p) => p.id === doc.projectCode) : null;
    const hay = [doc.title, doc.contractorName, doc.submitter, proj && proj.name, doc.portfolioOrConfigItem, doc.type]
      .filter(Boolean).join(" ").toLowerCase();
    if (!hay.includes(q)) return false;
  }
  if (state.project && doc.projectCode !== state.project) return false;
  if (state.contractor && doc.contractorName !== state.contractor) return false;
  if (state.reviewer) {
    const has = (doc.reviewers || []).some((reviewer) => normalizePersonKey(reviewer.email || reviewer.name) === state.reviewer);
    if (!has) return false;
  }
  if (state.signer) {
    const has = docSigners(doc).some((reviewer) => normalizePersonKey(reviewer.email || reviewer.name) === state.signer);
    if (!has) return false;
  }
  if (state.status && doc._column !== state.status) return false;
  if (state.dueBefore && (!doc.deadline || doc.deadline > state.dueBefore)) return false;
  return true;
}

PAGE_RENDERERS.docreview = function () {
  try {
    if (!window.AEWTTR.state.docToolView) window.AEWTTR.state.docToolView = "Review";
    if (!window.AEWTTR.state.docBoardView) window.AEWTTR.state.docBoardView = "Active";
    if (!window.AEWTTR.state.docSignBoardView) window.AEWTTR.state.docSignBoardView = "Needs Signature";
    if (!window.AEWTTR.state.docScope) window.AEWTTR.state.docScope = "My Documents";
    const db = window.AEWTTR.db;
    setTopbar("Document Review", "Track document revisions, reviewer decisions, signatures, and audit history.", `
      <button class="btn-aewttr" id="btn-upload-doc"${tip("Upload a new document for review")}><i class="bx bx-upload"></i> Upload Document</button>
    `);

  const scopeTabs = ["My Documents", "All Documents"];
  const isSignView = () => window.AEWTTR.state.docToolView === "Sign";

  async function runArchiveSweep() {
    if (_docReviewArchiveSweepStarted) return;
    _docReviewArchiveSweepStarted = true;
    try {
      if (typeof runDocReviewReminderSweep === "function") await runDocReviewReminderSweep();
    } catch (e) {}
    _docReviewArchiveSweepStarted = false;
    draw();
  }

  function scopedFilteredDocs() {
    const state = docReviewSearchState();
    const scoped = getAllDocReviewRecords().filter((doc) => {
      return window.AEWTTR.state.docScope === "All Documents" ? true : isCurrentUserOnDoc(doc);
    });
    return scoped.filter((doc) => {
      if (!docMatchesSearch(doc, state)) return false;
      if (isSignView()) return docSigners(doc).length > 0 && !doc.isArchived;
      return true;
    });
  }

  function groupedDocs(docs) {
    const groups = {};
    DOC_REVIEW_STATUSES.forEach((s) => { groups[s.key] = []; });
    docs.forEach((doc) => {
      const bucket = doc._column || "Not Started";
      if (!groups[bucket]) groups[bucket] = [];
      groups[bucket].push(doc);
    });
    return groups;
  }

  function visibleColumns() {
    if (isSignView()) {
      return DOC_REVIEW_SIGN_BOARD_VIEWS[window.AEWTTR.state.docSignBoardView] || DOC_REVIEW_SIGN_BOARD_VIEWS["Needs Signature"];
    }
    return DOC_REVIEW_BOARD_VIEWS[window.AEWTTR.state.docBoardView] || DOC_REVIEW_BOARD_VIEWS.Active;
  }

  function boardViewKeys() {
    return isSignView() ? Object.keys(DOC_REVIEW_SIGN_BOARD_VIEWS) : Object.keys(DOC_REVIEW_BOARD_VIEWS);
  }

  function activeBoardView() {
    return isSignView() ? window.AEWTTR.state.docSignBoardView : window.AEWTTR.state.docBoardView;
  }

  function docCardHtml(doc) {
    const summary = docStatusSummary(doc);
    const project = doc.projectCode ? db.projects.find((p) => p.id === doc.projectCode) : null;
    const needsReview = doesDocNeedCurrentUserReview(doc);
    const needsSign = canCurrentUserSignDoc(doc);
    const needsAction = needsReview || needsSign;
    const actionBadge = needsSign
      ? `<span class="docreview-need-review-badge docreview-need-sign-badge">Your signature</span>`
      : (needsReview ? `<span class="docreview-need-review-badge">Your review</span>` : "");
    const counts = isSignView()
      ? `${summary.signed}/${summary.signerTotal} signed`
      : `${summary.approved}/${summary.total} approved${summary.signerTotal ? ` · ${summary.signed}/${summary.signerTotal} signed` : ""}${summary.requested ? ` · ${summary.requested} changes` : ""}`;
    const metaBits = [
      project ? project.name : "",
      doc.submitter || "",
      doc.deadline ? `Due ${fmtDate(doc.deadline)}` : "",
      counts
    ].filter(Boolean);
    return `
      <button type="button" class="kanban-card docreview-row${needsAction ? " needs-user-review" : ""}${isSignView() ? " docreview-card--sign" : ""}" data-id="${doc.id}">
        <span class="docreview-row-top">
          <span class="docreview-row-name">${escapeHtml(doc.title || "Untitled")}</span>
          <span class="docreview-row-status">${escapeHtml(doc._column || doc.status || "")}</span>
          ${actionBadge}
        </span>
        <span class="docreview-row-meta">${metaBits.map((bit) => `<span>${escapeHtml(bit)}</span>`).join("")}</span>
      </button>
    `;
  }

  function emptyColHtml() {
    if (isSignView()) {
      return `<div class="empty-state docreview-empty-col">
        <strong>No documents here</strong>
        <span>Mark reviewers as Signers from a document’s Reviewers tab (Sign button on each person).</span>
      </div>`;
    }
    return `<div class="empty-state docreview-empty-col">
      <strong>No documents</strong>
      <span>Upload a document to start a review.</span>
    </div>`;
  }

  function draw() {
    const filtered = scopedFilteredDocs();
    const docsByColumn = groupedDocs(filtered);
    const cols = visibleColumns();
    const state = docReviewSearchState();
    const distinct = docReviewDistinctValues(getAllDocReviewRecords());
    const toolView = window.AEWTTR.state.docToolView;
    const boardKeys = boardViewKeys();
    const boardActive = activeBoardView();
    const signAwaitingCount = getAllDocReviewRecords().filter((doc) => {
      const inScope = window.AEWTTR.state.docScope === "All Documents" || isCurrentUserOnDoc(doc);
      return inScope && docSigners(doc).length > 0 && ["Awaiting Final Pack", "Signing in Progress"].includes(doc._column);
    }).length;
    $("#page-content").innerHTML = `
      <div class="docreview-page" style="--docreview-col-count:${cols.length};">
        <div class="docreview-view-tabs" role="tablist" aria-label="Document Review views">
          ${DOC_REVIEW_TOOL_VIEWS.map((view) => `
            <button type="button" class="docreview-view-tab ${toolView === view ? "active" : ""}" data-tool-view="${view}" role="tab" aria-selected="${toolView === view ? "true" : "false"}"${tip(view === "Review" ? "Submit, review, approve, comments, and revisions" : "Documents with assigned signers — pending and complete")}>
              <i class="bx ${view === "Review" ? "bx-file-find" : "bx-pen"}"></i>
              ${view}
              ${view === "Sign" && signAwaitingCount ? `<em>${signAwaitingCount}</em>` : ""}
            </button>`).join("")}
        </div>
        <div class="docreview-toolbar">
          <div class="docreview-toolbar-block">
            <div class="docreview-toolbar-label">Scope</div>
            <div class="filter-pills" id="doc-scope-filters">
              ${scopeTabs.map((scope) => `<button class="filter-pill ${window.AEWTTR.state.docScope === scope ? "active" : ""}" data-scope="${scope}"${tip(scope === "My Documents" ? "Documents you submitted, are reviewing, or must sign" : "All documents in the workspace")}>${scope}</button>`).join("")}
            </div>
          </div>
          <div class="docreview-toolbar-block">
            <div class="docreview-toolbar-label">${isSignView() ? "Signing board" : "Board"}</div>
            <div class="filter-pills" id="doc-filters">
              ${boardKeys.map((view) => `<button class="filter-pill ${boardActive === view ? "active" : ""}" data-f="${view}">${view}</button>`).join("")}
            </div>
          </div>
        </div>
        <div class="docreview-search-bar">
          <div class="docreview-search-input">
            <i class="bx bx-search"></i>
            <input class="input-aewttr" id="doc-search-q" placeholder="Search title, contractor, project, submitter…" value="${escapeHtml(state.q)}">
          </div>
          <select class="select-aewttr" id="doc-search-project"><option value="">Any project</option>${distinct.projects.map((p) => `<option value="${escapeHtml(p.id)}" ${state.project === p.id ? "selected" : ""}>${escapeHtml(p.name)}</option>`).join("")}</select>
          <select class="select-aewttr" id="doc-search-contractor"><option value="">Any contractor</option>${distinct.contractors.map((c) => `<option value="${escapeHtml(c)}" ${state.contractor === c ? "selected" : ""}>${escapeHtml(c)}</option>`).join("")}</select>
          ${isSignView()
            ? `<select class="select-aewttr" id="doc-search-signer"><option value="">Any signer</option>${distinct.signers.map((p) => `<option value="${escapeHtml(p.key)}" ${state.signer === p.key ? "selected" : ""}>${escapeHtml(p.name)}</option>`).join("")}</select>`
            : `<select class="select-aewttr" id="doc-search-reviewer"><option value="">Any reviewer</option>${distinct.people.map((p) => `<option value="${escapeHtml(p.key)}" ${state.reviewer === p.key ? "selected" : ""}>${escapeHtml(p.name)}</option>`).join("")}</select>
               <select class="select-aewttr" id="doc-search-signer"><option value="">Any signer</option>${distinct.signers.map((p) => `<option value="${escapeHtml(p.key)}" ${state.signer === p.key ? "selected" : ""}>${escapeHtml(p.name)}</option>`).join("")}</select>`}
          <select class="select-aewttr" id="doc-search-status"><option value="">Any status</option>${DOC_REVIEW_STATUSES.filter((s) => isSignView() ? s.key !== "Archived" && s.key !== "Review Complete" : true).map((s) => `<option value="${escapeHtml(s.key)}" ${state.status === s.key ? "selected" : ""}>${escapeHtml(s.key)}</option>`).join("")}</select>
          <input type="date" class="input-aewttr" id="doc-search-due" placeholder="Due before" value="${escapeHtml(state.dueBefore)}"${tip("Due before this date")}>
          <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="doc-search-clear"${tip("Clear filters")}><i class="bx bx-x"></i></button>
        </div>
        ${filtered.length || cols.some((col) => (docsByColumn[col] || []).length) ? `
          <div class="kanban-wrap">
            ${cols.map((col) => `
              <div class="kanban-col" style="cursor:default;">
                <div class="kanban-col-head"><h4${tip((DOC_REVIEW_STATUSES.find((s) => s.key === col) || {}).hint || "")}>${col}</h4><span class="kanban-col-count">${(docsByColumn[col] || []).length}</span></div>
                <div class="kanban-cards">
                  ${(docsByColumn[col] || []).map((doc) => docCardHtml(doc)).join("") || emptyColHtml()}
                </div>
              </div>
            `).join("")}
          </div>` : `
          <div class="docreview-empty-page">
            <div class="docreview-empty-page-icon"><i class="bx ${isSignView() ? "bx-pen" : "bx-file-blank"}"></i></div>
            <h3>${isSignView() ? "No documents with signers yet" : "No documents in this view"}</h3>
            <p>${isSignView()
              ? "Open a document and use Sign on a reviewer row to require their signature."
              : "Upload a document to start the review workflow, or widen scope and filters."}</p>
            <button type="button" class="btn-aewttr" id="doc-empty-upload"><i class="bx bx-upload"></i> Upload Document</button>
          </div>`}
      </div>
    `;
    $all("[data-tool-view]", $("#page-content")).forEach((button) => button.addEventListener("click", () => {
      window.AEWTTR.state.docToolView = button.dataset.toolView;
      draw();
    }));
    $all("[data-f]", $("#doc-filters")).forEach((button) => button.addEventListener("click", () => {
      if (isSignView()) window.AEWTTR.state.docSignBoardView = button.dataset.f;
      else window.AEWTTR.state.docBoardView = button.dataset.f;
      draw();
    }));
    $all("[data-scope]", $("#doc-scope-filters")).forEach((button) => button.addEventListener("click", () => {
      window.AEWTTR.state.docScope = button.dataset.scope;
      draw();
    }));
    const searchQ = $("#doc-search-q", $("#page-content"));
    if (searchQ) {
      let debounce = null;
      searchQ.addEventListener("input", () => {
        clearTimeout(debounce);
        debounce = setTimeout(() => { state.q = searchQ.value; draw(); searchQ.focus(); searchQ.setSelectionRange(searchQ.value.length, searchQ.value.length); }, 200);
      });
    }
    [
      ["doc-search-project", "project"], ["doc-search-contractor", "contractor"], ["doc-search-reviewer", "reviewer"],
      ["doc-search-signer", "signer"], ["doc-search-status", "status"], ["doc-search-due", "dueBefore"]
    ].forEach(([elId, key]) => {
      const el = $(`#${elId}`, $("#page-content"));
      if (el) el.addEventListener("change", () => { state[key] = el.value; draw(); });
    });
    const clearBtn = $("#doc-search-clear", $("#page-content"));
    if (clearBtn) clearBtn.addEventListener("click", () => {
      window.AEWTTR.state.docSearch = { q: "", project: "", contractor: "", reviewer: "", signer: "", status: "", dueBefore: "" };
      draw();
    });
    const emptyUpload = $("#doc-empty-upload", $("#page-content"));
    if (emptyUpload) emptyUpload.addEventListener("click", () => openUploadDocModal(draw));
    $all(".kanban-card", $("#page-content")).forEach((card) => {
      card.addEventListener("click", () => {
        if (isSignView()) openDocSigningModal(card.dataset.id, draw);
        else openDocReviewModal(card.dataset.id, draw);
      });
    });
  }

  const query = (typeof currentRoute === "function" ? currentRoute().query : {}) || {};
  const pending = typeof consumePendingRouteAction === "function" ? consumePendingRouteAction() : null;
  const docId = (pending && pending.doc) || query.doc;
  const mode = String((pending && pending.mode) || query.mode || "").toLowerCase();
  if (docId && getDocById(docId)) {
    window.AEWTTR.state.docToolView = mode === "sign" ? "Sign" : "Review";
  }

  draw();
  runArchiveSweep();
  // Background refresh (see wireBackgroundRefresh in app.js) updates
  // window.AEWTTR.db in place but deliberately never calls renderPage() —
  // pages have to opt in via "pulse:data-refreshed" to reflect it live.
  // Re-registers on every visit (removing the previous listener first)
  // instead of a register-once guard, so it always calls the current
  // draw() closure rather than one left over from an earlier page visit.
  if (window.AEWTTR._docReviewLiveRefreshHandler) {
    window.removeEventListener("pulse:data-refreshed", window.AEWTTR._docReviewLiveRefreshHandler);
  }
  window.AEWTTR._docReviewLiveRefreshHandler = () => {
    if (!document.querySelector(".docreview-page")) return;
    if (document.querySelector(".aewttr-modal-backdrop")) return;
    const active = document.activeElement;
    if (active && (active.tagName === "INPUT" || active.tagName === "TEXTAREA" || active.tagName === "SELECT" || active.isContentEditable)) return;
    const scroller = document.querySelector(".aewttr-content");
    const scrollTop = scroller ? scroller.scrollTop : 0;
    draw();
    if (scroller) scroller.scrollTop = scrollTop;
  };
  window.addEventListener("pulse:data-refreshed", window.AEWTTR._docReviewLiveRefreshHandler);
  $("#btn-upload-doc").addEventListener("click", () => openUploadDocModal(draw));
  if (docId && getDocById(docId)) {
    if (mode === "sign") openDocSigningModal(docId, draw);
    else openDocReviewModal(docId, draw);
  }
  } catch (e) {
    console.error("PULSE: doc review page failed", e);
    throw e;
  }
};

function openUploadDocModal(onDone) {
  const db = window.AEWTTR.db;
  const reviewers = [];
  const me = currentUserIdentity();
  let currentStep = 1;
  const TOTAL_STEPS = 3;

  const modal = openModal(`
    <div class="aewttr-modal-head docreview-modal-head-compact">
      <div class="docreview-modal-titleblock">
        <h3>Submit for review</h3>
      </div>
      <button class="aewttr-modal-close">&times;</button>
    </div>
    <div class="aewttr-modal-body docreview-submit-body-compact" style="display:flex;flex-direction:column;gap:0;padding-top:10px;">
      <div class="udm-steps">
        <div class="udm-step active" data-step="1">
          <div class="udm-step-num"><span>1</span></div>
          <span class="udm-step-label">Details</span>
        </div>
        <div class="udm-step-connector"></div>
        <div class="udm-step" data-step="2">
          <div class="udm-step-num"><span>2</span></div>
          <span class="udm-step-label">People</span>
        </div>
        <div class="udm-step-connector"></div>
        <div class="udm-step" data-step="3">
          <div class="udm-step-num"><span>3</span></div>
          <span class="udm-step-label">Review</span>
        </div>
      </div>

      <div class="udm-pages">
        <!-- Step 1: File & Details -->
        <div class="udm-page active" data-page="1">
          <div class="docreview-upload-strip" id="sd-upload-strip">
            <i class="bx bx-cloud-upload" aria-hidden="true"></i>
            <div class="docreview-upload-copy">
              <strong id="sd-file-label">Choose a file</strong>
              <span>Word, PDF, or text</span>
            </div>
            <button type="button" class="btn-aewttr-outline btn-aewttr-sm" id="sd-file-pick">Browse</button>
            <input type="file" accept="${DOC_REVIEW_UPLOAD_ACCEPT}" id="sd-file" hidden>
          </div>
          <div class="form-row"><label>Title</label><input class="input-aewttr" id="sd-title" placeholder="Document name"></div>
          <div class="form-grid-2">
            <div class="form-row"><label>Project</label>
              <select class="select-aewttr" id="sd-project">
                <option value="">General — no project</option>
                ${db.projects.map((p) => `<option value="${escapeHtml(p.id)}">${escapeHtml(p.name)}</option>`).join("")}
              </select>
            </div>
            <div class="form-row"><label>Deadline</label><input type="date" class="input-aewttr" id="sd-deadline"></div>
          </div>
          <div class="form-row"><label>Kind</label>
            <select class="select-aewttr" id="sd-kind">
              <option value="">Not specified</option>
              <option value="Government Document">Government document</option>
              <option value="Contractor Deliverable">Contractor deliverable</option>
            </select>
          </div>
          <div class="form-row" id="sd-contractor-row" hidden><label>Contractor</label>${typeof tagPickerHtml === "function" ? tagPickerHtml([], "sd-contractor", { emptyText: "No contractor set.", placeholder: "Search or add contractor…", hint: "New contractors are remembered for future documents." }) : `<input class="input-aewttr" id="sd-contractor-input" placeholder="Contractor / company">`}</div>
          <div class="form-row"><label>Portfolios</label>${typeof portfolioPickerHtml === "function" ? portfolioPickerHtml([], "sd-portfolios") : `<input class="input-aewttr" id="sd-portfolios-input" placeholder="Portfolio name…">`}</div>
          <div class="form-row"><label>Config end items</label>${typeof tagPickerHtml === "function" ? tagPickerHtml([], "sd-configenditem", { placeholder: "Search or add config end item…", emptyText: "No config end items linked.", hint: "" }) : `<input class="input-aewttr" id="sd-configenditem-input" placeholder="Config end item…">`}</div>
        </div>

        <!-- Step 2: People -->
        <div class="udm-page" data-page="2">
          <div>
            <div class="udm-section-label">Reviewers <span style="font-weight:400;text-transform:none;letter-spacing:0;font-size:10.5px;color:var(--aewttr-muted)">— you are included</span></div>
            <div class="traveler-picker traveler-picker--inline docreview-people-panel">
              <div id="sd-reviewers-selected" class="traveler-chip-list docreview-chip-list"></div>
              <input class="input-aewttr" id="sd-reviewers-input" placeholder="Add reviewer…">
              <div id="sd-reviewers-suggestions" class="traveler-suggestions"></div>
            </div>
          </div>
          <div style="margin-top:6px;">
            <div class="udm-section-label">Signers <span style="font-weight:400;text-transform:none;letter-spacing:0;font-size:10.5px;color:var(--aewttr-muted)">— optional</span></div>
            <div class="traveler-picker traveler-picker--inline docreview-people-panel">
              <div id="sd-signers-selected" class="traveler-chip-list docreview-chip-list"></div>
              <input class="input-aewttr" id="sd-signers-input" placeholder="Add signer…">
              <div id="sd-signers-suggestions" class="traveler-suggestions"></div>
            </div>
          </div>
        </div>

        <!-- Step 3: Review & Submit -->
        <div class="udm-page" data-page="3">
          <div class="udm-review-rows" id="sd-review-summary"></div>
        </div>
      </div>
    </div>
    <div class="aewttr-modal-foot docreview-modal-foot-compact">
      <button class="btn-aewttr-ghost" id="sd-back" hidden>Back</button>
      <div style="display:flex;gap:8px;align-items:center;margin-left:auto;">
        <button class="btn-aewttr-ghost" id="sd-cancel">Cancel</button>
        <button class="btn-aewttr" id="sd-next">Next</button>
      </div>
    </div>
  `, { docreview: true, className: "docreview-modal--submit" });

  reviewers.push({ id: `self-${normalizePersonKey(me.email || me.name)}`, name: me.name, email: me.email });
  wirePeoplePicker(modal, reviewers, { mount: "sd-reviewers-selected", input: "sd-reviewers-input", suggestions: "sd-reviewers-suggestions" }, {
    includeGroups: true,
    expandGroups: true
  });
  const signers = [];
  wirePeoplePicker(modal, signers, { mount: "sd-signers-selected", input: "sd-signers-input", suggestions: "sd-signers-suggestions" }, {
    includeGroups: false,
    expandGroups: false
  });

  const kindSelect = $("#sd-kind", modal);
  const contractorRow = $("#sd-contractor-row", modal);
  kindSelect.addEventListener("change", () => {
    contractorRow.hidden = kindSelect.value !== "Contractor Deliverable";
  });

  const selectedContractor = new Set();
  if (typeof wireTagPicker === "function") {
    wireTagPicker(modal, selectedContractor, "sd-contractor", {
      normalize: normalizeContractorName,
      getKnown: getKnownContractorNames,
      remember: rememberContractorNames,
      singleSelect: true
    });
  }

  const selectedPortfolios = new Set();
  if (typeof wirePortfolioPicker === "function") {
    wirePortfolioPicker(modal, selectedPortfolios, "sd-portfolios");
  }

  const selectedConfigEndItems = new Set();
  if (typeof wireTagPicker === "function") {
    wireTagPicker(modal, selectedConfigEndItems, "sd-configenditem", {
      normalize: (v) => String(v || "").trim(),
      getKnown: () => {
        const seen = new Set();
        ((window.AEWTTR && window.AEWTTR.db && window.AEWTTR.db.projects) || []).forEach((p) => {
          if (p.configEndItem) seen.add(String(p.configEndItem).trim());
        });
        return [...seen].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
      },
      remember: () => {}
    });
  }

  const fileInput = $("#sd-file", modal);
  const fileLabel = $("#sd-file-label", modal);
  const uploadStrip = $("#sd-upload-strip", modal);
  $("#sd-file-pick", modal).addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", () => {
    const file = fileInput.files[0];
    if (file) {
      fileLabel.textContent = file.name;
      uploadStrip.classList.add("has-file");
    } else {
      fileLabel.textContent = "Choose a file";
      uploadStrip.classList.remove("has-file");
    }
  });

  function setStep(n) {
    currentStep = n;
    $all(".udm-page", modal).forEach(p => p.classList.toggle("active", +p.dataset.page === n));
    $all(".udm-step", modal).forEach(s => {
      const sn = +s.dataset.step;
      s.classList.toggle("active", sn === n);
      s.classList.toggle("done", sn < n);
    });
    const backBtn = $("#sd-back", modal);
    const nextBtn = $("#sd-next", modal);
    backBtn.hidden = n === 1;
    if (n === TOTAL_STEPS) {
      nextBtn.textContent = "Upload";
    } else {
      nextBtn.textContent = "Next";
    }
  }

  function buildReviewSummary() {
    const file = fileInput.files[0];
    const title = $("#sd-title", modal).value.trim();
    const projectId = $("#sd-project", modal).value;
    const project = db.projects.find(p => p.id === projectId);
    const deadline = $("#sd-deadline", modal).value;
    const kind = kindSelect.value;
    const contractor = kind === "Contractor Deliverable" ? (Array.from(selectedContractor)[0] || "") : "";
    const flatReviewers = typeof expandPickerPeople === "function" ? expandPickerPeople(reviewers) : reviewers;
    const flatSigners = typeof expandPickerPeople === "function" ? expandPickerPeople(signers) : signers;

    function chipHtml(people) {
      if (!people.length) return `<span class="udm-review-value empty">None</span>`;
      return `<div class="udm-review-chips">${people.map(p => `<span class="udm-review-chip"><i class="bx bx-user" style="font-size:11px"></i>${escapeHtml(p.name)}</span>`).join("")}</div>`;
    }

    const portfoliosArr = Array.from(selectedPortfolios);
    const configEndItemsArr = Array.from(selectedConfigEndItems);
    function tagChipHtml(items) {
      if (!items.length) return `<span class="udm-review-value empty">None</span>`;
      return `<div class="udm-review-chips">${items.map(t => `<span class="udm-review-chip">${escapeHtml(t)}</span>`).join("")}</div>`;
    }
    const rows = [
      ["File", file ? `<span class="udm-review-file-name"><i class="bx bx-file"></i>${escapeHtml(file.name)}</span>` : `<span class="udm-review-value empty">No file selected</span>`],
      ["Title", title ? `<span class="udm-review-value">${escapeHtml(title)}</span>` : `<span class="udm-review-value empty">Not set</span>`],
      ["Project", project ? `<span class="udm-review-value">${escapeHtml(project.name)}</span>` : `<span class="udm-review-value empty">General</span>`],
      ["Deadline", deadline ? `<span class="udm-review-value">${deadline}</span>` : `<span class="udm-review-value empty">None</span>`],
      ["Kind", kind ? `<span class="udm-review-value">${escapeHtml(kind)}${contractor ? ` — ${escapeHtml(contractor)}` : ""}</span>` : `<span class="udm-review-value empty">Not specified</span>`],
      ["Portfolios", tagChipHtml(portfoliosArr)],
      ["Config end items", tagChipHtml(configEndItemsArr)],
      ["Reviewers", chipHtml(flatReviewers)],
      ["Signers", chipHtml(flatSigners)],
    ];

    $("#sd-review-summary", modal).innerHTML = rows.map(([label, val]) =>
      `<div class="udm-review-row"><span class="udm-review-label">${label}</span><span style="flex:1;min-width:0;">${val}</span></div>`
    ).join("");
  }

  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#sd-cancel", modal).addEventListener("click", closeModal);
  $("#sd-back", modal).addEventListener("click", () => setStep(currentStep - 1));

  $("#sd-next", modal).addEventListener("click", async () => {
    if (currentStep === 1) {
      const title = $("#sd-title", modal).value.trim();
      const file = fileInput.files[0];
      if (!title) { toast("Document title is required", "error"); return; }
      if (!file) { toast("Upload a document file to start the review record.", "error"); return; }
      setStep(2);
    } else if (currentStep === 2) {
      buildReviewSummary();
      setStep(3);
    } else {
      const nextBtn = $("#sd-next", modal);
      nextBtn.disabled = true;
      nextBtn.textContent = "Uploading…";
      try {
        const title = $("#sd-title", modal).value.trim();
        const file = fileInput.files[0];
        if (isSharePointMode()) {
          try { await ensureDocReviewList(window.AEWTTR.siteUrl); } catch (e) {}
        }
        const liveDb = window.AEWTTR.db;
        const flatReviewers = typeof expandPickerPeople === "function" ? expandPickerPeople(reviewers) : reviewers;
        const flatSigners = typeof expandPickerPeople === "function" ? expandPickerPeople(signers) : signers;
        const doc = normalizeDocReview({
          id: uid("D"),
          title: title,
          projectCode: $("#sd-project", modal).value || "",
          type: "Document",
          docKind: kindSelect.value || "",
          contractorName: kindSelect.value === "Contractor Deliverable" ? (Array.from(selectedContractor)[0] || "") : "",
          portfolios: Array.from(selectedPortfolios),
          configEndItems: Array.from(selectedConfigEndItems),
          submitter: (liveDb.user && liveDb.user.name) || (db.user && db.user.name) || "",
          submitterEmail: (liveDb.user && liveDb.user.email) || (db.user && db.user.email) || "",
          date: todayIsoDate(),
          deadline: $("#sd-deadline", modal).value || "",
          reviewers: flatReviewers.map((reviewer) => normalizeReviewer({ name: reviewer.name, email: reviewer.email })),
          signers: flatSigners.map((signer) => normalizeSigner({ name: signer.name, email: signer.email })),
          comments: [],
          revisions: [],
          reviewActivity: [],
          isArchived: false,
          archivedDate: "",
          fullyReviewedDate: "",
          finalPackUrl: "",
          signingSequenceLocked: false,
          _column: "Not Started"
        });
        if (typeof rememberPortfolioNames === "function" && doc.portfolios && doc.portfolios.length) {
          rememberPortfolioNames(doc.portfolios);
        }
        ensureSubmitterReviewer(doc.reviewers, doc.submitter, doc.submitterEmail);
        const initialRevision = await createRevisionFromUpload(doc, file, "Initial Upload");
        doc.revisions.unshift(initialRevision);
        setDocActiveRevision(doc, initialRevision.id);
        appendDocActivity(doc, "Uploaded", "Document uploaded for review.");
        syncDocReviewState(doc);
        moveDocBetweenBuckets(doc);
        await Repo.save("docReviewItem", doc, { column: doc._column });
        await notifyPendingReviewers(doc, {
          kind: "ready",
          subject: `PULSE Document Review — Ready for Review — ${doc.title}`,
          preview: `${doc.submitter} submitted ${doc.title} for your review`,
          body: `<p><strong>${escapeHtml(doc.submitter)}</strong> submitted <strong>${escapeHtml(doc.title)}</strong> for review.</p><p>Open PULSE → Document Review to review it.</p>`,
          excludeSubmitter: true,
          force: true
        });
        closeModal();
        toast("Document uploaded.", "success");
        if (onDone) onDone();
        if (typeof notifyLocalDataChanged === "function") notifyLocalDataChanged("doc-submit");
      } catch (e) {
        toast((e && e.friendly) || String((e && e.message) || e || "Document upload failed."), "error");
        const nextBtn = $("#sd-next", modal);
        nextBtn.disabled = false;
        nextBtn.textContent = "Upload";
      }
    }
  });
}

function openDocDetailModal(id, onDone, selectedRevisionId, options) {
  options = options || {};
  const mode = options.mode === "sign" ? "sign" : "review";
  if (mode === "sign") {
    openDocSigningModal(id, onDone);
    return;
  }
  openDocReviewModal(id, onDone, selectedRevisionId);
}

function openDocReviewModal(id, onDone, selectedRevisionId) {
  const doc = getDocById(id);
  if (!doc) {
    toast("Could not find that document record.", "error");
    return;
  }
  setDocActiveRevision(doc, selectedRevisionId || doc.activeRevisionId);
  const revision = getDocRevision(doc, doc.activeRevisionId);
  const me = currentUserIdentity();
  const canReview = canCurrentUserReviewDoc(doc);
  const canPack = canUploadFinalPack(doc);
  const canEditSigners = canEditDocSigners(doc);
  const canPublishRevision = canCurrentUserPublishDocRevision(doc);
  const canDelete = canCurrentUserDeleteDoc(doc);
  const canArchive = canCurrentUserArchiveDoc();
  const canAddReviewers = canCurrentUserAddReviewers(doc) && !doc.isArchived;
  const canDownload = canDownloadDoc(doc);
  const commentsLocked = !(canReview || canPack || (doc.reviewers || []).some((r) => samePersonByNameOrEmail(r, me)));
  const summary = docStatusSummary(doc);
  const activeTab = window.AEWTTR.state.docDetailTab === "comments" ? "overview" : (window.AEWTTR.state.docDetailTab || "overview");
  doc.comments = (doc.comments || []).map((c) => (c && c.id ? c : Object.assign({ id: uid("dc") }, c || {})));
  let editingCommentId = null;

  const project = doc.projectCode ? window.AEWTTR.db.projects.find((p) => p.id === doc.projectCode) : null;
  const inReviewPhase = doc._column === "Not Started" || doc._column === "In Review" || doc._column === "Changes Requested";
  const sortedRevisions = (doc.revisions || []).slice().sort((a, b) => (Number(b && b.number) || 0) - (Number(a && a.number) || 0));
  const currentRevision = sortedRevisions[0] || null;
  const isViewingPrevious = !!(revision && currentRevision && revision.id !== currentRevision.id);
  const _ddHasDesktop = revisionSupportsDesktopAppOpen(revision) && !!officeDesktopUriForRevision(revision);
  const _ddOpenFileBtn = revision && revision.fileUrl
    ? `<button type="button" class="btn-aewttr btn-aewttr-sm" id="dd-open-file"${tip(_ddHasDesktop ? "Open this file in Word" : sharePointOpenLabelForRevision(revision))}><i class="bx ${_ddHasDesktop ? "bx-desktop" : "bx-link-external"}"></i> ${_ddHasDesktop ? "Open in Word" : "Open"}</button>`
    : "";

  const modal = openModal(`
    <div class="aewttr-modal-head docreview-modal-head-compact">
      <div class="docreview-modal-titleblock">
        <div class="docreview-modal-kicker">Review · ${escapeHtml(doc.type || "Document")}${project ? ` · ${escapeHtml(project.name)}` : ""}${doc.docKind ? ` · ${escapeHtml(doc.docKind)}` : ""}</div>
        <h3>${escapeHtml(doc.title)}</h3>
        <div class="docreview-modal-subtitle">${escapeHtml(doc.submitter)} · ${fmtDate(doc.date)}${doc.deadline ? ` · Due ${fmtDate(doc.deadline)}` : ""}${doc.contractorName ? ` · ${escapeHtml(doc.contractorName)}` : ""}</div>
      </div>
      <div class="docreview-modal-head-actions">
        <span class="status-pill ${docColumnStatusClass(doc._column)}"${tip((DOC_REVIEW_STATUSES.find((s) => s.key === doc._column) || {}).hint || "")}>${escapeHtml(doc._column || "Not Started")}</span>
        <button class="aewttr-modal-close">&times;</button>
      </div>
    </div>
    <div class="aewttr-modal-body docreview-detail-layout">
      <div class="docreview-detail-main">
        ${renderDocWorkflowStages(doc)}
        <div class="docreview-tabbar">
          <button type="button" class="docreview-tab ${activeTab === "overview" ? "active" : ""}" data-doc-tab="overview">Overview</button>
          <button type="button" class="docreview-tab ${activeTab === "reviewers" ? "active" : ""}" data-doc-tab="reviewers">Reviewers <em id="dd-reviewers-tab-count">${summary.total}</em></button>
          <button type="button" class="docreview-tab ${activeTab === "history" ? "active" : ""}" data-doc-tab="history">History <em>${(doc.reviewActivity || []).length}</em></button>
        </div>
        <div class="docreview-tab-body">
          <div class="docreview-tab-panel ${activeTab === "overview" ? "active" : ""}" data-doc-panel="overview">
            ${(summary.stages.awaitingFinalPack || canPack) ? `
              <section class="docreview-pack-panel ${canPack ? "docreview-pack-panel--action" : ""}">
                <div class="docreview-pack-panel-copy">
                  <strong><i class="bx bx-package"></i> Final packed copy needed</strong>
                  <span>${summary.stages.awaitingFinalPack ? "Review is complete. Upload the" : "Upload the"} <em>final packed PDF</em> (not a revision) to start sequential signing.</span>
                </div>
                <div class="docreview-pack-panel-actions">
                  ${canPack ? `
                    <button type="button" class="btn-aewttr btn-aewttr-sm" id="dd-upload-final-pack"><i class="bx bx-upload"></i> Upload final PDF</button>
                    <input type="file" id="dd-final-pack-file" accept=".pdf,application/pdf" hidden>
                  ` : `<span class="docreview-hint">Document admins and the owner can upload the pack.</span>`}
                </div>
              </section>` : ""}
            <section class="docreview-file-panel${isViewingPrevious ? " is-viewing-previous" : ""}">
              <div class="docreview-file-panel-top">
                <div class="docreview-file-panel-identity">
                  <i class="bx ${isViewingPrevious ? "bx-history" : "bx-file"}" aria-hidden="true"></i>
                  <div class="docreview-file-panel-copy">
                    <strong>${escapeHtml(revision ? revisionFileLabel(revision) : "No file yet")}</strong>
                    <span>${isViewingPrevious
                      ? `Viewing older revision · current is Rev ${escapeHtml(String(currentRevision.number))}`
                      : `${summary.approved}/${summary.total} approved · ${summary.pending} pending`}</span>
                  </div>
                  ${!isViewingPrevious && revision ? `<span class="docreview-current-badge">Current</span>` : ""}
                  ${isViewingPrevious ? `<span class="docreview-viewing-badge">Rev ${escapeHtml(String(revision.number))}</span>` : ""}
                </div>
                ${sortedRevisions.length > 1 ? `
                  <div class="docreview-file-panel-rev">
                    <label class="sr-only" for="dd-rev-select">Revision</label>
                    <select class="select-aewttr docreview-rev-select" id="dd-rev-select"${tip("Switch which revision this modal shows")}>
                      ${sortedRevisions.map((item) => `
                        <option value="${escapeHtml(item.id)}" ${item.id === (revision && revision.id) ? "selected" : ""}>
                          Rev ${escapeHtml(String(item.number))}${item.id === (currentRevision && currentRevision.id) ? " · current" : ""}
                        </option>`).join("")}
                    </select>
                  </div>` : ""}
              </div>
              <div class="docreview-file-panel-actions">
                ${_ddOpenFileBtn}
                ${revision && sharePointFolderUrlForRevision(revision) ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="dd-open-location"${tip("Open this file's SharePoint folder")}><i class="bx bx-folder-open"></i> Location</button>` : ""}
                ${canPublishRevision && inReviewPhase ? `<button class="btn-aewttr-ghost btn-aewttr-sm" id="dd-upload-revision"${tip("Upload a new revision (prior file moves to the Previous Revisions folder)")}><i class="bx bx-upload"></i> New revision</button>` : ""}
              </div>
              <input type="file" id="dd-upload-revision-file" accept="${DOC_REVIEW_UPLOAD_ACCEPT}" hidden>
            </section>
            <div class="docreview-overview-grid">
              <div class="docreview-overview-stat"><span>Approved</span><strong>${summary.approved}</strong></div>
              <div class="docreview-overview-stat"><span>Pending</span><strong>${summary.pending}</strong></div>
              <div class="docreview-overview-stat"><span>Comments</span><strong id="dd-overview-comments-count">${doc.comments.length}</strong></div>
              <div class="docreview-overview-stat"><span>Revisions</span><strong>${doc.revisions.length}</strong></div>
            </div>
            <div class="docreview-kv-grid">
              ${doc.docKind ? `<div><dt>Document kind</dt><dd>${escapeHtml(doc.docKind)}</dd></div>` : ""}
              ${doc.contractorName ? `<div><dt>Contractor</dt><dd>${escapeHtml(doc.contractorName)}</dd></div>` : ""}
              ${doc.portfolios && doc.portfolios.length ? `<div><dt>Portfolios</dt><dd>${doc.portfolios.map(p => escapeHtml(p)).join(", ")}</dd></div>` : ""}
              ${doc.configEndItems && doc.configEndItems.length ? `<div><dt>Config end items</dt><dd>${doc.configEndItems.map(c => escapeHtml(c)).join(", ")}</dd></div>` : ""}
              ${!doc.portfolios?.length && !doc.configEndItems?.length && doc.portfolioOrConfigItem ? `<div><dt>Portfolio / config item</dt><dd>${escapeHtml(doc.portfolioOrConfigItem)}</dd></div>` : ""}
              ${!canDownload ? `<div class="docreview-download-note"><dt>Download</dt><dd>Restricted to Document Admins, leadership, and the current signer.</dd></div>` : ""}
            </div>
          </div>

          <div class="docreview-tab-panel ${activeTab === "reviewers" ? "active" : ""}" data-doc-panel="reviewers">
            ${isViewingPrevious ? `<p class="docreview-hint" style="margin:0 0 8px;font-size:12px;color:var(--aewttr-amber,#b45309);"><i class="bx bx-info-circle"></i> Review decisions below reflect the current revision (Rev ${escapeHtml(String(currentRevision ? currentRevision.number : "?"))}), not the older one you are previewing.</p>` : ""}
            <div class="docreview-block-head docreview-block-head--tab">
              <span>Review team</span>
              ${canAddReviewers ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="dd-toggle-add-reviewers"><i class="bx bx-user-plus"></i> Add</button>` : ""}
            </div>
            <div class="docreview-reviewer-list docreview-reviewer-list--compact docreview-scroll-panel" id="dd-reviewer-list"></div>
            ${canAddReviewers ? `
              <div id="dd-add-reviewers-panel" class="docreview-add-reviewers docreview-add-reviewers--compact" hidden>
                <div class="traveler-picker traveler-picker--inline">
                  <div id="dd-reviewers-selected" class="traveler-chip-list"></div>
                  <input class="input-aewttr" id="dd-reviewers-input" placeholder="Search people or groups…">
                  <div id="dd-reviewers-suggestions" class="traveler-suggestions"></div>
                </div>
                <div class="docreview-inline-actions">
                  <button type="button" class="btn-aewttr btn-aewttr-sm" id="dd-save-reviewers">Add</button>
                  <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="dd-cancel-add-reviewers">Cancel</button>
                </div>
              </div>` : ""}
            ${canEditSigners ? `
            <p class="docreview-hint docreview-signers-quiet-hint" style="margin-top:10px;">Use <strong>Sign</strong> on a reviewer to require their signature after review. Signing progress lives on the Sign tab.</p>` : ""}
          </div>

          <div class="docreview-tab-panel ${activeTab === "history" ? "active" : ""}" data-doc-panel="history">
            <div class="docreview-activity-list docreview-scroll-panel docreview-history-scroll">
              ${(doc.reviewActivity || []).length ? (doc.reviewActivity || []).map((item) => `
                <div class="docreview-history-item">
                  <div class="docreview-history-meta">${fmtDate(String(item.date || "").slice(0, 10))} · ${escapeHtml(item.author || "User")}</div>
                  <div class="docreview-history-text"><strong>${escapeHtml(item.action || "Update")}</strong>${item.note ? ` — ${escapeHtml(item.note)}` : ""}</div>
                </div>`).join("") : `<div class="empty-state" style="padding:12px 0;font-size:12px;">No history yet.</div>`}
            </div>
          </div>
        </div>
      </div>

      <aside class="docreview-comments-sidebar">
        <div class="docreview-comments-sidebar-head">
          <span><i class="bx bx-comment-detail"></i> Comments</span>
          <em id="dd-comments-count">${doc.comments.length}</em>
        </div>
        <div id="dd-comments" class="task-notes-chat-body docreview-comments-chat-body"></div>
        ${!commentsLocked ? `
          <div class="task-notes-input-row">
            <textarea class="task-notes-input" id="dd-comment-input" placeholder="Comment — Enter to post…" rows="1"></textarea>
            <button type="button" class="btn-aewttr btn-aewttr-sm task-notes-send" id="dd-comment-send"${tip("Post (Enter)")}><i class="bx bx-send"></i></button>
          </div>` : `<p class="docreview-hint docreview-comments-hint">${doc.isArchived ? "Archived — read only." : "Only assigned reviewers can comment."}</p>`}
      </aside>
    </div>
    <div class="aewttr-modal-foot docreview-modal-foot-compact docreview-modal-foot-review">
      <div class="docreview-action-battery">
        ${canReview && !doc.isArchived && inReviewPhase && !isViewingPrevious ? `
          <button type="button" class="docreview-battery-btn docreview-battery-btn--approve" id="dd-approve"><i class="bx bx-check"></i> Approve</button>
          <button type="button" class="docreview-battery-btn docreview-battery-btn--deny" id="dd-request-change"><i class="bx bx-x"></i> Deny</button>
          <span class="docreview-battery-sep" aria-hidden="true"></span>
        ` : (isViewingPrevious ? `<span class="docreview-hint" style="font-size:12px;padding:0 8px;">Viewing a previous revision — switch to current to submit your review.</span><span class="docreview-battery-sep" aria-hidden="true"></span>` : "")}
        ${canPack && !doc.isArchived ? `
          <button type="button" class="docreview-battery-btn docreview-battery-btn--sign" id="dd-foot-final-pack"><i class="bx bx-package"></i> Final pack</button>
          <span class="docreview-battery-sep" aria-hidden="true"></span>
        ` : ""}
        ${canArchive ? `
          <button type="button" class="docreview-battery-btn" id="dd-archive"><i class="bx bx-archive"></i> ${doc.isArchived ? "Restore" : "Archive"}</button>
        ` : ""}
        ${canDelete ? `
          <button type="button" class="docreview-battery-btn docreview-battery-btn--danger" id="dd-delete"><i class="bx bx-trash"></i> Delete</button>
        ` : ""}
        <button type="button" class="docreview-battery-btn" onclick="closeModal()"><i class="bx bx-x"></i> Close</button>
      </div>
    </div>
  `, { docreview: true, className: "docreview-modal--review docreview-modal--tabbed" });

  $all("[data-doc-tab]", modal).forEach((tabBtn) => tabBtn.addEventListener("click", () => {
    window.AEWTTR.state.docDetailTab = tabBtn.dataset.docTab;
    $all("[data-doc-tab]", modal).forEach((b) => b.classList.toggle("active", b === tabBtn));
    $all("[data-doc-panel]", modal).forEach((panel) => panel.classList.toggle("active", panel.dataset.docPanel === tabBtn.dataset.docTab));
  }));

  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  const revSelect = $("#dd-rev-select", modal);
  if (revSelect) {
    revSelect.addEventListener("change", () => {
      const nextId = revSelect.value;
      if (!nextId || nextId === (revision && revision.id)) return;
      openDocReviewModal(id, onDone, nextId);
    });
  }
  const openFileBtn = $("#dd-open-file", modal);
  if (openFileBtn) openFileBtn.addEventListener("click", () => {
    if (revisionSupportsDesktopAppOpen(revision) && officeDesktopUriForRevision(revision)) {
      openRevisionInDesktopApp(revision);
    } else {
      openRevisionInSharePoint(revision);
    }
    logDocDownload(doc, revision);
  });
  const openLocationBtn = $("#dd-open-location", modal);
  if (openLocationBtn) openLocationBtn.addEventListener("click", () => openRevisionFolderInSharePoint(revision));

  /* Reviewers render into their own in-place function (not the whole modal
     template) so add/remove feel immediate instead of a full modal reopen
     flashing the tab back to Overview-adjacent state. */
  function renderReviewers() {
    const listEl = $("#dd-reviewer-list", modal);
    if (!listEl) return;
    listEl.innerHTML = doc.reviewers.map((reviewer, index) => {
      let signerBadge = "";
      if (reviewer.isSigner && reviewer.signedAt) {
        signerBadge = `<span class="docreview-signer-chip is-signed is-static"><i class="bx bx-check-circle"></i><span>Signed</span></span>`;
      } else if (canEditSigners) {
        signerBadge = `<button type="button" class="docreview-signer-chip ${reviewer.isSigner ? "is-required" : ""}" data-toggle-signer="${index}" aria-pressed="${reviewer.isSigner ? "true" : "false"}"${tip(reviewer.isSigner ? "Remove as required signer" : "Mark as required signer")}>
          <i class="bx bx-pen"></i><span>${reviewer.isSigner ? "Signer" : "Sign"}</span>
        </button>`;
      } else if (reviewer.isSigner) {
        signerBadge = `<span class="docreview-signer-chip is-required is-static"><i class="bx bx-pen"></i><span>Signer</span></span>`;
      }
      return `
      <div class="docreview-reviewer-row docreview-reviewer-row--compact">
        ${userAvatarHtml({ name: reviewer.name, email: reviewer.email, className: "docreview-reviewer-avatar", size: 26 })}
        <div class="docreview-reviewer-copy">
          <strong>${escapeHtml(reviewer.name)}</strong>
          <span>${reviewer.reviewedAt ? `Reviewed ${fmtDate(reviewer.reviewedAt)}` : "Awaiting review"}${reviewer.isSigner && reviewer.signedAt ? ` · Signed ${fmtDate(reviewer.signedAt)}` : ""}</span>
        </div>
        <span class="status-pill ${reviewerStatusBadge(reviewer.decision)} docreview-reviewer-pill">${escapeHtml(reviewer.decision)}</span>
        ${signerBadge}
        ${canAddReviewers && doc.reviewers.length > 1 ? `<button type="button" class="docreview-reviewer-remove" data-remove-reviewer="${index}" aria-label="Remove ${escapeHtml(reviewer.name)}"${tip("Remove from review team")}><i class="bx bx-x"></i></button>` : ""}
      </div>`;
    }).join("");
    $all("[data-toggle-signer]", listEl).forEach((btn) => btn.addEventListener("click", async () => {
      const index = Number(btn.dataset.toggleSigner);
      const beforeColumn = doc._column;
      // Optimistic: mutate + re-paint the row chip before the network save.
      const reviewer = applyReviewerSignerToggle(doc, index);
      if (!reviewer) {
        toast("Could not update signer.", "error");
        return;
      }
      renderReviewers();
      try {
        await saveDocReview(doc, {
          action: reviewer.isSigner ? "Signer Assigned" : "Signer Removed",
          text: `${me.name || "A user"} ${reviewer.isSigner ? "assigned" : "removed"} ${reviewer.name} ${reviewer.isSigner ? "as a required signer" : "from required signers"}.`
        });
        moveDocBetweenBuckets(doc);
        await maybeNotifyFinalPackAfterReview(doc);
        // Status chip / pack banner may change when first/last signer is toggled.
        if (doc._column !== beforeColumn) {
          openDocReviewModal(id, onDone, doc.activeRevisionId);
        }
        if (onDone) onDone();
      } catch (e) {
        applyReviewerSignerToggle(doc, index); // revert
        renderReviewers();
        toast(String((e && e.message) || e || "Could not update signer."), "error");
      }
    }));
    $all("[data-remove-reviewer]", listEl).forEach((btn) => btn.addEventListener("click", async () => {
      const index = Number(btn.dataset.removeReviewer);
      const reviewer = doc.reviewers[index];
      if (!reviewer) return;
      const ok = await confirmDialog({
        title: "Remove reviewer",
        message: `Remove ${reviewer.name} from the review team for "${doc.title}"?`,
        confirmLabel: "Remove",
        danger: true
      });
      if (!ok) return;
      await removeReviewerFromDoc(doc, index, me.name);
      renderReviewers();
    }));
  }
  renderReviewers();

  const pendingReviewers = [];
  if (canAddReviewers) {
    const addPanel = $("#dd-add-reviewers-panel", modal);
    const toggleAddBtn = $("#dd-toggle-add-reviewers", modal);
    const reviewerPicker = wirePeoplePicker(modal, pendingReviewers, { mount: "dd-reviewers-selected", input: "dd-reviewers-input", suggestions: "dd-reviewers-suggestions" }, {
      includeGroups: true,
      expandGroups: true
    });
    if (toggleAddBtn && addPanel) {
      toggleAddBtn.addEventListener("click", () => {
        const show = addPanel.hidden;
        addPanel.hidden = !show;
        toggleAddBtn.innerHTML = show
          ? `<i class="bx bx-chevron-up"></i>`
          : `<i class="bx bx-user-plus"></i> Add`;
      });
    }
    const cancelAddBtn = $("#dd-cancel-add-reviewers", modal);
    if (cancelAddBtn && addPanel) {
      cancelAddBtn.addEventListener("click", () => {
        pendingReviewers.splice(0, pendingReviewers.length);
        reviewerPicker.refresh();
        addPanel.hidden = true;
        if (toggleAddBtn) toggleAddBtn.innerHTML = `<i class="bx bx-user-plus"></i> Add`;
      });
    }
    const saveReviewersBtn = $("#dd-save-reviewers", modal);
    if (saveReviewersBtn) {
      saveReviewersBtn.addEventListener("click", async () => {
        if (!pendingReviewers.length) {
          toast("Select at least one person to add.", "error");
          return;
        }
        saveReviewersBtn.disabled = true;
        try {
          const flat = typeof expandPickerPeople === "function" ? expandPickerPeople(pendingReviewers) : pendingReviewers;
          const added = await appendReviewersToDoc(doc, flat.slice(), me.name);
          if (!added) {
            toast("Those reviewers are already on this document.", "error");
            saveReviewersBtn.disabled = false;
            return;
          }
          toast(`Added ${added} reviewer${added === 1 ? "" : "s"}.`, "success");
          openDocReviewModal(id, onDone, doc.activeRevisionId);
          if (onDone) onDone();
        } catch (e) {
          toast(String((e && e.message) || e || "Could not add reviewers."), "error");
          saveReviewersBtn.disabled = false;
        }
      });
    }
  }

  /* Comments: persistent chat sidebar, same bubble/composer language as
     every other message-style tool in PULSE (task/subtask notes, the
     Project Notes hub, Meeting Notes) — Enter posts, Ctrl+Enter for an
     indented sub-line, own comments get inline Edit/Delete. Rendered in
     place rather than reopening the whole modal on every post/edit/delete. */
  function renderComments() {
    const chatBody = $("#dd-comments", modal);
    if (!chatBody) return;
    const comments = doc.comments;
    chatBody.innerHTML = comments.length
      ? comments.map((comment) => {
          const isMine = comment.author === me.name;
          const isEditing = editingCommentId === comment.id;
          return `
          <div class="task-notes-bubble-row ${isMine ? "mine" : ""}">
            <div class="task-notes-bubble">
              <div class="task-notes-bubble-meta">
                <strong>${escapeHtml(comment.author)}</strong><span>${fmtDate(comment.date)}</span>
                ${isMine && !isEditing ? `
                  <span class="task-notes-bubble-actions">
                    <button type="button" data-edit-comment="${comment.id}" aria-label="Edit"${tip("Edit")}><i class="bx bx-pencil"></i></button>
                    <button type="button" data-delete-comment="${comment.id}" aria-label="Delete"${tip("Delete")}><i class="bx bx-trash"></i></button>
                  </span>` : ""}
              </div>
              ${isEditing
                ? `<textarea class="task-notes-edit-input" id="dd-comment-edit-${comment.id}">${escapeHtml(comment.text)}</textarea>
                   <div class="task-notes-edit-actions">
                     <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-cancel-comment-edit>Cancel</button>
                     <button type="button" class="btn-aewttr btn-aewttr-sm" data-save-comment-edit="${comment.id}">Save</button>
                   </div>`
                : `<div class="task-notes-bubble-text">${escapeHtml(comment.text)}</div>`}
            </div>
          </div>`;
        }).join("")
      : `<div class="task-notes-empty">No comments yet.</div>`;
    if (!editingCommentId) chatBody.scrollTop = chatBody.scrollHeight;

    $all("[data-edit-comment]", chatBody).forEach((btn) => btn.addEventListener("click", () => {
      editingCommentId = btn.dataset.editComment;
      renderComments();
      const editInput = $(`#dd-comment-edit-${editingCommentId}`, chatBody);
      if (editInput) { editInput.focus(); editInput.setSelectionRange(editInput.value.length, editInput.value.length); }
    }));
    $all("[data-cancel-comment-edit]", chatBody).forEach((btn) => btn.addEventListener("click", () => {
      editingCommentId = null;
      renderComments();
    }));
    $all("[data-save-comment-edit]", chatBody).forEach((btn) => btn.addEventListener("click", async () => {
      const cid = btn.dataset.saveCommentEdit;
      const editInput = $(`#dd-comment-edit-${cid}`, chatBody);
      const text = editInput ? editInput.value.trim() : "";
      if (!text) { toast("Comment can't be empty.", "error"); return; }
      const comment = comments.find((c) => c.id === cid);
      if (comment) comment.text = text;
      editingCommentId = null;
      await saveDocReview(doc, null);
      renderComments();
      if (onDone) onDone();
    }));
    $all("[data-delete-comment]", chatBody).forEach((btn) => btn.addEventListener("click", async () => {
      const cid = btn.dataset.deleteComment;
      const ok = await confirmDialog({ title: "Delete comment", message: "Delete this comment? This cannot be undone.", confirmLabel: "Delete", danger: true });
      if (!ok) return;
      doc.comments = doc.comments.filter((c) => c.id !== cid);
      await saveDocReview(doc, null);
      renderComments();
      const countEl = $("#dd-comments-count", modal);
      if (countEl) countEl.textContent = String(doc.comments.length);
      const overviewCount = $("#dd-overview-comments-count", modal);
      if (overviewCount) overviewCount.textContent = String(doc.comments.length);
      if (onDone) onDone();
    }));
  }
  renderComments();

  const commentInput = $("#dd-comment-input", modal);
  const commentSendBtn = $("#dd-comment-send", modal);
  if (commentInput && commentSendBtn) {
    let commentPosting = false;
    wireChatComposer(commentInput, commentSendBtn, async (text) => {
      if (commentPosting) return;
      commentPosting = true;
      try {
        doc.comments.push({ id: uid("dc"), author: me.name, date: todayIsoDate(), text: text });
        await saveDocReview(doc, { action: "Comment Added", text: text });
        renderComments();
        const countEl = $("#dd-comments-count", modal);
        if (countEl) countEl.textContent = String(doc.comments.length);
        const overviewCount = $("#dd-overview-comments-count", modal);
        if (overviewCount) overviewCount.textContent = String(doc.comments.length);
        const submitterEmail = resolveReviewerEmail({ name: doc.submitter, email: doc.submitterEmail || "" });
        if (submitterEmail && !samePersonByNameOrEmail({ name: me.name, email: me.email }, { name: doc.submitter, email: doc.submitterEmail || "" })) {
          await notifyDocReviewStakeholders(doc, {
            emails: [submitterEmail],
            subject: `PULSE Document Review Comment — ${doc.title}`,
            preview: `${me.name} added a comment on ${doc.title}`,
            kind: "comment",
            body: `<p><strong>${escapeHtml(me.name)}</strong> added a comment on <strong>${escapeHtml(doc.title)}</strong>.</p><p>${escapeHtml(text)}</p>`
          });
        }
        if (onDone) onDone();
      } catch (e) {
        console.error("PULSE: comment post failed", e);
        toast(String((e && e.message) || e || "Could not post comment."), "error");
      } finally {
        commentPosting = false;
      }
    });
  }

  const approveBtn = $("#dd-approve", modal);
  if (approveBtn) approveBtn.addEventListener("click", async () => {
    approveBtn.disabled = true;
    try {
      const reviewer = ensureActorReviewerSlot(doc);
      reviewer.decision = "Approved";
      reviewer.reviewedAt = todayIsoDate();
      reviewer.note = "";
      syncDocReviewState(doc);
      await saveDocReview(doc, { action: "Approved", text: "Approved current revision." });
      moveDocBetweenBuckets(doc);
      const submitterEmail = resolveReviewerEmail({ name: doc.submitter, email: doc.submitterEmail || "" });
      if (submitterEmail) {
        await notifyDocReviewStakeholders(doc, {
          emails: [submitterEmail],
          subject: `PULSE Document Review Approved — ${doc.title}`,
          preview: `${me.name} approved ${doc.title}`,
          kind: "success",
          body: `<p><strong>${escapeHtml(me.name)}</strong> approved <strong>${escapeHtml(doc.title)}</strong>.</p>`
        });
      }
      await maybeNotifyFinalPackAfterReview(doc);
      toast("Approved.", "success");
      openDocReviewModal(id, onDone, doc.activeRevisionId);
      if (onDone) onDone();
      if (typeof renderNav === "function") renderNav();
      if (typeof refreshUserNotifications === "function") refreshUserNotifications();
    } catch (e) {
      console.error("PULSE: approve document failed", e);
      toast(String((e && e.message) || e || "Could not save your approval."), "error");
      approveBtn.disabled = false;
    }
  });

  const requestChangeBtn = $("#dd-request-change", modal);
  if (requestChangeBtn) requestChangeBtn.addEventListener("click", async () => {
    requestChangeBtn.disabled = true;
    try {
      let note = "";
      if (typeof promptDialog === "function") {
        const prompted = await promptDialog({
          title: "Deny review",
          label: "Optional note for the submitter",
          placeholder: "What needs to change?",
          confirmLabel: "Deny",
          cancelLabel: "Cancel"
        });
        if (prompted === null) {
          requestChangeBtn.disabled = false;
          return;
        }
        note = String(prompted || "").trim();
      }
      const reviewer = ensureActorReviewerSlot(doc);
      reviewer.decision = "Requested Changes";
      reviewer.reviewedAt = todayIsoDate();
      reviewer.note = note;
      syncDocReviewState(doc);
      await saveDocReview(doc, { action: "Requested Changes", text: note || "Changes requested on current revision." });
      moveDocBetweenBuckets(doc);
      const submitterEmail = resolveReviewerEmail({ name: doc.submitter, email: doc.submitterEmail || "" });
      if (submitterEmail) {
        await notifyDocReviewStakeholders(doc, {
          emails: [submitterEmail],
          subject: `PULSE Document Review Needs Changes — ${doc.title}`,
          preview: `${me.name} requested changes on ${doc.title}`,
          kind: "denied",
          body: `<p><strong>${escapeHtml(me.name)}</strong> requested changes on <strong>${escapeHtml(doc.title)}</strong>.</p>${note ? `<p>${escapeHtml(note)}</p>` : ""}`
        });
      }
      toast("Changes requested.", "success");
      openDocReviewModal(id, onDone, doc.activeRevisionId);
      if (onDone) onDone();
      if (typeof renderNav === "function") renderNav();
      if (typeof refreshUserNotifications === "function") refreshUserNotifications();
    } catch (e) {
      console.error("PULSE: request changes failed", e);
      toast(String((e && e.message) || e || "Could not save your review decision."), "error");
      requestChangeBtn.disabled = false;
    }
  });

  async function handleFinalPackUpload(file, triggerBtn) {
    if (!file) return;
    if (!isPdfFile(file)) {
      toast("Final packed copy must be a PDF.", "error");
      return;
    }
    if (triggerBtn) triggerBtn.disabled = true;
    try {
      await uploadFinalPackedCopy(doc, file, me.name);
      toast("Final pack uploaded. First signer notified.", "success");
      openDocReviewModal(id, onDone, doc.activeRevisionId);
      if (onDone) onDone();
      if (typeof renderNav === "function") renderNav();
      if (typeof refreshUserNotifications === "function") refreshUserNotifications();
    } catch (e) {
      toast(String((e && e.message) || e || "Could not upload final pack."), "error");
      if (triggerBtn) triggerBtn.disabled = false;
    }
  }

  const finalPackInput = $("#dd-final-pack-file", modal);
  const finalPackBtn = $("#dd-upload-final-pack", modal);
  const footFinalPackBtn = $("#dd-foot-final-pack", modal);
  function triggerFinalPackPick() {
    if (finalPackInput) finalPackInput.click();
    else {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".pdf,application/pdf";
      input.addEventListener("change", () => handleFinalPackUpload(input.files[0], footFinalPackBtn));
      input.click();
    }
  }
  if (finalPackBtn && finalPackInput) {
    finalPackBtn.addEventListener("click", () => finalPackInput.click());
    finalPackInput.addEventListener("change", () => handleFinalPackUpload(finalPackInput.files[0], finalPackBtn));
  }
  if (footFinalPackBtn) footFinalPackBtn.addEventListener("click", triggerFinalPackPick);

  const uploadInput = $("#dd-upload-revision-file", modal);
  const uploadBtn = $("#dd-upload-revision", modal);
  if (uploadBtn && uploadInput) {
    uploadBtn.addEventListener("click", () => uploadInput.click());
    uploadInput.addEventListener("change", async () => {
      const file = uploadInput.files[0];
      if (!file) return;
      uploadBtn.disabled = true;
      const uploadBtnHtml = uploadBtn.innerHTML;
      uploadBtn.innerHTML = `<i class="bx bx-loader-alt"></i>`;
      try {
        // New revision: keep every prior revision file (never deleted — see
        // createRevisionFromUpload, which uploads under a new -rev-N name
        // rather than overwriting), bump the revision number, reset every
        // reviewer's decision + any signatures, and restart the review from
        // whatever the new active decisions land on. Comments and history
        // are untouched by resetReviewersForNewRevision, so both carry
        // forward across revisions. After the new current file lands, prior
        // files still in the document folder are moved into the real SPO
        // "Previous revisions" subfolder so the live folder only shows current.
        const newRevision = await createRevisionFromUpload(doc, file, "Manual Revision Upload");
        resetReviewersForNewRevision(doc);
        doc.revisions.unshift(newRevision);
        setDocActiveRevision(doc, newRevision.id);
        await saveDocReview(doc, { action: "Revision Uploaded", text: `Uploaded ${newRevision.fileName}. Review restarted — all approvals and signatures reset.` });
        moveDocBetweenBuckets(doc);
        const relocate = await movePriorRevisionFilesToSharePointArchive(doc, newRevision.id);
        if (relocate.moved) await saveDocReview(doc, null);
        await notifyPendingReviewers(doc, {
          kind: "revision",
          subject: `PULSE Document Review — New Revision — ${doc.title}`,
          preview: `${me.name} uploaded a new revision of ${doc.title} — your review is needed again`,
          body: `<p><strong>${escapeHtml(me.name)}</strong> uploaded a new revision for <strong>${escapeHtml(doc.title)}</strong>.</p><p>${escapeHtml(newRevision.fileName)}</p>`,
          excludePeople: [{ name: me.name, email: me.email || "" }],
          force: true
        });
        openDocReviewModal(id, onDone, newRevision.id);
        if (onDone) onDone();
        if (typeof renderNav === "function") renderNav();
        if (typeof refreshUserNotifications === "function") refreshUserNotifications();
      } catch (e) {
        toast(String((e && e.message) || e || "Could not upload the new revision."), "error");
        uploadBtn.disabled = false;
        uploadBtn.innerHTML = uploadBtnHtml;
      }
    });
  }

  const archiveBtn = $("#dd-archive", modal);
  if (archiveBtn) archiveBtn.addEventListener("click", async () => {
    archiveBtn.disabled = true;
    try {
      doc.isArchived = !doc.isArchived;
      doc.archivedDate = doc.isArchived ? todayIsoDate() : "";
      syncDocReviewState(doc);
      await saveDocReview(doc, { action: doc.isArchived ? "Archived" : "Restored", text: "" });
      moveDocBetweenBuckets(doc);
      await notifyDocReviewStakeholders(doc, {
        emails: collectDocStakeholderEmails(doc, me.email),
        subject: `PULSE Document Review ${doc.isArchived ? "Archived" : "Restored"} — ${doc.title}`,
        preview: `${doc.title} was ${doc.isArchived ? "archived" : "restored"} by ${me.name}`,
        kind: "info",
        body: `<p><strong>${escapeHtml(doc.title)}</strong> was ${doc.isArchived ? "archived" : "restored"} by ${escapeHtml(me.name)}.</p>`
      });
      closeModal();
      if (onDone) onDone();
      if (typeof notifyLocalDataChanged === "function") notifyLocalDataChanged("doc-archive");
    } catch (e) {
      console.error("PULSE: archive document failed", e);
      toast(String((e && e.message) || e || "Could not update archive state."), "error");
      archiveBtn.disabled = false;
    }
  });

  const deleteBtn = $("#dd-delete", modal);
  if (deleteBtn) deleteBtn.addEventListener("click", async () => {
    const confirmed = await confirmDialog({
      title: "Delete document",
      message: `Delete "${doc.title}" and all of its review metadata?`,
      confirmLabel: "Delete",
      danger: true
    });
    if (!confirmed) return;
    deleteBtn.disabled = true;
    try {
      await notifyDocReviewStakeholders(doc, {
        emails: collectDocStakeholderEmails(doc, me.email),
        subject: `PULSE Document Review Deleted — ${doc.title}`,
        preview: `${doc.title} was deleted by ${me.name}`,
        kind: "info",
        body: `<p><strong>${escapeHtml(doc.title)}</strong> was deleted by ${escapeHtml(me.name)}.</p>`
      });
      // Live db — never mutate a closed-over orphan from before a refresh.
      const liveDb = window.AEWTTR.db;
      Object.keys(liveDb.docs || {}).forEach((column) => {
        liveDb.docs[column] = (liveDb.docs[column] || []).filter((entry) => entry.id !== doc.id);
      });
      await Repo.remove("docReviewItem", doc);
      closeModal();
      toast("Document review deleted.", "success");
      if (onDone) onDone();
      if (typeof notifyLocalDataChanged === "function") notifyLocalDataChanged("doc-delete");
    } catch (e) {
      console.error("PULSE: delete document failed", e);
      toast(String((e && e.message) || e || "Could not delete document."), "error");
      deleteBtn.disabled = false;
    }
  });

  // Live refresh: if someone else approves/comments/signs/adds a revision
  // while this modal is open, reopen in place with fresh data instead of
  // leaving it stale until you close and reopen it yourself. The list
  // page's own live-refresh handler deliberately skips redrawing while any
  // modal is open (see PAGE_RENDERERS.docreview), so without this the modal
  // — and everything behind it — just sat there stale. Self-unregisters
  // once the modal is no longer in the DOM instead of needing a close hook.
  if (window.AEWTTR._docDetailLiveRefreshHandler) {
    window.removeEventListener("pulse:data-refreshed", window.AEWTTR._docDetailLiveRefreshHandler);
  }
  window.AEWTTR._docDetailLiveRefreshHandler = () => {
    if (!document.querySelector(".docreview-modal--review")) {
      window.removeEventListener("pulse:data-refreshed", window.AEWTTR._docDetailLiveRefreshHandler);
      window.AEWTTR._docDetailLiveRefreshHandler = null;
      return;
    }
    const active = document.activeElement;
    if (active && (active.tagName === "INPUT" || active.tagName === "TEXTAREA" || active.tagName === "SELECT" || active.isContentEditable)) return;
    const fresh = getDocById(id);
    if (!fresh) { closeModal(); return; }
    openDocReviewModal(id, onDone, fresh.activeRevisionId);
  };
  window.addEventListener("pulse:data-refreshed", window.AEWTTR._docDetailLiveRefreshHandler);
}

function openDocSigningModal(id, onDone) {
  const doc = getDocById(id);
  if (!doc) {
    toast("Could not find that document record.", "error");
    return;
  }
  const me = currentUserIdentity();
  const canSign = canCurrentUserSignDoc(doc);
  const canDownload = canDownloadDoc(doc);
  const canArchive = canCurrentUserArchiveDoc();
  const canDelete = canCurrentUserDeleteDoc(doc);
  const summary = docStatusSummary(doc);
  const project = doc.projectCode ? window.AEWTTR.db.projects.find((p) => p.id === doc.projectCode) : null;
  const packDownloadUrl = doc.finalPackUrl || "";
  const packDownloadName = doc.finalPackFileName || "final-pack.pdf";
  const signers = docSignersOrdered(doc);
  const current = docCurrentSigner(doc);
  const stages = summary.stages;
  const showSignUpload = !!(canSign && !doc.isArchived && packDownloadUrl);

  function signerSequenceHtml() {
    if (!signers.length) {
      return `<div class="docreview-sign-empty">No required signers on this document. Mark Sign on reviewers in the Review popup first.</div>`;
    }
    return `
      <ol class="docreview-sign-sequence" aria-label="Signing sequence">
        ${signers.map((signer) => {
          const isCurrent = current && samePersonByNameOrEmail(signer, current);
          const state = signer.signedAt ? "signed" : (isCurrent ? "current" : "pending");
          const label = signer.signedAt
            ? `Signed ${fmtDate(signer.signedAt)}`
            : (isCurrent ? "Your turn" : (docHasFinalPack(doc) ? "Waiting" : "Pending"));
          const order = Number(signer.signOrder) || "";
          return `
          <li class="docreview-sign-sequence-item is-${state}">
            <span class="docreview-sign-sequence-ord">${order || "·"}</span>
            ${userAvatarHtml({ name: signer.name, email: signer.email, className: "kc-avatar docreview-sign-sequence-av", size: 32 })}
            <div class="docreview-sign-sequence-copy">
              <strong>${escapeHtml(signer.name)}</strong>
              <span>${escapeHtml(label)}</span>
            </div>
            <em class="docreview-sign-sequence-state">${escapeHtml(signer.signedAt ? "Signed" : (isCurrent ? "Now" : "Pending"))}</em>
          </li>`;
        }).join("")}
      </ol>`;
  }

  let packBody = "";
  if (stages.awaitingFinalPack || (!docHasFinalPack(doc) && stages.signatureRequired && stages.fullyReviewed)) {
    packBody = `
      <section class="docreview-sign-pack docreview-sign-pack--wait">
        <div class="docreview-sign-pack-copy">
          <strong><i class="bx bx-package"></i> Waiting on final pack</strong>
          <span>Signatures start after a Document Admin uploads the final packed PDF from the Review popup.</span>
        </div>
      </section>`;
  } else if (packDownloadUrl || stages.signingInProgress || stages.fullySigned) {
    const packActions = [];
    if (canDownload && packDownloadUrl) {
      packActions.push(`<a class="btn-aewttr-outline btn-aewttr-sm" id="ds-open-pack" href="${escapeHtml(packDownloadUrl)}" target="_blank" rel="noopener"${tip("Open packed PDF")}><i class="bx bx-link-external"></i> Open</a>`);
    } else if (!canDownload) {
      packActions.push(`<span class="docreview-hint">File access opens when it is your turn (or for admins).</span>`);
    }
    if (showSignUpload) {
      packActions.push(`<button type="button" class="btn-aewttr btn-aewttr-sm" id="ds-upload-signed"><i class="bx bx-pen"></i> Upload signed PDF</button>`);
      packActions.push(`<input type="file" id="ds-signed-pack-file" accept=".pdf,application/pdf" hidden>`);
    }
    packBody = `
      <section class="docreview-sign-pack ${stages.fullySigned ? "docreview-sign-pack--done" : ""}">
        <div class="docreview-sign-pack-copy">
          <strong><i class="bx bx-file"></i> ${escapeHtml(packDownloadName || "Packed PDF")}</strong>
          <span>${stages.fullySigned
            ? "Fully signed — all required signatures are on this package."
            : `Sequential signing · ${summary.signed} of ${summary.signerTotal}${current ? ` · current: ${escapeHtml(current.name)}` : ""}`}</span>
        </div>
        ${packActions.length ? `<div class="docreview-sign-pack-actions">${packActions.join("")}</div>` : ""}
      </section>`;
  } else if (stages.signatureRequired) {
    packBody = `
      <section class="docreview-sign-pack docreview-sign-pack--wait">
        <div class="docreview-sign-pack-copy">
          <strong><i class="bx bx-time-five"></i> Review still in progress</strong>
          <span>Signing begins after every reviewer approves and the final pack is uploaded.</span>
        </div>
      </section>`;
  } else {
    packBody = `
      <section class="docreview-sign-pack docreview-sign-pack--wait">
        <div class="docreview-sign-pack-copy">
          <strong><i class="bx bx-info-circle"></i> No signatures required</strong>
          <span>This document has no required signers.</span>
        </div>
      </section>`;
  }

  const modal = openModal(`
    <div class="aewttr-modal-head docreview-modal-head-compact">
      <div class="docreview-modal-titleblock">
        <div class="docreview-modal-kicker">Signing · ${escapeHtml(doc.type || "Document")}${project ? ` · ${escapeHtml(project.name)}` : ""}</div>
        <h3>${escapeHtml(doc.title)}</h3>
        <div class="docreview-modal-subtitle">${escapeHtml(doc.submitter)} · ${fmtDate(doc.date)}${summary.signerTotal ? ` · ${summary.signed}/${summary.signerTotal} signed` : ""}</div>
      </div>
      <div class="docreview-modal-head-actions">
        <span class="status-pill ${docColumnStatusClass(doc._column)}"${tip((DOC_REVIEW_STATUSES.find((s) => s.key === doc._column) || {}).hint || "")}>${escapeHtml(doc._column || "Not Started")}</span>
        <button type="button" class="aewttr-modal-close" aria-label="Close">&times;</button>
      </div>
    </div>
    <div class="aewttr-modal-body docreview-sign-layout">
      ${renderDocWorkflowStages(doc)}
      ${packBody}
      <section class="docreview-sign-sequence-block">
        <header class="docreview-sign-section-head">
          <span>Signing sequence</span>
          ${summary.signerTotal ? `<em>${summary.signed}/${summary.signerTotal}</em>` : ""}
        </header>
        ${signerSequenceHtml()}
        ${doc.signingSequenceLocked
          ? `<p class="docreview-hint">Order locked when the final pack was uploaded. One signer at a time.</p>`
          : (signers.length ? `<p class="docreview-hint">Order randomizes when the final pack is uploaded.</p>` : "")}
      </section>
    </div>
    <div class="aewttr-modal-foot docreview-modal-foot-compact docreview-modal-foot-signing">
      <div class="docreview-action-battery">
        ${showSignUpload ? `
          <button type="button" class="docreview-battery-btn docreview-battery-btn--sign" id="ds-foot-sign"><i class="bx bx-pen"></i> Upload signed</button>
          <span class="docreview-battery-sep" aria-hidden="true"></span>
        ` : ""}
        ${canArchive ? `
          <button type="button" class="docreview-battery-btn" id="ds-archive"><i class="bx bx-archive"></i> ${doc.isArchived ? "Restore" : "Archive"}</button>
        ` : ""}
        ${canDelete ? `
          <button type="button" class="docreview-battery-btn docreview-battery-btn--danger" id="ds-delete"><i class="bx bx-trash"></i> Delete</button>
        ` : ""}
        <button type="button" class="docreview-battery-btn" onclick="closeModal()"><i class="bx bx-x"></i> Close</button>
      </div>
    </div>
  `, { docreview: true, className: "docreview-modal--signing" });

  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);

  async function handleSignedPackUpload(file, triggerBtn) {
    if (!file) return;
    if (!isPdfFile(file)) {
      toast("Signed copy must be a PDF.", "error");
      return;
    }
    if (triggerBtn) triggerBtn.disabled = true;
    try {
      await uploadSignedPackedCopy(doc, file, me.name);
      toast(doc._column === "Signed" ? "Document fully signed." : "Signed PDF uploaded. Next signer notified.", "success");
      openDocSigningModal(id, onDone);
      if (onDone) onDone();
      if (typeof renderNav === "function") renderNav();
      if (typeof refreshUserNotifications === "function") refreshUserNotifications();
    } catch (e) {
      toast(String((e && e.message) || e || "Could not upload signed PDF."), "error");
      if (triggerBtn) triggerBtn.disabled = false;
    }
  }

  const signedInput = $("#ds-signed-pack-file", modal);
  const uploadBtn = $("#ds-upload-signed", modal);
  const footSignBtn = $("#ds-foot-sign", modal);
  function triggerSignedPick(btn) {
    if (signedInput) {
      signedInput.click();
    } else {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".pdf,application/pdf";
      input.addEventListener("change", () => handleSignedPackUpload(input.files[0], btn));
      input.click();
    }
  }
  if (uploadBtn && signedInput) {
    uploadBtn.addEventListener("click", () => triggerSignedPick(uploadBtn));
    signedInput.addEventListener("change", () => handleSignedPackUpload(signedInput.files[0], uploadBtn || footSignBtn));
  }
  if (footSignBtn) footSignBtn.addEventListener("click", () => triggerSignedPick(footSignBtn));

  const openPackLink = $("#ds-open-pack", modal);
  if (openPackLink) openPackLink.addEventListener("click", () => logDocDownload(doc, null));
  const downloadPackLink = $("#ds-download-pack", modal);
  if (downloadPackLink) downloadPackLink.addEventListener("click", () => logDocDownload(doc, null));

  const archiveBtn = $("#ds-archive", modal);
  if (archiveBtn) archiveBtn.addEventListener("click", async () => {
    archiveBtn.disabled = true;
    try {
      doc.isArchived = !doc.isArchived;
      doc.archivedDate = doc.isArchived ? todayIsoDate() : "";
      syncDocReviewState(doc);
      await saveDocReview(doc, { action: doc.isArchived ? "Archived" : "Restored", text: "" });
      moveDocBetweenBuckets(doc);
      await notifyDocReviewStakeholders(doc, {
        emails: collectDocStakeholderEmails(doc, me.email),
        subject: `PULSE Document Review ${doc.isArchived ? "Archived" : "Restored"} — ${doc.title}`,
        preview: `${doc.title} was ${doc.isArchived ? "archived" : "restored"} by ${me.name}`,
        kind: "info",
        mode: "sign",
        body: `<p><strong>${escapeHtml(doc.title)}</strong> was ${doc.isArchived ? "archived" : "restored"} by ${escapeHtml(me.name)}.</p>`
      });
      closeModal();
      if (onDone) onDone();
      if (typeof notifyLocalDataChanged === "function") notifyLocalDataChanged("doc-archive");
    } catch (e) {
      toast(String((e && e.message) || e || "Could not update archive state."), "error");
      archiveBtn.disabled = false;
    }
  });

  const deleteBtn = $("#ds-delete", modal);
  if (deleteBtn) deleteBtn.addEventListener("click", async () => {
    const confirmed = await confirmDialog({
      title: "Delete document",
      message: `Delete "${doc.title}" and all of its review metadata?`,
      confirmLabel: "Delete",
      danger: true
    });
    if (!confirmed) return;
    deleteBtn.disabled = true;
    try {
      await notifyDocReviewStakeholders(doc, {
        emails: collectDocStakeholderEmails(doc, me.email),
        subject: `PULSE Document Review Deleted — ${doc.title}`,
        preview: `${doc.title} was deleted by ${me.name}`,
        kind: "info",
        mode: "review",
        body: `<p><strong>${escapeHtml(doc.title)}</strong> was deleted by ${escapeHtml(me.name)}.</p>`
      });
      const liveDb = window.AEWTTR.db;
      Object.keys(liveDb.docs || {}).forEach((column) => {
        liveDb.docs[column] = (liveDb.docs[column] || []).filter((entry) => entry.id !== doc.id);
      });
      await Repo.remove("docReviewItem", doc);
      closeModal();
      toast("Document review deleted.", "success");
      if (onDone) onDone();
      if (typeof notifyLocalDataChanged === "function") notifyLocalDataChanged("doc-delete");
    } catch (e) {
      toast(String((e && e.message) || e || "Could not delete document."), "error");
      deleteBtn.disabled = false;
    }
  });

  if (window.AEWTTR._docSignLiveRefreshHandler) {
    window.removeEventListener("pulse:data-refreshed", window.AEWTTR._docSignLiveRefreshHandler);
  }
  window.AEWTTR._docSignLiveRefreshHandler = () => {
    if (!document.querySelector(".docreview-modal--signing")) {
      window.removeEventListener("pulse:data-refreshed", window.AEWTTR._docSignLiveRefreshHandler);
      window.AEWTTR._docSignLiveRefreshHandler = null;
      return;
    }
    const active = document.activeElement;
    if (active && (active.tagName === "INPUT" || active.tagName === "TEXTAREA" || active.tagName === "SELECT" || active.isContentEditable)) return;
    const fresh = getDocById(id);
    if (!fresh) { closeModal(); return; }
    openDocSigningModal(id, onDone);
  };
  window.addEventListener("pulse:data-refreshed", window.AEWTTR._docSignLiveRefreshHandler);
}
