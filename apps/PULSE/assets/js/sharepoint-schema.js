/* AEWTTR-PULSE — SharePoint list/field schema + setup orchestration.

   These lists cover every feature currently reachable from the app's
   navigation. Finance, Tickets, Data Import, and the standalone team
   Checklists app are intentionally NOT included here — none of them are
   linked from the nav today, so they stay on localStorage until a future
   pass migrates them too. */

const SHAREPOINT_SCHEMA = {
  "PULSE App Roles": {
    description: "App-level role assignments for AEWTTR-PULSE. Active records = people associated with PULSE (AEWTTR tab).",
    fields: [
      { name: "UserEmail", type: "Text" },
      { name: "UserDisplayName", type: "Text" },
      { name: "SharePointUserId", type: "Number" },
      { name: "LoginName", type: "Text" },
      { name: "PrincipalType", type: "Number" },
      { name: "IsSiteAdmin", type: "Boolean" },
      { name: "Role", type: "Choice", choices: ["Admin", "Meeting Admin", "Finance Admin", "Document Admin", "Manager", "Member", "Viewer"] },
      { name: "JobTitle", type: "Text" },
      { name: "SpoAccess", type: "Choice", choices: ["None", "View", "Full"] },
      { name: "PowerBiAccess", type: "Choice", choices: ["None", "View", "Full"] },
      { name: "ThemeMode", type: "Choice", choices: ["Light", "Dark"] },
      { name: "NotificationPrefsJson", type: "Note" },
      { name: "HideFromMeetings", type: "Boolean" },
      { name: "IsActive", type: "Boolean" },
      { name: "Source", type: "Choice", choices: ["Manual", "SharePoint Site Users", "SharePoint File Store", "Bootstrap Admin"] },
      { name: "CreatedByEmail", type: "Text" },
      { name: "CreatedDate", type: "DateTime" },
      { name: "UpdatedByEmail", type: "Text" },
      { name: "UpdatedDate", type: "DateTime" }
    ]
  },

  "PULSE Projects": {
    description: "AEWTTR-PULSE project portfolio.",
    fields: [
      { name: "ProjectCode", type: "Text", required: true },
      { name: "ProjectStatus", type: "Choice", choices: ["Not Started", "In Progress", "At Risk", "Off Track", "Blocked", "On Hold", "Complete", "Archived"] },
      { name: "ProjectType", type: "Choice", choices: ["Project", "Subproject"] },
      { name: "ParentProjectId", type: "Number" },
      { name: "Team", type: "Choice", choices: ["PM", "Engineering", "Contracts", "Cyber", "Support", "All"] },
      { name: "Priority", type: "Choice", choices: ["Immediate", "High", "Medium", "Lower", "Stretch", "Exploratory", "Ongoing", "Document"] },
      { name: "Tools", type: "Text" },
      { name: "Stakeholders", type: "Text" },
      { name: "Description", type: "Note" },
      { name: "Scope", type: "Note", numLines: 10 },
      { name: "Objectives", type: "Note", numLines: 10 },
      { name: "CoverImageUrl", type: "Text" },
      { name: "ProjectImagesJson", type: "Note", numLines: 30 },
      { name: "HistoryJson", type: "Note", numLines: 10 },
      { name: "RisksJson", type: "Note", numLines: 10 },
      { name: "Handoff", type: "Note" },
      { name: "MeetingNotes", type: "Note" },
      { name: "PeopleJson", type: "Note", numLines: 10 },
      { name: "BoardsJson", type: "Note", numLines: 30 },
      { name: "LinkSubtaskDates", type: "Boolean" },
      { name: "LastUpdated", type: "DateTime" },
      { name: "IsArchived", type: "Boolean" },
      { name: "Pm", type: "Text" },
      { name: "Engineer", type: "Text" },
      { name: "Isso", type: "Text" },
      { name: "RangePoc", type: "Text" },
      { name: "Contract", type: "Text" },
      { name: "Contractor", type: "Text" },
      { name: "TaskOrder", type: "Text" },
      { name: "FundingType", type: "Text" },
      { name: "FiscalYear", type: "Text" },
      { name: "FundingStatus", type: "Text" },
      { name: "ChangeRequestRequired", type: "Boolean" },
      { name: "ConfigEndItem", type: "Text" },
      { name: "Program", type: "Text" },
      { name: "Product", type: "Text" },
      { name: "LocationsJson", type: "Note" },
      { name: "PortfoliosJson", type: "Note" },
      { name: "MainFolderUrl", type: "Text" },
      { name: "ExportSelectionsJson", type: "Note", numLines: 20 },
      { name: "ReportConfigJson", type: "Note", numLines: 30 },
      { name: "ProjectStartDate", type: "DateTime" },
      { name: "ProjectDueDate", type: "DateTime" },
      { name: "ProjectCompletionDate", type: "DateTime" },
      { name: "NotesJson", type: "Note", numLines: 20 },
      { name: "FundedOnContractAmount", type: "Number" },
      { name: "ReimbursableAmount", type: "Number" },
      { name: "FundingNotes", type: "Note", numLines: 6 },
      { name: "TechnicalStatus", type: "Choice", choices: ["Green", "Amber", "Red"] },
      { name: "LifecycleStatus", type: "Choice", choices: ["Planned", "Awaiting Funding", "Active", "Paused", "Complete"] }
    ]
  },

  "PULSE Risks": {
    description: "Project and portfolio risks with likelihood/impact scoring, mitigation, and meeting review state.",
    fields: [
      { name: "ProjectId", type: "Number" },
      { name: "Portfolio", type: "Text" },
      { name: "RiskName", type: "Text", required: true },
      { name: "RiskDescription", type: "Note", numLines: 10 },
      { name: "OwnerName", type: "Text" },
      { name: "OwnerEmail", type: "Text" },
      { name: "Likelihood", type: "Number" },
      { name: "Impact", type: "Number" },
      { name: "RiskRating", type: "Choice", choices: ["Green", "Amber", "Red"] },
      { name: "Category", type: "Choice", choices: ["Schedule", "Cost", "Technical", "Cyber", "Contract", "Staffing", "Operational", "External"] },
      { name: "MitigationPlan", type: "Note", numLines: 10 },
      { name: "ResponseStrategy", type: "Choice", choices: ["Accept", "Avoid", "Mitigate", "Transfer", "Escalate"] },
      { name: "DueDate", type: "DateTime" },
      { name: "RiskStatus", type: "Choice", choices: ["Open", "Monitoring", "Mitigating", "Escalated", "Closed"] },
      { name: "LastReviewedDate", type: "DateTime" },
      { name: "ReviewNotesJson", type: "Note", numLines: 20 },
      { name: "MitigationsJson", type: "Note", numLines: 20 },
      { name: "ResidualLikelihood", type: "Number" },
      { name: "ResidualImpact", type: "Number" }
    ]
  },

  "PULSE Action Items": {
    description: "Tracker tasks and checklist items (see Source field). Subtasks are stored as nested JSON in SubtasksJson.",
    fields: [
      { name: "ProjectId", type: "Number", required: true },
      { name: "ParentActionItemId", type: "Number" },
      { name: "WorkItemLevel", type: "Choice", choices: ["Task", "Subtask"] },
      { name: "Source", type: "Choice", choices: ["Tracker", "Checklist"] },
      { name: "AssignedToName", type: "Text" },
      { name: "StartDate", type: "DateTime" },
      { name: "DueDate", type: "DateTime" },
      { name: "ActionStatus", type: "Choice", choices: ["Not Started", "In Progress", "Blocked", "On Hold", "Complete", "Done", "Cancelled"] },
      { name: "Health", type: "Choice", choices: ["On Track", "At Risk", "Off Track", "Blocked", "On Hold", "Complete"] },
      { name: "BlockedReason", type: "Note", numLines: 6 },
      { name: "OnHoldReason", type: "Note", numLines: 6 },
      { name: "Notes", type: "Note" },
      { name: "SubtasksJson", type: "Note", numLines: 10 },
      { name: "MetadataJson", type: "Note", numLines: 20 },
      { name: "IsMilestone", type: "Choice", choices: ["Yes", "No"] },
      { name: "Position", type: "Number" },
      { name: "SortOrder", type: "Number" }
    ]
  },

  "PULSE Meetings": {
    description: "Weekly Meeting sessions, global and per-project.",
    fields: [
      { name: "ProjectId", type: "Number" },
      { name: "MeetingSeries", type: "Text" },
      { name: "SeriesKey", type: "Text" },
      { name: "MeetingDate", type: "DateTime" },
      { name: "AgendaJson", type: "Note", numLines: 20 },
      { name: "DocBlocksJson", type: "Note", numLines: 30 },
      { name: "Attendees", type: "Note" },
      { name: "AttendanceJson", type: "Note" },
      { name: "Notes", type: "Note" },
      { name: "DecisionsJson", type: "Note", numLines: 20 },
      { name: "ActionItemsJson", type: "Note", numLines: 20 },
      { name: "CarryForwardJson", type: "Note", numLines: 20 },
      { name: "RiskChangesJson", type: "Note", numLines: 20 },
      { name: "ProjectStatusChangesJson", type: "Note", numLines: 20 },
      { name: "ActivityJson", type: "Note", numLines: 10 },
      { name: "SessionStatus", type: "Choice", choices: ["planned", "active", "ended"] }
    ]
  },

  "PULSE Travel Requests": {
    description: "Travel requests — standard or engineering submissions.",
    fields: [
      { name: "RequestCode", type: "Text", required: true },
      { name: "RequesterName", type: "Text" },
      { name: "RequesterEmail", type: "Text" },
      { name: "FormMode", type: "Choice", choices: ["Standard", "Engineering", "Leave", "Contractor"] },
      { name: "RequestType", type: "Choice", choices: ["Standard", "Engineering", "Project Travel", "Personal Leave", "Contractor Travel"] },
      { name: "TripTitle", type: "Text" },
      { name: "TravelersJson", type: "Note", numLines: 12 },
      { name: "ProjectIdsJson", type: "Note" },
      { name: "Destination", type: "Text" },
      { name: "StartDate", type: "DateTime" },
      { name: "EndDate", type: "DateTime" },
      { name: "AllDay", type: "Boolean" },
      { name: "StartTime", type: "Text" },
      { name: "EndTime", type: "Text" },
      { name: "Purpose", type: "Note" },
      { name: "ImpactIfNotApproved", type: "Note" },
      { name: "AlternativesJson", type: "Note", numLines: 12 },
      { name: "EngineeringFormJson", type: "Note", numLines: 20 },
      { name: "EstCost", type: "Number" },
      { name: "TravelType", type: "Choice", choices: ["TDY", "Conference", "Training", "Other"] },
      { name: "TravelNotes", type: "Note" },
      { name: "ReqStatus", type: "Choice", choices: ["Draft", "Submitted", "Withdrawn", "Cancelled", "Completed"] },
      { name: "ChargeObject", type: "Text" },
      { name: "ChargeObjectStatus", type: "Choice", choices: ["Pending", "Assigned"] },
      { name: "ChargeObjectAssignedBy", type: "Text" },
      { name: "ChargeObjectAssignedAt", type: "Text" },
      { name: "CustomerConcurrenceStatus", type: "Choice", choices: ["Pending", "Concurred"] },
      { name: "CustomerConcurredBy", type: "Text" },
      { name: "CustomerConcurredAt", type: "Text" },
      { name: "RequiresConcurrence", type: "Boolean" },
      { name: "DocReviewId", type: "Text" },
      { name: "DenyReason", type: "Note" },
      { name: "ExportFileUrl", type: "Text" },
      { name: "ExportFileName", type: "Text" },
      { name: "ExportMimeType", type: "Text" }
    ]
  },

  "PULSE Travel Debriefs": {
    description: "Post-trip debriefs linked to a travel request. One row per traveler — a multi-traveler request can have one debrief each.",
    fields: [
      { name: "RequestCode", type: "Text", required: true },
      { name: "TravelerName", type: "Text" },
      { name: "TravelerEmail", type: "Text" },
      { name: "TripDates", type: "Text" },
      { name: "Systems", type: "Text" },
      { name: "Classification", type: "Choice", choices: ["Non-Contractor-Forward", "Contractor-Forward"] },
      { name: "Summary", type: "Note" },
      { name: "Followup", type: "Note" }
    ]
  },

  "PULSE Team Events": {
    description: "One-day, non-travel team events displayed on the PULSE calendar.",
    fields: [
      { name: "EventCode", type: "Text", required: true },
      { name: "EventDate", type: "DateTime", required: true },
      { name: "AllDay", type: "Boolean" },
      { name: "StartTime", type: "Text" },
      { name: "EndTime", type: "Text" },
      { name: "Location", type: "Text" },
      { name: "Description", type: "Note", numLines: 8 },
      { name: "UpdatedAt", type: "DateTime" }
    ]
  },

  "PULSE Document Review": {
    description: "Documents moving through review, concurrence, and signature.",
    fields: [
      { name: "ProjectId", type: "Number" },
      { name: "DocType", type: "Text" },
      { name: "DocKind", type: "Choice", choices: ["Government Document", "Contractor Deliverable"] },
      { name: "ContractorName", type: "Text" },
      { name: "PortfolioOrConfigItem", type: "Text" },
      { name: "PortfoliosJson", type: "Note", numLines: 5 },
      { name: "ConfigEndItemsJson", type: "Note", numLines: 5 },
      { name: "ReviewerGroupName", type: "Text" },
      { name: "OwnerName", type: "Text" },
      { name: "SubmitterName", type: "Text" },
      { name: "SubmitterEmail", type: "Text" },
      { name: "SubmittedDate", type: "DateTime" },
      { name: "DeadlineDate", type: "DateTime" },
      { name: "ReviewColumn", type: "Choice", choices: ["Not Started", "In Review", "Changes Requested", "Review Complete", "Awaiting Final Pack", "Signing in Progress", "Signed", "Archived", "Concurrence", "Final / Signed", "Awaiting Signature"] },
      { name: "RevisionLabel", type: "Text" },
      { name: "SignatureStatus", type: "Choice", choices: ["Not Started", "Pending Review", "Awaiting Final Pack", "Pending Signature", "Signed", "Archived"] },
      { name: "PendingReviewersJson", type: "Note", numLines: 10 },
      { name: "ReviewersJson", type: "Note", numLines: 10 },
      { name: "SignersJson", type: "Note", numLines: 10 },
      { name: "CommentsJson", type: "Note", numLines: 10 },
      { name: "SignedDate", type: "DateTime" },
      { name: "FileUrl", type: "Text" },
      { name: "FileName", type: "Text" },
      { name: "FileType", type: "Text" },
      { name: "ExtractedText", type: "Note", numLines: 20 },
      { name: "RevisionHistoryJson", type: "Note", numLines: 30 },
      { name: "AiSuggestionsJson", type: "Note", numLines: 30 },
      { name: "ReviewActivityJson", type: "Note", numLines: 30 },
      { name: "CurrentRevisionNumber", type: "Number" },
      { name: "ActiveRevisionId", type: "Text" },
      { name: "IsArchived", type: "Boolean" },
      { name: "ArchivedDate", type: "DateTime" },
      { name: "FullyReviewedDate", type: "DateTime" },
      { name: "FinalPackUrl", type: "Text" },
      { name: "FinalPackFileName", type: "Text" },
      { name: "FinalPackUploadedAt", type: "DateTime" },
      { name: "FinalPackUploadedBy", type: "Text" },
      { name: "SigningSequenceLocked", type: "Boolean" },
      { name: "AwaitingFinalPackNotifiedAt", type: "Text" },
      { name: "CurrentSignerNotifiedAt", type: "Text" }
    ]
  },

  "PULSE Doc Reviewer Groups": {
    description: "Global people groups (Admin-managed + project auto-groups). Used in people pickers app-wide; Document Review expands a selected group into individual reviewers.",
    fields: [
      { name: "GroupName", type: "Text", required: true },
      { name: "MembersJson", type: "Note", numLines: 10 },
      { name: "GroupSource", type: "Text" },
      { name: "ProjectId", type: "Text" },
      { name: "CreatedByEmail", type: "Text" },
      { name: "CreatedDate", type: "DateTime" }
    ]
  },

  "PULSE Audit Log": {
    description: "Append-only record of user actions across AEWTTR-PULSE. Admin access only.",
    fields: [
      { name: "ActionTime", type: "DateTime" },
      { name: "ActorEmail", type: "Text" },
      { name: "ActorName", type: "Text" },
      { name: "ActorRole", type: "Text" },
      { name: "Action", type: "Choice", choices: ["Create", "Update", "Delete", "View", "Navigate", "Approve", "Deny", "Submit", "Export", "Login", "Other"] },
      { name: "Area", type: "Choice", choices: ["Travel", "Documents", "Projects", "Weekly", "Tickets", "Users", "Admin", "Navigation", "System"] },
      { name: "Summary", type: "Text" },
      { name: "DetailJson", type: "Note", numLines: 8 },
      { name: "Route", type: "Text" },
      { name: "RecordId", type: "Text" }
    ]
  },

  "PULSE Issues": {
    description: "User-submitted PULSE issue reports, including page context, diagnostics, and a full-resolution screenshot file. Admin access only.",
    fields: [
      { name: "IssueCode", type: "Text", required: true },
      { name: "IssueType", type: "Choice", choices: ["Bug / error", "Data / saving", "Access / permissions", "Display / usability", "Feature Request", "Other"] },
      { name: "IssueStatus", type: "Choice", choices: ["New", "Triaged", "In Progress", "Resolved", "Closed"] },
      { name: "ReportedByName", type: "Text" },
      { name: "ReporterEmail", type: "Text" },
      { name: "PageTitle", type: "Text" },
      { name: "Route", type: "Text" },
      { name: "PageUrl", type: "Note", numLines: 4 },
      { name: "OccurredAt", type: "DateTime" },
      { name: "Description", type: "Note", numLines: 16 },
      { name: "ExpectedBehavior", type: "Note", numLines: 10 },
      { name: "AdditionalContext", type: "Note", numLines: 10 },
      { name: "ErrorCodesJson", type: "Note", numLines: 10 },
      { name: "LogsJson", type: "Note", numLines: 30 },
      { name: "DiagnosticsJson", type: "Note", numLines: 30 },
      { name: "ScreenshotDataUrl", type: "Note", numLines: 40 },
      { name: "ScreenshotFileUrl", type: "Note", numLines: 4 },
      { name: "ScreenshotServerRelativeUrl", type: "Note", numLines: 4 },
      { name: "ScreenshotFileName", type: "Text" },
      { name: "ScreenshotMetaJson", type: "Note", numLines: 8 },
      { name: "ResolutionNote", type: "Note", numLines: 10 }
    ]
  },

  "PULSE Notifications": {
    description: "Outbound notification queue for Power Automate. BodyHtml stores Adaptive Card JSON for Teams.",
    fields: [
      { name: "ToEmails", type: "Text" },
      { name: "Subject", type: "Text" },
      { name: "BodyHtml", type: "Note", numLines: 12 },
      { name: "EmailBodyHtml", type: "Note", numLines: 10 },
      { name: "TeamsText", type: "Note", numLines: 10 },
      { name: "NotificationArea", type: "Choice", choices: ["Travel", "Documents", "Projects", "Weekly", "Tickets", "Admin", "General"] },
      { name: "DeliveryStatus", type: "Choice", choices: ["Pending", "Sent", "Failed"] },
      { name: "DeliveryError", type: "Note" }
    ]
  },

  "PULSE Notification Config": {
    description: "Single-row settings for outbound Teams/email notifications.",
    fields: [
      { name: "TeamsEnabled", type: "Boolean" },
      { name: "AlsoSendEmail", type: "Boolean" }
    ]
  },

  "PULSE App Settings": {
    description: "Single-row, admin-managed application-wide display settings.",
    fields: [
      { name: "CuiMarkingEnabled", type: "Boolean" }
    ]
  },

  "PULSE Location Config": {
    description: "Single-row, admin-managed catalogs for locations, portfolios, contractors, and config end items.",
    fields: [
      { name: "LocationsJson", type: "Note", numLines: 10 },
      { name: "PortfolioCatalogJson", type: "Note", numLines: 10 },
      { name: "ContractorCatalogJson", type: "Note", numLines: 10 },
      { name: "ConfigEndItemCatalogJson", type: "Note", numLines: 10 },
      { name: "HideUnaffiliatedPeople", type: "Boolean" }
    ]
  }
};

function normalizeSchemaAlias(value) {
  return String(value == null ? "" : value)
    .trim()
    .toLowerCase()
    .replace(/_x[0-9a-f]{4}_/gi, "")
    .replace(/[^a-z0-9]/g, "");
}

function buildSchemaFieldAliases(name) {
  const aliases = new Set();
  const raw = String(name == null ? "" : name).trim();
  if (!raw) return aliases;
  aliases.add(raw);
  aliases.add(normalizeSchemaAlias(raw));
  return aliases;
}

function isGenericInternalName(value) {
  return /^field_\d+$/i.test(String(value || "").trim());
}

function isExpectedFieldType(actualType, expectedType) {
  const actual = String(actualType || "").trim().toLowerCase();
  const expected = String(expectedType || "").trim().toLowerCase();
  if (!expected) return true;
  if (actual === expected) return true;
  if (expected === "datetime" && (actual === "datetime" || actual === "date")) return true;
  if (expected === "note" && actual === "note") return true;
  if (expected === "text" && actual === "text") return true;
  if (expected === "number" && (actual === "number" || actual === "integer" || actual === "counter")) return true;
  if (expected === "boolean" && actual === "boolean") return true;
  if (expected === "choice" && actual === "choice") return true;
  return false;
}

function scoreFieldDefinition(field, expectedField) {
  const expectedName = String(expectedField && expectedField.name || "");
  const expectedAlias = normalizeSchemaAlias(expectedName);
  let score = 0;
  if (String(field.InternalName || "") === expectedName) score += 100;
  if (String(field.Title || "") === expectedName) score += 90;
  if (String(field.StaticName || "") === expectedName) score += 80;
  if (String(field.EntityPropertyName || "") === expectedName) score += 70;
  const aliases = [field.InternalName, field.Title, field.StaticName, field.EntityPropertyName]
    .map((value) => normalizeSchemaAlias(value))
    .filter(Boolean);
  if (aliases.includes(expectedAlias)) score += 50;
  if (isExpectedFieldType(field.TypeAsString, expectedField.type)) score += 25;
  if (!field.Hidden) score += 5;
  if (!field.ReadOnlyField) score += 5;
  if (!field.Sealed) score += 5;
  if (isGenericInternalName(field.InternalName)) score -= 15;
  return score;
}

function findSchemaFieldMatch(fieldDefinitions, expectedField) {
  const expectedAliases = buildSchemaFieldAliases(expectedField && expectedField.name);
  const expectedName = String(expectedField && expectedField.name || "");
  const candidates = (fieldDefinitions || []).filter((field) => {
    // Exact internal-name match always counts, even for a sealed/system
    // field. Otherwise exclude SharePoint's own built-in/computed columns —
    // their display name can coincidentally normalize to the same alias as
    // a field we intend to create (e.g. built-in "Created_x0020_Date" vs.
    // our own "CreatedDate", or built-in "BaseName" vs. our own "FileName").
    // Without this exclusion the matcher reports a false "type mismatch"
    // against SharePoint's system column instead of finding/creating ours.
    if (String(field.InternalName || "") === expectedName) return true;
    if (field.Sealed || field.ReadOnlyField) return false;
    return [field.InternalName, field.Title, field.StaticName, field.EntityPropertyName]
      .some((value) => expectedAliases.has(value) || expectedAliases.has(normalizeSchemaAlias(value)));
  });
  if (!candidates.length) {
    return {
      expectedName: expectedField.name,
      expectedType: expectedField.type,
      found: false,
      typeMatches: false,
      suspicious: false,
      reasons: ["missing"]
    };
  }
  const best = candidates
    .map((field) => ({ field, score: scoreFieldDefinition(field, expectedField) }))
    .sort((left, right) => right.score - left.score)[0].field;
  const typeMatches = isExpectedFieldType(best.TypeAsString, expectedField.type);
  const suspicious = isGenericInternalName(best.InternalName) || !typeMatches;
  const reasons = [];
  if (isGenericInternalName(best.InternalName)) reasons.push("generic-internal-name");
  if (!typeMatches) reasons.push(`type-mismatch:${best.TypeAsString || 'unknown'}!=${expectedField.type}`);
  return {
    expectedName: expectedField.name,
    expectedType: expectedField.type,
    found: true,
    typeMatches,
    suspicious,
    reasons,
    field: {
      InternalName: best.InternalName,
      Title: best.Title,
      StaticName: best.StaticName,
      EntityPropertyName: best.EntityPropertyName,
      TypeAsString: best.TypeAsString,
      Hidden: !!best.Hidden,
      ReadOnlyField: !!best.ReadOnlyField,
      Sealed: !!best.Sealed
    }
  };
}

async function getSetupStatus(siteUrl) {
  const entries = Object.entries(SHAREPOINT_SCHEMA);
  const results = await Promise.all(entries.map(async ([listTitle, definition]) => {
    try {
      const info = await sharePointAdapter.getListInfo(siteUrl, listTitle);
      let missingFields = [];
      let typeMismatchFields = [];
      let suspiciousFields = [];
      if (info.exists) {
        const fieldDefinitions = await sharePointAdapter.getExistingFieldDefinitions(siteUrl, listTitle);
        (definition.fields || []).forEach((expectedField) => {
          const match = findSchemaFieldMatch(fieldDefinitions, expectedField);
          if (!match.found) { missingFields.push(expectedField.name); return; }
          if (!match.typeMatches) typeMismatchFields.push(`${expectedField.name} (${match.field.TypeAsString || "?"} ≠ ${expectedField.type})`);
          if (match.suspicious) suspiciousFields.push(`${expectedField.name} → ${match.field.InternalName} [${match.reasons.join(", ")}]`);
        });
      }
      return {
        listTitle,
        exists: info.exists,
        hidden: !!info.hidden,
        serverRelativeUrl: info.serverRelativeUrl || "",
        missingFields,
        typeMismatchFields,
        suspiciousFields
      };
    } catch (e) {
      return { listTitle, exists: false, hidden: false, serverRelativeUrl: "", missingFields: [], typeMismatchFields: [], suspiciousFields: [], error: e };
    }
  }));
  return {
    ready: results.every((list) => list.exists
      && (!list.missingFields || list.missingFields.length === 0)
      && (!list.typeMismatchFields || list.typeMismatchFields.length === 0)),
    lists: results
  };
}

async function runSharePointSetup(siteUrl) {
  const log = [];

  // Migration: Rename legacy lists (missing 'PULSE ' prefix) and delete unused lists
  try {
    const listTitles = Object.keys(SHAREPOINT_SCHEMA);
    for (const title of listTitles) {
      if (title.startsWith("PULSE ")) {
        const legacyTitle = title.replace("PULSE ", "");
        const legacyInfo = await sharePointAdapter.getListInfo(siteUrl, legacyTitle);
        if (legacyInfo && legacyInfo.exists) {
          const newInfo = await sharePointAdapter.getListInfo(siteUrl, title);
          if (!newInfo || !newInfo.exists) {
            await sharePointAdapter.renameList(siteUrl, legacyTitle, title);
            log.push(`Migrated legacy list: renamed '${legacyTitle}' to '${title}'`);
          } else {
            log.push(`Warning: Could not rename legacy list '${legacyTitle}' because '${title}' already exists.`);
          }
        }
      }
    }

    const deprecated = ["PULSE Decisions"];
    for (const title of deprecated) {
      const deleted = await sharePointAdapter.deleteList(siteUrl, title);
      if (deleted) log.push(`Cleaned up deprecated list: '${title}'`);
    }
  } catch (e) {
    log.push(`ERROR during schema migration: ${e && e.friendly ? e.friendly : String(e)}`);
  }

  try {
    const listResults = await sharePointAdapter.ensureListSchema(siteUrl, SHAREPOINT_SCHEMA);
    listResults.forEach((r) => {
      log.push(`${r.listTitle}: ${r.listCreated ? "list created" : "list already existed"}`);
      r.fields.forEach((f) => {
        if (f.created) log.push(`  + column ${f.name} created (${f.method || "xml"})`);
        if (f.error) log.push(`  ! column ${f.name} FAILED: ${f.error.friendly}`);
      });
      if (r.viewFieldsAdded && r.viewFieldsAdded.length) {
        log.push(`  ~ added to default view: ${r.viewFieldsAdded.join(", ")}`);
      }
    });
  } catch (e) {
    const friendly = e && e.friendly ? e.friendly : "Setup failed before lists/columns could be created.";
    log.push(`ERROR: ${friendly}`);
    if (e && e.status) log.push(`HTTP status: ${e.status}`);
    if (e && e.cause) log.push(`Cause: ${e.cause}`);
    if (e && e.raw) log.push(`Raw response: ${e.raw}`);
    return { ok: false, log, error: e };
  }

  try {
    const appSettingsRows = await sharePointAdapter.getItems(siteUrl, "PULSE App Settings", { top: 1 });
    if (!appSettingsRows.length) {
      await sharePointAdapter.createItem(siteUrl, "PULSE App Settings", {
        Title: "Default",
        CuiMarkingEnabled: false
      });
      log.push("PULSE App Settings: seeded default marking settings");
    } else {
      log.push("PULSE App Settings: row already exists");
    }
  } catch (e) {
    log.push(`ERROR seeding App Settings: ${e && e.friendly ? e.friendly : String(e)}`);
  }

  try {
    const siteUser = await sharePointAdapter.getCurrentUserFromRest(siteUrl);
    if (typeof sharePointAdapter.bootstrapAdminIfNeeded === "function") {
      const boot = await sharePointAdapter.bootstrapAdminIfNeeded(siteUrl, siteUser);
      if (boot && boot.bootstrapped) log.push("PULSE App Roles: bootstrap Admin created for current user");
    }
    const syncResult = await sharePointAdapter.syncSiteUsersToAppRoles(siteUrl, siteUser);
    log.push(`PULSE App Roles refresh: ${syncResult.updated.length} updated, ${syncResult.skipped.length} skipped (assign people from Users → All users)`);
  } catch (e) {
    log.push(`ERROR refreshing PULSE App Roles: ${e && e.friendly ? e.friendly : String(e)}`);
  }

  // Folded in here (used to be a separate "Fix Legacy Required Columns"
  // button) so a single Run Setup always leaves the site in a state where
  // saves won't fail on a stale Required column from an older schema.
  try {
    const orphans = await findOrphanRequiredFields(siteUrl);
    if (orphans.length) {
      const repaired = await repairOrphanRequiredFields(siteUrl);
      const fixed = repaired.filter((r) => r.ok).length;
      log.push(`Legacy required columns: fixed ${fixed}/${repaired.length}`);
      repaired.filter((r) => !r.ok).forEach((r) => log.push(`  ! ${r.listTitle}: ${r.title || r.internalName} FAILED: ${r.error}`));
    } else {
      log.push("Legacy required columns: none found");
    }
  } catch (e) {
    log.push(`ERROR checking legacy required columns: ${e && e.friendly ? e.friendly : String(e)}`);
  }

  const status = await getSetupStatus(siteUrl);
  return { ok: status.ready, log, status };
}

/* Finds columns on our lists that Required-block every create/update but
   aren't part of SHAREPOINT_SCHEMA anymore — leftovers from an earlier
   version of the schema (fields renamed, merged into a *Json blob, or
   dropped entirely, e.g. the old per-project Team/Priority/Effort/
   Stakeholders/ProjectPlan fields). The app never sends a value for these,
   so SharePoint rejects the write. We don't delete the column (it may hold
   real historical data); we just stop it from being Required. */
async function findOrphanRequiredFields(siteUrl) {
  const found = [];
  for (const listTitle of Object.keys(SHAREPOINT_SCHEMA)) {
    const def = SHAREPOINT_SCHEMA[listTitle];
    let fieldDefinitions;
    try {
      fieldDefinitions = await sharePointAdapter.getExistingFieldDefinitions(siteUrl, listTitle);
    } catch (e) {
      continue;
    }
    const matchedInternalNames = new Set();
    (def.fields || []).forEach((expectedField) => {
      const match = findSchemaFieldMatch(fieldDefinitions, expectedField);
      if (match.found && match.field) matchedInternalNames.add(match.field.InternalName);
    });
    fieldDefinitions.forEach((field) => {
      if (!field.Required) return;
      if (field.Hidden || field.ReadOnlyField || field.Sealed) return;
      if (field.InternalName === "Title") return;
      if (matchedInternalNames.has(field.InternalName)) return;
      found.push({ listTitle, internalName: field.InternalName, title: field.Title });
    });
  }
  return found;
}

async function repairOrphanRequiredFields(siteUrl) {
  const orphans = await findOrphanRequiredFields(siteUrl);
  const results = [];
  for (const orphan of orphans) {
    try {
      await sharePointAdapter.setFieldRequired(siteUrl, orphan.listTitle, orphan.internalName, false);
      results.push({ ...orphan, ok: true });
    } catch (e) {
      results.push({ ...orphan, ok: false, error: (e && e.friendly) || String(e) });
    }
  }
  return results;
}

async function autoEnsureSharePointSchema(siteUrl) {
  const result = await runSharePointSetup(siteUrl);
  const status = await getSetupStatus(siteUrl);
  return { ok: status.ready, log: result.log, status, results: [] };
}

async function ensureRolesListCore(siteUrl) {
  const rolesSchema = { "PULSE App Roles": SHAREPOINT_SCHEMA["PULSE App Roles"] };
  const results = await sharePointAdapter.ensureListSchema(siteUrl, rolesSchema);
  const rolesResult = results[0] || { listCreated: false, fields: [] };
  return {
    listCreated: !!rolesResult.listCreated,
    ensuredFields: rolesResult.fields.filter((f) => f.created).map((f) => f.name),
    skippedFields: rolesResult.fields.filter((f) => !f.created && !f.error).map((f) => f.name)
  };
}

/* Creates "PULSE Document Review" (list + columns) the first time anything
   needs it. Document
   submission was failing with a generic "SharePoint couldn't find the list
   or field" 404 on any tenant where an admin hadn't run full setup yet (or
   setup had only partially completed); this makes the list exist before the
   submit flow ever touches it, in the background at boot AND defensively
   right before a submit. Safe to call repeatedly — no-ops once ready. */
async function ensureDocReviewList(siteUrl) {
  const docSchema = { "PULSE Document Review": SHAREPOINT_SCHEMA["PULSE Document Review"] };
  const results = await sharePointAdapter.ensureListSchema(siteUrl, docSchema);
  const listResult = results[0] || { listCreated: false, fields: [] };
  return {
    listCreated: !!listResult.listCreated,
    ensuredFields: listResult.fields.filter((f) => f.created).map((f) => f.name)
  };
}

/* Same self-healing pattern, for the reusable
   reviewer-groups list — created on demand so "New group" never 404s on a
   tenant where full setup hasn't run yet. */
async function ensureDocReviewerGroupsList(siteUrl) {
  const groupSchema = { "PULSE Doc Reviewer Groups": SHAREPOINT_SCHEMA["PULSE Doc Reviewer Groups"] };
  const results = await sharePointAdapter.ensureListSchema(siteUrl, groupSchema);
  const listResult = results[0] || { listCreated: false, fields: [] };
  return {
    listCreated: !!listResult.listCreated,
    ensuredFields: listResult.fields.filter((f) => f.created).map((f) => f.name)
  };
}

/* Same pattern for "PULSE Risks" — the risk matrix/register/review-board on
   the Risks tab (and its compact copy on Project Home) all write here via
   Repo.save("risk", ...). Without this, "Save Risk" on a fresh or
   partially-set-up tenant hits the same generic 404 Document Review used to
   hit before ensureDocReviewList existed. Cached after the first success so
   inline risk creates don't re-probe list schema on every click. */
let _risksListReady = false;
let _risksListPromise = null;
async function ensureRisksList(siteUrl) {
  if (_risksListReady) {
    return { listCreated: false, ensuredFields: [] };
  }
  if (_risksListPromise) return _risksListPromise;
  _risksListPromise = (async () => {
    const riskSchema = { "PULSE Risks": SHAREPOINT_SCHEMA["PULSE Risks"] };
    const results = await sharePointAdapter.ensureListSchema(siteUrl, riskSchema);
    const listResult = results[0] || { listCreated: false, fields: [] };
    _risksListReady = true;
    return {
      listCreated: !!listResult.listCreated,
      ensuredFields: listResult.fields.filter((f) => f.created).map((f) => f.name)
    };
  })().finally(() => { _risksListPromise = null; });
  return _risksListPromise;
}
