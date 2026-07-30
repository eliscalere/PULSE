function projectDocumentsLabel(proj) {
  return `${proj.id} - ${proj.name}`;
}

function projectDocumentsState(projectId) {
  if (!window.AEWTTR.state.projectDocuments) window.AEWTTR.state.projectDocuments = {};
  if (!window.AEWTTR.state.projectDocuments[projectId]) {
    window.AEWTTR.state.projectDocuments[projectId] = { path: [] };
  }
  return window.AEWTTR.state.projectDocuments[projectId];
}

function formatFileSize(bytes) {
  const size = Number(bytes) || 0;
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(size < 10240 ? 1 : 0)} KB`;
  if (size < 1024 * 1024 * 1024) return `${(size / (1024 * 1024)).toFixed(size < 10485760 ? 1 : 0)} MB`;
  return `${(size / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}

function projectDocumentFileIcon(fileName) {
  const ext = String(fileName || "").toLowerCase().split(".").pop();
  if (["pdf"].includes(ext)) return "bxs-file-pdf";
  if (["doc", "docx"].includes(ext)) return "bxs-file-doc";
  if (["xls", "xlsx", "csv"].includes(ext)) return "bxs-spreadsheet";
  if (["ppt", "pptx"].includes(ext)) return "bxs-slideshow";
  if (["png", "jpg", "jpeg", "gif", "webp", "svg"].includes(ext)) return "bxs-image";
  if (["zip", "7z", "rar"].includes(ext)) return "bxs-file-archive";
  if (["txt", "md", "json", "xml", "html", "htm"].includes(ext)) return "bxs-file-txt";
  return "bxs-file";
}

function projectDocumentOpenUrl(file) {
  if (!file || !file.fileUrl) return "";
  const api = window.AEWTTR && window.AEWTTR.OfficeDesktop;
  if (api && typeof api.sharePointBrowserUrl === "function") {
    return api.sharePointBrowserUrl(file.fileUrl);
  }
  return file.fileUrl;
}

function openProjectDocumentByPolicy(fileUrl, fileName, mimeType) {
  if (!fileUrl) return;
  const api = window.AEWTTR && window.AEWTTR.OfficeDesktop;
  if (api && typeof api.openSharePointFileByPolicy === "function") {
    api.openSharePointFileByPolicy(fileUrl, fileName, mimeType);
    return;
  }
  window.open(projectDocumentOpenUrl({ fileUrl, name: fileName }), "_blank", "noopener");
}

async function resolveProjectDocumentsFolder(proj, state) {
  const siteUrl = typeof currentSiteUrl === "function" ? currentSiteUrl() : "";
  const root = await sharePointAdapter.ensureProjectDocumentsRoot(siteUrl, projectDocumentsLabel(proj));
  let folderUrl = root.projectFolderUrl;
  for (const segment of state.path || []) {
    folderUrl = combineServerRelativePath(folderUrl, segment);
  }
  return { siteUrl, root, folderUrl };
}

function combineServerRelativePath(base, ...segments) {
  return [base].concat(segments || []).map((part) => String(part || "").replace(/^\/+|\/+$/g, "")).filter(Boolean).join("/");
}

function linkedDocumentIconClass(url) {
  const u = String(url || "").toLowerCase();
  if (u.includes("sharepoint.com") || u.includes("/_layouts/")) return "bxl-microsoft";
  if (u.includes("drive.google.com") || u.includes("docs.google.com")) return "bxl-google";
  if (u.includes("confluence") || u.includes("atlassian")) return "bxl-confluence";
  return "bx-link";
}

function renderLinkedDocsList(mount, proj, canEdit, onRefresh) {
  const links = proj.linkedDocuments || [];
  mount.innerHTML = links.length ? links.map((link, i) => `
    <div class="linked-doc-row" data-link-idx="${i}">
      <i class="bx ${linkedDocumentIconClass(link.url)} linked-doc-icon"></i>
      <a class="linked-doc-name" href="${escapeHtml(link.url)}" target="_blank" rel="noopener">${escapeHtml(link.name || link.url)}</a>
      ${canEdit ? `<button class="btn-aewttr-ghost btn-aewttr-sm linked-doc-remove" data-link-idx="${i}"${tip("Remove link")}><i class="bx bx-x"></i></button>` : ""}
    </div>
  `).join("") : `<div class="linked-doc-empty">No linked files or folders yet.</div>`;
  if (canEdit) {
    $all(".linked-doc-remove", mount).forEach(btn => {
      btn.addEventListener("click", async () => {
        const idx = parseInt(btn.dataset.linkIdx, 10);
        proj.linkedDocuments.splice(idx, 1);
        await Repo.save("project", proj);
        onRefresh();
      });
    });
  }
}

function openAddLinkedDocModal(proj, onDone) {
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>Link a File or Folder</h3><button class="aewttr-modal-close" type="button">&times;</button></div>
    <div class="aewttr-modal-body">
      <div class="form-row"><label>URL</label><input class="input-aewttr" id="link-doc-url" placeholder="Paste SharePoint, Teams, or any URL…" autocomplete="off"></div>
      <div class="form-row" style="margin-bottom:0;"><label>Display name <span style="color:var(--aewttr-muted);font-weight:400;">(optional)</span></label><input class="input-aewttr" id="link-doc-name" placeholder="Leave blank to use the URL"></div>
    </div>
    <div class="aewttr-modal-foot">
      <button type="button" class="btn-aewttr-ghost" id="link-doc-cancel">Cancel</button>
      <button type="button" class="btn-aewttr" id="link-doc-save">Add Link</button>
    </div>
  `);
  const urlInput = $("#link-doc-url", modal);
  const nameInput = $("#link-doc-name", modal);
  const saveBtn = $("#link-doc-save", modal);
  const cancelBtn = $("#link-doc-cancel", modal);
  if (urlInput) urlInput.focus();

  function close() { modal.closest(".aewttr-modal-backdrop")?.remove(); }

  if (cancelBtn) cancelBtn.addEventListener("click", close);
  if (saveBtn) saveBtn.addEventListener("click", async () => {
    const url = (urlInput ? urlInput.value.trim() : "");
    if (!url) { if (urlInput) urlInput.focus(); return; }
    if (!proj.linkedDocuments) proj.linkedDocuments = [];
    proj.linkedDocuments.push({ id: uid("ld"), name: (nameInput ? nameInput.value.trim() : "") || url, url });
    saveBtn.disabled = true;
    try { await Repo.save("project", proj); } catch (e) { proj.linkedDocuments.pop(); }
    close();
    if (typeof onDone === "function") onDone();
  });
}

function drawProjectDocuments(body, proj) {
  const state = projectDocumentsState(proj.id);
  const canEdit = typeof canCurrentUserEdit === "function" ? canCurrentUserEdit() : true;

  body.innerHTML = `
    <div class="project-docs-shell">
      ${proj.sharepointFolderUrl ? `
      <div class="proj-spo-folder-banner">
        <div class="proj-spo-folder-banner-left">
          <i class="bx bxl-microsoft proj-spo-folder-icon" aria-hidden="true"></i>
          <div>
            <div class="proj-spo-folder-label">Project SharePoint folder</div>
            <div class="proj-spo-folder-url">${escapeHtml(proj.sharepointFolderUrl)}</div>
          </div>
        </div>
        <div style="display:flex;gap:8px;align-items:center;">
          <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="proj-spo-folder-refresh"><i class="bx bx-refresh"></i> Refresh</button>
          <a href="${escapeHtml(proj.sharepointFolderUrl)}" target="_blank" rel="noopener" class="btn-aewttr btn-aewttr-sm">
            <i class="bx bx-link-external"></i> Open in SharePoint
          </a>
        </div>
      </div>
      <div class="proj-spo-linked-files" id="proj-spo-linked-files">
        <div class="project-docs-loading">Loading SharePoint folder…</div>
      </div>` : canEdit ? `
      <div class="proj-spo-folder-banner proj-spo-folder-banner--unset">
        <i class="bx bx-folder-open proj-spo-folder-icon" aria-hidden="true"></i>
        <div class="proj-spo-folder-label">No SharePoint folder linked — <button type="button" class="btn-inline-link" id="proj-docs-go-settings">add one in Settings</button> to link this project's primary folder.</div>
      </div>` : ""}
      ${(proj.linkedDocuments && proj.linkedDocuments.length) || canEdit ? `
      <div class="linked-docs-section">
        <div class="linked-docs-head">
          <span class="linked-docs-title"><i class="bx bx-link-alt"></i> Linked Files &amp; Folders</span>
          ${canEdit ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="linked-doc-add"${tip("Paste a SharePoint or external URL to link")}><i class="bx bx-plus"></i> Add link</button>` : ""}
        </div>
        <div class="linked-docs-list" id="linked-docs-list"></div>
      </div>` : ""}
      <div class="project-docs-toolbar">
        <div class="project-docs-toolbar-left">
          <div class="project-docs-breadcrumb" id="project-docs-breadcrumb"></div>
          <div class="search-box project-docs-search"><i class="bx bx-search"></i><input id="project-docs-search" placeholder="Search this folder…"></div>
        </div>
        <div class="project-docs-toolbar-right">
          ${canEdit ? `<button type="button" class="btn-aewttr-outline btn-aewttr-sm" id="project-docs-new-folder"${tip("Create a new folder in this location")}><i class="bx bx-folder-plus"></i> New Folder</button>` : ""}
          ${canEdit ? `<button type="button" class="btn-aewttr btn-aewttr-sm" id="project-docs-upload"${tip("Upload files to this folder")}><i class="bx bx-upload"></i> Upload</button>` : ""}
          <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="project-docs-refresh"${tip("Refresh folder contents")}><i class="bx bx-refresh"></i></button>
          <input type="file" id="project-docs-file-input" multiple style="display:none;">
        </div>
      </div>
      <div class="project-docs-dropzone" id="project-docs-dropzone">
        <div class="project-docs-grid" id="project-docs-grid">
          <div class="project-docs-loading">Loading project documents…</div>
        </div>
      </div>
      <div class="project-docs-foot" id="project-docs-foot"></div>
    </div>
  `;

  let listing = { folders: [], files: [] };
  let searchTerm = "";

  function filteredListing() {
    const query = String(searchTerm || "").trim().toLowerCase();
    if (!query) return listing;
    return {
      folders: listing.folders.filter((item) => String(item.name || "").toLowerCase().includes(query)),
      files: listing.files.filter((item) => String(item.name || "").toLowerCase().includes(query))
    };
  }

  function breadcrumbHtml() {
    const crumbs = [{ label: "PULSE Documents", path: [] }];
    (state.path || []).forEach((segment, index) => {
      crumbs.push({ label: segment, path: state.path.slice(0, index + 1) });
    });
    return crumbs.map((crumb, index) => `
      ${index ? `<span class="project-docs-crumb-sep">/</span>` : ""}
      <button type="button" class="project-docs-crumb ${index === crumbs.length - 1 ? "is-current" : ""}" data-docs-path="${escapeHtml(JSON.stringify(crumb.path))}">${escapeHtml(crumb.label)}</button>
    `).join("");
  }

  function itemCardHtml(item) {
    if (item.type === "folder") {
      return `
        <button type="button" class="project-docs-item project-docs-item-folder" data-docs-folder="${escapeHtml(item.name)}">
          <span class="project-docs-item-icon folder"><i class="bx bxs-folder"></i></span>
          <span class="project-docs-item-name">${escapeHtml(item.name)}</span>
          <span class="project-docs-item-meta">${item.itemCount ? `${item.itemCount} item${item.itemCount === 1 ? "" : "s"}` : "Folder"}</span>
          ${canEdit ? `<span class="project-docs-item-actions"><button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-docs-delete-folder="${escapeHtml(item.serverRelativeUrl)}"${tip("Delete this folder")}><i class="bx bx-trash"></i></button></span>` : ""}
        </button>`;
    }
    const openUrl = item.fileUrl || "";
    return `
      <div class="project-docs-item project-docs-item-file">
        <button type="button" class="project-docs-item-open" data-docs-open="${escapeHtml(openUrl)}" data-docs-name="${escapeHtml(item.name || "")}" data-docs-mime="${escapeHtml(item.mimeType || "")}"${tip("Open this file")}>
          <span class="project-docs-item-icon file"><i class="bx ${projectDocumentFileIcon(item.name)}"></i></span>
          <span class="project-docs-item-name">${escapeHtml(item.name)}</span>
          <span class="project-docs-item-meta">${formatFileSize(item.size)}${item.modified ? ` · ${fmtDate(String(item.modified).slice(0, 10))}` : ""}</span>
        </button>
        <div class="project-docs-item-actions">
          <a class="btn-aewttr-ghost btn-aewttr-sm" href="${escapeHtml(openUrl)}" target="_blank" rel="noopener"${tip("Open in SharePoint")}><i class="bx bx-link-external"></i></a>
          ${canEdit ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-docs-delete-file="${escapeHtml(item.serverRelativeUrl)}"${tip("Delete this file")}><i class="bx bx-trash"></i></button>` : ""}
        </div>
      </div>`;
  }

  function renderGrid() {
    const grid = $("#project-docs-grid", body);
    const foot = $("#project-docs-foot", body);
    const breadcrumb = $("#project-docs-breadcrumb", body);
    if (breadcrumb) breadcrumb.innerHTML = breadcrumbHtml();
    const visible = filteredListing();
    const total = visible.folders.length + visible.files.length;
    if (!grid) return;
    if (!isSharePointMode()) {
      grid.innerHTML = `<div class="empty-state">Project documents are stored in the SharePoint <strong>PULSE Documents</strong> library. Open PULSE in SharePoint mode to browse and upload files here.</div>`;
      if (foot) foot.textContent = "";
      return;
    }
    grid.innerHTML = total ? `
      ${visible.folders.map(itemCardHtml).join("")}
      ${visible.files.map(itemCardHtml).join("")}
    ` : `<div class="empty-state">${searchTerm ? "No files or folders match your search." : "This folder is empty. Upload files or create a folder to get started."}</div>`;
    if (foot) {
      foot.textContent = `${visible.folders.length} folder${visible.folders.length === 1 ? "" : "s"} · ${visible.files.length} file${visible.files.length === 1 ? "" : "s"} · PULSE Documents / ${projectDocumentsLabel(proj)}${state.path.length ? ` / ${state.path.join(" / ")}` : ""}`;
    }
    wireGridEvents();
  }

  async function loadListing() {
    const grid = $("#project-docs-grid", body);
    if (grid) grid.innerHTML = `<div class="project-docs-loading">Loading project documents…</div>`;
    if (!isSharePointMode()) {
      renderGrid();
      return;
    }
    try {
      const { siteUrl, folderUrl } = await resolveProjectDocumentsFolder(proj, state);
      listing = await sharePointAdapter.listPulseDocumentsFolder(siteUrl, folderUrl);
      renderGrid();
    } catch (e) {
      if (grid) {
        grid.innerHTML = `<div class="empty-state">Could not load project documents. ${escapeHtml((e && e.friendly) || (e && e.message) || String(e))}</div>`;
      }
    }
  }

  async function uploadFiles(fileList) {
    if (!canEdit || !fileList || !fileList.length) return;
    if (!isSharePointMode()) {
      toast("Uploading project documents requires SharePoint mode.", "error");
      return;
    }
    const uploadBtn = $("#project-docs-upload", body);
    if (uploadBtn) {
      uploadBtn.disabled = true;
      uploadBtn.innerHTML = `<i class="bx bx-loader-alt bx-spin"></i> Uploading…`;
    }
    try {
      const { siteUrl, folderUrl } = await resolveProjectDocumentsFolder(proj, state);
      for (const file of Array.from(fileList)) {
        await sharePointAdapter.uploadPulseDocumentsFile(siteUrl, folderUrl, file);
      }
      if (typeof logUserAction === "function") {
        logUserAction({
          action: "Create",
          area: "Projects",
          summary: `Uploaded ${fileList.length} file${fileList.length === 1 ? "" : "s"} to ${proj.id} documents`,
          recordId: proj.id,
          route: `projects/${proj.id}/documents`
        });
      }
      toast(`Uploaded ${fileList.length} file${fileList.length === 1 ? "" : "s"}.`, "success");
      await loadListing();
    } catch (e) {
      toast((e && e.friendly) || (e && e.message) || "Upload failed.", "error");
    } finally {
      if (uploadBtn) {
        uploadBtn.disabled = false;
        uploadBtn.innerHTML = `<i class="bx bx-upload"></i> Upload`;
      }
    }
  }

  function wireGridEvents() {
    $all("[data-docs-folder]", body).forEach((button) => {
      button.addEventListener("click", (event) => {
        if (event.target.closest("[data-docs-delete-folder]")) return;
        state.path = (state.path || []).concat(button.dataset.docsFolder);
        searchTerm = "";
        const searchInput = $("#project-docs-search", body);
        if (searchInput) searchInput.value = "";
        loadListing();
      });
    });
    $all("[data-docs-open]", body).forEach((button) => {
      button.addEventListener("click", () => {
        const url = button.dataset.docsOpen;
        if (url) openProjectDocumentByPolicy(url, button.dataset.docsName, button.dataset.docsMime);
      });
    });
    $all("[data-docs-delete-file]", body).forEach((button) => {
      button.addEventListener("click", async (event) => {
        event.stopPropagation();
        const ok = await confirmDialog({
          title: "Delete file?",
          message: "This removes the file from the project folder in SharePoint.",
          confirmLabel: "Delete",
          danger: true
        });
        if (!ok) return;
        try {
          const siteUrl = typeof currentSiteUrl === "function" ? currentSiteUrl() : "";
          await sharePointAdapter.deletePulseDocumentsFile(siteUrl, button.dataset.docsDeleteFile);
          toast("File deleted.", "success");
          await loadListing();
        } catch (e) {
          toast((e && e.friendly) || (e && e.message) || "Could not delete file.", "error");
        }
      });
    });
    $all("[data-docs-delete-folder]", body).forEach((button) => {
      button.addEventListener("click", async (event) => {
        event.stopPropagation();
        const ok = await confirmDialog({
          title: "Delete folder?",
          message: "The folder must be empty before it can be deleted.",
          confirmLabel: "Delete",
          danger: true
        });
        if (!ok) return;
        try {
          const siteUrl = typeof currentSiteUrl === "function" ? currentSiteUrl() : "";
          await sharePointAdapter.deletePulseDocumentsFolder(siteUrl, button.dataset.docsDeleteFolder);
          toast("Folder deleted.", "success");
          await loadListing();
        } catch (e) {
          toast((e && e.friendly) || (e && e.message) || "Could not delete folder.", "error");
        }
      });
    });
    $all(".project-docs-crumb", body).forEach((button) => {
      button.addEventListener("click", () => {
        try {
          state.path = JSON.parse(button.dataset.docsPath || "[]");
        } catch (e) {
          state.path = [];
        }
        searchTerm = "";
        const searchInput = $("#project-docs-search", body);
        if (searchInput) searchInput.value = "";
        loadListing();
      });
    });
  }

  $("#project-docs-search", body).addEventListener("input", (event) => {
    searchTerm = event.target.value;
    renderGrid();
  });
  $("#project-docs-refresh", body).addEventListener("click", () => loadListing());
  const fileInput = $("#project-docs-file-input", body);
  const uploadBtn = $("#project-docs-upload", body);
  if (uploadBtn && fileInput) {
    uploadBtn.addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", async () => {
      await uploadFiles(fileInput.files);
      fileInput.value = "";
    });
  }
  const newFolderBtn = $("#project-docs-new-folder", body);
  if (newFolderBtn) {
    newFolderBtn.addEventListener("click", async () => {
      const name = window.prompt("New folder name");
      if (!name || !name.trim()) return;
      try {
        const { siteUrl, folderUrl } = await resolveProjectDocumentsFolder(proj, state);
        await sharePointAdapter.createPulseDocumentsSubfolder(siteUrl, folderUrl, name.trim());
        toast("Folder created.", "success");
        await loadListing();
      } catch (e) {
        toast((e && e.friendly) || (e && e.message) || "Could not create folder.", "error");
      }
    });
  }
  const dropzone = $("#project-docs-dropzone", body);
  if (dropzone && canEdit) {
    dropzone.addEventListener("dragover", (event) => {
      event.preventDefault();
      dropzone.classList.add("drag-over");
    });
    dropzone.addEventListener("dragleave", () => dropzone.classList.remove("drag-over"));
    dropzone.addEventListener("drop", async (event) => {
      event.preventDefault();
      dropzone.classList.remove("drag-over");
      await uploadFiles(event.dataTransfer && event.dataTransfer.files);
    });
  }

  const linkedDocsList = $("#linked-docs-list", body);
  if (linkedDocsList) {
    function refreshLinkedDocs() {
      renderLinkedDocsList(linkedDocsList, proj, canEdit, refreshLinkedDocs);
    }
    refreshLinkedDocs();
    const addLinkBtn = $("#linked-doc-add", body);
    if (addLinkBtn) {
      addLinkBtn.addEventListener("click", () => openAddLinkedDocModal(proj, refreshLinkedDocs));
    }
  }

  const goSettings = $("#proj-docs-go-settings", body);
  if (goSettings) goSettings.addEventListener("click", () => navigate(`projects/${projectRouteKey(proj)}/settings`));

  // Load linked SPO folder files
  const spoLinkedMount = $("#proj-spo-linked-files", body);
  if (spoLinkedMount && proj.sharepointFolderUrl) {
    async function loadSpoLinkedFiles() {
      spoLinkedMount.innerHTML = `<div class="project-docs-loading">Loading SharePoint folder…</div>`;
      if (!isSharePointMode()) {
        spoLinkedMount.innerHTML = `<div class="empty-state" style="padding:16px 0;">SharePoint mode required to browse files.</div>`;
        return;
      }
      try {
        const siteUrl = typeof currentSiteUrl === "function" ? currentSiteUrl() : "";
        // Extract server-relative path from the full SPO URL
        let serverRelPath = "";
        try { serverRelPath = new URL(proj.sharepointFolderUrl).pathname; } catch (e) { serverRelPath = proj.sharepointFolderUrl; }
        const result = await sharePointAdapter.listPulseDocumentsFolder(siteUrl, serverRelPath);
        const all = [...result.folders, ...result.files];
        if (!all.length) {
          spoLinkedMount.innerHTML = `<div class="empty-state" style="padding:16px 0;">No files found in this SharePoint folder.</div>`;
          return;
        }
        spoLinkedMount.innerHTML = `<div class="proj-spo-linked-list">${all.map(function(item) {
          if (item.type === "folder") {
            return `<a href="${escapeHtml(proj.sharepointFolderUrl)}" target="_blank" rel="noopener" class="proj-spo-linked-item">
              <i class="bx bxs-folder proj-spo-linked-icon proj-spo-linked-icon--folder"></i>
              <span class="proj-spo-linked-name">${escapeHtml(item.name)}</span>
              <span class="proj-spo-linked-meta">${item.itemCount || 0} items</span>
            </a>`;
          }
          const ext = String(item.name || "").split(".").pop().toLowerCase();
          const iconCls = ["pptx","ppt"].includes(ext) ? "bxl-microsoft" : ["docx","doc"].includes(ext) ? "bxl-microsoft" : ["xlsx","xls"].includes(ext) ? "bxl-microsoft" : "bx-file";
          const openUrl = item.fileUrl || proj.sharepointFolderUrl;
          const sizeTxt = item.size ? (item.size > 1048576 ? `${(item.size/1048576).toFixed(1)} MB` : `${Math.round(item.size/1024)} KB`) : "";
          return `<button type="button" class="proj-spo-linked-item" data-spo-file-url="${escapeHtml(openUrl)}" data-spo-file-name="${escapeHtml(item.name || "")}">
            <i class="bx ${iconCls} proj-spo-linked-icon"></i>
            <span class="proj-spo-linked-name">${escapeHtml(item.name)}</span>
            <span class="proj-spo-linked-meta">${sizeTxt}</span>
          </button>`;
        }).join("")}</div>`;
        $all("[data-spo-file-url]", spoLinkedMount).forEach((button) => {
          button.addEventListener("click", () => {
            openProjectDocumentByPolicy(button.dataset.spoFileUrl, button.dataset.spoFileName, "");
          });
        });
      } catch (e) {
        spoLinkedMount.innerHTML = `<div class="empty-state" style="padding:16px 0;">Could not load folder: ${escapeHtml((e && e.friendly) || (e && e.message) || String(e))}</div>`;
      }
    }
    loadSpoLinkedFiles();
    const refreshSpoBtn = $("#proj-spo-folder-refresh", body);
    if (refreshSpoBtn) refreshSpoBtn.addEventListener("click", loadSpoLinkedFiles);
  }

  loadListing();
}

window.drawProjectDocuments = drawProjectDocuments;
window.projectDocumentsLabel = projectDocumentsLabel;
