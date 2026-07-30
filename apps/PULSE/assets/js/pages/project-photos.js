/* Project Photos — a flat gallery for uploading and viewing images against a
   project. Rides the same "PULSE Documents" library used by regular project
   documents and the cover image, just in a dedicated "_photos" subfolder, so
   no new SharePoint list/schema is needed. */

const PHOTO_EXTENSIONS = ["png", "jpg", "jpeg", "gif", "webp", "svg", "bmp", "heic"];

function isPhotoFile(name) {
  const ext = String(name || "").toLowerCase().split(".").pop();
  return PHOTO_EXTENSIONS.includes(ext);
}

async function resolveProjectPhotosFolder(proj) {
  const siteUrl = typeof currentSiteUrl === "function" ? currentSiteUrl() : "";
  const root = await sharePointAdapter.ensureProjectDocumentsRoot(siteUrl, projectDocumentsLabel(proj));
  const folder = await sharePointAdapter.createPulseDocumentsSubfolder(siteUrl, root.projectFolderUrl, "_photos");
  return { siteUrl, folderUrl: folder.serverRelativeUrl };
}

/* The one way to ask "what photos does this project have".

   Cover images upload to _cover/ and project photos to _photos/, so a cover
   picked in project settings was invisible everywhere that showed photos —
   the Photos tab and the End Item Config picker both listed _photos alone.
   New cover uploads now also land in _photos, but covers uploaded before that
   only exist under _cover, so the cover is folded in here whenever the folder
   listing does not already contain it. That fixes existing projects with no
   migration and no duplicate entry once the file is in both places. */
async function listProjectPhotos(proj) {
  const { siteUrl, folderUrl } = await resolveProjectPhotosFolder(proj);
  const listing = await sharePointAdapter.listPulseDocumentsFolder(siteUrl, folderUrl);
  const photos = (listing.files || []).filter((file) => isPhotoFile(file.name));
  const cover = proj && proj.coverImage;
  if (cover && !photos.some((photo) => photo.fileUrl === cover)) {
    photos.unshift({
      name: String(cover).split("/").pop() || "cover",
      fileUrl: cover,
      serverRelativeUrl: "",
      isCover: true
    });
  }
  return photos;
}

function photoCardHtml(photo, selectedUrls) {
  const selected = Array.isArray(selectedUrls) && selectedUrls.includes(photo.fileUrl);
  return `
    <div class="project-photos-item${selected ? " is-export-selected" : ""}" data-photo-url="${escapeHtml(photo.fileUrl)}" data-photo-name="${escapeHtml(photo.name)}">
      <button type="button" class="project-photos-thumb-btn" data-photo-open="${escapeHtml(photo.fileUrl)}"${tip(photo.name)}>
        <img class="project-photos-thumb" src="${escapeHtml(photo.fileUrl)}" alt="${escapeHtml(photo.name)}" loading="lazy">
      </button>
      <div class="project-photos-item-foot">
        <span class="project-photos-item-name">${escapeHtml(photo.name)}</span>
        <div class="project-photos-item-actions">
          <button type="button" class="btn-aewttr-ghost btn-aewttr-sm project-photos-export-toggle${selected ? " is-on" : ""}" data-photo-export="${escapeHtml(photo.fileUrl)}"${tip(selected ? "Remove from status export" : "Include in status PowerPoint")}>
            <i class="bx ${selected ? "bxs-check-square" : "bx-square"}" aria-hidden="true"></i>
          </button>
          ${photo.serverRelativeUrl
            ? `<button type="button" class="btn-aewttr-ghost btn-aewttr-sm" data-photo-delete="${escapeHtml(photo.serverRelativeUrl)}"${tip("Delete this photo")}><i class="bx bx-trash"></i></button>`
            /* A cover folded in from _cover/ has no file in this folder to
               delete — it is removed from project settings instead. */
            : `<span class="project-photos-cover-tag"${tip("This is the project cover image — change it in Project settings")}>Cover</span>`}
        </div>
      </div>
    </div>`;
}

function openPhotoLightbox(photo) {
  const modal = openModal(`
    <div class="aewttr-modal-head">
      <h3>${escapeHtml(photo.name)}</h3>
      <button class="aewttr-modal-close">&times;</button>
    </div>
    <div class="aewttr-modal-body">
      <div class="project-photos-lightbox-toolbar">
        <a class="btn-aewttr-outline btn-aewttr-sm" href="${escapeHtml(photo.fileUrl)}" target="_blank" rel="noopener"><i class="bx bx-link-external"></i> Open In SharePoint</a>
      </div>
      <img class="project-photos-lightbox-img" src="${escapeHtml(photo.fileUrl)}" alt="${escapeHtml(photo.name)}">
    </div>
  `, { xwide: true });
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
}

function drawProjectPhotos(body, proj) {
  const canEdit = typeof canCurrentUserEdit === "function" ? canCurrentUserEdit() : true;

  body.innerHTML = `
    <div class="project-photos-shell">
      <div class="project-docs-toolbar">
        <div class="project-docs-toolbar-left">
          <div class="project-photos-heading">Photos</div>
        </div>
        <div class="project-docs-toolbar-right">
          ${canEdit ? `<button type="button" class="btn-aewttr btn-aewttr-sm" id="project-photos-upload"${tip("Upload photos")}><i class="bx bx-upload"></i> Upload</button>` : ""}
          <button type="button" class="btn-aewttr-ghost btn-aewttr-sm" id="project-photos-refresh"${tip("Refresh photos")}><i class="bx bx-refresh"></i></button>
          <input type="file" id="project-photos-file-input" accept="image/*" multiple style="display:none;">
        </div>
      </div>
      <div class="project-docs-dropzone project-photos-dropzone" id="project-photos-dropzone">
        <div class="project-photos-grid" id="project-photos-grid">
          <div class="project-docs-loading">Loading photos…</div>
        </div>
      </div>
      <div class="project-docs-foot" id="project-photos-foot"></div>
    </div>
  `;

  let photos = [];

  function exportSelectedUrls() {
    if (!proj.exportSelections || typeof proj.exportSelections !== "object") proj.exportSelections = {};
    if (!Array.isArray(proj.exportSelections.photos)) proj.exportSelections.photos = [];
    return proj.exportSelections.photos;
  }

  function renderGrid() {
    const grid = $("#project-photos-grid", body);
    const foot = $("#project-photos-foot", body);
    if (!grid) return;
    if (!isSharePointMode()) {
      grid.innerHTML = `<div class="empty-state">Photos are stored in the SharePoint <strong>PULSE Documents</strong> library. Open PULSE in SharePoint mode to upload and view them here.</div>`;
      if (foot) foot.textContent = "";
      return;
    }
    const selected = exportSelectedUrls();
    grid.innerHTML = photos.length
      ? photos.map((photo) => photoCardHtml(photo, selected)).join("")
      : `<div class="empty-state">No photos yet. Upload images or drop them here to get started.</div>`;
    const selectedCount = selected.filter((url) => photos.some((p) => p.fileUrl === url)).length;
    if (foot) {
      foot.textContent = `${photos.length} photo${photos.length === 1 ? "" : "s"}${selectedCount ? ` · ${selectedCount} selected for status export` : ""}`;
    }
    wireGridEvents();
  }

  async function loadPhotos() {
    const grid = $("#project-photos-grid", body);
    if (grid) grid.innerHTML = `<div class="project-docs-loading">Loading photos…</div>`;
    if (!isSharePointMode()) {
      renderGrid();
      return;
    }
    try {
      photos = await listProjectPhotos(proj);
      renderGrid();
    } catch (e) {
      if (grid) grid.innerHTML = `<div class="empty-state">Could not load photos. ${escapeHtml((e && e.friendly) || (e && e.message) || String(e))}</div>`;
    }
  }

  async function uploadPhotos(fileList) {
    if (!canEdit || !fileList || !fileList.length) return;
    if (!isSharePointMode()) {
      toast("Uploading photos requires SharePoint mode.", "error");
      return;
    }
    const files = Array.from(fileList).filter((file) => isPhotoFile(file.name) || String(file.type || "").startsWith("image/"));
    if (!files.length) {
      toast("Choose image files to upload.", "error");
      return;
    }
    const uploadBtn = $("#project-photos-upload", body);
    if (uploadBtn) {
      uploadBtn.disabled = true;
      uploadBtn.innerHTML = `<i class="bx bx-loader-alt bx-spin"></i> Uploading…`;
    }
    try {
      const { siteUrl, folderUrl } = await resolveProjectPhotosFolder(proj);
      for (const file of files) {
        await sharePointAdapter.uploadPulseDocumentsFile(siteUrl, folderUrl, file);
      }
      if (typeof logUserAction === "function") {
        logUserAction({
          action: "Create",
          area: "Projects",
          summary: `Uploaded ${files.length} photo${files.length === 1 ? "" : "s"} to ${proj.id}`,
          recordId: proj.id,
          route: `projects/${proj.id}/photos`
        });
      }
      toast(`Uploaded ${files.length} photo${files.length === 1 ? "" : "s"}.`, "success");
      await loadPhotos();
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
    $all("[data-photo-open]", body).forEach((btn) => btn.addEventListener("click", () => {
      const photo = photos.find((p) => p.fileUrl === btn.dataset.photoOpen);
      if (photo) openPhotoLightbox(photo);
    }));
    $all("[data-photo-export]", body).forEach((btn) => btn.addEventListener("click", async (event) => {
      event.stopPropagation();
      if (!canEdit) return;
      const url = btn.dataset.photoExport;
      const selected = exportSelectedUrls();
      const idx = selected.indexOf(url);
      if (idx >= 0) selected.splice(idx, 1);
      else selected.push(url);
      proj.exportSelections = proj.exportSelections || {};
      proj.exportSelections.photos = selected;
      try {
        await Repo.save("project", proj);
        renderGrid();
        toast(idx >= 0 ? "Removed from status export" : "Marked for status export", "success");
      } catch (e) {
        toast((e && e.friendly) || (e && e.message) || "Could not save export selection.", "error");
      }
    }));
    $all("[data-photo-delete]", body).forEach((btn) => btn.addEventListener("click", async (event) => {
      event.stopPropagation();
      const ok = await confirmDialog({
        title: "Delete photo?",
        message: "This removes the photo from the project folder in SharePoint.",
        confirmLabel: "Delete",
        danger: true
      });
      if (!ok) return;
      try {
        const siteUrl = typeof currentSiteUrl === "function" ? currentSiteUrl() : "";
        await sharePointAdapter.deletePulseDocumentsFile(siteUrl, btn.dataset.photoDelete);
        toast("Photo deleted.", "success");
        await loadPhotos();
      } catch (e) {
        toast((e && e.friendly) || (e && e.message) || "Could not delete photo.", "error");
      }
    }));
  }

  $("#project-photos-refresh", body).addEventListener("click", () => loadPhotos());
  const fileInput = $("#project-photos-file-input", body);
  const uploadBtn = $("#project-photos-upload", body);
  if (uploadBtn && fileInput) {
    uploadBtn.addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", async () => {
      await uploadPhotos(fileInput.files);
      fileInput.value = "";
    });
  }
  const dropzone = $("#project-photos-dropzone", body);
  if (dropzone && canEdit) {
    dropzone.addEventListener("dragover", (event) => {
      event.preventDefault();
      dropzone.classList.add("drag-over");
    });
    dropzone.addEventListener("dragleave", () => dropzone.classList.remove("drag-over"));
    dropzone.addEventListener("drop", async (event) => {
      event.preventDefault();
      dropzone.classList.remove("drag-over");
      await uploadPhotos(event.dataTransfer && event.dataTransfer.files);
    });
  }

  loadPhotos();
}

window.drawProjectPhotos = drawProjectPhotos;
