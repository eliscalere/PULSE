function boardsForProject(projectId) {
  return (window.AEWTTR.db.checklistBoards || []).filter((b) => b.projectId === projectId);
}

function projectForBoard(board) {
  const db = window.AEWTTR.db;
  return (db.projects || []).find((p) => p.id === (board && board.projectId)) || null;
}

async function persistProjectBoards(projectId) {
  const db = window.AEWTTR.db;
  const proj = (db.projects || []).find((p) => p.id === projectId);
  if (!proj) {
    if (typeof aewttrSaveStore === "function") aewttrSaveStore();
    return;
  }
  proj.updated = new Date().toISOString().slice(0, 10);
  // Stamp immediately so a background SharePoint poll cannot replace db
  // while the board write is still debouncing / in flight.
  if (window.AEWTTR) window.AEWTTR.lastLocalSpWriteAt = Date.now();
  if (typeof Repo !== "undefined" && Repo.save) {
    // Fire-and-forget: the debounced write settles in the background. Every
    // board interaction (drag, resize, rename, delete) used to await a full
    // Repo.flush() here — that forces ALL pending writes app-wide to fire
    // and complete before the UI could redraw, making every click feel
    // laggy. Local `db` state is already mutated by the caller, so the UI
    // can redraw immediately without waiting on the network round trip.
    Repo.save("project", proj);
  } else if (typeof aewttrSaveStore === "function") {
    aewttrSaveStore();
  }
}

function drawProjectBoardsTab(body, proj, boardId) {
  if (boardId) return renderProjectBoardDetail(body, proj, boardId);
  renderProjectBoardGallery(body, proj);
}

/* Legacy global route — boards live inside each project now */
PAGE_RENDERERS.boards = function () { navigate("projects"); };
PAGE_RENDERERS.checklists = PAGE_RENDERERS.boards;

if (!window.AEWTTR.state) window.AEWTTR.state = {};

function ensureBoardShape(board) {
  if (!board.statuses || !board.statuses.length) board.statuses = defaultBoardStatuses();
  if (!board.type) board.type = "table";
  if (!board.team) board.team = "All";
  ensureBoardColumnMeta(board);
  ensureBoardCustomFields(board);
  const db = window.AEWTTR.db;
  if (!db.checklistTasks[board.id]) db.checklistTasks[board.id] = {};
  board.statuses.forEach((s) => {
    if (!db.checklistTasks[board.id][s]) db.checklistTasks[board.id][s] = [];
  });
  return board;
}

function pctComplete(boardId) {
  const tasks = window.AEWTTR.db.checklistTasks[boardId] || {};
  const cols = Object.keys(tasks);
  const allTasks = cols.flatMap(c => tasks[c]);
  if (!allTasks.length) return 0;
  const board = window.AEWTTR.db.checklistBoards.find(b => b.id === boardId);
  const doneStatus = (board && board.statuses && board.statuses[board.statuses.length - 1]) || "Complete";
  const completeCount = (tasks[doneStatus] || []).length;
  return Math.round((completeCount / allTasks.length) * 100);
}

const BOARD_COLUMN_PALETTE = ["#8590A2", "#0C66E4", "#6E5DC6", "#E56910", "#22A06B", "#E34935"];

function defaultBoardStatuses() {
  return ["To Do", "In Progress", "Complete"];
}

function ensureBoardColumnMeta(board) {
  if (!board.columnMeta) board.columnMeta = {};
  (board.statuses || []).forEach((name, index) => {
    if (!board.columnMeta[name]) {
      board.columnMeta[name] = { color: BOARD_COLUMN_PALETTE[index % BOARD_COLUMN_PALETTE.length] };
    }
  });
  return board;
}

function renameBoardColumn(board, oldName, newName) {
  const trimmed = String(newName || "").trim();
  if (!trimmed || trimmed === oldName) return false;
  if (board.statuses.includes(trimmed)) { toast("A column with that name already exists.", "error"); return false; }
  const db = window.AEWTTR.db;
  const data = db.checklistTasks[board.id] || {};
  const idx = board.statuses.indexOf(oldName);
  if (idx < 0) return false;
  board.statuses[idx] = trimmed;
  if (data[oldName]) {
    data[trimmed] = data[oldName];
    delete data[oldName];
  } else if (!data[trimmed]) {
    data[trimmed] = [];
  }
  if (board.columnMeta && board.columnMeta[oldName]) {
    board.columnMeta[trimmed] = board.columnMeta[oldName];
    delete board.columnMeta[oldName];
  }
  ensureBoardColumnMeta(board);
  persistProjectBoards(board.projectId);
  return true;
}

function addBoardColumn(board, name) {
  const trimmed = String(name || "").trim();
  if (!trimmed) return false;
  if (board.statuses.includes(trimmed)) { toast("That column already exists.", "error"); return false; }
  board.statuses.push(trimmed);
  const db = window.AEWTTR.db;
  if (!db.checklistTasks[board.id]) db.checklistTasks[board.id] = {};
  db.checklistTasks[board.id][trimmed] = [];
  ensureBoardColumnMeta(board);
  persistProjectBoards(board.projectId);
  return true;
}

async function removeBoardColumn(board, name) {
  if (board.statuses.length <= 1) { toast("Boards need at least one column.", "error"); return false; }
  const db = window.AEWTTR.db;
  const data = db.checklistTasks[board.id] || {};
  const tasks = data[name] || [];
  const ok = tasks.length
    ? await confirmDialog({
      title: "Delete column",
      message: `"${name}" has ${tasks.length} item${tasks.length === 1 ? "" : "s"}. Delete the column and its items?`,
      confirmLabel: "Delete column",
      danger: true
    })
    : true;
  if (!ok) return false;
  const idx = board.statuses.indexOf(name);
  if (idx >= 0) board.statuses.splice(idx, 1);
  delete data[name];
  if (board.columnMeta) delete board.columnMeta[name];
  await persistProjectBoards(board.projectId);
  return true;
}

function boardColumnColor(board, columnName) {
  ensureBoardColumnMeta(board);
  return (board.columnMeta[columnName] && board.columnMeta[columnName].color) || BOARD_COLUMN_PALETTE[0];
}

/* Custom fields — the "make Boards fully customizable, like Excel" ask.
   A field is {key, label, type: "text"|"select", options}. Values live
   directly on the item as item[field.key], same as the built-in "priority"
   field — no separate values table, so they ride along with the item for
   free wherever items already get copied/serialized. */
function ensureBoardCustomFields(board) {
  if (!Array.isArray(board.customFields)) board.customFields = [];
  return board.customFields;
}

function boardCustomFieldKeyFromLabel(label) {
  const slug = String(label || "").toLowerCase().trim().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
  return `cf_${slug || uid("f")}`;
}

function addBoardCustomField(board, { label, type, options }) {
  const trimmed = String(label || "").trim();
  if (!trimmed) { toast("Field name is required.", "error"); return false; }
  ensureBoardCustomFields(board);
  const key = boardCustomFieldKeyFromLabel(trimmed);
  if (board.customFields.some((f) => f.key === key || f.label.toLowerCase() === trimmed.toLowerCase())) {
    toast("A field with that name already exists.", "error");
    return false;
  }
  const isSelect = type === "select";
  const cleanOptions = isSelect
    ? (options || []).map((o) => String(o).trim()).filter(Boolean)
    : undefined;
  if (isSelect && (!cleanOptions || !cleanOptions.length)) {
    toast("Add at least one option for a dropdown field.", "error");
    return false;
  }
  board.customFields.push({ key, label: trimmed, type: isSelect ? "select" : "text", options: cleanOptions });
  persistProjectBoards(board.projectId);
  return true;
}

async function removeBoardCustomField(board, key) {
  ensureBoardCustomFields(board);
  const field = board.customFields.find((f) => f.key === key);
  if (!field) return false;
  const ok = await confirmDialog({
    title: "Remove field",
    message: `Remove "${field.label}" from this board? Any values already entered for it are discarded.`,
    confirmLabel: "Remove field",
    danger: true
  });
  if (!ok) return false;
  board.customFields = board.customFields.filter((f) => f.key !== key);
  await persistProjectBoards(board.projectId);
  return true;
}

async function renameBoardCustomField(board, key) {
  const field = ensureBoardCustomFields(board).find((f) => f.key === key);
  if (!field) return false;
  const next = await promptDialog({ title: "Rename field", label: "Field name", value: field.label });
  if (!next || next === field.label) return false;
  if (board.customFields.some((f) => f.key !== key && f.label.toLowerCase() === next.toLowerCase())) {
    toast("A field with that name already exists.", "error");
    return false;
  }
  field.label = next;
  await persistProjectBoards(board.projectId);
  return true;
}

function openAddBoardFieldModal(board, onDone) {
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>Add field</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body">
      <div class="form-row"><label>Field name</label><input class="input-aewttr" id="bf-label" placeholder="e.g. Budget, Region, Risk Level"></div>
      <div class="form-row"><label>Type</label>
        <select class="select-aewttr" id="bf-type">
          <option value="text">Text</option>
          <option value="select">Dropdown</option>
        </select>
      </div>
      <div class="form-row" id="bf-options-row" style="display:none;">
        <label>Options <small style="font-weight:400;color:var(--aewttr-muted);">(one per line)</small></label>
        <textarea class="textarea-aewttr" id="bf-options" placeholder="Low\nMedium\nHigh"></textarea>
      </div>
    </div>
    <div class="aewttr-modal-foot">
      <button class="btn-aewttr-ghost" id="bf-cancel">Cancel</button>
      <button class="btn-aewttr" id="bf-save">Add field</button>
    </div>
  `);
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#bf-cancel", modal).addEventListener("click", closeModal);
  $("#bf-type", modal).addEventListener("change", (e) => {
    $("#bf-options-row", modal).style.display = e.target.value === "select" ? "" : "none";
  });
  $("#bf-label", modal).focus();
  $("#bf-save", modal).addEventListener("click", async () => {
    const label = $("#bf-label", modal).value;
    const type = $("#bf-type", modal).value;
    const options = $("#bf-options", modal).value.split("\n");
    if (addBoardCustomField(board, { label, type, options })) {
      closeModal();
      toast("Field added", "success");
      onDone();
    }
  });
}

function boardItemCount(board) {
  const tasks = window.AEWTTR.db.checklistTasks[board.id] || {};
  return Object.values(tasks).reduce((s, l) => s + l.length, 0);
}

async function renameBoard(board) {
  const next = await promptBoardColumnName("Rename board", board.name);
  if (!next || next === board.name) return false;
  board.name = next;
  await persistProjectBoards(board.projectId);
  toast("Board renamed", "success");
  return true;
}

async function deleteBoard(board) {
  const ok = await confirmDialog({
    title: "Delete board",
    message: `Delete "${board.name}"? This removes all its items and cannot be undone.`,
    confirmLabel: "Delete board",
    danger: true
  });
  if (!ok) return false;
  const db = window.AEWTTR.db;
  db.checklistBoards = db.checklistBoards.filter((b) => b.id !== board.id);
  delete db.checklistTasks[board.id];
  await persistProjectBoards(board.projectId);
  toast("Board deleted", "success");
  return true;
}

function boardMenuItems(board, { onRenamed, onDeleted }) {
  return [
    { label: "Rename board", onClick: async () => { if (await renameBoard(board)) onRenamed(); } },
    { label: "Delete board", danger: true, onClick: async () => { if (await deleteBoard(board)) onDeleted(); } }
  ];
}

function openNewBoardWizard(projectId) {
  const db = window.AEWTTR.db;
  let pickedType = null;
  const modal = openModal(`
    <div class="aewttr-modal-head board-wizard-head"><h3>New board</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body board-wizard-body">
      <div class="board-wizard-step" id="bw-step-type">
        <p class="board-wizard-lead">Pick how you want to organize work — you can switch between table and kanban anytime.</p>
        <div class="board-type-grid">
          <button type="button" class="board-type-card" data-pick="table">
            <div class="board-type-graphic board-type-graphic--table" aria-hidden="true">
              <span></span><span></span><span></span><span></span>
            </div>
            <strong>Main table</strong>
            <span>Spreadsheet-style rows with inline status, owner, priority, and expandable subitems.</span>
          </button>
          <button type="button" class="board-type-card" data-pick="kanban">
            <div class="board-type-graphic board-type-graphic--kanban" aria-hidden="true">
              <span></span><span></span><span></span>
            </div>
            <strong>Kanban</strong>
            <span>Drag cards across columns — great for flow, sprints, and standups.</span>
          </button>
        </div>
      </div>
      <div class="board-wizard-step" id="bw-step-name" hidden>
        <button type="button" class="board-wizard-back" id="bw-back"><i class="bx bx-arrow-back"></i> Back</button>
        <div class="board-wizard-picked">
          <span class="board-wizard-picked-label" id="bw-picked-label">Table board</span>
        </div>
        <div class="form-row"><label>Board name</label><input class="input-aewttr" id="nb-name" placeholder="e.g. Marketing launch"></div>
        <div class="form-row"><label>Team <small style="font-weight:400;color:var(--aewttr-muted);">(optional)</small></label><input class="input-aewttr" id="nb-team" placeholder="All"></div>
        <p class="board-wizard-hint">Starts with To Do, In Progress, and Complete — add or rename columns anytime on the board.</p>
      </div>
    </div>
    <div class="aewttr-modal-foot board-wizard-foot" id="bw-foot-type" style="justify-content:flex-end;">
      <button class="btn-aewttr-ghost" id="nb-cancel">Cancel</button>
    </div>
    <div class="aewttr-modal-foot board-wizard-foot" id="bw-foot-name" hidden>
      <button class="btn-aewttr-ghost" id="nb-cancel-2">Cancel</button>
      <button class="btn-aewttr" id="nb-save"><i class="bx bx-plus"></i> Create board</button>
    </div>
  `, { className: "board-wizard-modal" });

  function showNameStep(type) {
    pickedType = type;
    $("#bw-step-type", modal).hidden = true;
    $("#bw-step-name", modal).hidden = false;
    $("#bw-foot-type", modal).hidden = true;
    $("#bw-foot-name", modal).hidden = false;
    $("#bw-picked-label", modal).textContent = type === "kanban" ? "Kanban board" : "Main table";
    $("#nb-name", modal).focus();
  }

  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#nb-cancel", modal).addEventListener("click", closeModal);
  $("#nb-cancel-2", modal).addEventListener("click", closeModal);
  $("#bw-back", modal).addEventListener("click", () => {
    pickedType = null;
    $("#bw-step-name", modal).hidden = true;
    $("#bw-step-type", modal).hidden = false;
    $("#bw-foot-name", modal).hidden = true;
    $("#bw-foot-type", modal).hidden = false;
  });
  $all("[data-pick]", modal).forEach((btn) => btn.addEventListener("click", () => showNameStep(btn.dataset.pick)));
  $("#nb-save", modal).addEventListener("click", async () => {
    const name = $("#nb-name", modal).value.trim();
    if (!name) { toast("Board name is required", "error"); return; }
    if (!pickedType) return;
    const id = uid("CB");
    const statuses = defaultBoardStatuses();
    const board = { id, name, projectId, team: $("#nb-team", modal).value.trim() || "All", type: pickedType, statuses, columnMeta: {} };
    ensureBoardColumnMeta(board);
    db.checklistBoards.unshift(board);
    db.checklistTasks[id] = {};
    statuses.forEach((s) => { db.checklistTasks[id][s] = []; });
    await persistProjectBoards(projectId);
    closeModal();
    toast("Board created", "success");
    navigate(`projects/${projectId}/boards/${id}`);
  });
}

function closeBoardContextMenu() {
  const menu = document.getElementById("board-context-menu");
  if (menu) menu.remove();
}

function openBoardContextMenu(x, y, items) {
  closeBoardContextMenu();
  const menu = document.createElement("div");
  menu.id = "board-context-menu";
  menu.className = "gantt-context-menu board-context-menu";
  menu.style.left = `${Math.min(x, window.innerWidth - 220)}px`;
  menu.style.top = `${Math.min(y, window.innerHeight - 220)}px`;
  items.forEach((item) => {
    if (item.separator) {
      const sep = document.createElement("div");
      sep.className = "gantt-context-menu-sep";
      menu.appendChild(sep);
      return;
    }
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `gantt-context-menu-item${item.danger ? " is-danger" : ""}`;
    btn.textContent = item.label;
    btn.addEventListener("click", (ev) => {
      ev.stopPropagation();
      closeBoardContextMenu();
      item.onClick();
    });
    menu.appendChild(btn);
  });
  document.body.appendChild(menu);
  setTimeout(() => {
    document.addEventListener("click", closeBoardContextMenu, { once: true });
    document.addEventListener("contextmenu", closeBoardContextMenu, { once: true });
  }, 0);
}

async function promptBoardColumnName(title, initial) {
  return new Promise((resolve) => {
    const label = /board/i.test(title || "") ? "Board name" : "Column name";
    const placeholder = /board/i.test(title || "") ? "e.g. Marketing launch" : "e.g. Ready for review";
    const modal = openModal(`
      <div class="aewttr-modal-head"><h3>${escapeHtml(title)}</h3><button class="aewttr-modal-close">&times;</button></div>
      <div class="aewttr-modal-body">
        <div class="form-row"><label>${label}</label><input class="input-aewttr" id="bc-name" value="${escapeHtml(initial || "")}" placeholder="${escapeHtml(placeholder)}"></div>
      </div>
      <div class="aewttr-modal-foot">
        <button class="btn-aewttr-ghost" id="bc-cancel">Cancel</button>
        <button class="btn-aewttr" id="bc-save">Save</button>
      </div>
    `);
    const input = $("#bc-name", modal);
    input.focus();
    input.select();
    const finish = (value) => { closeModal(); resolve(value); };
    $(".aewttr-modal-close", modal).addEventListener("click", () => finish(null));
    $("#bc-cancel", modal).addEventListener("click", () => finish(null));
    $("#bc-save", modal).addEventListener("click", () => finish(input.value.trim() || null));
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); finish(input.value.trim() || null); }
      if (e.key === "Escape") finish(null);
    });
  });
}

function renderProjectBoardGallery(body, proj) {
  const db = window.AEWTTR.db;
  const boards = boardsForProject(proj.id);
  boards.forEach(ensureBoardShape);
  body.innerHTML = `
    <div class="board-gallery-hero">
      <div>
        <h3 class="board-gallery-title">Project boards</h3>
        <p class="board-gallery-sub">Custom table or kanban boards for ${escapeHtml(proj.id)} — separate from the tracker, for any workflow this project needs.</p>
      </div>
      <button class="btn-aewttr btn-aewttr-sm" id="btn-new-board"${tip("Create a new board for this project")}><i class="bx bx-plus"></i> New board</button>
    </div>
    <div class="board-gallery-grid">
      ${boards.map(b => {
        const count = boardItemCount(b);
        const pct = pctComplete(b.id);
        const previewCols = b.statuses.slice(0, 3).map((col) => `
          <span class="board-card-col" style="--col-color:${boardColumnColor(b, col)}">${escapeHtml(col)}</span>`).join("");
        return `
        <div class="board-card" data-id="${b.id}">
          <button type="button" class="board-card-menu-btn" data-board-menu="${b.id}"${tip("Board options")} aria-label="Board options"><i class="bx bx-dots-horizontal-rounded"></i></button>
          <div class="board-card-cover ${b.type === "kanban" ? "is-kanban" : "is-table"}">
            ${b.type === "kanban" ? `<div class="board-card-kanban-preview">${previewCols}</div>` : `<div class="board-card-table-preview"><span></span><span></span><span></span></div>`}
          </div>
          <div class="board-card-body">
            <div class="board-card-head">
              <span class="board-card-icon ${b.type === "kanban" ? "is-kanban" : "is-table"}"><i class="bx ${b.type === "kanban" ? "bx-columns" : "bx-table"}"></i></span>
              <div class="board-card-meta">
                <div class="board-card-name">${escapeHtml(b.name)}</div>
                <div class="board-card-team">${escapeHtml(b.team)} · ${count} item${count === 1 ? "" : "s"}</div>
              </div>
              <span class="board-card-type">${b.type === "kanban" ? "Kanban" : "Table"}</span>
            </div>
            <div class="rock-bar-track"><div class="rock-bar-fill" style="width:${pct}%;background:var(--aewttr-green);"></div></div>
            <div class="board-card-pct">${pct}% complete</div>
          </div>
        </div>`;
      }).join("") || `
        <div class="board-gallery-empty">
          <div class="board-gallery-empty-icon"><i class="bx bx-columns"></i></div>
          <h3>No boards yet</h3>
          <p>Create a table or kanban board for this project's work.</p>
          <button class="btn-aewttr" id="btn-new-board-empty"><i class="bx bx-plus"></i> New board</button>
        </div>`}
    </div>
  `;
  function openCardMenu(board, x, y) {
    openBoardContextMenu(x, y, boardMenuItems(board, {
      onRenamed: () => renderProjectBoardGallery(body, proj),
      onDeleted: () => renderProjectBoardGallery(body, proj)
    }));
  }
  $all(".board-card", body).forEach(c => {
    c.addEventListener("click", () => navigate(`projects/${proj.id}/boards/${c.dataset.id}`));
    c.addEventListener("contextmenu", (e) => {
      e.preventDefault();
      const board = boards.find((b) => b.id === c.dataset.id);
      if (board) openCardMenu(board, e.clientX, e.clientY);
    });
  });
  $all("[data-board-menu]", body).forEach((btn) => btn.addEventListener("click", (e) => {
    e.stopPropagation();
    const board = boards.find((b) => b.id === btn.dataset.boardMenu);
    const rect = btn.getBoundingClientRect();
    if (board) openCardMenu(board, rect.left, rect.bottom + 4);
  }));
  const newBtn = $("#btn-new-board", body);
  if (newBtn) newBtn.addEventListener("click", () => openNewBoardWizard(proj.id));
  const emptyBtn = $("#btn-new-board-empty", body);
  if (emptyBtn) emptyBtn.addEventListener("click", () => openNewBoardWizard(proj.id));
}

function renderBoardGallery() {
  navigate("projects");
}

function flattenBoardItems(board) {
  const db = window.AEWTTR.db;
  const data = db.checklistTasks[board.id];
  const items = [];
  board.statuses.forEach((status) => {
    (data[status] || []).forEach((item) => {
      item.status = status;
      items.push(item);
    });
  });
  return items;
}

function filterBoardItems(items, query) {
  const q = String(query || "").trim().toLowerCase();
  if (!q) return items;
  return items.filter((item) => {
    const hay = [item.title, item.owner, item.priority, item.status]
      .concat((item.subtasks || []).map((s) => s.text))
      .filter(Boolean)
      .join(" ")
      .toLowerCase();
    return hay.includes(q);
  });
}

function filterBoardData(board, query) {
  const db = window.AEWTTR.db;
  const data = db.checklistTasks[board.id] || {};
  const q = String(query || "").trim().toLowerCase();
  if (!q) return data;
  const filtered = {};
  board.statuses.forEach((col) => {
    filtered[col] = (data[col] || []).filter((item) => {
      const hay = [item.title, item.owner, item.priority]
        .concat((item.subtasks || []).map((s) => s.text))
        .filter(Boolean)
        .join(" ")
        .toLowerCase();
      return hay.includes(q);
    });
  });
  return filtered;
}

function getBoardFilter(boardId) {
  if (!window.AEWTTR.state.boardFilters) window.AEWTTR.state.boardFilters = {};
  return window.AEWTTR.state.boardFilters[boardId] || "";
}

function setBoardFilter(boardId, value) {
  if (!window.AEWTTR.state.boardFilters) window.AEWTTR.state.boardFilters = {};
  window.AEWTTR.state.boardFilters[boardId] = value;
}

function openBoardItemModal(board, itemId, subIndex, onDone) {
  const db = window.AEWTTR.db;
  const data = db.checklistTasks[board.id];
  let existing = null;
  let currentStatus = board.statuses[0];
  if (itemId) {
    for (const s of board.statuses) {
      const found = (data[s] || []).find((t) => t.id === itemId);
      if (found) { existing = found; currentStatus = s; break; }
    }
  }
  const modal = openModal(`
    <div class="aewttr-modal-head"><h3>${existing ? "Edit item" : "New item"}</h3><button class="aewttr-modal-close">&times;</button></div>
    <div class="aewttr-modal-body">
      <div class="form-row"><label>Title</label><input class="input-aewttr" id="ci-title" value="${existing ? escapeHtml(existing.title) : ""}"></div>
      <div class="form-grid-2">
        <div class="form-row"><label>Owner</label><input class="input-aewttr" id="ci-owner" value="${existing ? escapeHtml(existing.owner || "") : ""}"></div>
        <div class="form-row"><label>Due date</label><input type="date" class="input-aewttr" id="ci-due" value="${existing && existing.due ? existing.due : ""}"></div>
      </div>
      <div class="form-grid-2">
        <div class="form-row"><label>Status</label><select class="select-aewttr" id="ci-status">${board.statuses.map((s) => `<option ${currentStatus === s ? "selected" : ""}>${escapeHtml(s)}</option>`).join("")}</select></div>
        <div class="form-row"><label>Priority</label><select class="select-aewttr" id="ci-priority">${["Low", "Medium", "High", "Critical"].map((p) => `<option ${existing && existing.priority === p ? "selected" : ""}>${p}</option>`).join("")}</select></div>
      </div>
      <div class="form-row"><label>Linked Tracker Task <small style="font-weight:400;color:var(--aewttr-muted);">(optional)</small></label>
        <select class="select-aewttr" id="ci-linked-task">
          <option value="">-- None --</option>
          ${(window.AEWTTR.db.ganttTasks[board.projectId] || []).map(t => `<option value="${escapeHtml(t.id)}" ${existing && existing.linkedTaskId === t.id ? "selected" : ""}>${escapeHtml(t.title)}</option>`).join("")}
        </select>
      </div>
      <div class="form-row"><label>Subitems <small style="font-weight:400;color:var(--aewttr-muted);">(one per line)</small></label>
        <textarea class="textarea-aewttr" id="ci-subtasks">${existing && existing.subtasks ? existing.subtasks.map((s) => s.text).join("\n") : ""}</textarea>
      </div>
      ${(board.customFields || []).length ? `
      <div class="form-grid-2">
        ${(board.customFields || []).map((f) => `
          <div class="form-row"><label>${escapeHtml(f.label)}</label>
            ${f.type === "select"
              ? `<select class="select-aewttr" data-ci-field="${escapeHtml(f.key)}">${f.options.map((o) => `<option ${existing && existing[f.key] === o ? "selected" : ""}>${escapeHtml(o)}</option>`).join("")}</select>`
              : `<input class="input-aewttr" data-ci-field="${escapeHtml(f.key)}" value="${existing ? escapeHtml(existing[f.key] || "") : ""}">`}
          </div>
        `).join("")}
      </div>` : ""}
    </div>
    <div class="aewttr-modal-foot">
      ${existing ? `<button class="btn-danger-outline" id="ci-delete"${tip("Delete this item")}><i class="bx bx-trash"></i> Delete</button>` : ""}
      <button class="btn-aewttr-ghost" id="ci-cancel">Cancel</button>
      <button class="btn-aewttr" id="ci-save">${existing ? "Save changes" : "Add item"}</button>
    </div>
  `);
  const proj = window.AEWTTR.db.projects.find((p) => p.id === board.projectId);
  if (proj && typeof wireAssigneeAutocomplete === "function" && typeof projectRoleCandidateList === "function") {
    wireAssigneeAutocomplete(modal, "ci-owner", { directory: projectRoleCandidateList(proj) });
  }
  $(".aewttr-modal-close", modal).addEventListener("click", closeModal);
  $("#ci-cancel", modal).addEventListener("click", closeModal);
  if (existing) {
    $("#ci-delete", modal).addEventListener("click", async () => {
      const ok = await confirmDialog({ title: "Delete item", message: `Delete "${existing.title}"? This cannot be undone.`, confirmLabel: "Delete", danger: true });
      if (!ok) return;
      data[currentStatus] = (data[currentStatus] || []).filter((t) => t.id !== existing.id);
      await persistProjectBoards(board.projectId);
      closeModal();
      toast("Item deleted", "success");
      onDone();
    });
  }
  $("#ci-save", modal).addEventListener("click", async () => {
    const title = $("#ci-title", modal).value.trim();
    if (!title) { toast("Title is required", "error"); return; }
    const subtasks = $("#ci-subtasks", modal).value.split("\n").map((s) => s.trim()).filter(Boolean)
      .map((text) => ({ text, done: existing && existing.subtasks ? (existing.subtasks.find((s) => s.text === text) || {}).done || false : false }));
    const newStatus = $("#ci-status", modal).value;
    const payload = {
      title,
      owner: $("#ci-owner", modal).value.trim() || "Unassigned",
      due: $("#ci-due", modal).value || "",
      priority: $("#ci-priority", modal).value,
      linkedTaskId: $("#ci-linked-task", modal).value,
      subtasks
    };
    (board.customFields || []).forEach((f) => {
      const el = $(`[data-ci-field="${f.key}"]`, modal);
      if (el) payload[f.key] = el.value.trim ? el.value.trim() : el.value;
    });
    if (existing) {
      Object.assign(existing, payload);
      if (newStatus !== currentStatus) {
        data[currentStatus] = (data[currentStatus] || []).filter((t) => t.id !== existing.id);
        if (!data[newStatus]) data[newStatus] = [];
        data[newStatus].push(existing);
      }
      toast("Item updated", "success");
    } else {
      payload.id = uid("t");
      if (!data[newStatus]) data[newStatus] = [];
      data[newStatus].push(payload);
      toast("Item added", "success");
    }
    await persistProjectBoards(board.projectId);
    closeModal();
    onDone();
  });
}

function renderBoardTableView(mount, board, query, redraw) {
  const db = window.AEWTTR.db;
  if (!window.AEWTTR.state.checklistExpanded) window.AEWTTR.state.checklistExpanded = {};
  if (!window.AEWTTR.state.checklistExpanded[board.id]) window.AEWTTR.state.checklistExpanded[board.id] = {};
  const expanded = window.AEWTTR.state.checklistExpanded[board.id];
  if (!redraw) redraw = () => renderBoardTableView(mount, board, query, redraw);

  function moveItemStatus(item, fromStatus, toStatus) {
    const data = db.checklistTasks[board.id];
    const list = data[fromStatus] || [];
    const idx = list.findIndex((t) => t.id === item.id);
    if (idx < 0) return;
    list.splice(idx, 1);
    if (!data[toStatus]) data[toStatus] = [];
    data[toStatus].push(item);
  }

  const items = filterBoardItems(flattenBoardItems(board), query);
  mount.innerHTML = `<div id="cl-table-mount"></div>`;
  const tableMount = $("#cl-table-mount", mount);
  renderMondayTable(tableMount, {
    tasks: items,
    ownerField: "owner",
    dueField: "due",
    expanded,
    inlineOwner: true,
    groupHeader: { label: board.name, count: items.length },
    emptyMessage: query ? "No items match your search." : 'No items yet — use "New item" or "+ Add item" to start.',
    statusColumn: {
      key: "status",
      label: "Status",
      options: board.statuses,
      colorPrefix: "status",
      colorFor: (value) => boardColumnColor(board, value)
    },
    extraColumns: [
      { key: "priority", label: "Priority", options: ["Low", "Medium", "High", "Critical"], colorPrefix: "priority" },
      ...(board.customFields || []).map((f) => ({ key: f.key, label: f.label, type: f.type, options: f.options }))
    ],
    inlineAddLabel: "+ Add item",
    onInlineAdd: () => openBoardItemModal(board, null, null, redraw),
    onToggleExpand: (id) => { expanded[id] = !expanded[id]; redraw(); },
    onOpenEditor: (id) => openBoardItemModal(board, id, null, redraw),
    onFieldChange: async (item, key, value) => {
      if (key === "status") {
        const from = board.statuses.find((s) => (db.checklistTasks[board.id][s] || []).some((t) => t.id === item.id));
        if (from && from !== value) moveItemStatus(item, from, value);
      } else {
        item[key] = value;
      }
      await persistProjectBoards(board.projectId);
      redraw();
    },
    onToggleSubtaskDone: async (item, idx, checked) => {
      item.subtasks[idx].done = checked;
      await persistProjectBoards(board.projectId);
      redraw();
    }
  });

  function columnMenuItems() {
    const menuItems = [];
    menuItems.push({ label: "Add column", onClick: async () => {
      const name = await promptBoardColumnName("Add column");
      if (name && addBoardColumn(board, name)) redraw();
    }});
    board.statuses.forEach((col) => {
      menuItems.push({ label: `Rename “${col}”`, onClick: async () => {
        const next = await promptBoardColumnName("Rename column", col);
        if (next && renameBoardColumn(board, col, next)) redraw();
      }});
    });
    if (board.statuses.length > 1) {
      menuItems.push({ separator: true });
      board.statuses.forEach((col) => {
        menuItems.push({ label: `Delete “${col}”`, danger: true, onClick: async () => {
          if (await removeBoardColumn(board, col)) redraw();
        }});
      });
    }
    return menuItems;
  }

  tableMount.addEventListener("contextmenu", (e) => {
    const row = e.target.closest("tr[data-id], .monday-row[data-id]");
    const head = e.target.closest("th");
    if (!row && !head) return;
    e.preventDefault();
    const menuItems = [];
    if (row) {
      const itemId = row.dataset.id;
      const data = db.checklistTasks[board.id];
      let found = null;
      let status = null;
      board.statuses.forEach((s) => {
        const hit = (data[s] || []).find((t) => t.id === itemId);
        if (hit) { found = hit; status = s; }
      });
      menuItems.push({ label: "Edit item", onClick: () => openBoardItemModal(board, itemId, null, redraw) });
      board.statuses.filter((s) => s !== status).forEach((s) => {
        menuItems.push({ label: `Move to ${s}`, onClick: async () => {
          if (!found || !status) return;
          moveItemStatus(found, status, s);
          await persistProjectBoards(board.projectId);
          redraw();
        }});
      });
      menuItems.push({ label: "Delete item", danger: true, onClick: async () => {
        if (!found) return;
        const ok = await confirmDialog({ title: "Delete item", message: `Delete "${found.title}"?`, confirmLabel: "Delete", danger: true });
        if (!ok) return;
        data[status] = (data[status] || []).filter((t) => t.id !== itemId);
        await persistProjectBoards(board.projectId);
        redraw();
      }});
      menuItems.push({ separator: true });
    }
    openBoardContextMenu(e.clientX, e.clientY, menuItems.concat(columnMenuItems()));
  });

  // A "..." button next to the Status header makes column add/rename/
  // delete discoverable without needing to already know right-click works
  // there — right-click still does the same thing, this is just a visible
  // entry point to it.
  const statusHeader = $all("th", tableMount).find((th) => th.textContent.trim() === "Status");
  if (statusHeader) {
    statusHeader.classList.add("board-th-with-menu");
    const menuBtn = document.createElement("button");
    menuBtn.type = "button";
    menuBtn.className = "board-th-menu-btn";
    menuBtn.setAttribute("aria-label", "Column options");
    menuBtn.innerHTML = `<i class="bx bx-dots-horizontal-rounded"></i>`;
    menuBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      const rect = menuBtn.getBoundingClientRect();
      openBoardContextMenu(rect.left, rect.bottom + 4, columnMenuItems());
    });
    statusHeader.appendChild(menuBtn);
  }

  wireBoardColumnResize(tableMount, board.id);
}

/* Drag-to-resize table columns. Widths are session-local UI state (like
   the expanded/filter state above), not synced to SharePoint — resizing is
   about this table being usable, not something that needs to travel with
   the board data itself. */
function boardColumnWidthState(boardId) {
  if (!window.AEWTTR.state.boardColumnWidths) window.AEWTTR.state.boardColumnWidths = {};
  if (!window.AEWTTR.state.boardColumnWidths[boardId]) window.AEWTTR.state.boardColumnWidths[boardId] = {};
  return window.AEWTTR.state.boardColumnWidths[boardId];
}

function wireBoardColumnResize(tableMount, boardId) {
  const table = $(".monday-table", tableMount);
  if (!table) return;
  table.style.tableLayout = "fixed";
  const widths = boardColumnWidthState(boardId);
  const headers = $all("th", tableMount);
  headers.forEach((th, index) => {
    th.style.position = th.style.position || "relative";
    if (widths[index]) th.style.width = `${widths[index]}px`;
    if (index === 0) return; // skip the icon-only expand column
    const handle = document.createElement("span");
    handle.className = "board-col-resize-handle";
    th.appendChild(handle);
    handle.addEventListener("mousedown", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const startX = e.clientX;
      const startWidth = th.getBoundingClientRect().width;
      const onMove = (moveEvent) => {
        const next = Math.max(60, startWidth + (moveEvent.clientX - startX));
        th.style.width = `${next}px`;
        widths[index] = next;
      };
      const onUp = () => {
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
      };
      document.addEventListener("mousemove", onMove);
      document.addEventListener("mouseup", onUp);
    });
  });
}

function renderBoardKanbanView(mount, board, query, redraw) {
  if (!redraw) redraw = () => renderBoardKanbanView(mount, board, query, redraw);
  mount.innerHTML = `<div id="cl-kanban-mount"></div>`;
  buildKanban({
    mount: $("#cl-kanban-mount", mount),
    columns: board.statuses,
    data: filterBoardData(board, query),
    ownerField: "owner",
    showPriority: true,
    mondayStyle: true,
    customFields: board.customFields || [],
    addButtonLabel: "Add item",
    columnColor: (col) => boardColumnColor(board, col),
    allowColumnEdit: true,
    onAddColumn: (name) => addBoardColumn(board, name),
    onRenameColumn: (oldName, newName) => renameBoardColumn(board, oldName, newName),
    onRemoveColumn: async (name) => removeBoardColumn(board, name),
    onOpenItem: (id) => openBoardItemModal(board, id, null, redraw),
    onSave: async () => { await persistProjectBoards(board.projectId); },
    onColumnsChange: redraw
  });
}

function renderProjectBoardDetail(body, proj, id) {
  const board = boardsForProject(proj.id).find((b) => b.id === id);
  if (!board) { navigate(`projects/${proj.id}/boards`); return; }
  renderBoardDetailInto(body, proj, board);
}

function renderBoardDetail(id) {
  const board = window.AEWTTR.db.checklistBoards.find((b) => b.id === id);
  if (!board || !board.projectId) { navigate("projects"); return; }
  navigate(`projects/${board.projectId}/boards/${id}`);
}

function renderBoardDetailInto(body, proj, board) {
  const id = board.id;
  ensureBoardShape(board);
  const query = getBoardFilter(board.id);
  const itemCount = boardItemCount(board);
  const pct = pctComplete(board.id);
  function redraw() { renderBoardDetailInto(body, proj, board); }

  body.innerHTML = `
    <div class="board-shell">
      <div class="board-shell-head">
        <div class="board-shell-title-row">
          <div class="board-shell-title">
            <button type="button" class="board-back-link" id="btn-back-boards"><i class="bx bx-arrow-back"></i> Boards</button>
            <h2>${escapeHtml(board.name)}</h2>
            <span class="board-team-chip">${escapeHtml(board.team)}</span>
            <button type="button" class="mrc-icon-btn" id="board-menu-btn"${tip("Board options")} aria-label="Board options"><i class="bx bx-dots-horizontal-rounded"></i></button>
          </div>
          <div class="board-shell-stats">
            <span class="board-stat-pill">${itemCount} item${itemCount === 1 ? "" : "s"}</span>
            <span class="board-stat-pill is-green">${pct}% done</span>
          </div>
        </div>
        <div class="board-view-tabs" role="tablist">
          <button type="button" class="board-view-tab ${board.type === "table" ? "active" : ""}" data-view="table" role="tab">
            <i class="bx bx-table"></i> Main table
          </button>
          <button type="button" class="board-view-tab ${board.type === "kanban" ? "active" : ""}" data-view="kanban" role="tab">
            <i class="bx bx-columns"></i> Kanban
          </button>
        </div>
      </div>
      <div class="board-toolbar">
        <button type="button" class="btn-aewttr btn-aewttr-sm" id="board-new-item"><i class="bx bx-plus"></i> New item</button>
        <div class="board-toolbar-search">
          <i class="bx bx-search"></i>
          <input type="search" id="board-search" placeholder="Search items" value="${escapeHtml(query)}" aria-label="Search board items">
        </div>
        <div class="board-fields-row" id="board-fields-row">
          ${(board.customFields || []).map((f) => `
            <span class="board-field-pill" data-field-pill="${escapeHtml(f.key)}"${tip(f.type === "select" ? `Dropdown: ${f.options.join(", ")}` : "Text field")}>
              <button type="button" class="board-field-pill-label" data-rename-field="${escapeHtml(f.key)}">${escapeHtml(f.label)}</button>
              <button type="button" class="board-field-pill-remove" data-remove-field="${escapeHtml(f.key)}" aria-label="Remove field">&times;</button>
            </span>
          `).join("")}
          <button type="button" class="board-toolbar-btn" id="board-add-field"${tip("Add a custom column — text or dropdown")}><i class="bx bx-plus"></i> Field</button>
        </div>
      </div>
      <div class="board-body" id="board-body-mount"></div>
    </div>
  `;

  $("#btn-back-boards", body).addEventListener("click", () => navigate(`projects/${proj.id}/boards`));
  $("#board-menu-btn", body).addEventListener("click", (e) => {
    const rect = e.currentTarget.getBoundingClientRect();
    openBoardContextMenu(rect.left, rect.bottom + 4, boardMenuItems(board, {
      onRenamed: redraw,
      onDeleted: () => navigate(`projects/${proj.id}/boards`)
    }));
  });
  $("#board-new-item", body).addEventListener("click", () => openBoardItemModal(board, null, null, redraw));
  $("#board-search", body).addEventListener("input", (e) => {
    setBoardFilter(board.id, e.target.value);
    redraw();
  });
  $all(".board-view-tab", body).forEach((tab) => tab.addEventListener("click", async () => {
    const next = tab.dataset.view;
    if (next === board.type) return;
    board.type = next;
    await persistProjectBoards(board.projectId);
    redraw();
  }));
  const addFieldBtn = $("#board-add-field", body);
  if (addFieldBtn) addFieldBtn.addEventListener("click", () => openAddBoardFieldModal(board, redraw));
  $all("[data-rename-field]", body).forEach((btn) => btn.addEventListener("click", async () => {
    if (await renameBoardCustomField(board, btn.dataset.renameField)) redraw();
  }));
  $all("[data-remove-field]", body).forEach((btn) => btn.addEventListener("click", async (e) => {
    e.stopPropagation();
    if (await removeBoardCustomField(board, btn.dataset.removeField)) redraw();
  }));

  const bodyMount = $("#board-body-mount", body);
  if (board.type === "kanban") renderBoardKanbanView(bodyMount, board, query, redraw);
  else renderBoardTableView(bodyMount, board, query, redraw);
}
