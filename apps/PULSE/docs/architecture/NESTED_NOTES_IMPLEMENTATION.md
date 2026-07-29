# Nested Notes Collection Feature - Implementation Summary

## Overview
Implemented a feature on the **Project Notes page** (`drawProjectNotes` in `assets/js/pages/projects.js`) that displays all notes from a task/subtask **and all notes from nested children** when clicking on that item in the sidebar.

## Current Behavior (Before)
- Clicking a task or subtask showed **only notes directly on that item**
- Notes on nested children were only visible by clicking into each child individually

## New Behavior (After)
- **Click a task** → Shows all notes on that task + all notes recursively from all its subitems
- **Click a subtask** → Shows all notes on that subtask + all notes recursively from all its nested children
- **Click "Project Notes" or "Overview"** → Shows project-level notes only (unchanged)
- Notes are **grouped by source item** with a visual header showing which task/subtask each group belongs to
- A helpful info banner appears: "Showing notes from this item and all nested subitems"

## Implementation Details

### Files Modified
- `assets/js/pages/projects.js` - Main implementation in the `drawProjectNotes()` function

### Key Functions Added

#### 1. `collectNotesRecursive(item)`
**Purpose:** Recursively collects all notes from an item and its nested children
```
- Collects direct notes from the item
- Recursively gathers notes from all subtasks
- Returns a flat array of all collected notes
```

#### 2. `collectNotesWithSource(item, task, pathPrefix)`
**Purpose:** Collects notes WITH SOURCE TRACKING for grouped display
```
- Creates rows for each note with metadata about its source
- Tracks the source item and a human-readable source label
- Recursively processes subtasks with breadcrumb-style labels
- Example labels:
  - "Task · My Task Name"
  - "Task · My Task Name › Subtask 1"
  - "Task · My Task Name › Subtask 1 › Subtask 1.1"
```

#### 3. Modified `resolveFolder(folderId)`
**Changes:**
- For task folders: Now calls `collectNotesRecursive(task)` instead of `task.notes`
- For sub folders: Now calls `collectNotesRecursive(sub)` instead of `sub.notes`
- Adds `hasNestedNotes` flag: true if collected notes > direct notes (indicates grouping needed)
- Adds `displayMode: "nested"` to signal the UI to use grouped display
- Notes are sorted in descending chronological order via `sortNotesDesc()`

### Display Logic in `renderMain()`

#### When `hasNestedNotes === false`
- Shows notes in a flat list (original behavior)
- No grouping overhead

#### When `hasNestedNotes === true`
- Groups notes by source item
- Each group has:
  - **Header** with link icon, source label, and note count
  - **Visual separator** (border-bottom)
  - **Note cards** below the header
- Groups maintain reverse chronological order overall
- Each note displays with its source label in the `showAssigned` parameter

### Example Display
```
Task Title
Task · My Task Name
ℹ Showing notes from this item and all nested subitems

[Group Header] 🔗 Task · My Task Name [3]
  [Note 1]
  [Note 2]
  [Note 3]

[Group Header] 🔗 Task · My Task Name › Subtask A [2]
  [Note 4]
  [Note 5]

[Group Header] 🔗 Task · My Task Name › Subtask B › Nested [1]
  [Note 6]
```

## Click Handler Integration
The existing click handler on sidebar task/subtask rows (`[data-folder]` buttons in `renderSidebar()`) automatically triggers the new behavior:
- Calls `setActiveFolder(btn.dataset.folder)` 
- Which calls `renderMain()`
- Which now uses the enhanced `resolveFolder()` logic

## Backward Compatibility
- **Overview tab** continues to work unchanged - shows all notes across all tasks and subitems in a hierarchical tree
- **Project Notes** tab continues to work unchanged - shows only project-level notes
- Direct clicks on notes still allow edit/delete with proper ownership checks
- The code gracefully handles items with no nested subitems (doesn't show grouping)

## Performance Considerations
- **collectNotesRecursive()**: O(n) where n = total number of notes in the tree
- **collectNotesWithSource()**: O(n log n) due to sorting
- Collection happens on each render, but:
  - Only when a task/sub folder is selected (not for overview)
  - Notes arrays are typically small per item
  - Acceptable for typical use cases

## CSS Classes
The implementation uses these CSS classes for styling:
- `.proj-notes-nested-group` - Container for grouped notes
- `.proj-notes-card` - Individual note card (existing)
- `.bx-link-alt` - Link icon from Boxicons library (existing)

**Note:** Uses inline styles for the group headers to avoid CSS file changes; can be moved to stylesheet if needed.

## Testing Checklist
- [ ] Click a task with nested subitems → shows all nested notes grouped by source
- [ ] Click a subtask with children → shows all child notes grouped by source
- [ ] Click a task with no subitems → shows only direct notes (no grouping)
- [ ] Click Overview → unchanged behavior
- [ ] Click Project Notes → unchanged behavior
- [ ] Edit/delete notes from grouped display → works correctly
- [ ] Add new note to task → appears in task's notes group
- [ ] Tree collapse/expand → sidebar works correctly

## Rebuild Command
```bash
node scripts/build-sharepoint-package.js
```
Output: `fs-packages/PULSE-v1.0.0.html`

## Future Enhancements
- Add a toggle to "Show only direct notes" vs "Show all nested notes"
- Add filtering/search within nested note groups
- Add option to collapse groups that have been reviewed
- Show visual indicators (badges) in sidebar for items with nested notes
