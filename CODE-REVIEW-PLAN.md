# Code Review Remediation Plan

**Target:** SharePoint Smart Gantt Chart (SPFx 1.20.0, React 17, TypeScript 4.7, Fluent UI 8, PnPjs 3.26, date-fns 2, SheetJS 0.20.3)
**Produced by:** full-codebase review, July 2026 (post v1.2.1). All findings were verified against the code at review time with file:line citations.

---

## Ground rules for the executor

1. **Verify before editing.** Read the cited code first. If a finding does not reproduce against the current code (line numbers may have drifted), skip it and note it in your summary — do not force a change.
2. **No new npm dependencies.** Everything here is fixable with what's installed. For memoization in the class component, use instance fields with manual input-comparison (or `memoize-one` ONLY if it is already in package.json — check first).
3. **Preserve conventions.** Schedule dates are canonical `YYYY-MM-DD` strings via `utils/dateUtils.ts` (`toDateOnly`/`parseDateOnly`); Created/Modified are real timestamps parsed with `new Date()`. Never introduce `new Date('YYYY-MM-DD')` for schedule dates (UTC shift). American English spelling everywhere.
4. **Match surrounding style** — comment density, naming, SCSS module patterns. `SmartGantt.tsx` is a class component; `GanttChart.tsx` and views are function components. Keep them that way.
5. **Build gate:** after each workstream, run `gulp build` and fix any TypeScript errors before moving on. There is no unit test suite; verification is compile + the acceptance criteria listed per item.
6. **Commit per workstream** (A, B, C, …) with a descriptive message, so changes are reviewable in coherent chunks.
7. Work the workstreams **in order** — they are sequenced by user impact and to minimize conflicts (A and B touch the same import files; do A fully before B's ImportPanel items).

Severity legend: 🔴 critical/high (visible wrong behavior or data corruption) · 🟠 medium (edge case, robustness, perf) · 🟡 low (polish).

---

## Workstream A — Import pipeline correctness (ImportService + ImportPanel)

These interact; implement A1–A6 together, then test the whole import flow.

### A1 🔴 Excel date cells import as garbage — serial-number branch is unreachable
**Files:** `src/webparts/smartGantt/services/ImportService.ts:198, 207-209, 118-124, 145, 245-246`
`parseExcelFile` reads the sheet with `raw: true`, so real Excel date cells arrive as serial numbers (e.g. `45123`), but every cell is then stringified: `cleaned[h] = String(r[h] ?? '')`. `parseExcelDate`'s `typeof value === 'number'` branch is therefore dead code, and the string `"45123"` falls through to `new Date("45123")` → Invalid Date. **Net effect: importing a normal .xlsx with date-formatted cells silently drops all Start/Due dates.**
**Fix:** in `parseExcelDate`, before the `new Date(str)` fallback, add a numeric-string branch:
```ts
if (/^\d+(\.\d+)?$/.test(str)) {
  const dc = XLSX.SSF.parse_date_code(parseFloat(str));
  return dc ? `${dc.y}-${String(dc.m).padStart(2, '0')}-${String(dc.d).padStart(2, '0')}` : '';
}
```
Guard against unreasonable serials (e.g. require `parseFloat(str) > 59 && < 200000` so plain integers like an ID column mapped by mistake don't become dates).
**Accept:** an .xlsx with a date-formatted Start Date column imports with correct `YYYY-MM-DD` dates; a text-date column (`03/15/2026`) still imports correctly.

### A2 🔴 Numeric dependency values written as raw SharePoint item IDs at create time
**Files:** `ImportService.ts:260-267` (`applyMapping`), interacts with `resolveDependencies` (450-516)
`applyMapping` parses numeric dependency cells (MS Project row numbers) and stores them directly in `task.dependencies`, which `batchImport` → `createTask` writes to the `Dependencies` field as if they were item IDs. If the target list already has tasks, or if `resolveDependencies` later fails or can't resolve a row, tasks permanently point at arbitrary unrelated items. Also `parseInt("3rd Party Review")` → dependency ID `3` (the strict `String(rowNum) === name` guard exists only in `resolveDependencies:495`, not here).
**Fix:** remove the `dependencies` case from `applyMapping` entirely — create tasks with empty `dependencies` and keep the raw dependency strings on the parsed rows for `resolveDependencies` to handle (it already implements both numeric-row and title-based resolution).
**Accept:** after import, no task's Dependencies field contains an ID that wasn't produced by `resolveDependencies`.

### A3 🔴 Dependency resolution is wrong for duplicate titles and fragile on failure
**Files:** `ImportService.ts:463-516`, called from `ImportPanel.tsx:244`
Three related defects:
- `titleToId` (line 467) is last-wins: two tasks titled "Kickoff" → both name- and row-number-based deps resolve to the wrong task. `rowToId` (472-476) is built via title lookup so it inherits the bug.
- One failed `updateTask` rejects the whole `Promise.all` (511-515), abandoning remaining updates; the function returns `void` so the caller can't report what happened.
- `ImportPanel.startImport` (ImportPanel.tsx:233-247) has no try/catch around `resolveDependencies`; a rejection strands the panel on the "importing" spinner forever (footer renders nothing on that step, line 523) even though all tasks were created.

**Fix (do all three as one change):**
1. Make `batchImport` return the created SharePoint item ID per row (in row order, `null` for failed rows) alongside its existing result.
2. Change `resolveDependencies` to accept that array and build `rowToId` **positionally** (row N → createdIds[N-1]); use the title map only for name-based references, and track duplicate titles in a `Set` — skip resolving any dep that references an ambiguous title, recording a warning string instead.
3. Replace `Promise.all` with chunked `Promise.allSettled` (chunks of ~10) and return `{ resolved: number; warnings: string[] }`.
4. In `ImportPanel.startImport`, wrap the `resolveDependencies` call in try/catch; merge warnings (or the catch message, e.g. "Tasks imported, but some dependencies could not be linked") into `importResult.errors`, and **always** reach `setStep('done')`.

**Accept:** import a sheet with duplicate titles and numeric predecessors → deps attach to the positionally correct rows; kill the network mid-resolution → panel reaches the done step with a warning listed.

### A4 🔴 Switching import source Excel → Planner keeps the stale Excel source
**File:** `ImportPanel.tsx:266, 276`
The Excel card's onClick clears state; the Planner card's is just `setSourceType('planner')`. A user who uploads an Excel file then clicks Planner sees `tasks found in "undefined"` and importing proceeds with the Excel rows under a Planner label.
**Fix:** create one `selectSource(type)` helper that sets `sourceType` and clears `importSource`, `mapping`, and `selectedPlan`; use it in both cards.
**Accept:** upload Excel → click Planner card → the wizard requires picking a plan before Next enables.

### A5 🔴 Out-of-order plan selection race
**File:** `ImportPanel.tsx:142-155`
`handlePlanSelect` sets `selectedPlan` synchronously then awaits `fetchPlannerTasks` and unconditionally `setImportSource(...)`. Click plan A (slow) then plan B (fast): A's response lands last and overwrites the source while B shows the checkmark — user imports A's tasks believing they picked B.
**Fix:** keep `const latestPlanIdRef = useRef<string>()`; set `latestPlanIdRef.current = plan.id` at the top of the handler and after the await bail if `latestPlanIdRef.current !== plan.id` (also guard the catch/finally that clears `planTasksLoading`).
**Accept:** rapid-click two plans → the source always matches the plan showing as selected.

### A6 🟠 Percent-formatted Excel cells import as fractions (50% → 0.5%)
**File:** `ImportService.ts:251-254` (with `raw: true` at 198)
A cell displayed as "50%" has raw value `0.5` → clamped to 0.5% complete.
**Fix:** in `applyMapping`'s percent handling, after `parseFloat`, if **every** non-empty value in the mapped percent column is > 0 and ≤ 1, multiply all by 100 (column-level heuristic — compute once, not per cell). Round to integer.
**Accept:** percent-formatted column imports as 50, not 0.5; a column already using 0–100 integers is untouched.

### A7 🟠 Planner import silently truncates — no `@odata.nextLink` paging
**File:** `ImportService.ts:343-348` (plan tasks), `297-301` (`/me/memberOf` capped at `.top(50)`)
Graph pages planner tasks (~400/page) regardless of `$top`; plans beyond one page import partially with no warning. Users in >50 groups won't see some plans.
**Fix:** loop following `resp['@odata.nextLink']` until exhausted in both calls, accumulating results.
**Accept:** code follows nextLink; verified by inspection (no big tenant needed).

### A8 🟠 `batchImport` is one round trip per task
**File:** `ImportService.ts:422-438`
500 rows = 500 sequential awaited `items.add` calls — minutes of wall clock. The codebase already uses `this.sp.batched()` (SharePointService.ts:201, 431).
**Fix:** add a batched create path in `SharePointService` (chunks of ~50 per `execute()`), capturing per-item rejections into the existing error-reporting shape and **preserving row-order of created IDs** (required by A3). Call `onProgress` per chunk.
**Accept:** import of 100 rows performs ~2 batch executes, per-row failures still land in `importResult.errors`, and A3's positional dependency resolution still passes.

### A9 🟠 Importing into an existing project scrambles sort order
**Files:** `ImportService.ts:428`, caller `ImportPanel.tsx:233`
`sortOrder: i` starts at 0, colliding with existing tasks' sortOrder (list is sorted by `sortOrder, id` — SharePointService.ts:268), interleaving imported rows into the existing manual order.
**Fix:** `batchImport` takes a `sortOrderBase` param; in ImportPanel's append mode, compute `max(existing sortOrder) + 1` from the already-loaded project tasks (pass them in or fetch once) and pass it through. Create mode passes 0.
**Accept:** importing into a project with existing tasks appends after them in List view.

### A10 🟡 Import polish (do all, they're small)
- `ImportService.ts:222-224` — `parseExcelFile`'s `catch {}` discards the real SheetJS error. Append `e instanceof Error ? `: ${e.message}` : ''` to the thrown message.
- `ImportService.ts:196-210` — headerless columns surface as `__EMPTY*` in the mapping UI. Auto-map any `__EMPTY*` header to `skip` before showing ColumnMapper.
- `ImportPanel.tsx:486-531` — when `createProject` throws in create mode, the done step renders "0 tasks added to ‌" (empty title) and an "Open Project" button firing `onImportComplete(undefined)`. When `createdProject` is null in create mode, render an error-only result and a single "Close" button.

---

## Workstream B — Error surfacing & portfolio reliability

### B1 🔴 `getProjectTaskStats` swallows all errors → confidently wrong portfolio data
**File:** `src/webparts/smartGantt/services/SharePointService.ts:345-347`, amplified by `getAllProjectStats:350-357`
The bare `catch { return empty; }` maps 403/429/network to `totalTasks: 0, health: 'on-track'`. `getAllProjectStats` fires an unbounded `Promise.all` (one `getAll()` per project), inviting throttling at portfolio scale — throttled projects then render as "0 tasks, On Track" in PortfolioView **and in the Excel/PowerPoint portfolio exports**.
**Fix:**
1. In the catch, return `empty` only for not-found (`isNotFoundError(e)` — already exists in the file); otherwise rethrow.
2. In `getAllProjectStats`, run in chunks of 4–6 concurrent requests. Wrap each project in try/catch; on failure, produce a stats object flagged as errored. Add `statsError?: boolean` to `IProjectTaskStats` (models/index.ts) for this.
3. In `PortfolioView` and the two portfolio exports (`ExportService`), render errored projects as "—" / "Stats unavailable" rather than 0/on-track.
**Accept:** a project whose list was deleted out from under the registry shows "unavailable", not healthy-empty; exports show "—" for it.

### B2 🔴 Portfolio stats load: stuck spinner, missing mount load, empty-projects cache
**File:** `src/webparts/smartGantt/components/SmartGantt.tsx:86, 112-114, 211-222`
Three defects in `_loadPortfolioStats` / `_handleViewChange`:
- No try/catch/finally — a rejection leaves `portfolioLoading: true` forever and PortfolioView's Refresh is `disabled={loading}` (PortfolioView.tsx:361): unrecoverable without reload.
- Persisted `viewMode: 'portfolio'` never loads stats on mount (`componentDidMount` only calls `_loadProjects`) — cards render statless until manual Refresh.
- If invoked while `_loadProjects` is still in flight, stats compute over `[]` and are cached as a non-null empty Map; the `portfolioStats === null` guard then never reloads.
**Fix:** wrap the body in try/catch/finally (`finally` clears `portfolioLoading`; catch sets an error message via the existing `_errMessage` pattern). Add a sequence-token stale guard like `_taskLoadSeq`. In `componentDidMount` after projects resolve, `if (this.state.viewMode === 'portfolio') void this._loadPortfolioStats();` — and skip/invalidate the stats cache when it was computed against an empty projects array.
**Accept:** load the page with portfolio persisted as the view → stats appear without clicking Refresh; simulate a rejection → spinner clears and an error message shows, Refresh works.

### B3 🔴 Save failures are invisible while a panel is open
**Files:** `SmartGantt.tsx:266-270, 374-378, 510-518`; `TaskPanel.tsx:82-95`; `ProjectPanel.tsx:59-74`
`_handleTaskSave`/`_handleProjectSave` catch errors into `saveError`, whose MessageBar renders **behind** the open Fluent Panel overlay. Inside the panel the spinner just stops — the user assumes the save worked, hits Cancel, and loses their edits.
**Fix:** have `_handleTaskSave`/`_handleProjectSave` rethrow after setting state. In each panel's `handleSave`, catch, set a local `saveError` string, and render an inline `MessageBar (error)` above the form; keep the panel open. Clear it on the next save attempt.
**Accept:** force `updateTask` to reject → an error bar appears inside the open panel; the panel stays open with edits intact.

### B4 🟡 Small error-surfacing gaps (do all)
- `SmartGantt.tsx:401-406` — `_handleExportImage` is the only export path without try/catch. Wrap like `_handleExportPowerPoint` (408-418).
- `SharePointService.ts:182-189` — `deleteProject` swallows non-404 failures recycling the task list, silently orphaning it. `catch (e) { if (!isNotFoundError(e)) throw e; }`.
- `ExportService.ts:363-364, 381-402` — `canvas.toBlob` can resolve null (canvas over browser size cap: multi-year charts exceed ~16,384 px at scale 2) and `downloadPNG` resolves as success; PowerPoint export then embeds a blank image. Reject with a clear "chart too large" message when the blob is null, and clamp `scale` so `width×scale ≤ 16000` (drop to 1, then reduce further if needed).

---

## Workstream C — Gantt chart fixes

### C1 🔴 Timeline scroll position resets after every task edit/drag
**File:** `src/webparts/smartGantt/components/gantt/GanttChart.tsx:156-167 (range memo), 201-208 (scroll effect)`
The `rangeStart` memo depends on `tasks`, so any task mutation creates a new `Date` object even when its value is unchanged; the scroll effect deps contain `rangeStart` as an object → effect refires → viewport snaps back to the earliest task after every single edit or drag commit.
**Fix:** depend on the primitive: use `rangeStart.getTime()` in the effect dep array (and keep a `hasScrolledRef` so the auto-scroll only runs on mount/project change, not on genuine range growth from an edit — scroll once per task-set identity change coming from project switch; simplest robust form: only auto-scroll when the component mounts or `tasks` transitions from empty to non-empty).
**Accept:** scroll to a future month, drag a bar → viewport stays put.

### C2 🔴 "Critical path highlight" toggle has no effect (red outlines always on)
**File:** `GanttChart.tsx:134-137, 408, 424, 448-450`; defaults in `models/index.ts:135-140`
`criticalIds` is computed when `showCriticalPath || showCriticalPathOnly`; bar/milestone red-dashed decoration is gated only on membership in `criticalIds`. With defaults (`showCriticalPath: false`, `showCriticalPathOnly: true`) the red outlines render even though the toggle is OFF, and toggling changes nothing.
**Fix:** in `renderTaskBar`, gate bar/milestone decoration on `settings.showCriticalPath && criticalIds.has(task.id)`; keep the raw `criticalIds` membership only for the dependency-arrow logic (`renderDependencyArrows:554, 561-563`).
**Accept:** toggle off → no red dashed outlines; toggle on → outlines appear on critical tasks.

### C3 🔴 `barStyle` setting ("Flat") never applied to the on-screen chart
**File:** `GanttChart.tsx:431-439, 466`; setting UI `GanttSettings.tsx:123-137`; export honors it (`ExportService.ts:261`)
`settings.barStyle` appears nowhere in GanttChart.tsx — the live chart always renders the gradient, so screen and export don't match.
**Fix:** when `settings.barStyle === 'flat'`, fill the progress rect with `color` directly and skip the per-bar `<linearGradient>` defs; keep `url(#…)` for `'gradient'`. (Coordinate with C7: hoist gradients while you're in this code.)
**Accept:** selecting "Flat" visibly flattens on-screen bars and matches the export.

### C4 🔴 Deep sub-tasks silently disappear (Gantt + ListView) / 3-level nesting is creatable
**Files:** `GanttChart.tsx:1015-1045 (buildVisibleRows)`; `components/views/ListView.tsx:79-119`; `components/panels/TaskPanel.tsx:101-106`
Both renderers emit top-level tasks plus **direct** children only. The parent dropdown prevents picking a parented task as a parent, but not assigning a parent to a task that already **has children** — doing so makes grandchildren, which vanish from Gantt and List (they still exist in SharePoint). Also reachable via list edits/imports.
**Fix (all three together, choose the flatten strategy consistently):**
1. `TaskPanel`: when the edited task has children (`tasks.some(t => t.parentTaskId === task?.id)`), offer only "None" in the parent dropdown with a hint "This task has sub-tasks and cannot be nested."
2. `buildVisibleRows` (Gantt): make `addTask` recurse through `subtaskMap` with a visited-set cycle guard, rendering depth ≥ 2 at depth-1 indentation (flatten display, never drop).
3. `ListView`: same — recurse `children.get(c.id)` in `pushTask` (or flatten to the nearest top-level ancestor), with a cycle guard.
**Accept:** hand-edit `ParentTaskId` in the SharePoint list to create A→B→C → C is visible in both Gantt and List; the panel no longer lets you create that state.

### C5 🟠 Milestone with empty startDate renders at *today*, detached from its arrows
**File:** `GanttChart.tsx:397, 410-411` vs `:570`
Diamond position falls back to `today` while the outgoing arrow anchors at the milestone's dueDate.
**Fix:** in `renderTaskBar`, for milestones use `parseDateOnly(task.startDate) || parseDateOnly(task.dueDate) || today` — matching line 570's fallback.
**Accept:** milestone with only a due date renders at that date with arrows attached.

### C6 🟠 Gantt render performance (do as one pass)
**File:** `GanttChart.tsx`
- `:249-282` — `handleMouseMove` allocates a new Map and setStates on **every mousemove pixel**, re-rendering the entire SVG even when `deltaDays` didn't change. Keep the last applied `deltaDays` in a ref; return early when unchanged.
- `:171` — `visibleTasks = buildVisibleRows(tasks, collapsedPhases)` runs unmemoized every render (including every tooltip enter/leave). Wrap in `React.useMemo(..., [tasks, collapsedPhases])`; also memoize the `taskById`/`rowIndexById` maps built per-render in `renderDependencyArrows:529-535`.
- `:139-143` + `utils/healthUtils.ts:135` — `violationIds` calls `hasDependencyViolation(t, tasks)` per task, and each call rebuilds a `Map` of all tasks → O(n²). Change `hasDependencyViolation` to accept a prebuilt `Map<number, ITask>` and build it once in the memo (update all call sites — grep for it).
**Accept:** `gulp build` clean; dragging a bar on a 200-task project is visibly smooth.

### C7 🟡 Gantt polish (do all)
- `:434-439` — one `<linearGradient>` def per task bar. Hoist a single `<defs>` emitting one gradient per **distinct color** (id keyed by color hex + uid); reference by color. (Coordinates with C3.)
- `:76-81, 447` — `hexToRgba` yields `rgba(NaN,…)` for non-6-digit hex (task.color is free-form). Validate `/^#[0-9a-fA-F]{6}$/`; expand 3-digit hex; else fall back to `#0078D4`.
- `:148-150` — project-week header labels clamp pre-project weeks to "W1" (~4 duplicate bands). Return `''` when `diff < 0`.
- `:83, 103` — `React.useRef(`sg${++ganttInstanceCounter}`)` increments the module counter on every render. Lazy-init: `const uidRef = React.useRef<string>(); if (!uidRef.current) uidRef.current = ...`.
- `GanttChart.module.scss:297` — `.dependencyArrow { marker-end: url(#arrowhead); }` references a nonexistent id (real markers are `${uid}-arrow`, set inline). Delete the rule (then regenerate the `.module.scss.ts` if the build doesn't do it automatically).

---

## Workstream D — SharePointService robustness

### D1 🟠 Orphaned list when field provisioning fails mid-create
**File:** `SharePointService.ts:191-230 (_createProjectList), 141-144`
Cleanup only wraps the registry add. If the field-provisioning batch `execute()` (line 227) rejects, `createProject` throws before the registry add — the list is orphaned (invisible, and a retry allocates `_2`).
**Fix:** in `_createProjectList`, wrap everything after `lists.add(...)` in try/catch; on failure, recycle the just-created list, then rethrow.
**Accept:** by inspection — every failure path after list creation recycles the list.

### D2 🟠 `deleteTask` sub-task promotion: 500-item cap, non-indexed column, proceeds on failure
**File:** `SharePointService.ts:426-441`, field def `:218` area
Children are promoted via `.filter(`ParentTaskId eq ${id}`).top(500)()`. Beyond 500 children they're skipped; on a >5,000-item list the filtered query on the **non-indexed** ParentTaskId throws, is merely `console.warn`ed, and the parent is recycled anyway → invisible orphans (contradicting the code's own comment).
**Fix:** (1) add `Indexed: true` to the ParentTaskId field definition in `_createProjectList` (only helps new lists — fine); (2) page the children query in a loop (`top(500)` + `Id gt lastId` filter) until exhausted; (3) if promotion throws, **rethrow instead of recycling the parent**, so the delete fails loudly and no orphans are created.
**Accept:** deleting a parent either promotes all children or fails without recycling.

### D3 🟠 `_buildListName` treats any probe error as "name free" + duplicate-name race
**File:** `SharePointService.ts:449-466`
Both catches assume free on any error (403/429 → duplicate-name crash later); two simultaneous same-title creates race check-then-create.
**Fix:** in the probe catches, treat only `isNotFoundError(e)` as free; rethrow otherwise. In `_createProjectList`, if `lists.add` fails with an already-exists error, retry once with the next numeric suffix.
**Accept:** by inspection; a 403 during probe surfaces as an error, not a duplicate-name crash.

### D4 🟠 No optimistic concurrency — concurrent editors silently clobber each other
**File:** `SharePointService.ts:402-421 (updateTask), 162-172 (updateProject)`, `getProjectTasks:242+`
PnPjs `update()` defaults to `IF-Match: *`.
**Fix:** select `odata.etag` in `getProjectTasks`/`getProjects`, carry `etag?: string` on `ITask`/`IProject` (models/index.ts), pass it as the eTag argument to `update()`, and on HTTP 412 throw a distinct error the UI maps to "This item was changed by someone else — refresh and retry" (surfaces via the B3 panel error bar). Don't send etag on create.
**Accept:** simulate a 412 → user sees the conflict message; normal saves unaffected.

### D5 🟡 Dependencies field: 500-char TextField cap can reject saves
**File:** `SharePointService.ts:218` (field XML MaxLength 500) vs `createTask:371`/`updateTask:414`
A long dependency list overflows and SharePoint throws a raw error.
**Fix:** cheapest robust fix — in `createTask`/`updateTask`, if the joined string exceeds 500 chars, throw a clear Error ("Too many dependencies on one task (limit ~60)") before hitting SharePoint. (Changing the field to Note for new lists is optional; don't migrate existing lists.)

### D6 🟡 Export edge: empty project → Excel sheet with no header row
**File:** `ExportService.ts:40-50`
`json_to_sheet([])` yields a fully empty sheet.
**Fix:** when rows are empty, seed headers with `XLSX.utils.sheet_add_aoa(ws, [headerArray])` (derive the header array from the same field list used to build row objects).

---

## Workstream E — Shell & view behavior fixes

### E1 🟠 `showDependencies` preference force-overridden to true on every load
**File:** `SmartGantt.tsx:105`
`ganttSettings: { ...DEFAULT_GANTT_SETTINGS, ...prefs.ganttSettings, showDependencies: true }` — the only setting that never sticks; GanttSettings exposes a toggle for exactly this.
**Fix:** delete the trailing `showDependencies: true` override (the DEFAULT spread already covers legacy prefs missing the key).
**Accept:** turn arrows off, reload → still off.

### E2 🟠 Optimistic task update races across a project switch
**File:** `SmartGantt.tsx:434-450`
(a) The optimistic patch matches by numeric task `id` — IDs are per-list, so switching projects mid-flight can patch the *wrong project's* same-ID task. (b) The failure rollback calls `_loadTasks(oldProject.listName)`, which bumps `_taskLoadSeq` and **wins** over the new project's completed load — project A's tasks display under project B's header.
**Fix:** capture `const projectId = this.state.selectedProject.id` before the await; in both the optimistic setState and the catch, bail if `this.state.selectedProject?.id !== projectId`.
**Accept:** by inspection of both guards.

### E3 🟠 Stale task filter hides tasks in a newly created project
**File:** `SmartGantt.tsx:259-263 (reset on select)` vs `:374-378 (create path)`
Project create selects the new project without resetting `taskFilter`; FilterBar is hidden while the project has 0 tasks (Toolbar.tsx:347), so the active filter is invisible and a just-created task can silently not appear.
**Fix:** include `taskFilter: EMPTY_TASK_FILTER` in the project-create selection setState.

### E4 🟠 ListView sorting semantics
**File:** `components/views/ListView.tsx:79-119`
- `:96` — Status/Priority sort alphabetically (Critical, High, Low, Medium). Special-case: `cmp = TASK_PRIORITY_OPTIONS.indexOf(a.priority) - TASK_PRIORITY_OPTIONS.indexOf(b.priority)` (and TASK_STATUS_OPTIONS for status — both already imported).
- `:90-92` — undated tasks sort as epoch 0, piling at the top ascending. Use a sentinel so undated always lands **last** after the direction negation is applied (i.e. `?? (sortDir === 'asc' ? Number.MAX_SAFE_INTEGER : -Number.MAX_SAFE_INTEGER)`).
**Accept:** Priority asc = Critical→Low order (or Low→Critical, matching the arrow); undated rows last in both directions.

### E5 🟠 Kanban drag-over flicker
**File:** `components/views/KanbanView.tsx:92-94, 255-258, 291`
Column `dragleave` fires when the pointer enters a child card, clearing `dragOverCol` → highlight/placeholder flickers.
**Fix:** in `handleDragLeave(e)`, `if (e.currentTarget.contains(e.relatedTarget as Node)) return;` before clearing (pass the event through; apply to the Cancelled column too).

### E6 🟠 Dashboard "recent" panels aren't recency-sorted
**File:** `components/views/DashboardView.tsx:144-152`
"Updated this week"/"Completed this week" slice the first 6/4 in sortOrder order.
**Fix:** sort by `modified` descending (it's a real timestamp — `new Date(x.modified).getTime()`, per the documented convention) before slicing.

### E7 🟠 Autocomplete keyboard highlight can select an invisible suggestion
**File:** `components/common/AutocompleteField.tsx:44, 50, 73`
Navigation clamps to `filtered.length - 1` but only `filtered.slice(0, 12)` renders; Enter can commit an option the user never saw.
**Fix:** compute `const visible = filtered.slice(0, 12)` once; use it for rendering, clamping, and the Enter commit.

### E8 🟡 Shell/view polish (do all)
- `SmartGantt.tsx:169-174` — persisted selection can restore an **archived** (hidden) project while archived are not shown. Prefer `visible.find(p => p.id === wantedId)` (fall back to the full list only when `showArchivedProjects`), and final fallback `visible[0] || null`.
- `SmartGanttWebPart.ts:36, 60-63` + `SmartGantt.tsx` — the "Web Part Title" property-pane field does nothing (`props.title` never rendered). Render it as a heading above the toolbar when non-empty (preferred over removing an existing pane field).
- `SmartGantt.tsx:162-232` — add an `_isMounted` guard (set in `componentDidMount`/`componentWillUnmount`) around post-await setStates in `_loadProjects`/`_loadTasks`/`_loadPortfolioStats`, and clear `_handleScrollToToday`'s 100 ms timeout in `componentWillUnmount`.
- `TaskPanel.tsx:167, 179` / `ProjectPanel.tsx:83, 95` — panels discard unsaved edits on ESC/X silently. Track a `dirty` flag (set in the shared `set()` helper); in `onDismiss`, if dirty and not saving, `confirm('Discard unsaved changes?')` before closing.
- `TaskPanel.tsx:253-260, 296-307` — status→percent sync is one-way. Mirror it in the slider handler: slider→100 sets status "Completed" (if not Cancelled); slider>0 while "Not Started" sets "In Progress"; slider<100 while "Completed" sets "In Progress".
- `DashboardView.tsx:157, 162` — comparators never return 0 for ties. Use `a.dueDate.localeCompare(b.dueDate)` (YYYY-MM-DD compares correctly).
- `DashboardView.tsx:195-196` — project status chip uses task `STATUS_COLORS`, so 'Planning'/'Active' fall to gray while PortfolioView (`:24-38`) has its own correct local maps. Move PortfolioView's maps into `models/index.ts` as `PROJECT_STATUS_COLORS`/`PROJECT_STATUS_LIGHT_COLORS`; use in both views.
- `PortfolioView.tsx:65-66` — `MiniTimeline` returns null for a single-day project (`total <= 0`). Use `Math.max(1, differenceInCalendarDays(end, start))` and keep only the negative bail-out.

---

## Workstream F — App-shell render performance

### F1 🟠 Derived data recomputed on every render; every search keystroke re-renders the whole app
**File:** `SmartGantt.tsx:463-470, 492-506`; `FilterBar.tsx:111`
`visibleProjects`, `knownPhases`, `knownUsers` (two Set+sort passes over all tasks), and `filterTasks(tasks, taskFilter)` run on every render, and ~10 inline arrow props to Toolbar change identity each time. Every keystroke re-renders the full Gantt SVG.
**Fix (class-component-friendly, in this order):**
1. Cache derived values on instance fields with manual input checks, e.g. `private _filteredCache: { tasks: ITask[]; filter: ITaskFilter; result: ITask[] }` — recompute only when inputs differ by reference. Same pattern for knownPhases/knownUsers (keyed on `tasks`) and visibleProjects (keyed on `projects`, `showArchivedProjects`).
2. Hoist the inline arrows passed to Toolbar into bound instance methods (most already exist — pass them directly).
3. Give FilterBar local state for the text input, debounced ~150 ms before calling `onFilterChange`.
4. After C6 lands, optionally wrap `Toolbar`/`FilterBar` in `React.memo`.
**Accept:** typing in search doesn't lag on a 200-task project; `gulp build` clean.

### F2 🟡 Minor render hygiene
- `ListView.tsx:130-143` — `SortTh` component type is defined inside the render body → header `<th>`s remount every render. Hoist to module scope (props: field, label, sortField, sortDir, onSort) or convert to a plain render function.
- `DashboardView.tsx:118-173` — wrap the aggregation blocks in `useMemo(..., [tasks])`, hoisting `today`/`in14`/`weekAgo` inside the memo.

---

## Workstream G — Accessibility & deduplication (batch, lowest priority)

### G1 🟡 Keyboard accessibility
- `Toolbar.tsx:138-181, 234-239, 308-337` — items inside the project/export callouts are plain `div onClick` — a keyboard user can open the callout but select nothing. Render items as restyled `<button>`s (preferred) or add `role="menuitem"`, `tabIndex={0}`, Enter/Space handlers.
- `TaskPanel.tsx:346-370` / `ProjectPanel.tsx:126-138` — color swatch divs: add `role="radio"`, `aria-checked`, `tabIndex={0}`, key handling (do together with G2's extraction).
- `GanttSettings.tsx:86-96` — header-theme swatches → `<button type="button">` with `aria-pressed` + `aria-label`.
- `GanttSettings.tsx:166-194` — Fluent `Toggle`s have no accessible name (label is a sibling span). Pass `ariaLabel={label}`.
- `ListView.tsx:131-137` — sortable headers: `tabIndex={0}` + Enter/Space → `handleSort`.
- `KanbanView.tsx` — cards are mouse-only. Minimum: `tabIndex={0}` on cards and a keyboard status-move (left/right arrows call `onTaskUpdate` with the same %-complete side effects as `handleDrop:102-108`).
- `GanttChart.tsx:298-303, 452-521` — drag/resize is mouse-event-only (dead on touch devices) and bars have no ARIA. Switch to Pointer Events (`onPointerDown` + `setPointerCapture`, `pointermove`/`pointerup` window listeners — keep the existing cleanup pattern) and add `role="img"` + `aria-label` (title, dates, % complete) per bar group. *This is the largest G item — do it last, test drag/resize thoroughly with a mouse afterward.*

### G2 🟡 Deduplication
- Color-swatch picker (swatch grid + custom color input + selected ring) is near-verbatim duplicated between `TaskPanel.tsx:343-400` and `ProjectPanel.tsx:123-159`. Extract `components/common/ColorSwatchPicker.tsx` (props: colors, value, onChange) and use in both — fold in G1's a11y.
- `isOverdue`, `initials`, `stringToColor` are copy-pasted between `ListView.tsx:25-41` and `KanbanView.tsx:34-49`. Move to a shared util (e.g. `utils/taskDisplayUtils.ts`) and import.
- `KanbanView.tsx:27-32` — `COLUMNS` hardcodes hexes duplicating `STATUS_COLORS`. Reference `STATUS_COLORS[status]`.
- (Covered by E8) `PROJECT_STATUS_COLORS` maps move to models.

---

## Final verification checklist (run after all workstreams)

1. `gulp build` — zero errors/warnings introduced.
2. Import flow: .xlsx with date cells, percent column, duplicate titles, numeric predecessors → dates correct, percents correct, deps positionally correct, done step always reached.
3. Gantt: scroll persists across edits; critical-path toggle works both ways; Flat bar style applies; drag smooth; drag with mouse still works after pointer-event migration.
4. Panels: failed save shows inline error, edits preserved; ESC on dirty form prompts.
5. Portfolio: loads on mount when persisted; failure recoverable via Refresh; unavailable projects show "—" in view and exports.
6. Filters: create project while filtered → new tasks visible.
7. List/Kanban: priority sort in semantic order; undated last; no drag flicker.
8. Update `CHANGELOG`/version only if the repo's release process asks for it — otherwise leave versioning alone.
